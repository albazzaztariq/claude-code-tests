"""CPU worker processes - PDF rendering and text extraction (no CUDA)."""

from multiprocessing import Queue
from pathlib import Path
import fitz
import numpy as np
from PIL import Image
from typing import List
from shared_types import PageBatch, DetectionResult, GPU_DONE
import time
import sys
import os

# Import extraction functions at module level for Windows spawn
sys.path.insert(0, str(Path(__file__).parent))
from NativeTextExtract_pipeline import (
    image_rects_to_pdf,
    extract_text_with_banned_regions,
    extract_metric_sentences,
    SEARCHABLE_TERMS,
    MASK_LABELS,
    CLASS_NAMES,
    ACRONYM_MAPPING,
)


DPI = 150


def render_pages_to_batches(pdf_path: Path,
                            pdf_num: int,
                            worker_id: int,
                            batch_size: int = 20) -> List[PageBatch]:
    """
    Render all pages of PDF into batches of images.

    Args:
        pdf_path: Path to PDF
        pdf_num: PDF number
        worker_id: Worker ID for routing results
        batch_size: Pages per batch

    Returns:
        List of PageBatch objects
    """
    doc = fitz.open(pdf_path)
    num_pages = len(doc)
    batches = []
    batch_id = 0

    zoom = DPI / 72.0
    mat = fitz.Matrix(zoom, zoom)

    for start in range(0, num_pages, batch_size):
        end = min(num_pages, start + batch_size)
        imgs = []
        idxs = list(range(start, end))

        for page_idx in idxs:
            page = doc[page_idx]
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            imgs.append(np.array(img))  # Convert to numpy array

        batches.append(PageBatch(
            batch_id=batch_id,
            pdf_num=pdf_num,
            worker_id=worker_id,
            page_indices=idxs,
            images=imgs
        ))
        batch_id += 1

    doc.close()
    return batches


def cpu_worker(worker_id: int,
               pdf_queue: Queue,
               gpu_in: Queue,
               gpu_out: Queue,
               result_queue: Queue,
               stop_token: str = "__STOP__"):
    """
    CPU process:
      - receives (pdf_num, pdf_path, output_dir)
      - renders pages -> PageBatch
      - sends batches to GPU, waits for DetectionResult
      - runs text extraction + saving

    Args:
        worker_id: Worker ID for routing results
        pdf_queue: Queue receiving (pdf_num, pdf_path_str, output_dir_str) tuples
        gpu_in: Queue sending PageBatch to GPU worker
        gpu_out: Dedicated output queue for THIS worker only (no routing needed)
        result_queue: Queue sending results back to orchestrator
        stop_token: String to signal shutdown
    """
    worker_pid = os.getpid()
    pdfs_processed = 0

    while True:
        item = pdf_queue.get()
        if item == stop_token:
            break

        pdf_num, pdf_path_str, output_dir_str = item
        pdf_path = Path(pdf_path_str)
        output_dir = Path(output_dir_str)
        output_dir.mkdir(exist_ok=True)

        extraction_start = time.time()

        # Render pages and send batches to GPU
        batches = render_pages_to_batches(pdf_path, pdf_num, worker_id)
        pending = {b.batch_id for b in batches}  # Set of batch_ids, not dict

        for b in batches:
            gpu_in.put(b)

        # Open PDF for text extraction
        doc = fitz.open(pdf_path)
        num_pages = len(doc)
        all_page_texts = [""] * num_pages
        table_detections_for_pdf = []

        # Collect results from GPU (dedicated queue - no routing needed!)
        while pending:
            msg = gpu_out.get()

            # Check for GPU_DONE sentinel (signals no more results coming)
            if isinstance(msg, tuple) and msg[0] == GPU_DONE:
                break

            # Must be DetectionResult (already routed to our dedicated queue)
            det: DetectionResult = msg

            # Check if we're still waiting for this batch
            if det.batch_id not in pending:
                continue

            # Remove from pending
            pending.remove(det.batch_id)

            # Get the corresponding batch (need to reconstruct for images)
            # Find batch by batch_id
            batch = None
            for b in batches:
                if b.batch_id == det.batch_id:
                    batch = b
                    break

            if batch is None:
                continue

            # Process each page in the batch
            for local_idx, page_idx in enumerate(det.page_indices):
                page = doc[page_idx]
                img_np = batch.images[local_idx]
                img_h, img_w, _ = img_np.shape

                banned_img_rects = []

                # Process detections for this page
                for x1, y1, x2, y2, cls_id, conf in det.boxes_per_page[local_idx]:
                    label = CLASS_NAMES.get(cls_id, 'unknown')

                    # Save table detections
                    if label == 'Table':
                        table_detections_for_pdf.append({
                            "pdf_num": pdf_num,
                            "page_num": page_idx,
                            "bbox": [float(x1), float(y1), float(x2), float(y2)],
                            "confidence": conf
                        })

                    # Keep only banned labels for text filtering
                    if label in MASK_LABELS:
                        banned_img_rects.append((x1, y1, x2, y2))

                # Convert to PDF coordinates
                page_rect = page.rect
                banned_pdf_rects = image_rects_to_pdf(
                    banned_img_rects,
                    page_rect.width,
                    page_rect.height,
                    img_w,
                    img_h
                )

                # Extract text
                page_text = extract_text_with_banned_regions(page, banned_pdf_rects)
                all_page_texts[page_idx] = page_text

        doc.close()

        # Combine text
        extracted_text = "\n\n".join(all_page_texts)
        extraction_time = time.time() - extraction_start

        # Skip if no text
        if not extracted_text or len(extracted_text.strip()) == 0:
            result_queue.put({
                "pdf_num": pdf_num,
                "extracted_text": "",
                "table_detections": table_detections_for_pdf,
                "metric_sentences": [],
                "extraction_time": extraction_time,
                "metrics_time": 0.0,
                "error": True
            })
            pdfs_processed += 1
            continue

        # Extract metric sentences
        filter_start = time.time()
        metric_sentences = extract_metric_sentences(extracted_text, SEARCHABLE_TERMS)
        filter_time = time.time() - filter_start

        # CRITICAL: Write text files to disk (Pipeline 3 expects them)
        text_path = output_dir / f"{pdf_num}.txt"
        filtered_path = output_dir / f"{pdf_num}_filtered.txt"

        with open(text_path, 'w', encoding='utf-8') as f:
            f.write(extracted_text)

        # Write filtered text with "Metric Match:" labels
        with open(filtered_path, 'w', encoding='utf-8') as f:
            for metric_term, sentence in metric_sentences:
                # Format based on whether it's an acronym or phrase
                if metric_term in ACRONYM_MAPPING:
                    # It's an acronym - show "Acronym, which means Phrase"
                    phrase = ACRONYM_MAPPING[metric_term]
                    f.write(f"Metric Match: {metric_term}, which means {phrase}\n")
                else:
                    # It's a phrase - show just the phrase
                    f.write(f"Metric Match: {metric_term}\n")
                f.write(sentence + "\n\n")

        # Send results back
        result_queue.put({
            "pdf_num": pdf_num,
            "extracted_text": extracted_text,
            "table_detections": table_detections_for_pdf,
            "metric_sentences": metric_sentences,
            "extraction_time": extraction_time,
            "metrics_time": filter_time,
            "error": False
        })

        pdfs_processed += 1
