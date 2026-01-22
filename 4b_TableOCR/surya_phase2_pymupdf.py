"""
Phase 2: Text Extraction with PyMuPDF Hybrid Method (GPU + CPU)

HYBRID APPROACH:
1. Surya TableRecPredictor (GPU) - Table structure
2. Surya DetectionPredictor (GPU) - Text bbox detection
3. PyMuPDF (CPU) - Native PDF text extraction

CRITICAL: This method ONLY works on native text tables (96.3% of academic PDFs).
Image-embedded tables are SKIPPED (3.7% of tables).

Benefits:
- 4-5x faster than full Surya (0.5-0.6s vs 2.6s per table)
  * GPU time (TableRec + Detection): ~0.517s per table
  * PyMuPDF extraction: ~0.078s per table
  * Total: ~0.595s per table
- 97-100% accuracy on native text
- 0.537 GB VRAM (vs 5-6 GB for full Surya)
- Enables Llama 3.2 3B coexistence (7.5 GB free)

See: Docs-ProjectScripts/PyMuPDF_Hybrid_Method.md
"""

import sys
import os
import json
import re
import time
import fitz  # PyMuPDF
from pathlib import Path
from PIL import Image
from collections import defaultdict

# Set Surya environment variables
os.environ['DISABLE_MODEL_SOURCE_CHECK'] = 'True'

# CRITICAL: Optimize batch sizes for RTX 4070 Laptop (8GB VRAM)
# Default batch sizes cause GPU memory swapping (10x slowdown)
os.environ["DETECTOR_BATCH_SIZE"] = "4"      # Default 36 (16GB) -> 4 (1.76GB)
os.environ["TABLE_REC_BATCH_SIZE"] = "8"     # Default 64 (10GB) -> 8 (1.2GB)

def clean_html_tags(text: str) -> str:
    """Remove HTML/markdown tags from text"""
    cleaned = re.sub(r'<[^>]+>', '', text)
    return cleaned

def filter_columns_by_headers(structured_table: list, metrics_list: list):
    """
    Filter table columns based on header row keywords.

    Keeps only columns where the header (row 0) contains a keyword from MetricsSearchable.txt.
    If no columns match, returns empty table.
    """
    if not structured_table or len(structured_table) == 0:
        return [], [], []

    headers = structured_table[0]  # First row is headers
    matching_columns = []
    found_metrics = []

    # Check each header for metric keywords
    for col_idx, header in enumerate(headers):
        header_lower = header.lower()

        # Check if this header contains any metric keyword
        for metric in metrics_list:
            metric_pattern = r'\b' + re.escape(metric.lower()) + r'\b'
            if re.search(metric_pattern, header_lower):
                matching_columns.append(col_idx)
                if metric not in found_metrics:
                    found_metrics.append(metric)
                break  # Found match for this column, move to next column

    # If no columns match, return empty table
    if not matching_columns:
        return [], [], []

    # Build filtered table with only matching columns
    filtered_table = []
    for row in structured_table:
        filtered_row = [row[col_idx] for col_idx in matching_columns if col_idx < len(row)]
        filtered_table.append(filtered_row)

    return filtered_table, matching_columns, found_metrics

def check_if_image_table(pdf_path: str, page_num: int, table_bbox: tuple, render_dpi: int = 300) -> bool:
    """
    Check if table is image-embedded (no text layer).

    Args:
        pdf_path: Path to PDF file
        page_num: Page number (0-indexed)
        table_bbox: (x1, y1, x2, y2) in rendered coordinates
        render_dpi: DPI used for rendering (default 300)

    Returns:
        True if image-embedded (no text), False if native text
    """
    doc = fitz.open(pdf_path)
    page = doc[page_num]

    # Convert from rendered coordinates to PDF coordinates
    x1, y1, x2, y2 = table_bbox
    scale = render_dpi / 72.0
    pdf_x1 = x1 / scale
    pdf_y1 = y1 / scale
    pdf_x2 = x2 / scale
    pdf_y2 = y2 / scale

    # Check if text layer exists
    clip_rect = fitz.Rect(pdf_x1, pdf_y1, pdf_x2, pdf_y2)
    text = page.get_text("text", clip=clip_rect).strip()

    doc.close()

    return len(text) == 0  # True if no text (image-embedded)

def extract_text_pymupdf(pdf_path: str, page_num: int, table_bbox: tuple, text_bboxes: list, render_dpi: int = 300) -> list:
    """
    Extract text from PDF using PyMuPDF for each text bbox.

    Args:
        pdf_path: Path to PDF file
        page_num: Page number (0-indexed)
        table_bbox: (x1, y1, x2, y2) table bbox in rendered coordinates
        text_bboxes: List of text bboxes from Surya DetectionPredictor
        render_dpi: DPI used for rendering (default 300)

    Returns:
        List of extracted text strings (one per bbox)
    """
    doc = fitz.open(pdf_path)
    page = doc[page_num]

    scale = render_dpi / 72.0
    table_x1, table_y1, table_x2, table_y2 = table_bbox

    extracted_texts = []

    for bbox in text_bboxes:
        # bbox.polygon is [[x1,y1], [x2,y2], [x3,y3], [x4,y4]] in rendered coords
        xs = [p[0] for p in bbox.polygon]
        ys = [p[1] for p in bbox.polygon]
        x_min, x_max = int(min(xs)), int(max(xs))
        y_min, y_max = int(min(ys)), int(max(ys))

        # Convert to PDF coordinates (relative to page, not table crop)
        pdf_x1 = (x_min / scale) + (table_x1 / scale)
        pdf_y1 = (y_min / scale) + (table_y1 / scale)
        pdf_x2 = (x_max / scale) + (table_x1 / scale)
        pdf_y2 = (y_max / scale) + (table_y1 / scale)

        # Extract text from bbox
        clip_rect = fitz.Rect(pdf_x1, pdf_y1, pdf_x2, pdf_y2)
        text = page.get_text("text", clip=clip_rect).strip()

        # Clean HTML tags (Surya sometimes outputs these)
        text = clean_html_tags(text)

        extracted_texts.append(text)

    doc.close()

    return extracted_texts

def bbox_overlap(bbox1, bbox2):
    """Calculate overlap area between two bboxes"""
    x1_min, y1_min, x1_max, y1_max = bbox1
    x2_min, y2_min, x2_max, y2_max = bbox2

    x_overlap = max(0, min(x1_max, x2_max) - max(x1_min, x2_min))
    y_overlap = max(0, min(y1_max, y2_max) - max(y1_min, y2_min))

    return x_overlap * y_overlap

def process_tables(metadata_json: Path, metrics_file: Path, results_json: Path):
    """
    Process table crops using Surya + PyMuPDF hybrid method.

    Args:
        metadata_json: Path to phase1_metadata.json (crop metadata)
        metrics_file: Path to MetricsSearchable.txt
        results_json: Path to save phase2_results.json

    Returns:
        None (saves results to results_json)
    """

    print("[PHASE 2] PyMuPDF Hybrid - Text Extraction (GPU + CPU)")
    print(f"Metadata JSON: {metadata_json}")
    print(f"Metrics File: {metrics_file}")
    print(f"Results JSON: {results_json}")

    # Load metrics
    print("\nLoading metrics list...")
    with open(metrics_file, 'r', encoding='utf-8') as f:
        metrics_list = [line.strip() for line in f if line.strip()]
    print(f"  Loaded {len(metrics_list)} metrics")

    # Load crop metadata
    print("\nLoading crop metadata...")
    with open(metadata_json, 'r', encoding='utf-8') as f:
        table_crop_metadata = json.load(f)
    print(f"  Loaded {len(table_crop_metadata)} crops")

    # Import PyTorch and Surya with timing
    print("\nImporting PyTorch...")
    t0 = time.perf_counter()
    import torch
    t1 = time.perf_counter()
    print(f"  PyTorch {torch.__version__} (CUDA: {torch.cuda.is_available()}) - {t1-t0:.2f}s")

    print("Importing Surya...")
    t2 = time.perf_counter()
    from surya.table_rec import TableRecPredictor
    from surya.detection import DetectionPredictor
    t3 = time.perf_counter()
    print(f"  Surya imports - {t3-t2:.2f}s")

    print("Loading TableRecPredictor model...")
    t4 = time.perf_counter()
    table_rec = TableRecPredictor()
    t5 = time.perf_counter()
    print(f"  TableRecPredictor loaded - {t5-t4:.2f}s")

    print("Loading DetectionPredictor model...")
    t6 = time.perf_counter()
    det_predictor = DetectionPredictor()
    t7 = time.perf_counter()
    print(f"  DetectionPredictor loaded - {t7-t6:.2f}s")

    print(f"\nTotal overhead: {t7-t0:.2f}s (PyTorch: {t1-t0:.2f}s, Surya imports: {t3-t2:.2f}s, TableRec: {t5-t4:.2f}s, Detection: {t7-t6:.2f}s)")
    print("NOTE: PyMuPDF used for text extraction (no RecognitionPredictor)")
    print()

    # PHASE A: Collect ALL valid crops first (single-batch processing)
    print("\n[PHASE A] Collecting all valid crops...")
    all_valid_metas = []
    all_crops = []
    skipped_image_tables = 0

    # Start timing
    start_time = time.perf_counter()

    for meta in table_crop_metadata:
        # Get PDF path from metadata
        pdf_path = meta.get("pdf_path", "")
        if not pdf_path or not Path(pdf_path).exists():
            print(f"  WARNING: PDF path not found in metadata for {meta['crop_filename']}")
            continue

        # Check if table is image-embedded (5ms overhead)
        table_bbox = meta.get("table_bbox", None)
        page_num = meta.get("page_num")
        if table_bbox is None:
            print(f"  WARNING: Missing table_bbox in metadata for {meta['crop_filename']}")
            continue

        is_image = check_if_image_table(pdf_path, page_num, table_bbox)

        if is_image:
            print(f"  SKIPPING {meta['crop_filename']} - Image-embedded table (no text layer)")
            skipped_image_tables += 1
            continue

        crop = Image.open(meta["crop_path"])
        all_valid_metas.append(meta)
        all_crops.append(crop)

    print(f"  Collected {len(all_crops)} valid crops (skipped {skipped_image_tables} image tables)")

    # Skip if no valid crops
    if not all_crops:
        print("\n[PHASE 2 COMPLETE]")
        print(f"Total tables through GPU: 0")
        print(f"Tables with matching metrics: 0")
        print(f"Image tables skipped: {skipped_image_tables}")
        print(f"Phase 2 Runtime: 0.00s")
        return

    # PHASE B: Single GPU batch call for ALL crops
    print(f"\n[PHASE B] Processing all {len(all_crops)} tables in single batch...")
    total_tables_processed = len(all_crops)

    print(f"  Running TableRecPredictor on {len(all_crops)} crops...")
    tb0 = time.perf_counter()
    table_results = table_rec(all_crops)
    tb1 = time.perf_counter()
    print(f"    TableRec GPU time: {tb1-tb0:.2f}s ({(tb1-tb0)/len(all_crops):.3f}s per table)")

    # Clear GPU cache after TableRec to prevent memory leak
    import torch
    torch.cuda.empty_cache()

    print(f"  Running DetectionPredictor on {len(all_crops)} crops...")
    tb2 = time.perf_counter()
    det_predictions = det_predictor(all_crops)
    tb3 = time.perf_counter()
    print(f"    Detection GPU time: {tb3-tb2:.2f}s ({(tb3-tb2)/len(all_crops):.3f}s per table)")

    # Clear GPU cache after Detection to prevent memory leak
    torch.cuda.empty_cache()

    print(f"  GPU processing complete - Total: {tb3-tb0:.2f}s")

    # PHASE C: Process results and organize by PDF/page
    print(f"\n[PHASE C] Organizing results by PDF/page...")

    # Group results by PDF
    crops_by_pdf = defaultdict(lambda: defaultdict(list))
    for i, meta in enumerate(all_valid_metas):
        crops_by_pdf[meta["pdf_num"]][(meta["pdf_name"], meta["page_num"])].append((i, meta))

    # Process crops
    all_results = []
    total_tables_filtered = 0  # Tables that passed column filtering

    for pdf_num in sorted(crops_by_pdf.keys()):
        pdf_data = {
            "pdf_number": pdf_num,
            "filename": None,
            "pages": [],
            "table_count": 0,
            "crops": []
        }

        for (pdf_name, page_num), indexed_metas in sorted(crops_by_pdf[pdf_num].items()):
            if pdf_data["filename"] is None:
                pdf_data["filename"] = pdf_name

            page_data = {
                "page_number": page_num,
                "tables": []
            }

            # Process results for each crop on this page
            for i, meta in indexed_metas:
                print(f"  [{i+1}/{total_tables_processed}] {meta['crop_filename']}...")

                table_result = table_results[i]

                if not det_predictions[i].bboxes:
                    print(f"    WARNING: No text bboxes detected")
                    continue

                text_bboxes = det_predictions[i].bboxes

                # Step 3: Extract text from PDF using PyMuPDF
                pdf_path = meta["pdf_path"]
                table_bbox = meta["table_bbox"]
                crop_page_num = meta["page_num"]
                extracted_texts = extract_text_pymupdf(pdf_path, crop_page_num, table_bbox, text_bboxes)

                # Create text_lines compatible with old format
                text_lines = []
                for bbox, text in zip(text_bboxes, extracted_texts):
                    # Convert bbox.polygon to [x1, y1, x2, y2]
                    xs = [p[0] for p in bbox.polygon]
                    ys = [p[1] for p in bbox.polygon]
                    bbox_rect = [min(xs), min(ys), max(xs), max(ys)]

                    # Create text_line object
                    class TextLine:
                        def __init__(self, text, bbox):
                            self.text = text
                            self.bbox = bbox

                    text_lines.append(TextLine(text, bbox_rect))

                # Step 4: Match text lines to cells by bounding box overlap
                rows_dict = {}
                for cell in table_result.cells:
                    row_id = cell.row_id
                    if row_id not in rows_dict:
                        rows_dict[row_id] = []

                    # Find text lines that overlap with this cell
                    cell_text_parts = []
                    for text_line in text_lines:
                        overlap = bbox_overlap(cell.bbox, text_line.bbox)
                        if overlap > 0:
                            cell_text_parts.append(text_line.text)

                    cell_text = ' '.join(cell_text_parts)

                    rows_dict[row_id].append({
                        "col_id": cell.col_id,
                        "text": cell_text
                    })

                # Build structured table
                structured_table = []
                for row_id in sorted(rows_dict.keys()):
                    row_cells = sorted(rows_dict[row_id], key=lambda c: c["col_id"] if c["col_id"] is not None else 0)
                    structured_table.append([c["text"] for c in row_cells])

                # Filter columns based on header keywords
                filtered_table, matching_columns, found_metrics = filter_columns_by_headers(structured_table, metrics_list)

                # Only save table if it has at least one matching column
                if not filtered_table or len(matching_columns) == 0:
                    print(f"    Skipping - no column headers match MetricsSearchable.txt")
                    continue

                table_data = {
                    "table_number": meta["table_num"],
                    "crop_filename": meta["crop_filename"],
                    "structured_table": filtered_table,
                    "original_columns": len(structured_table[0]) if structured_table else 0,
                    "filtered_columns": len(matching_columns),
                    "matching_column_indices": matching_columns,
                    "found_metrics": found_metrics,
                    "has_metrics": True
                }

                page_data["tables"].append(table_data)
                pdf_data["table_count"] += 1
                pdf_data["crops"].append(meta["crop_filename"])

                print(f"    {len(filtered_table)} rows x {len(filtered_table[0]) if filtered_table else 0} cols (kept {len(matching_columns)}/{len(structured_table[0]) if structured_table else 0}), metrics: {', '.join(found_metrics)}")

            if page_data["tables"]:
                pdf_data["pages"].append(page_data)

        all_results.append(pdf_data)
        total_tables_filtered += pdf_data["table_count"]

    # End timing
    elapsed = time.perf_counter() - start_time

    # Save results
    with open(results_json, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, indent=2, ensure_ascii=False)

    print(f"\n[PHASE 2 COMPLETE]")
    print(f"Total tables through GPU: {total_tables_processed}")
    print(f"Tables with matching metrics: {total_tables_filtered}")
    print(f"Image tables skipped: {skipped_image_tables}")
    print(f"Phase 2 Runtime: {elapsed:.2f}s ({elapsed/total_tables_processed if total_tables_processed > 0 else 0:.3f}s per table)")
    print(f"Results saved: {results_json}")

def main():
    """CLI entry point."""
    if len(sys.argv) != 4:
        print("Usage: surya_phase2_pymupdf.py <metadata_json> <metrics_file> <results_json>")
        sys.exit(1)

    metadata_json = Path(sys.argv[1])
    metrics_file = Path(sys.argv[2])
    results_json = Path(sys.argv[3])

    process_tables(metadata_json, metrics_file, results_json)

if __name__ == "__main__":
    main()
