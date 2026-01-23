"""
Phase 1: Table Cropping from YOLO Detections

Reads table detections from text extraction (YOLO), crops tables, saves metadata.
NO detection overhead - detections already done during text extraction.
"""

import sys
import json
import time
from pathlib import Path
from PIL import Image
import fitz

DPI = 150  # Match text extraction DPI

def crop_table_images(page_image, table_boxes):
    """Crop table regions from page image."""
    crops = []
    for x1, y1, x2, y2 in table_boxes:
        crop = page_image.crop((x1, y1, x2, y2))
        crops.append(crop)
    return crops

def main():
    if len(sys.argv) != 4:
        print("Usage: yolo_phase1.py <pdf_folder> <output_folder> <metadata_json>")
        sys.exit(1)

    pdf_folder = Path(sys.argv[1])
    output_folder = Path(sys.argv[2])
    metadata_json_out = Path(sys.argv[3])

    # Read YOLO detections from text extraction
    query_folder = pdf_folder.parent
    yolo_detections_json = query_folder / "yolo_table_detections.json"

    print("[PHASE 1] YOLO Table Cropping (NO detection - reusing from text extraction)")
    print(f"PDF Folder: {pdf_folder}")
    print(f"Output Folder: {output_folder}")
    print(f"Detections JSON: {yolo_detections_json}")

    if not yolo_detections_json.exists():
        print(f"\nERROR: YOLO detections not found: {yolo_detections_json}")
        print("       Text extraction must run first to generate table detections.")
        sys.exit(1)

    # Load YOLO table detections
    print("\nLoading YOLO table detections...")
    with open(yolo_detections_json, 'r', encoding='utf-8') as f:
        all_detections = json.load(f)
    print(f"Loaded {len(all_detections)} table detections\n")

    # Group detections by PDF and page
    detections_by_pdf_page = {}
    for det in all_detections:
        key = (det["pdf_num"], det["page_num"])
        if key not in detections_by_pdf_page:
            detections_by_pdf_page[key] = []
        detections_by_pdf_page[key].append(det)

    # Process PDFs and crop tables
    table_crop_metadata = []
    start_time = time.perf_counter()

    for (pdf_num, page_num), page_detections in sorted(detections_by_pdf_page.items()):
        pdf_path = pdf_folder / f"{pdf_num}.pdf"

        if not pdf_path.exists():
            print(f"WARNING: PDF not found: {pdf_path}")
            continue

        doc = fitz.open(pdf_path)

        if page_num >= len(doc):
            print(f"WARNING: Page {page_num} out of range for {pdf_path.name}")
            doc.close()
            continue

        page = doc[page_num]

        # Render page at same DPI as text extraction
        mat = fitz.Matrix(DPI / 72, DPI / 72)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Extract bboxes
        table_boxes = [det["bbox"] for det in page_detections]

        print(f"Processing PDF {pdf_num}, Page {page_num + 1}: {len(table_boxes)} tables")

        # Crop and save tables
        table_crops = crop_table_images(img, table_boxes)

        for table_idx, (crop, bbox) in enumerate(zip(table_crops, table_boxes), 1):
            crop_filename = f"{pdf_num}_page{page_num + 1}_table{table_idx}.png"
            crop_path = output_folder / crop_filename
            crop.save(crop_path)

            # Store metadata for Phase 2
            table_crop_metadata.append({
                "pdf_num": pdf_num,
                "pdf_name": pdf_path.name,
                "pdf_path": str(pdf_path.resolve()),  # Absolute path for PyMuPDF
                "page_num": page_num,  # 0-indexed for PyMuPDF
                "table_num": table_idx,
                "table_bbox": bbox,  # [x1, y1, x2, y2] for image detection
                "crop_filename": crop_filename,
                "crop_path": str(crop_path)
            })

        doc.close()

    # End timing
    elapsed = time.perf_counter() - start_time

    # Save metadata
    with open(metadata_json_out, 'w', encoding='utf-8') as f:
        json.dump(table_crop_metadata, f, indent=2, ensure_ascii=False)

    print(f"\n[PHASE 1 COMPLETE]")
    print(f"Total crops: {len(table_crop_metadata)}")

    # Calculate unique PDFs processed
    unique_pdfs = len(set(det["pdf_num"] for det in all_detections))

    if unique_pdfs > 0:
        print(f"Phase 1 Runtime: {elapsed:.2f}s ({unique_pdfs} PDFs, {elapsed/unique_pdfs:.2f}s per PDF)")
    else:
        print(f"Phase 1 Runtime: {elapsed:.2f}s (no PDFs to process)")

    print(f"Metadata saved: {metadata_json_out}")

if __name__ == "__main__":
    main()
