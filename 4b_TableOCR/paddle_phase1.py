"""
Phase 1: Table Detection with PaddleOCR PPStructure (GPU)

Runs in isolated subprocess to avoid GPU conflicts with PyTorch/Surya.
Detects tables in PDFs, saves crops and metadata.

Uses PPStructure instead of LayoutPredictor (fixes model/dict mismatch).
"""

import sys
import json
import time
from pathlib import Path
from PIL import Image
import fitz

# Import PaddleOCR PPStructure (handles layout detection robustly)
from paddleocr import PPStructure

DPI = 300

def detect_table_regions_ppstructure(image_path, ppstructure):
    """Detect table regions using PaddleOCR PPStructure."""
    # PPStructure expects image file path, not numpy array
    results = ppstructure(str(image_path))

    table_boxes = []
    for item in results:
        if isinstance(item, dict) and item.get('type') == 'table':
            bbox = item.get('bbox')
            if bbox:
                x1, y1, x2, y2 = map(int, bbox)
                table_boxes.append((x1, y1, x2, y2))

    return table_boxes

def crop_table_images(page_image, table_boxes):
    """Crop table regions from page image."""
    crops = []
    for x1, y1, x2, y2 in table_boxes:
        crop = page_image.crop((x1, y1, x2, y2))
        crops.append(crop)
    return crops

def main():
    if len(sys.argv) != 4:
        print("Usage: paddle_phase1.py <pdf_folder> <output_folder> <metadata_json>")
        sys.exit(1)

    pdf_folder = Path(sys.argv[1])
    output_folder = Path(sys.argv[2])
    metadata_json = Path(sys.argv[3])

    print("[PHASE 1] PaddleOCR PPStructure - Table Detection (GPU)")
    print(f"PDF Folder: {pdf_folder}")
    print(f"Output Folder: {output_folder}")
    print(f"Metadata JSON: {metadata_json}")

    # Load PaddleOCR PPStructure
    print("\nLoading PaddleOCR PPStructure...")
    ppstructure = PPStructure(
        layout=True,
        ocr=False,
        table=False,
        show_log=False,
        use_gpu=True
    )
    print("PaddleOCR PPStructure loaded\n")

    # Process PDFs
    pdf_files = sorted(pdf_folder.glob("*.pdf"), key=lambda p: int(p.stem))
    print(f"Found {len(pdf_files)} PDFs to process\n")

    table_crop_metadata = []

    # Create temp folder for page images
    temp_folder = output_folder / "temp_pages"
    temp_folder.mkdir(exist_ok=True)

    # Start timing
    start_time = time.perf_counter()

    for pdf_path in pdf_files:
        pdf_num = int(pdf_path.stem)
        doc = fitz.open(pdf_path)

        print(f"Processing PDF {pdf_num}: {pdf_path.name}")
        print(f"  Pages: {len(doc)}")

        for page_num in range(len(doc)):
            page = doc[page_num]

            # Render page
            mat = fitz.Matrix(DPI / 72, DPI / 72)
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Save temp image for PPStructure (it needs file path)
            temp_img_path = temp_folder / f"temp_{pdf_num}_p{page_num + 1}.png"
            img.save(temp_img_path)

            # Detect tables
            table_boxes = detect_table_regions_ppstructure(temp_img_path, ppstructure)

            # Delete temp image
            temp_img_path.unlink()

            if not table_boxes:
                continue

            print(f"  Page {page_num + 1}: Found {len(table_boxes)} tables")

            # Crop and save tables
            table_crops = crop_table_images(img, table_boxes)

            for table_idx, crop in enumerate(table_crops, 1):
                crop_filename = f"{pdf_num}_page{page_num + 1}_table{table_idx}.png"
                crop_path = output_folder / crop_filename
                crop.save(crop_path)

                # Store metadata for Phase 2
                table_crop_metadata.append({
                    "pdf_num": pdf_num,
                    "pdf_name": pdf_path.name,
                    "page_num": page_num + 1,
                    "table_num": table_idx,
                    "crop_filename": crop_filename,
                    "crop_path": str(crop_path)
                })

        doc.close()

    # End timing
    elapsed = time.perf_counter() - start_time

    # Clean up temp folder
    temp_folder.rmdir()

    # Save metadata
    with open(metadata_json, 'w', encoding='utf-8') as f:
        json.dump(table_crop_metadata, f, indent=2, ensure_ascii=False)

    print(f"\n[PHASE 1 COMPLETE]")
    print(f"Total crops: {len(table_crop_metadata)}")
    print(f"Phase 1 Runtime: {elapsed:.2f}s ({len(pdf_files)} PDFs, {elapsed/len(pdf_files):.2f}s per PDF)")
    print(f"Metadata saved: {metadata_json}")

if __name__ == "__main__":
    main()
