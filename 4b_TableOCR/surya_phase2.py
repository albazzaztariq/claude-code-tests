"""
Phase 2: Text Extraction with Surya OCR (GPU)

Runs in isolated subprocess to avoid GPU conflicts with PaddleOCR.
Processes table crops with Surya OCR, filters by metrics.
"""

import sys
import os
import json
import re
import time
from pathlib import Path
from PIL import Image
from collections import defaultdict

# Set Surya environment variables
os.environ['DISABLE_MODEL_SOURCE_CHECK'] = 'True'
os.environ['RECOGNITION_BATCH_SIZE'] = '512'  # Surya batching for 10x speedup

def bbox_overlap(bbox1, bbox2):
    """Calculate overlap area between two bboxes"""
    x1_min, y1_min, x1_max, y1_max = bbox1
    x2_min, y2_min, x2_max, y2_max = bbox2

    x_overlap = max(0, min(x1_max, x2_max) - max(x1_min, x2_min))
    y_overlap = max(0, min(y1_max, y2_max) - max(y1_min, y2_min))

    return x_overlap * y_overlap

def clean_html_tags(text: str) -> str:
    """Remove HTML/markdown tags from text"""
    # Remove HTML tags like <b>, <i>, </b>, </i>, etc.
    cleaned = re.sub(r'<[^>]+>', '', text)
    return cleaned

def filter_columns_by_headers(structured_table: list, metrics_list: list):
    """
    Filter table columns based on header row keywords.

    Keeps only columns where the header (row 0) contains a keyword from MetricsSearchable.txt.
    If no columns match, returns empty table.

    Args:
        structured_table: 2D list of table data (rows x columns)
        metrics_list: List of keywords/phrases from MetricsSearchable.txt

    Returns:
        (filtered_table, matching_columns, found_metrics)
        - filtered_table: 2D list with only matching columns
        - matching_columns: List of column indices that matched
        - found_metrics: List of metrics found in headers
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

def main():
    if len(sys.argv) != 4:
        print("Usage: surya_phase2.py <metadata_json> <metrics_file> <results_json>")
        sys.exit(1)

    metadata_json = Path(sys.argv[1])
    metrics_file = Path(sys.argv[2])
    results_json = Path(sys.argv[3])

    print("[PHASE 2] Surya OCR - Text Extraction (GPU)")
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

    # Import PyTorch and Surya
    print("\nImporting PyTorch and Surya...")
    import torch
    print(f"  PyTorch {torch.__version__} (CUDA: {torch.cuda.is_available()})")

    from surya.table_rec import TableRecPredictor
    from surya.recognition import FoundationPredictor, RecognitionPredictor
    from surya.detection import DetectionPredictor

    print("Loading Surya models...")
    table_rec = TableRecPredictor()
    fp = FoundationPredictor()
    dp = DetectionPredictor()
    rp = RecognitionPredictor(fp)
    print("  All models loaded\n")

    # Group crops by PDF
    crops_by_pdf = defaultdict(lambda: defaultdict(list))
    for meta in table_crop_metadata:
        crops_by_pdf[meta["pdf_num"]][(meta["pdf_name"], meta["page_num"])].append(meta)

    # Process crops
    all_results = []
    total_tables = 0

    # Start timing
    start_time = time.perf_counter()

    for pdf_num in sorted(crops_by_pdf.keys()):
        pdf_data = {
            "pdf_number": pdf_num,
            "filename": None,
            "pages": [],
            "table_count": 0,
            "crops": []
        }

        for (pdf_name, page_num), page_crops in sorted(crops_by_pdf[pdf_num].items()):
            if pdf_data["filename"] is None:
                pdf_data["filename"] = pdf_name

            page_data = {
                "page_number": page_num,
                "tables": []
            }

            for meta in page_crops:
                crop = Image.open(meta["crop_path"])

                print(f"  Processing {meta['crop_filename']}...")

                # Step 1: Detect table structure
                table_results = table_rec([crop])
                table_result = table_results[0]

                # Step 2: Run OCR on WHOLE table image
                recognitions = rp([crop], det_predictor=dp)

                if not recognitions or not recognitions[0].text_lines:
                    print(f"    WARNING: No text detected in {meta['crop_filename']}")
                    continue

                text_lines = recognitions[0].text_lines

                # Step 3: Match text lines to cells by bounding box overlap
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
                            # Clean HTML tags from text
                            cleaned_text = clean_html_tags(text_line.text)
                            cell_text_parts.append(cleaned_text)

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
                    print(f"    Skipping table - no column headers match MetricsSearchable.txt")
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

                print(f"    Table: {len(filtered_table)} rows x {len(filtered_table[0]) if filtered_table else 0} cols (kept {len(matching_columns)}/{len(structured_table[0]) if structured_table else 0}), metrics: {', '.join(found_metrics)}")

            if page_data["tables"]:
                pdf_data["pages"].append(page_data)

        all_results.append(pdf_data)
        total_tables += pdf_data["table_count"]

    # End timing
    elapsed = time.perf_counter() - start_time

    # Save results
    with open(results_json, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, indent=2, ensure_ascii=False)

    print(f"\n[PHASE 2 COMPLETE]")
    print(f"Total tables processed: {total_tables}")
    print(f"Phase 2 Runtime: {elapsed:.2f}s ({total_tables} tables, {elapsed/total_tables if total_tables > 0 else 0:.2f}s per table)")
    print(f"Results saved: {results_json}")

if __name__ == "__main__":
    main()
