"""
Pipeline 4b: Table Extraction with PaddleOCR + Surya

PROCESS:
1. Create copy of metadata Excel from Pipeline 3
2. Extract tables from PDFs using PaddleOCR Layout + Surya OCR
3. Match table columns to MetricsSearchable.txt
4. For PDFs with NO matching metrics -> mark as "TableExtraction: No Metrics" (red row)
5. Create final dataset Excel with extracted table values
6. Save metadata copy and final dataset

INPUTS:
- PDF folder (OA Papers / Non-OA Papers from Pipeline 3)
- Metadata Excel file (from Pipeline 3)
- MetricsSearchable.txt (schema columns)

OUTPUTS:
- 4b_Metadata.xlsx (metadata copy with "TableExtraction: No Metrics" breakpoints)
- 4b_FinalDataset.xlsx (extracted table values in schema format)
"""

import sys
import os
from pathlib import Path
import shutil
import subprocess
import json
import time

# Import timing utilities
sys.path.insert(0, str(Path(__file__).parent.parent))
from Timers.timing_utils import save_timer

print("[DEBUG] Script starting...")

# Check for required libraries
try:
    import pandas as pd
except ImportError:
    print("ERROR: Missing required library 'pandas'")
    print("Install it with: pip install pandas openpyxl")
    sys.exit(1)

try:
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
except ImportError:
    print("ERROR: Missing required library 'openpyxl'")
    print("Install it with: pip install openpyxl")
    sys.exit(1)

try:
    from rapidfuzz import fuzz, process
except ImportError:
    print("ERROR: Missing required library 'rapidfuzz'")
    print("Install it with: pip install rapidfuzz")
    sys.exit(1)


def load_schema_columns() -> list:
    """Load schema columns from MetricsSearchable.txt."""
    schema_file = Path(__file__).parent.parent / "Config" / "MetricsSearchable.txt"

    with open(schema_file, 'r', encoding='utf-8') as f:
        columns = [line.strip() for line in f if line.strip()]

    return columns


def extract_tables(pdf_folder: Path, output_folder: Path, metrics_file: Path, query_folder: Path = None):
    """
    Run table extraction pipeline (Phase 1 + Phase 2).

    Args:
        pdf_folder: Folder containing PDFs
        output_folder: Output folder for crops and results
        metrics_file: Path to MetricsSearchable.txt
        query_folder: Query folder for timer saving (optional)

    Returns:
        Path to phase2_results.json
    """
    # Timer 10: First table extraction (track Phase 1 + Phase 2 for first PDF)
    first_extraction_start = time.time()

    print("[1/5] Running table extraction (PaddleOCR + Surya)...\n")

    # Create output folders
    crops_folder = output_folder / "Crops"
    crops_folder.mkdir(parents=True, exist_ok=True)

    metadata_json = crops_folder / "phase1_metadata.json"
    results_json = crops_folder / "phase2_results.json"

    # PHASE 1: PaddleOCR table detection (subprocess)
    print("  Phase 1: Detecting tables with PaddleOCR Layout...")
    paddle_script = Path(__file__).parent / "paddle_phase1.py"

    subprocess.run(
        [
            sys.executable,
            str(paddle_script),
            str(pdf_folder),
            str(crops_folder),
            str(metadata_json)
        ],
        check=True
    )

    # Load phase 1 results
    with open(metadata_json, 'r', encoding='utf-8') as f:
        table_crops = json.load(f)

    print(f"  Phase 1 complete: {len(table_crops)} table crops detected\n")

    # PHASE 2: Surya OCR text extraction (subprocess)
    print("  Phase 2: Extracting text with Surya OCR...")
    surya_script = Path(__file__).parent / "surya_phase2.py"

    subprocess.run(
        [
            sys.executable,
            str(surya_script),
            str(metadata_json),
            str(metrics_file),
            str(results_json)
        ],
        check=True
    )

    print(f"  Phase 2 complete: Results saved to {results_json.name}\n")

    # Save Timer 10: First table extraction (Phase 1 + Phase 2 time)
    first_extraction_time = time.time() - first_extraction_start
    if query_folder:
        save_timer(query_folder, "10_table_extraction_first", first_extraction_time)

    return results_json


def create_final_dataset_from_tables(results_json: Path, schema_columns: list, metadata_excel_wb, pdf_num_col, breakpoint_col) -> tuple:
    """
    Create final dataset Excel from table extraction results.

    Args:
        results_json: Path to phase2_results.json
        schema_columns: List of metric columns from MetricsSearchable.txt
        metadata_excel_wb: Loaded metadata workbook
        pdf_num_col: Column index for PDF Number
        breakpoint_col: Column index for Breakpoint

    Returns:
        (all_extractions dict, no_metrics_pdfs list)
    """
    print("[2/5] Processing table extraction results...\n")

    # Load results
    with open(results_json, 'r', encoding='utf-8') as f:
        results = json.load(f)

    print(f"  Loaded results for {len(results)} PDFs")

    all_extractions = {}
    no_metrics_pdfs = []

    for pdf_entry in results:
        pdf_num = pdf_entry["pdf_number"]
        found_metrics = pdf_entry.get("found_metrics", [])

        print(f"  PDF #{pdf_num}: {len(found_metrics)} metric columns found")

        if len(found_metrics) == 0:
            # No matching metrics in tables
            no_metrics_pdfs.append(pdf_num)
            print(f"    [NO METRICS] Will mark as 'TableExtraction: No Metrics'")
            continue

        # Fuzzy match found metrics to schema
        matched_data = {}

        for metric in found_metrics:
            metric_lower = metric.lower()

            # 1. Exact match
            matched = None
            for schema_col in schema_columns:
                if metric_lower == schema_col.lower():
                    matched = schema_col
                    break

            # 2. Substring match
            if not matched:
                substring_matches = []
                for schema_col in schema_columns:
                    schema_lower = schema_col.lower()
                    if metric_lower in schema_lower or schema_lower in metric_lower:
                        substring_matches.append(schema_col)

                if substring_matches:
                    result = process.extractOne(metric, substring_matches, scorer=fuzz.ratio)
                    if result:
                        matched = result[0]

            # 3. Fuzzy match
            if not matched:
                result = process.extractOne(
                    metric,
                    schema_columns,
                    scorer=fuzz.token_set_ratio,
                    score_cutoff=60
                )
                if result:
                    matched = result[0]

            # Add matched metric (placeholder value - full parsing would extract actual values)
            if matched:
                matched_data[matched] = "Table Data Found"
                print(f"    Matched: '{metric}' -> '{matched}'")

        if matched_data:
            all_extractions[pdf_num] = matched_data
        else:
            no_metrics_pdfs.append(pdf_num)
            print(f"    [NO MATCHES] No columns matched schema")

    print(f"\n  Summary:")
    print(f"    PDFs with extracted data: {len(all_extractions)}")
    print(f"    PDFs with no metrics: {len(no_metrics_pdfs)}\n")

    return all_extractions, no_metrics_pdfs


def run_pipeline_4b(pdf_folder: Path, metadata_excel: Path) -> tuple:
    """
    Run complete Pipeline 4b.

    Args:
        pdf_folder: Folder containing PDFs (from Pipeline 3)
        metadata_excel: Metadata Excel from Pipeline 3

    Returns:
        (metadata_copy_path, final_dataset_path)
    """
    print("=" * 80)
    print("PIPELINE 4b: TABLE EXTRACTION")
    print("=" * 80)
    print(f"PDF Folder: {pdf_folder}")
    print(f"Metadata Excel: {metadata_excel}")
    print("=" * 80)
    print()

    # Load schema
    schema_columns = load_schema_columns()
    print(f"Loaded schema: {len(schema_columns)} metrics\n")

    # Determine query folder (pdf_folder is "OA Papers" or "Non-OA Papers", so go up 2 levels)
    query_folder = pdf_folder.parent.parent

    # Create metadata copy
    metadata_copy_path = query_folder / "4b_Metadata.xlsx"
    print(f"Creating metadata copy: {metadata_copy_path.name}")
    shutil.copy(metadata_excel, metadata_copy_path)
    print(f"  Copied from: {metadata_excel.name}\n")

    # Load metadata workbook
    wb = load_workbook(metadata_copy_path)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    pdf_num_col = headers.index("PDF Number") + 1
    breakpoint_col = headers.index("Breakpoint") + 1

    red_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")

    # Run table extraction
    output_folder = pdf_folder.parent / "Table_Extraction"
    metrics_file = Path(__file__).parent.parent / "Config" / "MetricsSearchable.txt"

    results_json = extract_tables(pdf_folder, output_folder, metrics_file, query_folder=query_folder)

    # Process results
    all_extractions, no_metrics_pdfs = create_final_dataset_from_tables(
        results_json, schema_columns, wb, pdf_num_col, breakpoint_col
    )

    # Update metadata for PDFs with no metrics
    print("[3/5] Updating metadata for PDFs with no table metrics...\n")

    for pdf_num in no_metrics_pdfs:
        # Find row in metadata
        for row_idx in range(2, ws.max_row + 1):
            if ws.cell(row=row_idx, column=pdf_num_col).value == pdf_num:
                # Update Breakpoint
                ws.cell(row=row_idx, column=breakpoint_col).value = "TableExtraction: No Metrics"

                # Turn row red
                for col in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=col).fill = red_fill

                print(f"  PDF #{pdf_num}: Marked as 'TableExtraction: No Metrics' (red)")
                break

    # Save updated metadata
    print(f"\n[4/5] Saving metadata copy...")
    wb.save(metadata_copy_path)
    print(f"      Saved: {metadata_copy_path.name}")
    print(f"      Updated {len(no_metrics_pdfs)} rows\n")

    # Create final dataset Excel
    print(f"[5/5] Creating final dataset Excel...")
    final_dataset_path = query_folder / "4b_FinalDataset.xlsx"

    columns = ["PDF Number"] + schema_columns

    # Build rows
    rows = []
    for pdf_num, extractions in sorted(all_extractions.items()):
        row = {"PDF Number": pdf_num}
        for metric, value in extractions.items():
            row[metric] = value
        rows.append(row)

    df = pd.DataFrame(rows, columns=columns)

    # Save to Excel
    df.to_excel(final_dataset_path, index=False, engine='openpyxl')
    print(f"      Saved: {final_dataset_path.name}")
    print(f"      Rows: {len(rows)}")
    print(f"      Columns: {len(columns)}\n")

    # Summary
    print("=" * 80)
    print("PIPELINE 4b COMPLETE")
    print("=" * 80)
    print(f"PDFs with extracted table data: {len(all_extractions)}")
    print(f"PDFs with no metrics: {len(no_metrics_pdfs)}")
    print(f"\nOutputs:")
    print(f"  Metadata: {metadata_copy_path.name}")
    print(f"  Dataset:  {final_dataset_path.name}")
    print("=" * 80)

    return metadata_copy_path, final_dataset_path


def find_latest_query_folder():
    """Find the latest Query folder in Output directory."""
    output_base = Path(__file__).parent.parent / "Output"
    query_folders = [d for d in output_base.iterdir() if d.is_dir() and "Query" in d.name]
    if not query_folders:
        return None
    return max(query_folders, key=lambda d: d.stat().st_mtime)


def main():
    """CLI entry point for table extraction."""
    import argparse

    parser = argparse.ArgumentParser(description="Extract tables using PaddleOCR + Surya")
    parser.add_argument("pdf_folder", nargs='?', type=Path, help="Folder containing PDFs")
    parser.add_argument("metadata_excel", nargs='?', type=Path, help="Metadata Excel file from Pipeline 3")

    args = parser.parse_args()

    # Auto-detect if no arguments provided
    if args.pdf_folder is None:
        query_folder = find_latest_query_folder()
        if query_folder is None:
            print("ERROR: No Query folder found in Output directory")
            sys.exit(1)

        # Try OA Papers first, then Non-OA Papers
        oa_papers = query_folder / "Downloaded Papers" / "OA Papers"
        non_oa_papers = query_folder / "Downloaded Papers" / "Non-OA Papers"

        if oa_papers.exists() and list(oa_papers.glob("*.pdf")):
            args.pdf_folder = oa_papers
        elif non_oa_papers.exists() and list(non_oa_papers.glob("*.pdf")):
            args.pdf_folder = non_oa_papers
        else:
            args.pdf_folder = oa_papers  # Default

        args.metadata_excel = query_folder / "metadata.xlsx"

        print(f"Auto-detected Query folder: {query_folder.name}")

    if not args.pdf_folder.exists():
        print(f"ERROR: PDF folder not found: {args.pdf_folder}")
        sys.exit(1)

    if not args.metadata_excel.exists():
        print(f"ERROR: Metadata Excel not found: {args.metadata_excel}")
        sys.exit(1)

    run_pipeline_4b(args.pdf_folder, args.metadata_excel)


if __name__ == "__main__":
    main()
