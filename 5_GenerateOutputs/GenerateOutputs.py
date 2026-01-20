"""
ScholarSweep Final Output Combiner

Combines outputs from Pipelines 1-4 into a single final workbook.

PROCESS:
1. Add "Extractor" column to 4a_Metadata.xlsx (mark all rows as "LLM")
2. Add "Extractor" column to 4b_Metadata.xlsx (mark all rows as "TableExtraction")
3. Combine 4a and 4b metadata into single "Extraction Results" sheet (each study has 2 rows)
4. Add "Extractor" column to 4a_FinalDataset.xlsx (mark as "LLM")
5. Add "Extractor" column to 4b_FinalDataset.xlsx (mark as "TableExtraction")
6. Combine both datasets into single "Metrics" sheet
7. Create final 3-sheet workbook:
   - Sheet 1: "Corpus" (original metadata from Pipelines 1+3)
   - Sheet 2: "Extraction Results" (combined metadata from 4a+4b)
   - Sheet 3: "Metrics" (combined dataset from 4a+4b)
8. Delete intermediate files

INPUTS:
- Corpus Metadata.xlsx (from Pipeline 1+3)
- 4a_Metadata.xlsx
- 4a_FinalDataset.xlsx
- 4b_Metadata.xlsx
- 4b_FinalDataset.xlsx

OUTPUTS:
- Query Output [mmddyy-hh]H[mm]M.xlsx (3 sheets)
"""

import sys
from pathlib import Path
import time

# Import timing utilities
sys.path.insert(0, str(Path(__file__).parent.parent))
from Timers.timing_utils import save_timer, display_timers

# Check for required libraries
try:
    import pandas as pd
except ImportError:
    print("ERROR: Missing required library 'pandas'")
    print("Install it with: pip install pandas openpyxl")
    sys.exit(1)

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
except ImportError:
    print("ERROR: Missing required library 'openpyxl'")
    print("Install it with: pip install openpyxl")
    sys.exit(1)


def combine_metadata(metadata_4a_path: Path, metadata_4b_path: Path) -> pd.DataFrame:
    """
    Combine 4a and 4b metadata into single DataFrame with "Extractor" column.
    Each study will have 2 rows - one from 4a (LLM), one from 4b (TableExtraction).

    Args:
        metadata_4a_path: Path to 4a_Metadata.xlsx
        metadata_4b_path: Path to 4b_Metadata.xlsx

    Returns:
        Combined DataFrame with "Extractor" column at beginning
    """
    print("\n[1/4] Combining 4a and 4b metadata...")

    # Load metadata files
    df_4a = pd.read_excel(metadata_4a_path, engine='openpyxl')
    df_4b = pd.read_excel(metadata_4b_path, engine='openpyxl')

    print(f"  4a metadata: {len(df_4a)} rows")
    print(f"  4b metadata: {len(df_4b)} rows")

    # Add "Extractor" column at the beginning
    df_4a.insert(0, "Extractor", "LLM")
    df_4b.insert(0, "Extractor", "TableExtraction")

    print(f"  Added 'Extractor' column to both metadata files")

    # Combine (each study now has 2 rows)
    combined = pd.concat([df_4a, df_4b], ignore_index=True)

    # Sort by PDF Number so both rows for each study are together
    if "PDF Number" in combined.columns:
        combined = combined.sort_values("PDF Number")

    print(f"  Combined metadata: {len(combined)} rows (each study has 2 rows)\n")

    return combined


def combine_datasets(dataset_4a_path: Path, dataset_4b_path: Path) -> pd.DataFrame:
    """
    Combine 4a and 4b final datasets into single DataFrame with "Extractor" column.

    Args:
        dataset_4a_path: Path to 4a_FinalDataset.xlsx
        dataset_4b_path: Path to 4b_FinalDataset.xlsx

    Returns:
        Combined DataFrame with "Extractor" column at beginning
    """
    print("[2/4] Combining 4a and 4b datasets...")

    # Load datasets
    df_4a = pd.read_excel(dataset_4a_path, engine='openpyxl')
    df_4b = pd.read_excel(dataset_4b_path, engine='openpyxl')

    print(f"  4a dataset: {len(df_4a)} rows")
    print(f"  4b dataset: {len(df_4b)} rows")

    # Add "Extractor" column
    df_4a.insert(0, "Extractor", "LLM")
    df_4b.insert(0, "Extractor", "TableExtraction")

    print(f"  Added 'Extractor' column to both datasets")

    # Combine
    combined = pd.concat([df_4a, df_4b], ignore_index=True)

    print(f"  Combined dataset: {len(combined)} rows\n")

    return combined


def create_final_workbook(
    query_folder: Path,
    original_metadata: Path,
    combined_metadata: pd.DataFrame,
    combined_dataset: pd.DataFrame
) -> Path:
    """
    Create final 3-sheet workbook.

    Args:
        query_folder: Query folder path
        original_metadata: Original metadata Excel from Pipeline 1+3
        combined_metadata: Combined metadata from 4a+4b
        combined_dataset: Combined metrics dataset from 4a+4b

    Returns:
        Path to final workbook
    """
    print("[3/4] Creating final 3-sheet workbook...")

    # Extract timestamp from query folder name (format: "mm-dd-yy-hhmm Query")
    # Convert to output format: "Query Output mm-dd-yy-hhHmmM.xlsx"
    folder_name = query_folder.name
    if " Query" in folder_name:
        timestamp_part = folder_name.replace(" Query", "")
        parts = timestamp_part.split("-")
        if len(parts) == 4:
            mm, dd, yy, hhmm = parts
            hh = hhmm[:2]
            min_part = hhmm[2:]
            output_filename = f"Query Output {mm}-{dd}-{yy}-{hh}H{min_part}M.xlsx"
        else:
            output_filename = "Query Output.xlsx"  # Fallback
    else:
        output_filename = "Query Output.xlsx"  # Fallback

    final_workbook_path = query_folder / output_filename

    # Create new workbook
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Sheet 1: "Corpus" (original metadata from Pipeline 1+3)
    print("  [1/3] Adding 'Corpus' sheet...")
    wb_original = load_workbook(original_metadata)
    ws_original = wb_original.active

    ws1 = wb.create_sheet("Corpus")
    for row in ws_original.iter_rows():
        ws1.append([cell.value for cell in row])

    # Copy formatting (colors)
    from openpyxl.styles import PatternFill
    for row_idx, row in enumerate(ws_original.iter_rows(), 1):
        for col_idx, cell in enumerate(row, 1):
            if cell.fill and cell.fill.start_color:
                # Create a new PatternFill instead of copying the object
                new_fill = PatternFill(
                    start_color=cell.fill.start_color.rgb if hasattr(cell.fill.start_color, 'rgb') else cell.fill.start_color.index,
                    end_color=cell.fill.end_color.rgb if hasattr(cell.fill.end_color, 'rgb') else cell.fill.end_color.index,
                    fill_type=cell.fill.fill_type
                )
                ws1.cell(row=row_idx, column=col_idx).fill = new_fill

    print(f"      Copied {ws_original.max_row} rows")

    # Sheet 2: "Extraction Results" (combined metadata from 4a+4b)
    print("  [2/3] Adding 'Extraction Results' sheet...")
    ws2 = wb.create_sheet("Extraction Results")

    # Filter out rows where PDF was never successfully downloaded
    # Download failures from Pipeline 1: "Download Failed", "HTML", "XML", "JSON"
    download_failures = ["Download Failed", "HTML", "XML", "JSON"]
    filtered_metadata = combined_metadata[~combined_metadata["Breakpoint"].isin(download_failures)]

    filtered_count = len(combined_metadata) - len(filtered_metadata)
    if filtered_count > 0:
        print(f"      Filtered out {filtered_count} rows (PDFs that were never downloaded)")

    # Add rows to sheet
    for row in dataframe_to_rows(filtered_metadata, index=False, header=True):
        ws2.append(row)

    # Color rows red where Breakpoint contains "Error:"
    from openpyxl.styles import PatternFill
    red_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")

    # Find Breakpoint column index
    headers = [cell.value for cell in ws2[1]]
    if "Breakpoint" in headers:
        breakpoint_col_idx = headers.index("Breakpoint") + 1

        error_count = 0
        for row_idx in range(2, ws2.max_row + 1):
            breakpoint_value = ws2.cell(row=row_idx, column=breakpoint_col_idx).value
            if breakpoint_value and isinstance(breakpoint_value, str) and "Error:" in breakpoint_value:
                # Color entire row red
                for col in range(1, len(headers) + 1):
                    ws2.cell(row=row_idx, column=col).fill = red_fill
                error_count += 1

        if error_count > 0:
            print(f"      Colored {error_count} error rows red")

    print(f"      Added {len(filtered_metadata)} rows (each study has 2 rows)")

    # Sheet 3: "Metrics" (combined dataset from 4a+4b)
    print("  [3/3] Adding 'Metrics' sheet...")
    ws3 = wb.create_sheet("Metrics")

    for row in dataframe_to_rows(combined_dataset, index=False, header=True):
        ws3.append(row)

    print(f"      Added {len(combined_dataset)} data rows")

    # Save workbook
    wb.save(final_workbook_path)
    print(f"\n  Saved: {final_workbook_path.name}\n")

    return final_workbook_path


def cleanup_intermediate_files(metadata_4a: Path, metadata_4b: Path, dataset_4a: Path, dataset_4b: Path):
    """
    Delete intermediate files.

    Args:
        metadata_4a: Path to 4a_Metadata.xlsx
        metadata_4b: Path to 4b_Metadata.xlsx
        dataset_4a: Path to 4a_FinalDataset.xlsx
        dataset_4b: Path to 4b_FinalDataset.xlsx
    """
    print("[4/4] Cleaning up intermediate files...")

    for file_path in [metadata_4a, metadata_4b, dataset_4a, dataset_4b]:
        if file_path.exists():
            file_path.unlink()
            print(f"  Deleted: {file_path.name}")

    print()


def run_combiner(query_folder: Path):
    """
    Run final output combiner.

    Args:
        query_folder: Path to Query folder containing all outputs
    """
    # Timer 11: Start timing final output creation
    combiner_start = time.time()

    print("=" * 80)
    print("SCHOLARSWEEP FINAL OUTPUT COMBINER")
    print("=" * 80)
    print(f"Query Folder: {query_folder}")
    print("=" * 80)
    print()

    # Find input files
    original_metadata = query_folder / "Corpus Metadata.xlsx"  # From Pipeline 1+3
    metadata_4a = query_folder / "4a_Metadata.xlsx"
    metadata_4b = query_folder / "4b_Metadata.xlsx"
    dataset_4a = query_folder / "4a_FinalDataset.xlsx"
    dataset_4b = query_folder / "4b_FinalDataset.xlsx"

    # Verify all files exist
    missing_files = []
    for file_path in [original_metadata, metadata_4a, metadata_4b, dataset_4a, dataset_4b]:
        if not file_path.exists():
            missing_files.append(file_path.name)

    if missing_files:
        print("ERROR: Missing required files:")
        for filename in missing_files:
            print(f"  - {filename}")
        print("\nEnsure Pipelines 1-4 have completed successfully.")
        sys.exit(1)

    print("All input files found:\n")
    print(f"  Original Metadata: {original_metadata.name}")
    print(f"  4a Metadata: {metadata_4a.name}")
    print(f"  4b Metadata: {metadata_4b.name}")
    print(f"  4a Dataset: {dataset_4a.name}")
    print(f"  4b Dataset: {dataset_4b.name}")

    # Combine metadata (4a + 4b -> single Corpus with 2 rows per study)
    combined_metadata = combine_metadata(metadata_4a, metadata_4b)

    # Combine datasets (4a + 4b -> single Metrics sheet)
    combined_dataset = combine_datasets(dataset_4a, dataset_4b)

    # Create final workbook (3 sheets)
    final_workbook = create_final_workbook(
        query_folder,
        original_metadata,
        combined_metadata,
        combined_dataset
    )

    # Cleanup
    cleanup_intermediate_files(metadata_4a, metadata_4b, dataset_4a, dataset_4b)

    # Summary
    print("Pipeline combination complete!\n")

    # Timer 11: Save final output creation time
    combiner_time = time.time() - combiner_start
    save_timer(query_folder, "11_final_output", combiner_time)

    print("=" * 80)
    print("FINAL OUTPUT COMPLETE")
    print("=" * 80)
    print(f"\nFinal Output: {final_workbook.name}")
    print("\nWorkbook Structure:")
    print("  Sheet 1: Corpus (original metadata from Pipelines 1+3)")
    print("  Sheet 2: Extraction Results (combined metadata from 4a+4b, 2 rows per study)")
    print("  Sheet 3: Metrics (combined dataset from 4a+4b)")
    print("=" * 80)

    # Display all 11 timers
    display_timers(query_folder)


def find_latest_query_folder():
    """Find the latest Query folder in Output directory."""
    output_base = Path(__file__).parent.parent / "Output"
    query_folders = [d for d in output_base.iterdir() if d.is_dir() and "Query" in d.name]
    if not query_folders:
        return None
    return max(query_folders, key=lambda d: d.stat().st_mtime)


def main():
    """CLI entry point."""
    import argparse

    parser = argparse.ArgumentParser(description="Combine ScholarSweep pipeline outputs")
    parser.add_argument("query_folder", nargs='?', type=Path, help="Query folder path")

    args = parser.parse_args()

    # Auto-detect if no arguments provided
    if args.query_folder is None:
        args.query_folder = find_latest_query_folder()
        if args.query_folder is None:
            print("ERROR: No Query folder found in Output directory")
            sys.exit(1)

        print(f"Auto-detected Query folder: {args.query_folder.name}")

    if not args.query_folder.exists():
        print(f"ERROR: Query folder not found: {args.query_folder}")
        sys.exit(1)

    run_combiner(args.query_folder)


if __name__ == "__main__":
    main()
