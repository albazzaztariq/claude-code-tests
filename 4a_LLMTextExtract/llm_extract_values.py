"""
LLM Text Extract - Value Extraction with SambaNova

Extracts metric values from filtered text using LLM.

PROCESS:
1. Create copy of metadata Excel from Pipeline 3
2. Read all _filtered.txt files in Extracted Text folder
3. For each filtered text file:
   - Parse metric sentences
   - Send to SambaNova LLM for value extraction
   - If NO values extracted -> mark metadata as "LLM: No Metrics" (red row)
4. Create final dataset Excel with extracted values
5. Save metadata copy and final dataset

INPUTS:
- Metadata Excel file (from Pipeline 3)
- Filtered text files (*_filtered.txt from Pipeline 2)
- LLMTextExtractPrompt.txt (system prompt)
- MetricsSearchable.txt (schema columns)

OUTPUTS:
- 4a_Metadata.xlsx (metadata copy with "LLM: No Metrics" breakpoints)
- 4a_FinalDataset.xlsx (extracted values in schema format)
"""

import os
import sys
from pathlib import Path
from typing import Dict, List, Tuple
import time
import shutil

# Import timing utilities
sys.path.insert(0, str(Path(__file__).parent.parent))
from Timers.timing_utils import save_timer

# Load environment variables
from dotenv import load_dotenv
load_dotenv()

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
    import requests
except ImportError:
    print("ERROR: Missing required library 'requests'")
    print("Install it with: pip install requests")
    sys.exit(1)

try:
    from rapidfuzz import fuzz, process
except ImportError:
    print("ERROR: Missing required library 'rapidfuzz'")
    print("Install it with: pip install rapidfuzz")
    sys.exit(1)

# SambaNova API configuration
SAMBANOVA_API_KEY = os.getenv("SAMBANOVA_API_KEY")
SAMBANOVA_CHAT_ENDPOINT = os.getenv("SAMBANOVA_CHAT_ENDPOINT")

if not SAMBANOVA_API_KEY or not SAMBANOVA_CHAT_ENDPOINT:
    print("ERROR: SambaNova API credentials not found in .env file")
    print("Required: SAMBANOVA_API_KEY, SAMBANOVA_CHAT_ENDPOINT")
    sys.exit(1)


def load_extract_prompt() -> str:
    """Load the extraction prompt from Config/LLMTextExtractPrompt.txt."""
    prompt_file = Path(__file__).parent.parent / "Config" / "LLMTextExtractPrompt.txt"

    if not prompt_file.exists():
        print(f"ERROR: Extract prompt file not found: {prompt_file}")
        sys.exit(1)

    with open(prompt_file, 'r', encoding='utf-8') as f:
        return f.read()


# Load extraction prompt from config
EXTRACT_PROMPT = load_extract_prompt()


def load_schema_columns() -> List[str]:
    """Load schema columns from MetricsSearchable.txt."""
    schema_file = Path(__file__).parent.parent / "Config" / "MetricsSearchable.txt"

    with open(schema_file, 'r', encoding='utf-8') as f:
        columns = [line.strip() for line in f if line.strip()]

    return columns


def fuzzy_match_metric(metric_name: str, schema_cols: List[str], threshold: int = 60) -> Tuple[str, float]:
    """
    Fuzzy match metric name to schema columns.

    Returns:
        (matched_column, confidence) or (metric_name, 0.0) if no match
    """
    metric_lower = metric_name.lower()

    # 1. Exact match
    for schema_col in schema_cols:
        if metric_lower == schema_col.lower():
            return schema_col, 1.0

    # 2. Substring match
    substring_matches = []
    for schema_col in schema_cols:
        schema_lower = schema_col.lower()
        if metric_lower in schema_lower or schema_lower in metric_lower:
            substring_matches.append(schema_col)

    if substring_matches:
        result = process.extractOne(metric_name, substring_matches, scorer=fuzz.ratio)
        if result:
            return result[0], 0.9

    # 3. Fuzzy match
    result = process.extractOne(
        metric_name,
        schema_cols,
        scorer=fuzz.token_set_ratio,
        score_cutoff=threshold
    )

    if result:
        matched_col, score, _ = result
        return matched_col, score / 100.0
    else:
        # No match - return original name
        return metric_name, 0.0


# Load schema once at startup
SCHEMA_COLUMNS = load_schema_columns()


def call_sambanova(metric_sentence: str, max_retries: int = 3) -> str:
    """
    Call SambaNova API to extract metric value from sentence.

    Args:
        metric_sentence: Text containing metric match line and sentence
        max_retries: Number of retry attempts on failure

    Returns:
        Extracted value or "No value found"
    """
    # Combine prompt with metric sentence
    full_prompt = f"{EXTRACT_PROMPT}\n\n{metric_sentence}"

    headers = {
        "Authorization": f"Bearer {SAMBANOVA_API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": "Meta-Llama-3.3-70B-Instruct",
        "messages": [
            {"role": "user", "content": full_prompt}
        ],
        "max_tokens": 50,  # Short response needed (just the value)
        "temperature": 0.1  # Low temperature for consistent extraction
    }

    for attempt in range(max_retries):
        try:
            response = requests.post(
                SAMBANOVA_CHAT_ENDPOINT,
                headers=headers,
                json=payload,
                timeout=30
            )

            if response.status_code == 200:
                result = response.json()
                extracted = result["choices"][0]["message"]["content"].strip()
                return extracted
            else:
                print(f"  [Attempt {attempt+1}/{max_retries}] API error: {response.status_code}")
                if attempt < max_retries - 1:
                    time.sleep(2)

        except Exception as e:
            print(f"  [Attempt {attempt+1}/{max_retries}] Error: {e}")
            if attempt < max_retries - 1:
                time.sleep(2)

    return f"Error: Failed after {max_retries} attempts"


def parse_filtered_text(text: str) -> List[Tuple[str, str]]:
    """
    Parse filtered text file into metric-sentence pairs.

    Args:
        text: Content of filtered text file

    Returns:
        List of (metric_label, sentence) tuples
    """
    pairs = []
    lines = text.split('\n')

    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # Look for "Metric Match:" lines
        if line.startswith("Metric Match:"):
            metric_label = line.replace("Metric Match:", "").strip()

            # Get the next non-empty line(s) as the sentence
            i += 1
            sentence_parts = []
            while i < len(lines) and lines[i].strip() and not lines[i].startswith("Metric Match:"):
                sentence_parts.append(lines[i].strip())
                i += 1

            if sentence_parts:
                sentence = " ".join(sentence_parts)
                # Combine metric label and sentence for LLM
                combined = f"Metric Match: {metric_label}\n{sentence}"
                pairs.append((metric_label, combined))
        else:
            i += 1

    return pairs


def process_filtered_file(filtered_path: Path, track_first_call: bool = False) -> tuple[Dict[str, any], float]:
    """
    Process a single filtered text file.

    Args:
        filtered_path: Path to _filtered.txt file
        track_first_call: Whether to track the first LLM call time

    Returns:
        Tuple of (dict with normalized metric names as keys and extracted values, first_call_time or None)
    """
    print(f"  Processing: {filtered_path.name}")

    # Read filtered text
    with open(filtered_path, 'r', encoding='utf-8') as f:
        text = f.read()

    # Parse into metric-sentence pairs
    pairs = parse_filtered_text(text)
    print(f"    Found {len(pairs)} metric sentences")

    if len(pairs) == 0:
        print(f"    No metric sentences found")
        return {}, None

    results = {}
    first_call_time = None

    for idx, (metric_label, combined_text) in enumerate(pairs, 1):
        print(f"    [{idx}/{len(pairs)}] Extracting: {metric_label[:40]}...")

        # Call LLM to extract value (track first call if requested)
        if track_first_call and first_call_time is None:
            call_start = time.time()
            extracted = call_sambanova(combined_text)
            first_call_time = time.time() - call_start
        else:
            extracted = call_sambanova(combined_text)

        # Skip if no value found
        if "no value" in extracted.lower() or "error" in extracted.lower():
            print(f"        Result: {extracted}")
            continue

        # Fuzzy match metric name to schema
        normalized_metric, confidence = fuzzy_match_metric(metric_label, SCHEMA_COLUMNS)

        print(f"        Value: {extracted}")
        print(f"        Matched to: '{normalized_metric}' (confidence: {confidence:.2f})")

        # Store normalized metric name and value
        results[normalized_metric] = extracted

        # Rate limiting (avoid overwhelming API)
        time.sleep(0.5)

    return results, first_call_time


def extract_all_values(extracted_text_folder: Path, metadata_excel: Path) -> Tuple[Path, Path]:
    """
    Extract values from all filtered text files and create outputs.

    Args:
        extracted_text_folder: Folder containing _filtered.txt files
        metadata_excel: Original metadata Excel from Pipeline 3

    Returns:
        (metadata_copy_path, final_dataset_path)
    """
    print("=" * 80)
    print("PIPELINE 4a: LLM TEXT EXTRACTION")
    print("=" * 80)
    print(f"Extracted Text Folder: {extracted_text_folder}")
    print(f"Metadata Excel: {metadata_excel}")
    print("=" * 80)
    print()

    # Create metadata copy
    metadata_copy_path = extracted_text_folder.parent / "4a_Metadata.xlsx"
    print(f"[1/5] Creating metadata copy: {metadata_copy_path.name}")
    shutil.copy(metadata_excel, metadata_copy_path)
    print(f"      Copied from: {metadata_excel.name}\n")

    # Load metadata workbook
    wb = load_workbook(metadata_copy_path)
    ws = wb.active

    # Get headers
    headers = [cell.value for cell in ws[1]]
    pdf_num_col = headers.index("PDF Number") + 1
    breakpoint_col = headers.index("Breakpoint") + 1

    red_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")

    # Find all filtered text files
    print(f"[2/5] Finding filtered text files...")
    filtered_files = sorted(extracted_text_folder.glob("*_filtered.txt"))
    print(f"      Found {len(filtered_files)} files\n")

    # Process each file
    print(f"[3/5] Extracting values from filtered text...\n")
    all_extractions = {}
    no_metrics_count = 0
    first_llm_extraction_time = None

    total_start = time.time()

    for idx, filtered_path in enumerate(filtered_files, 1):
        pdf_num = int(filtered_path.stem.replace("_filtered", ""))

        print(f"PDF {idx}/{len(filtered_files)} (PDF #{pdf_num})")
        print(f"{'-' * 80}")

        # Track first LLM call time (Timer 9)
        track_first = first_llm_extraction_time is None
        results, call_time = process_filtered_file(filtered_path, track_first_call=track_first)

        # Save first LLM extraction call time
        if first_llm_extraction_time is None and call_time is not None:
            first_llm_extraction_time = call_time
            query_folder = extracted_text_folder.parent
            save_timer(query_folder, "9_llm_extraction_first", first_llm_extraction_time)

        if len(results) == 0:
            # No metrics extracted - mark in metadata
            print(f"    [NO METRICS] Updating metadata with 'LLM: No Metrics'\n")

            # Find row in metadata
            for row_idx in range(2, ws.max_row + 1):
                if ws.cell(row=row_idx, column=pdf_num_col).value == pdf_num:
                    # Update Breakpoint
                    ws.cell(row=row_idx, column=breakpoint_col).value = "LLM: No Metrics"

                    # Turn row red
                    for col in range(1, len(headers) + 1):
                        ws.cell(row=row_idx, column=col).fill = red_fill

                    no_metrics_count += 1
                    break
        else:
            # Metrics extracted - store for final dataset
            all_extractions[pdf_num] = results
            print(f"    [SUCCESS] Extracted {len(results)} values\n")

    total_time = time.time() - total_start

    # Save updated metadata
    print(f"[4/5] Saving metadata copy...")
    wb.save(metadata_copy_path)
    print(f"      Saved: {metadata_copy_path.name}")
    print(f"      Updated {no_metrics_count} rows with 'LLM: No Metrics'\n")

    # Create final dataset Excel
    print(f"[5/5] Creating final dataset Excel...")
    final_dataset_path = extracted_text_folder.parent / "4a_FinalDataset.xlsx"

    # Create DataFrame with schema columns
    columns = ["PDF Number"] + SCHEMA_COLUMNS

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
    print("PIPELINE 4a COMPLETE")
    print("=" * 80)
    print(f"Total PDFs processed: {len(filtered_files)}")
    print(f"PDFs with extracted values: {len(all_extractions)}")
    print(f"PDFs with no metrics: {no_metrics_count}")
    print(f"\nOutputs:")
    print(f"  Metadata: {metadata_copy_path.name}")
    print(f"  Dataset:  {final_dataset_path.name}")
    print(f"\nRuntime: {total_time:.1f}s ({total_time/len(filtered_files):.1f}s per PDF)")
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
    """CLI entry point for LLM value extraction."""
    import argparse

    parser = argparse.ArgumentParser(description="Extract metric values using SambaNova LLM")
    parser.add_argument("extracted_text_folder", nargs='?', type=Path, help="Folder containing _filtered.txt files")
    parser.add_argument("metadata_excel", nargs='?', type=Path, help="Metadata Excel file from Pipeline 3")

    args = parser.parse_args()

    # Auto-detect if no arguments provided
    if args.extracted_text_folder is None:
        query_folder = find_latest_query_folder()
        if query_folder is None:
            print("ERROR: No Query folder found in Output directory")
            sys.exit(1)

        args.extracted_text_folder = query_folder / "Extracted Text"
        args.metadata_excel = query_folder / "metadata.xlsx"

        print(f"Auto-detected Query folder: {query_folder.name}")

    if not args.extracted_text_folder.exists():
        print(f"ERROR: Extracted text folder not found: {args.extracted_text_folder}")
        sys.exit(1)

    if not args.metadata_excel.exists():
        print(f"ERROR: Metadata Excel not found: {args.metadata_excel}")
        sys.exit(1)

    extract_all_values(args.extracted_text_folder, args.metadata_excel)


if __name__ == "__main__":
    main()
