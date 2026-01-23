"""
LLM Filter - Relevance Checking with Qwen 2.5 7B

Filters extracted text to determine if papers are relevant to fabric research.

PROCESS:
1. Read metadata JSON file (contains PDF Number and empty Breakpoint field)
2. For each PDF, read the full text extraction file
3. Send to Qwen 2.5 7B LLM (via Ollama) with filtering prompt
4. Update Breakpoint field with response
5. Save updated JSON file and delete irrelevant paper files

INPUTS:
- Metadata Excel file (.xlsx)
- Full text extraction files (.txt)

OUTPUTS:
- Updated Excel file with Breakpoint column filled

BREAKPOINT VALUES:
- "Passes" - Paper performs experiment with knit/woven fabrics or yarns
- "No experiment" - Paper does not perform an experiment
- "No knit/woven/yarn" - Paper performs experiment but not with relevant materials
- "Error: <message>" - Processing error occurred
"""

import os
import sys
from pathlib import Path
from typing import Dict, List, Tuple
import time
import json

# Import timing and corpus utilities
sys.path.insert(0, str(Path(__file__).parent.parent))
from Timers.timing_utils import save_timer
from Utils.corpus_utils import load_corpus_json, save_corpus_json

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

import subprocess
import re

def load_filter_prompt() -> str:
    """Load the filtering prompt from Config/LLMRelevancePrompt.txt."""
    prompt_file = Path(__file__).parent.parent / "Config" / "LLMRelevancePrompt.txt"

    if not prompt_file.exists():
        print(f"ERROR: Filter prompt file not found: {prompt_file.name}")
        sys.exit(1)

    with open(prompt_file, 'r', encoding='utf-8') as f:
        return f.read()


# Load filtering prompt from config
FILTER_PROMPT = load_filter_prompt()


def call_qwen(text: str, max_retries: int = 3) -> str:
    """
    Call Qwen 2.5 7B via Ollama for relevance classification.

    Args:
        text: Extracted text from PDF (first 3 pages, ~5000 chars)
        max_retries: Number of retry attempts if call fails

    Returns:
        Filter decision: "Passes", "No experiment", "No knit/woven/yarn", or "Error: <message>"
    """
    # Limit text to first 5000 characters (first 3 pages usually)
    text_sample = text[:5000]

    # Build prompt with text
    prompt = FILTER_PROMPT.replace("{text}", text_sample)

    for attempt in range(max_retries):
        try:
            # Call Ollama
            cmd = [
                "ollama", "run", "qwen2.5:7b-instruct-q5_K_M",
                prompt
            ]

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=30  # 30 second timeout
            )

            if result.returncode != 0:
                error_msg = result.stderr.strip() if result.stderr else "Unknown error"
                if attempt < max_retries - 1:
                    print(f"    [RETRY {attempt + 1}/{max_retries}] Ollama error: {error_msg}")
                    time.sleep(1)
                    continue
                return f"Error: Ollama failed - {error_msg}"

            # Clean response (remove ANSI codes and whitespace)
            response = result.stdout.strip()
            response = re.sub(r'\x1b\[[0-9;]*[a-zA-Z]', '', response)  # Remove ANSI codes
            response = re.sub(r'\[?[0-9]+[hlmKG]', '', response)  # Remove spinner codes
            response = response.strip()

            # Clean up common formatting issues
            # Sometimes model adds extra text before/after the decision
            lines = [line.strip() for line in response.split('\n') if line.strip()]

            # Look for the decision in the response
            valid_decisions = ["Passes", "No experiment", "No knit/woven/yarn"]

            for line in lines:
                # Check if line contains any valid decision
                for decision in valid_decisions:
                    if decision.lower() in line.lower():
                        return decision

            # If no valid decision found, try the last non-empty line
            if lines:
                last_line = lines[-1]
                # Check if it's close to a valid decision
                for decision in valid_decisions:
                    if decision.lower() in last_line.lower():
                        return decision

            # If still nothing, return the cleaned response as-is
            # (might be a valid decision with slight formatting)
            return response if response else "Error: Empty response from LLM"

        except subprocess.TimeoutExpired:
            if attempt < max_retries - 1:
                print(f"    [RETRY {attempt + 1}/{max_retries}] LLM call timed out")
                time.sleep(1)
                continue
            return "Error: LLM call timed out"

        except Exception as e:
            if attempt < max_retries - 1:
                print(f"    [RETRY {attempt + 1}/{max_retries}] Exception: {str(e)}")
                time.sleep(1)
                continue
            return f"Error: {str(e)}"

    return "Error: Max retries reached"


def filter_papers(pdf_folder: Path) -> None:
    """
    Filter papers using Lambda AI LLM and update corpus metadata JSON.
    Deletes irrelevant papers and their files.

    Args:
        pdf_folder: Folder containing PDFs and text extraction files
    """
    query_folder = pdf_folder.parent

    all_papers = load_corpus_json(query_folder)

    # Filter to only processable papers (exclude failed and unreadable)
    processable_statuses = ["downloaded", "html", "xml", "json", "unknown"]
    processable_papers = [p for p in all_papers if p.get("status", "") in processable_statuses]

    total_papers = len(processable_papers)

    # Process each paper
    total_start = time.time()
    deleted_count = 0
    total_llm_time = 0.0
    filtered_out_pdfs = set()  # Track PDFs that were filtered out

    for idx, paper in enumerate(processable_papers, 1):
        pdf_num = paper.get("study_number")

        # Check if paper already has a breakpoint from Pipeline 1 (shouldn't happen for processable papers)
        existing_breakpoint = paper.get("breakpoint", "")
        if existing_breakpoint and existing_breakpoint != "":
            continue

        # Find text file (in Extracted Text folder)
        text_file = query_folder / "Extracted Text" / f"{pdf_num}.txt"

        if not text_file.exists():
            paper["breakpoint"] = "Error: Text file not found"
            continue

        # Read full text
        try:
            with open(text_file, 'r', encoding='utf-8') as f:
                text = f.read()

            if not text.strip():
                paper["breakpoint"] = "Error: Empty text file"
                continue

        except Exception as e:
            paper["breakpoint"] = f"Error: {str(e)}"
            continue

        # Call Qwen 2.5 7B LLM (Timer 7: Track LLM time)
        start_time = time.time()

        decision = call_qwen(text)

        elapsed = time.time() - start_time

        # Accumulate LLM time
        total_llm_time += elapsed

        # Update paper with decision
        paper["breakpoint"] = decision

        # If irrelevant, MOVE files to Filtered Papers folder
        if decision in ["No experiment", "No knit/woven/yarn"]:
            # Create Filtered Papers folder
            filtered_folder = query_folder / "Filtered Papers"
            filtered_folder.mkdir(exist_ok=True)

            # Move files
            extracted_text_folder = query_folder / "Extracted Text"
            files_to_move = [
                (pdf_folder / f"{pdf_num}.pdf", filtered_folder / f"{pdf_num}.pdf"),
                (pdf_folder / f"{pdf_num}.html", filtered_folder / f"{pdf_num}.html"),
                (pdf_folder / f"{pdf_num}.xml", filtered_folder / f"{pdf_num}.xml"),
                (pdf_folder / f"{pdf_num}.json", filtered_folder / f"{pdf_num}.json"),
                (extracted_text_folder / f"{pdf_num}.txt", filtered_folder / f"{pdf_num}.txt"),
                (extracted_text_folder / f"{pdf_num}_filtered.txt", filtered_folder / f"{pdf_num}_filtered.txt"),
                (extracted_text_folder / f"{pdf_num}_filtered_NEW.txt", filtered_folder / f"{pdf_num}_filtered_NEW.txt"),
            ]

            for src_path, dest_path in files_to_move:
                if src_path.exists():
                    try:
                        src_path.rename(dest_path)
                    except Exception:
                        pass

            # Track this PDF as filtered out (for table detections cleanup)
            filtered_out_pdfs.add(pdf_num)
            deleted_count += 1

        elif not decision.startswith("Error:"):
            # Passes - clear breakpoint, keep files
            paper["breakpoint"] = ""  # Clear breakpoint for papers that pass

        # Small delay between papers
        time.sleep(0.5)

    # Save updated JSON
    total_time = time.time() - total_start

    # Save Timer 7: Total LLM filter time
    if total_llm_time > 0:
        save_timer(query_folder, "7_llm_filter_total", total_llm_time)

    # Clean up table detections JSON (remove detections for filtered-out PDFs)
    if filtered_out_pdfs:
        table_detections_json = query_folder / "yolo_table_detections.json"

        if table_detections_json.exists():
            # Load table detections
            with open(table_detections_json, 'r', encoding='utf-8') as f:
                all_detections = json.load(f)

            # Filter out detections for filtered-out PDFs
            kept_detections = [
                det for det in all_detections
                if det.get("pdf_num") not in filtered_out_pdfs
            ]

            # Save cleaned detections
            with open(table_detections_json, 'w', encoding='utf-8') as f:
                json.dump(kept_detections, f, indent=2)

    save_corpus_json(all_papers, query_folder)

    # Count papers that passed filter (no breakpoint or empty breakpoint)
    passed_papers = [p for p in all_papers if not p.get("breakpoint", "").strip()]
    print(f"Papers filtered for relevance: {deleted_count}")
    print(f"Final Corpus Size: {len(passed_papers)}")


# =============================================================================
# QUEUE-BASED SINGLE PAPER PROCESSING (For Pipeline Parallelization)
# =============================================================================

def filter_single_paper_for_queue(pdf_num: int, text_path: Path, query_folder: Path, all_papers: list) -> dict:
    """
    Filter a single paper for queue-based parallelization.

    Args:
        pdf_num: PDF number (study number)
        text_path: Path to extracted text file
        query_folder: Query folder
        all_papers: List of all papers (metadata) - will be updated in place

    Returns:
        dict: {
            "pdf_num": int,
            "decision": str,
            "llm_time": float,
            "filtered_out": bool,
            "success": bool
        }
    """
    # Find paper in metadata
    paper = None
    for p in all_papers:
        if p.get("study_number") == pdf_num:
            paper = p
            break

    if not paper:
        return {
            "pdf_num": pdf_num,
            "decision": "Error: Not in metadata",
            "llm_time": 0.0,
            "filtered_out": False,
            "success": False
        }

    # Check if paper already has breakpoint
    existing_breakpoint = paper.get("breakpoint", "")
    if existing_breakpoint and existing_breakpoint != "":
        return {
            "pdf_num": pdf_num,
            "decision": existing_breakpoint,
            "llm_time": 0.0,
            "filtered_out": False,
            "success": True
        }

    # Read full text
    if not text_path.exists():
        paper["breakpoint"] = "Error: Text file not found"
        return {
            "pdf_num": pdf_num,
            "decision": "Error: Text file not found",
            "llm_time": 0.0,
            "filtered_out": False,
            "success": False
        }

    try:
        with open(text_path, 'r', encoding='utf-8') as f:
            text = f.read()

        if not text.strip():
            paper["breakpoint"] = "Error: Empty text file"
            return {
                "pdf_num": pdf_num,
                "decision": "Error: Empty text file",
                "llm_time": 0.0,
                "filtered_out": False,
                "success": False
            }

    except Exception as e:
        paper["breakpoint"] = f"Error: {str(e)}"
        return {
            "pdf_num": pdf_num,
            "decision": f"Error: {str(e)}",
            "llm_time": 0.0,
            "filtered_out": False,
            "success": False
        }

    # Call LLM
    start_time = time.time()
    decision = call_qwen(text)
    llm_time = time.time() - start_time

    # Update paper metadata
    paper["breakpoint"] = decision

    # Move files if irrelevant
    filtered_out = False
    if decision in ["No experiment", "No knit/woven/yarn"]:
        # Create Filtered Papers folder
        filtered_folder = query_folder / "Filtered Papers"
        filtered_folder.mkdir(exist_ok=True)

        # Move files
        pdf_folder = query_folder / "Corpus PDFs"
        extracted_text_folder = query_folder / "Extracted Text"
        files_to_move = [
            (pdf_folder / f"{pdf_num}.pdf", filtered_folder / f"{pdf_num}.pdf"),
            (pdf_folder / f"{pdf_num}.html", filtered_folder / f"{pdf_num}.html"),
            (pdf_folder / f"{pdf_num}.xml", filtered_folder / f"{pdf_num}.xml"),
            (pdf_folder / f"{pdf_num}.json", filtered_folder / f"{pdf_num}.json"),
            (extracted_text_folder / f"{pdf_num}.txt", filtered_folder / f"{pdf_num}.txt"),
            (extracted_text_folder / f"{pdf_num}_filtered.txt", filtered_folder / f"{pdf_num}_filtered.txt"),
            (extracted_text_folder / f"{pdf_num}_filtered_NEW.txt", filtered_folder / f"{pdf_num}_filtered_NEW.txt"),
        ]

        for src_path, dest_path in files_to_move:
            if src_path.exists():
                try:
                    src_path.rename(dest_path)
                except Exception:
                    pass

        filtered_out = True

    elif decision == "Passes" or decision.startswith("Passes"):
        # Clear breakpoint for papers that pass
        paper["breakpoint"] = ""

    return {
        "pdf_num": pdf_num,
        "decision": decision,
        "llm_time": llm_time,
        "filtered_out": filtered_out,
        "success": True
    }


def find_latest_query_folder():
    """Find the latest Query folder in Output directory."""
    output_base = Path(__file__).parent.parent / "Output"
    query_folders = [d for d in output_base.iterdir() if d.is_dir() and "Query" in d.name]
    if not query_folders:
        return None
    return max(query_folders, key=lambda d: d.stat().st_mtime)


def main():
    """CLI entry point for LLM filtering."""
    import argparse

    parser = argparse.ArgumentParser(description="Filter papers using Qwen 2.5 7B LLM")
    parser.add_argument("pdf_folder", nargs='?', type=Path, help="Folder containing PDFs and text files")

    args = parser.parse_args()

    # Auto-detect if no arguments provided
    if args.pdf_folder is None:
        query_folder = find_latest_query_folder()
        if query_folder is None:
            print("ERROR: No Query folder found in Output directory")
            sys.exit(1)

        args.pdf_folder = query_folder / "Corpus PDFs"

        print(f"Auto-detected Query folder: {query_folder.name}")

    if not args.pdf_folder.exists():
        print(f"ERROR: PDF folder not found: {args.pdf_folder.name}")
        sys.exit(1)

    print("=" * 80)
    print("LLM FILTER - Qwen 2.5 7B Paper Relevance Checking")
    print("=" * 80)
    print(f"PDF Folder: {args.pdf_folder.name}")
    print("=" * 80 + "\n")

    filter_papers(args.pdf_folder)


if __name__ == "__main__":
    main()
