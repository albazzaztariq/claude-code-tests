"""
Timing Utilities for ScholarSweep Pipeline

Shared timing mechanism to track performance across all pipelines.
All timers are stored in Runtimes.txt in the Query folder.
"""

import json
from pathlib import Path
from typing import Optional


def get_timers_file(query_folder: Path) -> Path:
    """Get path to .timers_temp.json file in Query folder (internal use only)."""
    return query_folder / ".timers_temp.json"


def get_runtimes_file(query_folder: Path) -> Path:
    """Get path to Runtimes.txt file in Query folder."""
    return query_folder / "Runtimes.txt"


def load_timers(query_folder: Path) -> dict:
    """Load timers from temporary JSON file (internal use only).

    Returns:
        dict with timer names as keys and values in seconds
    """
    timers_file = get_timers_file(query_folder)

    if timers_file.exists():
        with open(timers_file, 'r') as f:
            return json.load(f)
    else:
        # Initialize with all 10 timers (pagination timer removed)
        return {
            "1_query_generation": None,
            "2_openalex_total": None,
            "3_unpaywall_total": None,
            "4_corpus_download_total": None,
            "5_text_extraction_total": None,
            "6_metrics_match_total": None,
            "7_llm_filter_total": None,
            "8_llm_extraction_total": None,
            "9_table_extraction_total": None,
            "10_final_output": None
        }


def save_timer(query_folder: Path, timer_name: str, value: float):
    """Save a single timer value to temporary JSON file.

    Args:
        query_folder: Path to Query folder
        timer_name: Name of timer (e.g., "1_query_generation")
        value: Time value in seconds
    """
    timers = load_timers(query_folder)
    timers[timer_name] = value

    timers_file = get_timers_file(query_folder)
    with open(timers_file, 'w') as f:
        json.dump(timers, f, indent=2)


def display_timers(query_folder: Path):
    """Display all timers in formatted output and save to Runtimes.txt."""
    timers = load_timers(query_folder)

    # Format each timer (add space before single digits for alignment)
    timer_labels = {
        "1_query_generation": " 1.  Query Generation",
        "2_openalex_total": " 2.  OpenAlex API",
        "3_unpaywall_total": " 3.  UnPaywall API",
        "4_corpus_download_total": " 4.  Corpus Download",
        "5_text_extraction_total": " 5.  Text Extraction Pipeline",
        "6_metrics_match_total": " 6.  Metrics Match Pipeline",
        "7_llm_filter_total": " 7.  LLM Relevance Filter",
        "8_llm_extraction_total": " 8.  LLM Text Extraction",
        "9_table_extraction_total": " 9.  Table Extraction",
        "10_final_output": "10.  Create Final Output"
    }

    # Sort by numeric prefix (1_, 2_, ..., 10_, 11_)
    sorted_keys = sorted(timers.keys(), key=lambda x: int(x.split('_')[0]))

    # Build output lines
    output_lines = []
    output_lines.append("=" * 80)
    output_lines.append("PIPELINE PERFORMANCE TIMERS")
    output_lines.append("=" * 80)

    for key in sorted_keys:
        label = timer_labels.get(key, key)
        value = timers[key]

        if value is not None:
            if value < 1:
                # Sub-second: show milliseconds with decimal places
                time_str = f"{value*1000:.2f}ms"
            else:
                # Seconds with 2 decimal places
                time_str = f"{value:.2f}s"
            output_lines.append(f"  {label:45} {time_str:>10}")
        else:
            output_lines.append(f"  {label:45} {'N/A':>10}")

    output_lines.append("=" * 80)

    # Print to terminal
    print("\n" + "\n".join(output_lines))

    # Save to Runtimes.txt
    runtimes_file = get_runtimes_file(query_folder)
    with open(runtimes_file, 'w', encoding='utf-8') as f:
        f.write("\n".join(output_lines))

    # Delete temporary JSON file (with retry for Windows file locks)
    timers_file = get_timers_file(query_folder)
    if timers_file.exists():
        for attempt in range(3):
            try:
                timers_file.unlink()
                break
            except (PermissionError, OSError) as e:
                if attempt < 2:
                    import time
                    time.sleep(0.5 * (attempt + 1))  # 0.5s, 1.0s, then give up
                else:
                    print(f"WARNING: Could not delete {timers_file.name}: {e}")
