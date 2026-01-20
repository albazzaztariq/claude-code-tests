"""
Timing Utilities for ScholarSweep Pipeline

Shared timing mechanism to track performance across all pipelines.
All timers are stored in timers.json in the Query folder.
"""

import json
from pathlib import Path
from typing import Optional


def get_timers_file(query_folder: Path) -> Path:
    """Get path to timers.json file in Query folder."""
    return query_folder / "timers.json"


def load_timers(query_folder: Path) -> dict:
    """Load timers from timers.json file.

    Returns:
        dict with timer names as keys and values in seconds
    """
    timers_file = get_timers_file(query_folder)

    if timers_file.exists():
        with open(timers_file, 'r') as f:
            return json.load(f)
    else:
        # Initialize with all 11 timers
        return {
            "1_query_generation": None,
            "2_openalex_first_call": None,
            "3_openalex_pagination_avg": None,
            "4_unpaywall_first_call": None,
            "5_corpus_download_total": None,
            "6_text_extraction_first": None,
            "7_metrics_match_first": None,
            "8_llm_filter_first": None,
            "9_llm_extraction_first": None,
            "10_table_extraction_first": None,
            "11_final_output": None
        }


def save_timer(query_folder: Path, timer_name: str, value: float):
    """Save a single timer value to timers.json.

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
    """Display all timers in formatted output."""
    timers = load_timers(query_folder)

    print("\n" + "=" * 80)
    print("PIPELINE PERFORMANCE TIMERS")
    print("=" * 80)

    # Format each timer (add space before single digits for alignment)
    timer_labels = {
        "1_query_generation": " 1.  Query Generation",
        "2_openalex_first_call": " 2.  OpenAlex API (first call)",
        "3_openalex_pagination_avg": " 3.  OpenAlex Pagination (average)",
        "4_unpaywall_first_call": " 4.  UnPaywall API (first call)",
        "5_corpus_download_total": " 5.  Corpus Download (total)",
        "6_text_extraction_first": " 6.  Full Text Extraction (first)",
        "7_metrics_match_first": " 7.  Metrics Match Textfile (first)",
        "8_llm_filter_first": " 8.  LLM Relevance Filter (first call)",
        "9_llm_extraction_first": " 9.  LLM Text Extraction (first call)",
        "10_table_extraction_first": "10.  Table Extraction (first)",
        "11_final_output": "11.  Create Final Output"
    }

    # Sort by numeric prefix (1_, 2_, ..., 10_, 11_)
    sorted_keys = sorted(timers.keys(), key=lambda x: int(x.split('_')[0]))

    for key in sorted_keys:
        label = timer_labels.get(key, key)
        value = timers[key]

        if value is not None:
            if value < 1:
                time_str = f"{value*1000:.0f}ms"
            else:
                time_str = f"{value:.2f}s"
            print(f"  {label:45} {time_str:>10}")
        else:
            print(f"  {label:45} {'N/A':>10}")

    print("=" * 80)
