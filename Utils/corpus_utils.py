"""
Utilities for saving/loading corpus metadata as JSON.

Replaces "Corpus Metadata.xlsx" with "corpus_metadata.json" to avoid
creating unnecessary Excel files during pipeline execution.

Final Excel worksheet is only created in Pipeline 5 (GenerateOutputs).
"""

import json
from pathlib import Path


def get_corpus_json_path(query_folder: Path) -> Path:
    """Get path to corpus_metadata.json file."""
    return query_folder / "corpus_metadata.json"


def save_corpus_json(all_papers: list[dict], query_folder: Path):
    """Save corpus metadata to JSON file.

    Args:
        all_papers: List of paper metadata dicts
        query_folder: Query folder path
    """
    json_path = get_corpus_json_path(query_folder)

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(all_papers, f, indent=2, ensure_ascii=False)


def load_corpus_json(query_folder: Path) -> list[dict]:
    """Load corpus metadata from JSON file.

    Args:
        query_folder: Query folder path

    Returns:
        List of paper metadata dicts
    """
    json_path = get_corpus_json_path(query_folder)

    if not json_path.exists():
        raise FileNotFoundError(f"Corpus metadata not found: {json_path}")

    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def convert_corpus_json_to_excel_data(all_papers: list[dict]) -> list[list]:
    """Convert corpus JSON data to Excel rows format (for Pipeline 5).

    Returns data rows ready to write to Excel, with color codes embedded.

    Args:
        all_papers: List of paper metadata dicts

    Returns:
        List of rows: [header_row, data_row_1, data_row_2, ...]
        Each data row is: [values..., color_code]
        color_code: "green", "orange", or "red"
    """
    # Header row
    headers = ["PDF Number", "FileType", "Breakpoint", "Access", "Download Source",
               "Sources", "Author", "Title", "Year", "DOI", "DOI URL", "PDF URL", "Abstract"]

    rows = [headers]

    for paper in all_papers:
        status = paper.get("status", "pending")

        # FileType column
        if status == "failed":
            filetype_value = ""  # No file downloaded
        elif status == "Not PDF":
            # Show the actual format (HTML, XML, JSON, Unknown)
            file_format = paper.get("file_format", "Unknown")
            filetype_value = file_format.upper() if file_format != "UNKNOWN" else "Unknown"
        elif status == "unreadable":
            filetype_value = "Non-text/Unknown"
        elif status == "unknown":
            filetype_value = "Unknown"
        elif status in ["html", "xml", "json"]:
            filetype_value = status.upper()  # "HTML", "XML", or "JSON"
        else:
            filetype_value = "PDF"  # Valid PDF

        # Breakpoint column
        if status == "failed":
            breakpoint_value = "Download Failed"
        elif status == "Not PDF":
            breakpoint_value = "Not PDF"
        elif status == "unreadable":
            breakpoint_value = "Not Readable"
        else:
            breakpoint_value = ""  # No error

        # Row color
        if status == "downloaded":
            color = "green"
        elif status == "Not PDF":
            color = "red"  # Not PDF files are red
        elif status in ["html", "xml", "json", "unknown"]:
            color = "orange"
        elif status == "failed" or status == "unreadable":
            color = "red"
        else:
            color = ""  # No color

        # Build row
        row = [
            paper.get("study_number", ""),
            filetype_value,
            breakpoint_value,
            paper.get("access", ""),
            paper.get("download_source", ""),
            paper.get("sources", ""),
            paper.get("author", ""),
            paper.get("title", ""),
            paper.get("year", ""),
            paper.get("doi", ""),
            paper.get("doi_url", ""),
            paper.get("pdf_url", ""),
            paper.get("abstract", "")[:32000],  # Truncate to Excel limit
            color  # Color code at end
        ]

        rows.append(row)

    return rows
