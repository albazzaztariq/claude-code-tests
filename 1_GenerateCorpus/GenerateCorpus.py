"""
GenerateCorpus - Research Paper Collection Module
API search → Deduplicate → OA check → Excel → Mirror Testing → Download

FLOW:
1. Read config file (terms, journals, fields)
2. Build OpenAlex query with AND/OR/NOT logic
3. Call API and paginate through results
4. Deduplicate by DOI
5. Check OA status via Unpaywall (get all mirrors)
6. Save to Excel (2 sheets: OA / No OA)
7. Test all unique mirrors in parallel (50KB sample, 32 concurrent)
8. Download PDFs using fastest mirrors (sorted by measured speed)

MIRROR TESTING:
- Extracts unique mirrors from all papers
- Tests in batches of 50, max 32 concurrent connections
- Downloads 50KB sample from each to measure bandwidth
- Sorts mirrors by actual speed (not ping or static rankings)
- Downloads use fastest available mirror for each paper

Usage:
    python GenerateCorpus.py [config_file.txt]

TEMPORARY: Limited to first 250 PDFs for testing
"""

import sys
import re
import time
import json
import asyncio

# Check for required third-party libraries
try:
    import requests
except ImportError:
    print("ERROR: Missing required library 'requests'")
    print("Install it with: pip install requests")
    sys.exit(1)

try:
    import aiohttp
except ImportError:
    print("ERROR: Missing required library 'aiohttp'")
    print("Install it with: pip install aiohttp")
    sys.exit(1)

try:
    from dotenv import load_dotenv
except ImportError:
    print("ERROR: Missing required library 'python-dotenv'")
    print("Install it with: pip install python-dotenv")
    sys.exit(1)

from pathlib import Path
from datetime import datetime
from urllib.parse import urlencode, urlparse
from urllib.request import urlretrieve
import os

# Import timing utilities
sys.path.insert(0, str(Path(__file__).parent.parent))
from Timers.timing_utils import save_timer

# =============================================================================
# CONFIGURATION
# =============================================================================

# Load .env from FabricETL base directory
load_dotenv(Path(__file__).parent.parent / ".env")

OPENALEX_API_URL = "https://api.openalex.org/works"
OPENALEX_EMAIL = os.getenv("OPENALEX_EMAIL")  # Polite pool access (10 req/sec)
UNPAYWALL_EMAIL = os.getenv("UNPAYWALL_EMAIL")

# Default config file location
BASE_DIR = Path(__file__).parent
DEFAULT_CONFIG = BASE_DIR.parent / "Config" / "query.txt"  # DEV/Config or SBX/Config or PROD/Config
OUTPUT_DIR = Path(__file__).parent.parent / "Output"  # DEV/Output or SBX/Output or PROD/Output

# Journal list import (optional)
try:
    sys.path.insert(0, str(BASE_DIR.parent.parent / "v02" / "v02-2" / "src"))
    from JournalList import TEXTILE_JOURNALS
except ImportError:
    TEXTILE_JOURNALS = {}

# =============================================================================
# CONFIG FILE PARSER
# =============================================================================

def parse_config(filepath: Path) -> dict:
    """Parse config file and return search parameters.

    Returns:
        dict with keys: terms, journals, fields, max_results
    """
    if not filepath.exists():
        raise FileNotFoundError(f"Config file not found: {filepath}")

    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # Parse sections
    config = {
        "terms": "",
        "journals": [],
        "fields": ["title", "abstract"],
        "max_results": None
    }

    current_section = None

    for line in lines:
        line = line.strip()

        # Skip comments and blanks
        if not line or line.startswith("#"):
            continue

        # Section headers
        if line.startswith("[") and line.endswith("]"):
            current_section = line[1:-1].lower()
            continue

        # Parse based on section
        if current_section == "openalex":
            if line.startswith("field:"):
                field_str = line.split(":", 1)[1].strip()
                config["fields"] = [f.strip() for f in field_str.split(",")]
            elif line.startswith("terms:"):
                config["terms"] = line.split(":", 1)[1].strip()

        elif current_section == "journals":
            # Expand acronyms to full names
            journal = line.strip()
            if journal in TEXTILE_JOURNALS.values():
                # Find full name from acronym
                for full_name, acronym in TEXTILE_JOURNALS.items():
                    if acronym == journal:
                        config["journals"].append(full_name)
                        break
            else:
                config["journals"].append(journal)

        elif current_section == "max results":
            try:
                config["max_results"] = int(line.replace(",", ""))
            except ValueError:
                pass

    if not config["terms"]:
        raise ValueError("No search terms found in config file")

    return config

# =============================================================================
# OPENALEX QUERY BUILDER
# =============================================================================

def build_openalex_query(config: dict) -> str:
    """Build OpenAlex API URL from config.

    Handles AND/OR/NOT logic in search terms.

    CRITICAL: OpenAlex requires comma-separated filters for AND logic.
    Example: (woven OR knit) AND moisture-wicking becomes:
      filter=fulltext.search:woven OR knit,fulltext.search:moisture-wicking

    Args:
        config: Dict with terms, journals, fields, max_results

    Returns:
        Full API URL with query parameters
    """
    terms = config["terms"]
    fields = config["fields"]

    # Parse terms: /AND/ splits into separate filter parts (comma-separated)
    # Within each AND group, /// = OR
    and_groups = [g.strip() for g in terms.split("/AND/") if g.strip()]

    filter_parts = []

    for and_group in and_groups:
        # Split by /// for OR terms within this AND group
        or_terms = [t.strip() for t in and_group.split("///") if t.strip()]

        # Format each term (quote multi-word phrases)
        formatted_terms = []
        for term in or_terms:
            term = term.strip()
            if ' ' in term:
                # Multi-word phrase: needs quotes
                formatted_terms.append(f'"{term}"')
            else:
                # Single word or hyphenated: no quotes
                formatted_terms.append(term)

        # Join OR terms with " OR "
        or_query = " OR ".join(formatted_terms)

        # Create filter part for this AND group
        if "fulltext" in fields:
            filter_parts.append(f"fulltext.search:{or_query}")
        elif "title" in fields and "abstract" in fields:
            filter_parts.append(f"title_and_abstract.search:{or_query}")
        elif "title" in fields:
            filter_parts.append(f"title.search:{or_query}")
        else:
            # Fallback to search param (not recommended)
            filter_parts.append(f"search:{or_query}")

    # Add journal filter if specified
    if config["journals"]:
        journal_filter = "|".join(config["journals"])
        filter_parts.append(f"primary_location.source.display_name.search:{journal_filter}")

    # Build params
    params = {
        "per_page": 200,
        "cursor": "*",
        "mailto": OPENALEX_EMAIL
    }

    if filter_parts:
        params["filter"] = ",".join(filter_parts)

    return f"{OPENALEX_API_URL}?{urlencode(params, safe=':,|')}"

# =============================================================================
# OPENALEX API CALLER
# =============================================================================

def call_openalex(url: str, max_results: int = None, query_folder: Path = None) -> list[dict]:
    """Call OpenAlex API and paginate through all results.

    Args:
        url: OpenAlex API URL with cursor=*
        max_results: Maximum papers to retrieve (None = unlimited)
        query_folder: Query folder for timer saving (optional)

    Returns:
        List of paper metadata dicts
    """
    papers = []
    current_url = url

    # Timing variables for pagination
    first_call_time = None
    pagination_times = []

    print(f"\n{'='*60}")
    print("CALLING OPENALEX API")
    print(f"{'='*60}")

    while current_url:
        print(f"  Fetching batch {len(papers)//200 + 1}...", end='', flush=True)

        # Timer 2 & 3: Track first call and pagination times
        call_start = time.time()

        try:
            response = requests.get(current_url, timeout=30)
            response.raise_for_status()
            data = response.json()
        except Exception as e:
            print(f"\n  ERROR: {e}")
            break

        call_time = time.time() - call_start

        # Save first call time (Timer 2)
        if first_call_time is None:
            first_call_time = call_time
            if query_folder:
                save_timer(query_folder, "2_openalex_first_call", first_call_time)
        else:
            # Track pagination times (Timer 3)
            pagination_times.append(call_time)

        results = data.get("results", [])
        print(f" {len(results)} papers")

        # Parse results
        for item in results:
            paper = {
                "title": item.get("title", ""),
                "author": (item.get("authorships", [{}])[0].get("author", {}).get("display_name", "") if item.get("authorships") else ""),
                "year": item.get("publication_year", ""),
                "doi": item.get("doi", "").replace("https://doi.org/", "") if item.get("doi") else "",
                "abstract": (item.get("abstract_inverted_index", None) and "Abstract available" or ""),
                "pdf_url": "",
                "sources": "OpenAlex"
            }

            # Get DOI URL
            if paper["doi"]:
                paper["doi_url"] = f"https://doi.org/{paper['doi']}"
            else:
                paper["doi_url"] = ""

            # Get PDF URL if available
            if item.get("open_access", {}).get("oa_url"):
                paper["pdf_url"] = item["open_access"]["oa_url"]

            papers.append(paper)

            # Check max results
            if max_results and len(papers) >= max_results:
                print(f"  Reached max results limit ({max_results})")
                return papers

        # Get next cursor
        meta = data.get("meta", {})
        next_cursor = meta.get("next_cursor")

        if next_cursor:
            current_url = url.replace("cursor=*", f"cursor={next_cursor}")
            time.sleep(0.1)  # Rate limiting
        else:
            current_url = None

    print(f"\n  Total papers fetched: {len(papers):,}")

    # Save pagination average (Timer 3)
    if pagination_times and query_folder:
        avg_pagination = sum(pagination_times) / len(pagination_times)
        save_timer(query_folder, "3_openalex_pagination_avg", avg_pagination)

    return papers

# =============================================================================
# DEDUPLICATION
# =============================================================================

def deduplicate(papers: list[dict]) -> list[dict]:
    """Deduplicate papers by DOI and title.

    Args:
        papers: List of paper metadata dicts

    Returns:
        Deduplicated list
    """
    print(f"\n{'='*60}")
    print("DEDUPLICATING")
    print(f"{'='*60}")
    print(f"  Before: {len(papers):,}")

    seen_dois = set()
    seen_titles = set()
    unique = []

    for paper in papers:
        doi = paper.get("doi", "")
        title = paper.get("title", None)

        # Handle None titles
        if title:
            title = title.lower().strip()
        else:
            title = ""

        # Skip if DOI or title already seen
        if doi and doi in seen_dois:
            continue
        if title and title in seen_titles:
            continue

        # Add to unique list
        unique.append(paper)

        if doi:
            seen_dois.add(doi)
        if title:
            seen_titles.add(title)

    print(f"  After: {len(unique):,}")

    # Assign study numbers
    for i, paper in enumerate(unique, 1):
        paper["study_number"] = i
        paper["status"] = "pending"

    return unique

# =============================================================================
# DYNAMIC MIRROR SPEED TESTING
# =============================================================================

def extract_domain(url: str) -> str:
    """Extract domain from URL for mirror identification."""
    try:
        parsed = urlparse(url)
        return parsed.netloc.replace("www.", "")
    except:
        return ""

async def test_mirror_speed(session, url: str, semaphore, timeout=2):
    """Test download speed from a mirror by downloading 50KB sample.

    Args:
        session: aiohttp ClientSession
        url: Mirror URL to test
        semaphore: Asyncio semaphore for concurrency limiting
        timeout: Timeout in seconds

    Returns:
        Tuple of (domain, speed_bytes_per_sec)
    """
    domain = extract_domain(url)

    async with semaphore:
        try:
            start = time.time()
            async with session.get(url, timeout=timeout) as resp:
                if resp.status != 200:
                    return domain, 0

                # Download 50KB sample
                data = await resp.content.read(51200)
                elapsed = time.time() - start

                if elapsed > 0:
                    speed = len(data) / elapsed  # bytes/sec
                    return domain, speed
                else:
                    return domain, 0
        except:
            return domain, 0

async def test_mirrors_batched(mirrors: list[str], batch_size=50, concurrency=32):
    """Test all mirrors in batches with concurrency limits.

    Args:
        mirrors: List of mirror URLs to test
        batch_size: Number of mirrors per batch
        concurrency: Max concurrent connections

    Returns:
        Dict mapping domain -> speed (bytes/sec)
    """
    semaphore = asyncio.Semaphore(concurrency)
    speeds = {}

    async with aiohttp.ClientSession(
        connector=aiohttp.TCPConnector(limit=concurrency),
        timeout=aiohttp.ClientTimeout(total=3)
    ) as session:
        # Process in batches
        for i in range(0, len(mirrors), batch_size):
            batch = mirrors[i:i+batch_size]

            # Test batch in parallel (up to concurrency limit)
            tasks = [test_mirror_speed(session, url, semaphore) for url in batch]
            batch_results = await asyncio.gather(*tasks, return_exceptions=True)

            # Collect results
            for result in batch_results:
                if isinstance(result, tuple):
                    domain, speed = result
                    if domain and speed > 0:
                        # Keep best speed if domain tested multiple times
                        speeds[domain] = max(speeds.get(domain, 0), speed)

            # Rate limit courtesy between batches
            if i + batch_size < len(mirrors):
                await asyncio.sleep(1.5)

    return speeds

def rank_mirrors_by_speed(oa_papers: list[dict]) -> dict:
    """Extract unique mirrors from all papers and test speeds.

    Args:
        oa_papers: List of OA papers with oa_locations

    Returns:
        Dict mapping domain -> speed (bytes/sec)
    """
    print(f"\n{'='*60}")
    print("TESTING MIRROR SPEEDS")
    print(f"{'='*60}")

    # Extract all unique mirror URLs
    unique_mirrors = set()
    for paper in oa_papers:
        for loc in paper.get("oa_locations", []):
            url = loc.get("url_for_pdf") or loc.get("url")
            if url:
                unique_mirrors.add(url)

    if not unique_mirrors:
        print("  No mirrors to test")
        return {}

    print(f"  Found {len(unique_mirrors)} unique mirrors across {len(oa_papers)} papers")
    print(f"  Testing in batches of 50 (max 32 concurrent)...")

    start = time.time()

    # Test mirrors asynchronously
    speeds = asyncio.run(test_mirrors_batched(list(unique_mirrors)))

    elapsed = time.time() - start

    print(f"  Tested {len(speeds)} mirrors in {elapsed:.1f}s")

    # Show top 5 fastest
    if speeds:
        top_5 = sorted(speeds.items(), key=lambda x: x[1], reverse=True)[:5]
        print(f"\n  Top 5 fastest mirrors:")
        for domain, speed in top_5:
            mbps = (speed * 8) / (1024 * 1024)  # Convert to Mbps
            print(f"    {domain}: {mbps:.1f} Mbps")

    return speeds

# =============================================================================
# UNPAYWALL OA CHECKER
# =============================================================================

def check_oa_status(papers: list[dict], query_folder: Path = None) -> tuple[list[dict], list[dict]]:
    """Check OA status via Unpaywall and split into OA/non-OA.

    Args:
        papers: List of paper metadata dicts
        query_folder: Query folder for timer saving (optional)

    Returns:
        Tuple of (oa_papers, non_oa_papers)
    """
    print(f"\n{'='*60}")
    print("CHECKING OA STATUS (Unpaywall)")
    print(f"{'='*60}")

    oa_papers = []
    non_oa_papers = []
    first_call_time = None

    for i, paper in enumerate(papers, 1):
        if i % 50 == 0:
            print(f"  Checked {i}/{len(papers)}...", flush=True)

        doi = paper.get("doi", "")

        if not doi:
            non_oa_papers.append(paper)
            continue

        # Check Unpaywall (Timer 4: First call)
        try:
            url = f"https://api.unpaywall.org/v2/{doi}?email={UNPAYWALL_EMAIL}"

            # Track first call time
            if first_call_time is None:
                call_start = time.time()
                response = requests.get(url, timeout=5)
                first_call_time = time.time() - call_start
                if query_folder:
                    save_timer(query_folder, "4_unpaywall_first_call", first_call_time)
            else:
                response = requests.get(url, timeout=5)

            if response.status_code == 200:
                data = response.json()

                # Get ALL OA locations (not sorted yet - will sort by measured speed later)
                oa_locations = data.get("oa_locations", [])

                if oa_locations:
                    # Add OpenAlex URL to locations if not already present
                    openalex_url = paper.get("pdf_url", "")
                    if openalex_url:
                        # Check if OpenAlex URL is already in Unpaywall list
                        unpaywall_urls = [loc.get("url_for_pdf") or loc.get("url") for loc in oa_locations]
                        if openalex_url not in unpaywall_urls:
                            # Add OpenAlex URL as first mirror to try
                            oa_locations.insert(0, {"url_for_pdf": openalex_url, "host_type": "publisher"})

                    # Save ALL locations (unsorted)
                    paper["oa_locations"] = oa_locations

                    # Set pdf_url to first location (will be re-sorted by speed before download)
                    first_loc = oa_locations[0]
                    if first_loc:
                        paper["pdf_url"] = first_loc.get("url_for_pdf") or first_loc.get("url")

                    oa_papers.append(paper)
                else:
                    non_oa_papers.append(paper)
            else:
                non_oa_papers.append(paper)

        except Exception:
            non_oa_papers.append(paper)

        time.sleep(0.02)  # Unpaywall rate limit: 100k/day

    print(f"\n  OA papers: {len(oa_papers):,}")
    print(f"  Non-OA papers: {len(non_oa_papers):,}")

    return oa_papers, non_oa_papers

# =============================================================================
# EXCEL SAVER
# =============================================================================

def save_to_excel(oa_papers: list[dict], non_oa_papers: list[dict], filepath: Path):
    """Save papers to Excel with OA and No OA worksheets.

    Args:
        oa_papers: List of OA paper metadata
        non_oa_papers: List of non-OA paper metadata
        filepath: Path to save Excel file
    """
    print(f"\n{'='*60}")
    print("SAVING TO EXCEL")
    print(f"{'='*60}")

    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill

        wb = Workbook()
        headers = ["PDF Number", "Breakpoint", "Download Source", "Sources", "Author", "Title", "Year", "DOI", "DOI URL", "PDF URL", "Abstract"]

        def write_sheet(ws, papers, sheet_name):
            ws.title = sheet_name

            # Headers
            for col, header in enumerate(headers, 1):
                ws.cell(row=1, column=col, value=header).font = Font(bold=True)

            # Data
            for row, paper in enumerate(papers, 2):
                status = paper.get("status", "pending")

                ws.cell(row=row, column=1, value=paper.get("study_number", ""))
                # Breakpoint: "Download Failed" if failed, "HTML"/"XML"/"JSON" for non-PDF formats, blank if downloaded/pending
                if status == "failed":
                    breakpoint_value = "Download Failed"
                elif status == "html":
                    breakpoint_value = "HTML"
                elif status == "xml":
                    breakpoint_value = "XML"
                elif status == "json":
                    breakpoint_value = "JSON"
                else:
                    breakpoint_value = ""
                ws.cell(row=row, column=2, value=breakpoint_value)
                ws.cell(row=row, column=3, value=paper.get("download_source", ""))  # Which mirror succeeded
                ws.cell(row=row, column=4, value=paper.get("sources", ""))
                ws.cell(row=row, column=5, value=paper.get("author", ""))
                ws.cell(row=row, column=6, value=paper.get("title", ""))
                ws.cell(row=row, column=7, value=paper.get("year", ""))
                ws.cell(row=row, column=8, value=paper.get("doi", ""))
                ws.cell(row=row, column=9, value=paper.get("doi_url", ""))
                ws.cell(row=row, column=10, value=paper.get("pdf_url", ""))
                ws.cell(row=row, column=11, value=paper.get("abstract", "")[:32000])

                # Apply color based on status
                if status in ["downloaded", "html", "xml", "json"]:
                    fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Green (we have the data)
                    for c in range(1, 12):
                        ws.cell(row=row, column=c).fill = fill
                elif status == "failed":
                    fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")  # Red
                    for c in range(1, 12):
                        ws.cell(row=row, column=c).fill = fill

            # Column widths
            ws.column_dimensions['A'].width = 10
            ws.column_dimensions['B'].width = 12
            ws.column_dimensions['C'].width = 20  # Download Source
            ws.column_dimensions['D'].width = 15  # Sources
            ws.column_dimensions['E'].width = 25  # Author
            ws.column_dimensions['F'].width = 50  # Title
            ws.column_dimensions['G'].width = 8   # Year
            ws.column_dimensions['H'].width = 25  # DOI
            ws.column_dimensions['I'].width = 35  # DOI URL
            ws.column_dimensions['J'].width = 40  # PDF URL
            ws.column_dimensions['K'].width = 80  # Abstract

        # Write sheets
        ws1 = wb.active
        write_sheet(ws1, oa_papers, "OA")

        ws2 = wb.create_sheet()
        write_sheet(ws2, non_oa_papers, "No OA")

        # Save
        wb.save(filepath)

        print(f"  Saved: {filepath}")
        print(f"    OA sheet: {len(oa_papers):,} papers")
        print(f"    No OA sheet: {len(non_oa_papers):,} papers")

    except ImportError:
        print("  ERROR: openpyxl not installed")
        print("  Install with: pip install openpyxl")

# =============================================================================
# PDF DOWNLOADER WITH PROGRESS BAR
# =============================================================================

def format_time(seconds):
    """Format seconds into HH:MM:SS."""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    if hours > 0:
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    else:
        return f"{minutes:02d}:{secs:02d}"

def format_size(bytes_val):
    """Format bytes into MB."""
    return f"{bytes_val / 1024 / 1024:.2f} MB"

def download_with_progress(url, filepath, study_num, paper_type, max_retries=3):
    """Download file with progress tracking and retry logic.

    Args:
        url: URL to download from
        filepath: Path to save file
        study_num: Study number for tracking
        paper_type: "OA" or "Non-OA"
        max_retries: Maximum number of download attempts (default: 3)

    Returns:
        tuple: (success: bool, size: int, file_format: str or None, source_name: str or None)
               file_format is "HTML", "XML", "JSON", or None (for PDF/unknown)
               source_name is extracted from URL (e.g., "nature.com", "doaj.org")
    """
    for attempt in range(max_retries):
        try:
            response = requests.get(url, stream=True, timeout=30)
            if response.status_code != 200:
                continue  # Try again

            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            first_chunk = None

            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        if first_chunk is None:
                            first_chunk = chunk
                        f.write(chunk)
                        downloaded += len(chunk)

            # Detect file format (PDF, HTML, XML, JSON)
            file_format = None
            if first_chunk:
                header = first_chunk[:100].lower()
                stripped_header = header.strip()

                # Detect specific formats
                if header.startswith(b'<!doctype html') or header.startswith(b'<html'):
                    file_format = "HTML"
                elif header.startswith(b'<?xml') or header.startswith(b'<article'):
                    file_format = "XML"
                elif stripped_header.startswith(b'{') or stripped_header.startswith(b'['):
                    file_format = "JSON"
                # If none of above, assume PDF (or unknown)

            # Extract source name from URL
            source_name = None
            try:
                from urllib.parse import urlparse
                parsed = urlparse(url)
                source_name = parsed.netloc.replace('www.', '')
            except:
                source_name = "unknown"

            return True, downloaded, file_format, source_name
        except Exception as e:
            if attempt == max_retries - 1:
                return False, 0, None, None
            time.sleep(1)  # Wait 1 second before retry

    return False, 0, None, None


def download_with_multiple_sources(locations, filepath, study_num, paper_type, mirror_speeds):
    """Download file trying multiple OA sources sorted by measured speed.

    Strategy: Sort mirrors by measured speed, try each once, then retry failures.

    Args:
        locations: List of OA location dicts from UnPaywall
        filepath: Path to save file
        study_num: Study number for tracking
        paper_type: "OA" or "Non-OA"
        mirror_speeds: Dict mapping domain -> speed (bytes/sec)

    Returns:
        tuple: (success: bool, size: int, file_format: str or None, source_name: str or None)
    """
    if not locations:
        return False, 0, None, None

    # Sort locations by measured speed (fastest first)
    def get_speed(loc):
        url = loc.get("url_for_pdf") or loc.get("url") or ""
        domain = extract_domain(url)
        return mirror_speeds.get(domain, 0)

    sorted_locations = sorted(locations, key=get_speed, reverse=True)

    # Try each location once (fastest first)
    for loc in sorted_locations:
        url = loc.get("url_for_pdf") or loc.get("url")
        if not url:
            continue

        success, size, file_format, source_name = download_with_progress(url, filepath, study_num, paper_type, max_retries=1)
        if success:
            return success, size, file_format, source_name

    # Retry failed locations once more (fastest first)
    for loc in sorted_locations:
        url = loc.get("url_for_pdf") or loc.get("url")
        if not url:
            continue

        success, size, file_format, source_name = download_with_progress(url, filepath, study_num, paper_type, max_retries=1)
        if success:
            return success, size, file_format, source_name

    return False, 0, None, None

def download_pdfs(oa_papers: list[dict], non_oa_papers: list[dict],
                  oa_folder: Path, non_oa_folder: Path, excel_path: Path,
                  api_start_time: float, mirror_speeds: dict, query_folder: Path = None):
    """Download PDFs to OA and Non-OA folders with progress bar.

    Args:
        oa_papers: List of OA papers with pdf_url
        non_oa_papers: List of non-OA papers with DOI
        oa_folder: Directory for OA PDFs
        non_oa_folder: Directory for non-OA PDFs
        excel_path: Path to Excel file for status updates
        api_start_time: Time when API calls started (for runtime tracking)
        mirror_speeds: Dict mapping domain -> speed (bytes/sec) from testing
        query_folder: Query folder for timer saving (optional)
    """
    print(f"\n{'='*60}")
    print("DOWNLOADING PDFs")
    print(f"{'='*60}\n")

    # Count total downloads needed
    oa_downloadable = [p for p in oa_papers if p.get("pdf_url") and p.get("status") == "pending"]
    non_oa_downloadable = [p for p in non_oa_papers if p.get("doi") and p.get("status") == "pending"]

    total_downloads = len(oa_downloadable) + len(non_oa_downloadable)

    if total_downloads == 0:
        print("  No papers to download.")
        return

    print(f"Total to download: {total_downloads}")
    print(f"  OA: {len(oa_downloadable)}")
    print(f"  Non-OA: {len(non_oa_downloadable)}")
    print()

    # Download tracking
    downloaded_count = 0
    failed_count = 0
    total_bytes = 0
    download_start_time = time.time()

    def update_progress():
        """Update progress bar display."""
        elapsed = time.time() - download_start_time
        api_elapsed = time.time() - api_start_time

        # Calculate progress
        progress = downloaded_count / total_downloads if total_downloads > 0 else 0
        percent = progress * 100

        # Calculate speed
        if elapsed > 0:
            mbps = (total_bytes * 8 / 1024 / 1024) / elapsed  # Megabits per second
            avg_time_per_file = elapsed / downloaded_count if downloaded_count > 0 else 0
            remaining_files = total_downloads - downloaded_count
            eta = avg_time_per_file * remaining_files if downloaded_count > 0 else 0
        else:
            mbps = 0
            eta = 0

        # Progress bar (50 chars) - ASCII for Windows compatibility
        bar_length = 50
        filled = int(bar_length * progress)
        bar = '=' * filled + '-' * (bar_length - filled)

        # Print on same line
        runtime_str = format_time(api_elapsed)
        eta_str = format_time(eta) if eta > 0 else "--:--"

        print(f"\r[{bar}] {percent:5.1f}% | {downloaded_count}/{total_downloads} | "
              f"{mbps:6.2f} Mbps | ETA: {eta_str} | Runtime: {runtime_str}", end='', flush=True)

    # Download OA papers
    for paper in oa_downloadable:
        study_num = paper.get("study_number", "unknown")
        filename = f"{study_num}.pdf"
        filepath = oa_folder / filename

        # Use multi-source download with all OA locations
        locations = paper.get("oa_locations", [])
        if not locations:
            # Fallback to single URL if no locations saved
            pdf_url = paper.get("pdf_url", "")
            if pdf_url:
                locations = [{"url_for_pdf": pdf_url}]

        success, size, file_format, source_name = download_with_multiple_sources(locations, filepath, study_num, "OA", mirror_speeds)

        if success:
            if file_format:
                paper["status"] = file_format.lower()  # "html", "xml", or "json"
            else:
                paper["status"] = "downloaded"  # Valid PDF
            paper["download_source"] = source_name  # Track which mirror succeeded
            downloaded_count += 1
            total_bytes += size
        else:
            paper["status"] = "failed"  # Mark as failed after trying all mirrors
            paper["download_source"] = None
            failed_count += 1

        update_progress()

    # Download non-OA papers (SciHub)
    for paper in non_oa_downloadable:
        doi = paper.get("doi", "")
        study_num = paper.get("study_number", "unknown")
        filename = f"{study_num}.pdf"
        filepath = non_oa_folder / filename

        # Try SciHub mirrors
        scihub_urls = [
            f"https://sci-hub.se/{doi}",
            f"https://sci-hub.st/{doi}",
            f"https://sci-hub.ru/{doi}"
        ]

        success = False
        for scihub_url in scihub_urls:
            result, size, file_format = download_with_progress(scihub_url, filepath, study_num, "Non-OA")
            if result:
                if file_format:
                    paper["status"] = file_format.lower()  # "html", "xml", or "json"
                else:
                    paper["status"] = "downloaded"  # Valid PDF
                downloaded_count += 1
                total_bytes += size
                success = True
                break

        if not success:
            paper["status"] = "failed"  # Mark as failed after trying all mirrors
            failed_count += 1

        update_progress()

    # Final progress update
    print()  # New line after progress bar

    # Summary
    elapsed = time.time() - download_start_time
    api_elapsed = time.time() - api_start_time

    # Timer 5: Total download time
    if query_folder:
        save_timer(query_folder, "5_corpus_download_total", elapsed)

    print(f"\n{'='*60}")
    print("DOWNLOAD SUMMARY")
    print(f"{'='*60}")
    print(f"  Downloaded: {downloaded_count}/{total_downloads}")
    print(f"  Failed: {failed_count}")
    print(f"  Total size: {format_size(total_bytes)}")
    print(f"  Download time: {format_time(elapsed)}")
    print(f"  Total runtime: {format_time(api_elapsed)}")
    if elapsed > 0:
        avg_mbps = (total_bytes * 8 / 1024 / 1024) / elapsed
        print(f"  Average speed: {avg_mbps:.2f} Mbps")
    print(f"{'='*60}")

    # Update Excel with status changes (green for success, red for failed)
    if downloaded_count > 0 or failed_count > 0:
        save_to_excel(oa_papers, non_oa_papers, excel_path)

# =============================================================================
# MAIN ORCHESTRATOR
# =============================================================================

def main():
    """Main orchestrator."""
    print("="*60)
    print("SCHOLARSWEEP v01-1")
    print("="*60)

    # Track API start time
    api_start_time = time.time()

    # Get config file path
    if len(sys.argv) > 1:
        config_path = Path(sys.argv[1])
    else:
        config_path = DEFAULT_CONFIG

    # Parse config
    print(f"\nConfig file: {config_path}")
    config = parse_config(config_path)

    print(f"  Terms: {config['terms'][:80]}...")
    print(f"  Fields: {', '.join(config['fields'])}")
    print(f"  Journals: {len(config['journals'])} specified")
    print(f"  Max results: {config['max_results'] or 'Unlimited'}")

    # TEMPORARY: Limit to 250 for testing
    print(f"\n  *** TEST MODE: Limiting to first 250 PDFs ***")

    # Create output folders
    timestamp = datetime.now().strftime("%m-%d-%y-%H%M")
    query_folder = OUTPUT_DIR / f"{timestamp} Query"
    oa_folder = query_folder / "Downloaded Papers" / "OA Papers"
    non_oa_folder = query_folder / "Downloaded Papers" / "Non-OA Papers"

    oa_folder.mkdir(parents=True, exist_ok=True)
    non_oa_folder.mkdir(parents=True, exist_ok=True)

    excel_path = query_folder / f"{timestamp} Papers.xlsx"

    print(f"\n{'='*60}")
    print("API CALLS BEING BUILT...")
    print(f"{'='*60}")

    # Build query (Timer 1: Query Generation)
    query_start = time.time()
    url = build_openalex_query(config)
    query_time = time.time() - query_start
    save_timer(query_folder, "1_query_generation", query_time)
    print(f"  Query built successfully ({query_time*1000:.0f}ms).\n")

    print(f"{'='*60}")
    print("API CALLS BEING MADE...")
    print(f"{'='*60}")

    # Call API (limit to 250) - pass query_folder for timer saving
    papers = call_openalex(url, max_results=250, query_folder=query_folder)

    if not papers:
        print("\nResults: 0")
        print("\nNo papers found.")
        return

    print(f"\nResults: {len(papers):,} papers retrieved")

    # Deduplicate
    papers = deduplicate(papers)

    # Limit to max_results from config (for testing)
    max_pdfs = config.get('max_results', 250)
    if len(papers) > max_pdfs:
        print(f"\n  Trimming from {len(papers)} to {max_pdfs} papers (from config)...")
        papers = papers[:max_pdfs]
        # Re-assign study numbers
        for i, paper in enumerate(papers, 1):
            paper["study_number"] = i

    # Check OA status (pass query_folder for timer saving)
    oa_papers, non_oa_papers = check_oa_status(papers, query_folder=query_folder)

    # Save to Excel
    save_to_excel(oa_papers, non_oa_papers, excel_path)

    # Test mirror speeds (before downloading)
    mirror_speeds = rank_mirrors_by_speed(oa_papers)

    # Download PDFs (with progress tracking, using measured mirror speeds)
    download_pdfs(oa_papers, non_oa_papers, oa_folder, non_oa_folder, excel_path, api_start_time, mirror_speeds, query_folder=query_folder)

    # Create metadata.xlsx copy for other pipelines to use
    import shutil
    metadata_copy = query_folder / "metadata.xlsx"
    shutil.copy(excel_path, metadata_copy)

    print(f"\n{'='*60}")
    print("COMPLETE")
    print(f"{'='*60}")
    print(f"  Results: {excel_path}")
    print(f"  Metadata: {metadata_copy}")
    print(f"  PDFs: {query_folder / 'Downloaded Papers'}")

if __name__ == "__main__":
    main()
