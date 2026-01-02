import os
import re
import json
import requests
import pandas as pd
import pdfplumber

# ================== CONFIG ==================

EXCEL_PATH = r"C:\Users\azt12\OneDrive\Documents\Wrestling Robe\Materials Science - Wickability\Studies for Analysis by LLM AI\Simplified Table Format.xlsx"
PDF_FOLDER = r"C:\Users\azt12\OneDrive\Documents\Wrestling Robe\Materials Science - Wickability\Studies for Analysis by LLM AI"
OUTPUT_PATH = os.path.join(os.path.dirname(EXCEL_PATH), "Sample-Level-Metadata.xlsx")

OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "gemma3:1b"

STUDY_ID_COL_INDEX = 0  # column with Study Number


# ================== HELPERS ==================


def call_local_llm(prompt: str) -> str:
    payload = {
        "model": OLLAMA_MODEL,
        "prompt": prompt,
        "stream": False,
    }
    resp = requests.post(OLLAMA_URL, json=payload, timeout=240)
    resp.raise_for_status()
    data = resp.json()
    return data.get("response", "")


def extract_json_object(text: str) -> str:
    """
    From an LLM response, try to extract the first JSON object substring.
    More robust - handles malformed JSON better.
    """
    t = text.strip()
    lines = t.splitlines()

    # Strip markdown fences
    if lines and lines[0].strip().startswith("```"):
        lines = lines[1:]
    if lines and lines[-1].strip().startswith("```"):
        lines = lines[:-1]

    t = "\n".join(lines).strip()

    start = t.find("{")
    end = t.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("No JSON braces found")

    json_str = t[start : end + 1]

    # Try to fix common JSON issues
    # Remove trailing commas before closing braces
    json_str = re.sub(r",(\s*[}\]])", r"\1", json_str)

    return json_str


def extract_title_from_first_pages(full_text: str) -> str:
    """
    Improved title extraction focusing on actual paper structure.
    """
    lines = full_text.split("\n")

    # Find where Abstract starts
    abstract_idx = -1
    for i, line in enumerate(lines[:100]):
        if re.match(r"^\s*abstract\s*$", line.lower().strip()):
            abstract_idx = i
            break

    # If we found Abstract, look in lines before it
    search_end = abstract_idx if abstract_idx > 0 else 50

    potential_titles = []
    title_lines = []

    for i in range(min(search_end, len(lines))):
        line = lines[i].strip()

        # Skip empty lines
        if not line:
            if title_lines:  # End of potential multi-line title
                combined = " ".join(title_lines)
                if 20 < len(combined) < 300:
                    potential_titles.append(combined)
                title_lines = []
            continue

        # Skip obvious non-title lines
        skip_patterns = [
            r"^\d+$",  # Just a number
            r"page\s+\d",
            r"doi:",
            r"http",
            r"www\.",
            r"volume",
            r"issue",
            r"copyright",
            r"©",
            r"\d{4}$",  # Just a year
            r"journal",
            r"received:",
            r"accepted:",
            r"published:",
            r"correspondence",
            r"email",
            r"@",
            r"university",
            r"department",
            r"school\s+of",
            r"college",
            r"institute",
        ]
        if any(re.search(pattern, line.lower()) for pattern in skip_patterns):
            title_lines = []
            continue

        # Skip lines that start with numbers (like "1. Introduction")
        if re.match(r"^\d+[\.\)]\s", line):
            title_lines = []
            continue

        # Skip very short lines
        if len(line) < 15:
            title_lines = []
            continue

        # This could be part of the title
        title_lines.append(line)

    # Check any remaining title_lines
    if title_lines:
        combined = " ".join(title_lines)
        if 20 < len(combined) < 300:
            potential_titles.append(combined)

    # Return the first reasonable title
    if potential_titles:
        # Prefer titles with good capitalization
        for title in potential_titles:
            capital_count = sum(1 for c in title if c.isupper())
            if capital_count >= 5 and capital_count < len(title) * 0.7:
                return title
        return potential_titles[0]

    return None


def extract_year_from_text(text: str) -> int:
    """
    Much more aggressive year extraction.
    """
    # Look for years in reasonable range (1990-2025)
    year_pattern = r"\b(19[9]\d|20[0-2]\d)\b"

    # First 6000 chars should contain publication info
    first_part = text[:6000]

    # Strategy 1: Look near publication-related keywords (most reliable)
    keywords_patterns = [
        (r"published[:\s]+.*?(\d{4})", 1),
        (r"received[:\s]+.*?(\d{4})", 1),
        (r"accepted[:\s]+.*?(\d{4})", 1),
        (r"copyright[:\s]+.*?(\d{4})", 1),
        (r"©[:\s]*(\d{4})", 1),
        (r"(\d{4})\s+published", 1),
        (r"online:?\s+.*?(\d{4})", 1),
    ]

    for pattern, group in keywords_patterns:
        matches = re.finditer(pattern, first_part, re.IGNORECASE)
        for match in matches:
            year_str = match.group(group)
            year_int = int(year_str)
            if 1990 <= year_int <= 2025:
                return year_int

    # Strategy 2: Look for standalone years near the top
    lines = first_part.split("\n")
    for i, line in enumerate(lines[:50]):
        # Look for lines that are mostly just a year
        if re.match(r"^\s*\d{4}\s*$", line):
            year_int = int(line.strip())
            if 2000 <= year_int <= 2025:  # More restrictive for standalone years
                return year_int

    # Strategy 3: Find all years in first part and return most common recent one
    all_years = re.findall(year_pattern, first_part)
    if all_years:
        valid_years = [int(y) for y in all_years if 1990 <= int(y) <= 2025]
        if valid_years:
            # Count occurrences and prefer years that appear multiple times
            from collections import Counter

            year_counts = Counter(valid_years)
            # Get years that appear more than once, otherwise just get most recent
            multiple_years = [y for y, count in year_counts.items() if count > 1]
            if multiple_years:
                return max(multiple_years)
            return max(valid_years)

    return None


def extract_first_author(text: str) -> str:
    """
    Improved author extraction with multiple strategies.
    """
    # Get first 4000 characters where author names appear
    first_part = text[:4000]
    lines = first_part.split("\n")

    # Find where title likely ends (at Abstract)
    title_end = 0
    for i, line in enumerate(lines[:40]):
        if re.match(r"^\s*abstract\s*$", line.lower().strip()):
            title_end = i
            break

    if title_end == 0:
        # Try to find title by looking for longer capitalized lines
        for i, line in enumerate(lines[:20]):
            if len(line.strip()) > 30 and sum(1 for c in line if c.isupper()) > 5:
                title_end = i + 1
                break

    if title_end == 0:
        title_end = 5  # Fallback: assume title is in first 5 lines

    # Look for author names in the lines after title
    for i in range(title_end, min(title_end + 20, len(lines))):
        line = lines[i].strip()

        # Skip empty lines
        if not line:
            continue

        # Skip lines with common non-author content
        skip_patterns = [
            r"school",
            r"department",
            r"university",
            r"college",
            r"institute",
            r"email",
            r"@",
            r"http",
            r"www\.",
            r"abstract",
            r"keywords",
            r"^\d+$",
            r"page\s+\d",
        ]
        if any(re.search(pattern, line, re.IGNORECASE) for pattern in skip_patterns):
            continue

        # Pattern 1: "FirstName LastName" (e.g., "Jiahui Ou", "John Smith")
        # Look for capitalized words
        match = re.search(r"\b([A-Z][a-z]{2,})\s+([A-Z][a-z]{2,})\b", line)
        if match:
            # Return last name (second group)
            return match.group(2)

        # Pattern 2: "LastName, FirstInitial" (e.g., "Smith, J.")
        match = re.search(r"\b([A-Z][a-z]{2,}),\s*[A-Z]\.", line)
        if match:
            return match.group(1)

        # Pattern 3: Name with superscript markers (e.g., "Jiahui Ou^a", "John Smith*")
        match = re.search(r"\b([A-Z][a-z]{2,})\s+([A-Z][a-z]{2,})[\*\^a-z0-9,]+", line)
        if match:
            return match.group(2)

        # Pattern 4: All caps names (e.g., "JOHN SMITH")
        match = re.search(r"\b([A-Z]{2,})\s+([A-Z]{2,})\b", line)
        if match and len(match.group(1)) > 2 and len(match.group(2)) > 2:
            # Convert to title case
            return match.group(2).title()

    return None


# ================== LOAD SOURCE EXCEL ==================

df_src = pd.read_excel(EXCEL_PATH, header=0, engine="openpyxl")
all_headers = list(df_src.columns)
study_id_column = all_headers[STUDY_ID_COL_INDEX]

print("Study ID column:", study_id_column)
print("Number of source rows:", len(df_src))


# ================== PDF NAME HELPER ==================


def pdf_for_row(row_index: int) -> str:
    study_num = df_src.at[row_index, study_id_column]
    try:
        n = int(study_num)
    except Exception:
        n = row_index + 1
    return f"{n}.pdf"


# ================== OUTPUT STRUCTURE ==================

output_columns = [
    "Study Number",
    "Study title",
    "Year of Publish",
    "Name of first-listed author",
    "MATERIALS, STRUCTURES, AND TREATMENTS & WATER AFFINITY CHANGES (LEAVE THIS COLUMN BLANK)",
    "Materials, Structures, and Treatments & Change in Water Affinity (Material Name, Woven or Knit, Treatments Performed and Change to Water Affinity) (For all layers) (List Water Affinity Change if Water Affinity Change Treatment was Applied)",
    "Number of Sample Fabrics",
    "Number of Fabric Layers of each sample",
]

output_rows = []


# ================== MAIN LOOP (FIRST 5 STUDIES ONLY) ==================

n_rows_to_process = min(5, len(df_src))

for row_idx in range(n_rows_to_process):
    study_id = df_src.at[row_idx, study_id_column]
    pdf_name = pdf_for_row(row_idx)
    pdf_path = os.path.join(PDF_FOLDER, pdf_name)

    print(f"\n=== Row {row_idx} (Study {study_id}), PDF: {pdf_name} ===")

    if not os.path.exists(pdf_path):
        print(f"  PDF not found -> {pdf_path}")
        continue

    # read full PDF text
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception as e:
        print("  Error reading PDF:", e)
        continue

    # --------- 1) METADATA - Try rule-based extraction FIRST ---------

    print("  Attempting rule-based extraction...")
    meta_title = extract_title_from_first_pages(full_text)
    meta_year = extract_year_from_text(full_text)
    meta_first_last = extract_first_author(full_text)

    print("  Rule-based results:")
    print(f"    Title: {meta_title}")
    print(f"    Year: {meta_year}")
    print(f"    Author: {meta_first_last}")

    # Track what's still missing
    missing = []
    if not meta_title:
        missing.append("title")
    if not meta_year:
        missing.append("year")
    if not meta_first_last:
        missing.append("author")

    # Try LLM for missing items
    if missing:
        print(f"  Missing: {', '.join(missing)}. Trying LLM...")

        # Use first 4000 chars
        short_text = full_text[:4000]

        # Simpler prompt - ask for each item separately if needed
        if "title" in missing:
            title_prompt = f"What is the main article title in this text? Return only the title, nothing else.\n\nTEXT:\n{short_text[:2000]}"
            try:
                llm_title = call_local_llm(title_prompt).strip()
                # Clean up the response
                llm_title = llm_title.replace('"', "").replace("'", "").strip()
                if (
                    llm_title
                    and len(llm_title) > 15
                    and "title" not in llm_title.lower()[:20]
                ):
                    meta_title = llm_title
                    print(f"    LLM title: {meta_title}")
            except Exception as e:
                print(f"    LLM title error: {e}")

        if "year" in missing:
            year_prompt = f"What year was this paper published? Return only a 4-digit year number.\n\nTEXT:\n{short_text[:2000]}"
            try:
                llm_year = call_local_llm(year_prompt).strip()
                # Extract just the year number
                year_match = re.search(r"\b(19[9]\d|20[0-2]\d)\b", llm_year)
                if year_match:
                    year_int = int(year_match.group(1))
                    if 1990 <= year_int <= 2025:
                        meta_year = year_int
                        print(f"    LLM year: {meta_year}")
            except Exception as e:
                print(f"    LLM year error: {e}")

        if "author" in missing:
            author_prompt = f"Who is the first author of this paper? Return only the last name (family name).\n\nTEXT:\n{short_text[:2000]}"
            try:
                llm_author = call_local_llm(author_prompt).strip()
                # Clean up the response
                llm_author = llm_author.replace('"', "").replace("'", "").strip()
                # Extract just a name (single word, capitalized)
                author_match = re.search(r"\b([A-Z][a-z]{2,})\b", llm_author)
                if author_match:
                    meta_first_last = author_match.group(1)
                    print(f"    LLM author: {meta_first_last}")
            except Exception as e:
                print(f"    LLM author error: {e}")

    # Set final values with defaults
    if not meta_title:
        meta_title = "Title not extracted"
    if not meta_year:
        meta_year = "Year not found"
    if not meta_first_last:
        meta_first_last = "Author not found"

    print("\n  FINAL METADATA:")
    print(f"    Title: {meta_title}")
    print(f"    Year: {meta_year}")
    print(f"    Author: {meta_first_last}")

    # --------- 2) SAMPLES & LAYERS ---------

    sample_prompt = (
        "List all fabric samples tested in this study. "
        "For each sample provide the sample name/label and number of fabric layers.\n"
        f"TEXT:\n{full_text[:8000]}\n\n"
        "Return only a simple list, one per line: SampleName LayerCount"
    )

    samples = []

    try:
        s_resp_raw = call_local_llm(sample_prompt)
        print(f"  Sample response: {s_resp_raw[:200]}")

        # Try to parse as simple text list first (more reliable than JSON for this LLM)
        lines = s_resp_raw.split("\n")
        for line in lines:
            line = line.strip()
            if not line or len(line) < 3:
                continue
            # Look for patterns like "Sample 1: 1 layer" or "S1 1" etc.
            match = re.search(r"([^\d:]+)\s*:?\s*(\d+)", line)
            if match:
                name = match.group(1).strip()
                layers = int(match.group(2))
                if 1 <= layers <= 3:
                    samples.append({"name": name, "layers": layers})

        # If that didn't work, try JSON
        if not samples:
            try:
                s_json = extract_json_object(s_resp_raw)
                s_data = json.loads(s_json)
                samples = s_data.get("samples", []) if isinstance(s_data, dict) else []
            except:
                pass

    except Exception as e:
        print(f"  Sample extraction error: {e}")

    if samples:
        print(f"  Samples found ({len(samples)}):")
        for s in samples:
            print(f"    - {s.get('name', 'Unknown')}: {s.get('layers', '?')} layer(s)")
    else:
        print("  No samples identified")

    # --------- 3) BUILD OUTPUT ROWS ---------

    if samples:
        for sample in samples:
            s_name = sample.get("name", "Unknown")
            n_layers = sample.get("layers", "unknown")

            if isinstance(n_layers, int):
                layers_desc = n_layers
            elif isinstance(n_layers, str) and n_layers.isdigit():
                layers_desc = int(n_layers)
            else:
                layers_desc = "unknown"

            row = {
                "Study Number": study_id,
                "Study title": meta_title,
                "Year of Publish": meta_year,
                "Name of first-listed author": meta_first_last,
                "MATERIALS, STRUCTURES, AND TREATMENTS & WATER AFFINITY CHANGES (LEAVE THIS COLUMN BLANK)": "",
                "Materials, Structures, and Treatments & Change in Water Affinity (Material Name, Woven or Knit, Treatments Performed and Change to Water Affinity) (For all layers) (List Water Affinity Change if Water Affinity Change Treatment was Applied)": "",
                "Number of Sample Fabrics": s_name,
                "Number of Fabric Layers of each sample": layers_desc,
            }
            output_rows.append(row)
    else:
        row = {
            "Study Number": study_id,
            "Study title": meta_title,
            "Year of Publish": meta_year,
            "Name of first-listed author": meta_first_last,
            "MATERIALS, STRUCTURES, AND TREATMENTS & WATER AFFINITY CHANGES (LEAVE THIS COLUMN BLANK)": "",
            "Materials, Structures, and Treatments & Change in Water Affinity (Material Name, Woven or Knit, Treatments Performed and Change to Water Affinity) (For all layers) (List Water Affinity Change if Water Affinity Change Treatment was Applied)": "",
            "Number of Sample Fabrics": "NA",
            "Number of Fabric Layers of each sample": "NA",
        }
        output_rows.append(row)


# ================== WRITE OUTPUT ==================

df_out = pd.DataFrame(output_rows, columns=output_columns)

print(f"\nProcessed rows: {n_rows_to_process}")
print(f"Total output rows: {len(df_out)}")

df_out.to_excel(OUTPUT_PATH, index=False)
print(f"\nWrote: {OUTPUT_PATH}")
