import os
import re
import sys
from pathlib import Path
import fitz  # PyMuPDF
import openpyxl
from openpyxl import Workbook
from PIL import Image
import io

# Nougat OCR for table extraction
# Note: Requires albumentations==1.3.1 for nougat-ocr compatibility
try:
    from nougat import NougatModel
    from nougat.utils.checkpoint import get_checkpoint
    NOUGAT_AVAILABLE = True

    # Monkey-patch nougat's BARTDecoder to handle new transformers cache_position arg
    from nougat.model import BARTDecoder
    _original_prepare = BARTDecoder.prepare_inputs_for_inference
    def _patched_prepare(self, input_ids, encoder_outputs, past=None, past_key_values=None,
                         use_cache=None, attention_mask=None, cache_position=None, **kwargs):
        # Ignore cache_position - nougat doesn't use it
        return _original_prepare(self, input_ids, encoder_outputs, past, past_key_values,
                                 use_cache, attention_mask)
    BARTDecoder.prepare_inputs_for_inference = _patched_prepare
    print("Nougat loaded with transformers compatibility patch")
except Exception as e:
    NOUGAT_AVAILABLE = False
    print(f"Warning: Nougat OCR not available: {e}")

# ================== CONFIGURATION ==================
BASE_DIR = Path(
    r"C:\Users\azt12\OneDrive\Documents\Wrestling Robe\Materials Science - Wickability\Studies for Analysis by LLM AI\Datafiles & Python Scripts"
)
INPUT_EXCEL = BASE_DIR / "Simplified Table Format.xlsx"
OUTPUT_EXCEL = BASE_DIR / "Sample-Level-Metadata-V2.xlsx"
PDF_FOLDER = Path(
    r"C:\Users\azt12\OneDrive\Documents\Wrestling Robe\Materials Science - Wickability\Studies for Analysis by LLM AI"
)

# How many studies to process (set to None for all)
MAX_STUDIES = 10

# ================== HELPER FUNCTIONS ==================

def extract_title_with_formatting(pdf_path: str) -> str:
    """Extract title by finding the largest font size in first page."""
    doc = fitz.open(pdf_path)
    page = doc[0]
    blocks = page.get_text("dict")["blocks"]
    text_items = []
    for block in blocks:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    size = span["size"]
                    if text and len(text) > 3:
                        text_items.append({
                            "text": text,
                            "size": size,
                            "y": span["bbox"][1],
                            "x": span["bbox"][0],
                        })
    if not text_items:
        doc.close()
        return None
    text_items.sort(key=lambda x: (x["y"], x["x"]))
    skip_phrases = [
        "preprint", "not peer reviewed", "copyright", "license", "doi:",
        "arxiv", "biorxiv", "medrxiv", "journal of", "page ", "volume ", "issue ",
    ]
    candidates = []
    for item in text_items[:30]:
        if any(phrase in item["text"].lower() for phrase in skip_phrases):
            continue
        if len(item["text"]) < 10 and item["size"] < 15:
            continue
        candidates.append(item)
    if not candidates:
        doc.close()
        return None
    max_size = max(c["size"] for c in candidates)
    title_parts = []
    last_y = -999
    for item in candidates:
        if item["size"] >= max_size - 2:
            if abs(item["y"] - last_y) > 5:
                if title_parts and abs(item["y"] - last_y) > 20:
                    break
            title_parts.append(item["text"])
            last_y = item["y"]
        elif title_parts:
            break
    doc.close()
    if title_parts:
        full_title = " ".join(title_parts)
        full_title = re.sub(r"\s+", " ", full_title)
        return full_title.strip()
    return None

def extract_first_author_with_formatting(pdf_path: str) -> str:
    """Extract first author by finding text after the title."""
    doc = fitz.open(pdf_path)
    page = doc[0]
    blocks = page.get_text("dict")["blocks"]
    text_items = []
    for block in blocks:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    if text:
                        text_items.append({
                            "text": text,
                            "size": span["size"],
                            "y": span["bbox"][1],
                        })
    if not text_items:
        doc.close()
        return None
    text_items.sort(key=lambda x: x["y"])
    max_size = max(item["size"] for item in text_items[:30])
    title_end_idx = 0
    for i, item in enumerate(text_items[:30]):
        if item["size"] >= max_size - 2:
            title_end_idx = i + 1
    for i in range(title_end_idx, min(title_end_idx + 10, len(text_items))):
        text = text_items[i]["text"]
        if any(word in text.lower() for word in ["university", "school", "department", "@", "http"]):
            continue
        words = text.split()
        for j, word in enumerate(words):
            clean_word = re.sub(r"[^A-Za-z]", "", word)
            if clean_word and clean_word[0].isupper() and len(clean_word) >= 3:
                if j + 1 < len(words):
                    next_word = re.sub(r"[^A-Za-z]", "", words[j + 1])
                    if next_word and next_word[0].isupper() and len(next_word) >= 3:
                        doc.close()
                        return next_word
                doc.close()
                return clean_word
    doc.close()
    return None

def extract_year_from_text(text: str) -> int:
    """Extract publication year."""
    year_pattern = r"\b(19[9]\d|20[0-2]\d)\b"
    first_part = text[:6000]
    keywords_patterns = [
        (r"published[:\s]+.*?(\d{4})", 1),
        (r"received[:\s]+.*?(\d{4})", 1),
        (r"accepted[:\s]+.*?(\d{4})", 1),
        (r"copyright[:\s]+.*?(\d{4})", 1),
        (r"©[:\s]*(\d{4})", 1),
        (r"(\d{4})\s+published", 1),
        (r"online:?\s+.*?(\d{4})", 1),
        (r"\b(\d{4})\s+elsevier", 1),
        (r"\b(\d{4})\s+wiley", 1),
        (r"\b(\d{4})\s+springer", 1),
    ]
    for pattern, group in keywords_patterns:
        matches = re.finditer(pattern, first_part, re.IGNORECASE)
        for match in matches:
            year_str = match.group(group)
            year_int = int(year_str)
            if 1990 <= year_int <= 2025:
                return year_int
    lines = first_part.split("\n")
    for i, line in enumerate(lines[:50]):
        if re.match(r"^\s*\d{4}\s*$", line):
            year_int = int(line.strip())
            if 2000 <= year_int <= 2025:
                return year_int
    all_years = re.findall(year_pattern, first_part)
    if all_years:
        valid_years = [int(y) for y in all_years if 1990 <= int(y) <= 2025]
        if valid_years:
            from collections import Counter
            year_counts = Counter(valid_years)
            multiple_years = [y for y, count in year_counts.items() if count > 1]
            if multiple_years:
                return max(multiple_years)
            return max(valid_years)
    return None

def extract_table_with_nougat(pdf_path: str, table_number: int):
    """
    Use Nougat OCR to extract tables from PDF.
    Uses Python API with PyMuPDF for rendering.
    Returns tuple of (row_count, table_data) if found, None otherwise.
    """
    try:
        print(f"    Extracting with Nougat OCR...")
        import torch
        import torchvision.transforms as T

        # Load model (cached after first load)
        if not hasattr(extract_table_with_nougat, '_model'):
            print(f"    Loading Nougat model...")
            extract_table_with_nougat._model = NougatModel.from_pretrained('facebook/nougat-base')
            if torch.cuda.is_available():
                extract_table_with_nougat._model = extract_table_with_nougat._model.to('cuda')
                print(f"    Model loaded on GPU")
            else:
                print(f"    Model loaded on CPU (slower)")
            extract_table_with_nougat._model.eval()

        model = extract_table_with_nougat._model
        device = next(model.parameters()).device

        doc = fitz.open(pdf_path)
        print(f"    Processing {len(doc)} pages...")

        transform = T.Compose([
            T.ToTensor(),
            T.Normalize(mean=[0.485, 0.456, 0.406], std=[0.229, 0.224, 0.225])
        ])

        predictions = []
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            pix = page.get_pixmap(matrix=fitz.Matrix(150/72, 150/72))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img = img.resize((672, 896), Image.Resampling.LANCZOS)
            image_tensor = transform(img).unsqueeze(0).to(device)

            with torch.no_grad():
                output = model.inference(image_tensors=image_tensor)

            # Handle different output formats from Nougat
            text = None
            if output is not None:
                if isinstance(output, str):
                    text = output
                elif isinstance(output, dict):
                    text = output.get('predictions', output.get('text', output.get('generated_text', '')))
                    if isinstance(text, list) and len(text) > 0:
                        text = text[0]
                elif isinstance(output, (list, tuple)) and len(output) > 0:
                    text = output[0]

            if text and len(str(text)) > 10:
                predictions.append(str(text))
                print(f"    Page {page_idx + 1}: {len(str(text))} chars")

        doc.close()

        if not predictions:
            print("    No text extracted")
            return None

        markdown_text = "\n\n".join(predictions)

        # Find tables
        tables = []

        # LaTeX tabular environments
        latex_pattern = r'\\begin\{tabular\}.*?\\end\{tabular\}'
        latex_tables = re.findall(latex_pattern, markdown_text, re.DOTALL)
        if latex_tables:
            tables.extend([('latex', t) for t in latex_tables])

        # Nougat-style tables
        nougat_table_pattern = r'\\begin\{table\}.*?\\end\{table\}'
        nougat_tables = re.findall(nougat_table_pattern, markdown_text, re.DOTALL)
        if nougat_tables:
            tables.extend([('nougat', t) for t in nougat_tables])

        # Markdown tables
        md_table_pattern = r'\|[^\n]+\|\n\|[-:\s|]+\|\n(?:\|[^\n]+\|\n)+'
        md_tables = re.findall(md_table_pattern, markdown_text)
        if md_tables:
            tables.extend([('markdown', t) for t in md_tables])

        if not tables:
            return None

        if table_number > len(tables):
            print(f"    Table {table_number} not found (found {len(tables)} tables)")
            return None

        table_type, table_content = tables[table_number - 1]
        print(f"    Found table {table_number} (type: {table_type})")

        # Parse table
        table_data = []

        if table_type == 'markdown':
            lines = [line.strip() for line in table_content.split('\n') if line.strip()]
            data_lines = [line for line in lines if not all(c in '-:|' for c in line.replace(' ', ''))]
            for line in data_lines:
                cells = [cell.strip() for cell in line.split('|')]
                cells = [c for c in cells if c]
                if cells:
                    table_data.append(cells)

        elif table_type in ('latex', 'nougat'):
            content_match = re.search(r'\\begin\{tabular\}\{[^}]*\}(.*?)\\end\{tabular\}', table_content, re.DOTALL)
            if content_match:
                content = content_match.group(1)
            else:
                content = table_content
            rows = re.split(r'\\\\|\\hline|\\tabularnewline', content)
            for row in rows:
                row = row.strip()
                if row and not row.startswith('\\'):
                    cells = [cell.strip() for cell in row.split('&')]
                    cells = [c for c in cells if c]
                    if cells:
                        table_data.append(cells)

        if len(table_data) > 0:
            row_count = len(table_data) - 1  # Subtract header
            print(f"    ✓ Extracted {row_count} data rows")
            return (row_count, table_data)

        return None

    except Exception as e:
        import traceback
        print(f"    Nougat ERROR: {e}")
        return None

def extract_table_row_count(pdf_path: str, table_number: int) -> int:
    """Extract row count from a specific table using Nougat OCR."""
    print(f"\n    --- Extracting Table {table_number} ---")

    if NOUGAT_AVAILABLE:
        try:
            nougat_result = extract_table_with_nougat(pdf_path, table_number)
            if nougat_result:
                rows, table_data = nougat_result
                print(f"    Table {table_number}: {rows} data rows")
                # Show first few rows
                for idx, row in enumerate(table_data[:5]):
                    row_str = " | ".join([str(cell)[:30] for cell in row])
                    print(f"      Row {idx}: {row_str}")
                if len(table_data) > 5:
                    print(f"      ... and {len(table_data) - 5} more rows")
                return rows
        except Exception as e:
            print(f"    Nougat ERROR: {e}")
    else:
        print(f"    Nougat not available")

    return None

def extract_sample_count(pdf_path: str, full_text: str) -> int:
    """
    Extract number of samples - TABLE EXTRACTION FIRST, text detection as fallback.
    """
    print("\n    === SAMPLE DETECTION ===")

    sentences = re.split(r'[.!?](?=\s+[A-Z])', full_text)

    # ==================================================================
    # PRIORITY 0: TABLE EXTRACTION (PRIMARY METHOD)
    # ==================================================================
    print("    Looking for tables with sample/fabric/material keywords...")

    group2_words = ["fabric", "fabrics", "material", "materials", "garment", "garments",
                    "sample", "samples", "specimen", "specimens"]

    extracted_tables = set()

    for sentence in sentences:
        sentence_lower = sentence.lower()
        group2_found = [w for w in group2_words if w in sentence_lower]
        table_match = re.search(r"table\s+(\d+)", sentence_lower, re.IGNORECASE)

        if group2_found and table_match:
            table_num = int(table_match.group(1))
            if table_num in extracted_tables:
                continue

            print(f"    Found reference to Table {table_num}")
            result = extract_table_row_count(pdf_path, table_num)
            extracted_tables.add(table_num)

            if result and result > 0:
                print(f"    ✓ DECISION: {result} samples from Table {table_num}")
                return result

    # ==================================================================
    # PRIORITY 1: TEXT DETECTION (FALLBACK)
    # ==================================================================
    print("    No table found - checking text for explicit counts...")

    # Look for patterns like "total of X fabrics" or "X samples were tested"
    patterns = [
        r"total\s+of\s+(\d+)\s+(?:fabrics?|materials?|samples?|garments?)",
        r"(\d+)\s+(?:different\s+)?(?:fabrics?|materials?|samples?)\s+(?:were|was)\s+(?:tested|used|analyzed)",
        r"(\d+)\s+(?:fabrics?|materials?|samples?)\s+(?:tested|used|analyzed)",
    ]

    for pattern in patterns:
        match = re.search(pattern, full_text, re.IGNORECASE)
        if match:
            count = int(match.group(1))
            if 1 <= count <= 100:
                print(f"    ✓ DECISION: {count} samples from text pattern")
                return count

    # ==================================================================
    # FINAL FALLBACK
    # ==================================================================
    print("    No indicators found - defaulting to 1")
    return 1

# ================== MAIN PROCESSING ==================

def process_studies():
    """Process all PDFs and extract metadata."""
    print(f"Loading input: {INPUT_EXCEL}")

    if not INPUT_EXCEL.exists():
        print(f"ERROR: Input file not found -> {INPUT_EXCEL}")
        sys.exit(1)

    wb_input = openpyxl.load_workbook(INPUT_EXCEL)
    ws_input = wb_input.active

    study_ids = []
    for row in ws_input.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0]:
            study_ids.append(str(row[0]).strip())

    print(f"Found {len(study_ids)} studies")

    output_columns = [
        "Study Number",
        "Study title",
        "Year of Publish",
        "Name of first-listed author",
        "Number of Sample Fabrics",
    ]
    output_rows = []

    for study_idx, study_id in enumerate(study_ids):
        study_num_match = re.search(r'(\d+)', study_id)
        study_num = int(study_num_match.group(1)) if study_num_match else study_idx + 1

        # Stop after MAX_STUDIES
        if MAX_STUDIES and study_num > MAX_STUDIES:
            break

        print(f"\n{'=' * 50}")
        print(f"Study {study_id}")
        print(f"{'=' * 50}")

        pdf_path = PDF_FOLDER / f"{study_id}.pdf"
        if not os.path.exists(pdf_path):
            print(f"  PDF not found")
            output_rows.append({
                "Study Number": study_id,
                "Study title": "PDF not found",
                "Year of Publish": "N/A",
                "Name of first-listed author": "N/A",
                "Number of Sample Fabrics": 0,
            })
            continue

        # Extract text
        doc = fitz.open(pdf_path)
        full_text = ""
        for page in doc:
            full_text += page.get_text()
        doc.close()

        # Extract metadata
        title = extract_title_with_formatting(str(pdf_path))
        year = extract_year_from_text(full_text)
        author = extract_first_author_with_formatting(str(pdf_path))
        samples = extract_sample_count(str(pdf_path), full_text)

        print(f"  Title: {title}")
        print(f"  Year: {year}")
        print(f"  Author: {author}")
        print(f"  Samples: {samples}")

        # Set defaults
        if not title:
            title = "Title not extracted"
        if not year:
            year = "Year not found"
        if not author:
            author = "Author not found"

        # Truncate long values
        if len(str(title)) > 100:
            title = str(title)[:100]

        output_rows.append({
            "Study Number": study_id,
            "Study title": title,
            "Year of Publish": year,
            "Name of first-listed author": author,
            "Number of Sample Fabrics": samples,
        })

    # Print summary
    print(f"\n{'=' * 60}")
    print("SUMMARY TABLE")
    print(f"{'=' * 60}")
    print(f"{'Study':<15} {'Samples':>10}")
    print("-" * 60)
    for row in output_rows:
        study = row.get("Study Number", "N/A")
        samples = row.get("Number of Sample Fabrics", "N/A")
        print(f"{study:<15} {samples:>10}")
    print("=" * 60)

    # Save to Excel
    print(f"\nWriting output: {OUTPUT_EXCEL}")
    wb_output = Workbook()
    ws_output = wb_output.active
    ws_output.title = "Metadata"
    ws_output.append(output_columns)
    for row_dict in output_rows:
        row_values = [row_dict.get(col, "") for col in output_columns]
        ws_output.append(row_values)
    wb_output.save(OUTPUT_EXCEL)
    print(f"✓ Saved {len(output_rows)} studies to {OUTPUT_EXCEL}")

if __name__ == "__main__":
    process_studies()
