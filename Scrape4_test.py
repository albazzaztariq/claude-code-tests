import os
import re
import sys
from pathlib import Path
import fitz  # PyMuPDF
# import requests  # Disabled - not using LLM
import openpyxl
from openpyxl import Workbook
# import pdfplumber  # Commented out - using Nougat instead
# import camelot  # Commented out - using Nougat instead
# import pytest  # Not needed for direct execution
# from azure.ai.formrecognizer import DocumentAnalysisClient  # Disabled - using Nougat only
# from azure.core.credentials import AzureKeyCredential  # Disabled - using Nougat only
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
# OLLAMA_URL = "http://localhost:11434/api/generate"  # Disabled
# OLLAMA_MODEL = "gemma2:2b"  # Disabled
# AZURE_VISION_KEY = os.getenv("AZURE_VISION_KEY", "YOUR_AZURE_VISION_KEY_HERE")  # Disabled
# AZURE_VISION_ENDPOINT = os.getenv("AZURE_VISION_ENDPOINT", "YOUR_AZURE_ENDPOINT_HERE")  # Disabled
BASE_DIR = Path(
    r"C:\Users\azt12\OneDrive\Documents\Wrestling Robe\Materials Science - Wickability\Studies for Analysis by LLM AI\Datafiles & Python Scripts"
)
INPUT_EXCEL = BASE_DIR / "Simplified Table Format.xlsx"
OUTPUT_EXCEL = BASE_DIR / "Sample-Level-Metadata.xlsx"
PDF_FOLDER = Path(
    r"C:\Users\azt12\OneDrive\Documents\Wrestling Robe\Materials Science - Wickability\Studies for Analysis by LLM AI"
)

# ================== HELPER FUNCTIONS (not tests) ==================
# call_local_llm - DISABLED (not using LLM)
# def call_local_llm(prompt: str) -> str:
#     ... (commented out)

def extract_title_with_formatting(pdf_path: str) -> str:
    """
    Extract title by finding the largest font size in first page.
    """
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
                        text_items.append(
                            {
                                "text": text,
                                "size": size,
                                "y": span["bbox"][1],
                                "x": span["bbox"][0],
                            }
                        )
    if not text_items:
        doc.close()
        return None
    text_items.sort(key=lambda x: (x["y"], x["x"]))
    skip_phrases = [
        "preprint",
        "not peer reviewed",
        "copyright",
        "license",
        "doi:",
        "arxiv",
        "biorxiv",
        "medrxiv",
        "journal of",
        "page ",
        "volume ",
        "issue ",
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
    """
    Extract first author by finding text after the title.
    """
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
                        text_items.append(
                            {
                                "text": text,
                                "size": span["size"],
                                "y": span["bbox"][1],
                            }
                        )
    if not text_items:
        doc.close()
        return None
    text_items.sort(key=lambda x: x["y"])
    # Find max font size (title)
    max_size = max(item["size"] for item in text_items[:30])
    # Find where title ends
    title_end_idx = 0
    for i, item in enumerate(text_items[:30]):
        if item["size"] >= max_size - 2:
            title_end_idx = i + 1
    # Look for author in next 10 items
    for i in range(title_end_idx, min(title_end_idx + 10, len(text_items))):
        text = text_items[i]["text"]
        # Skip institutional lines
        if any(
            word in text.lower()
            for word in ["university", "school", "department", "@", "http"]
        ):
            continue
        # Look for any capitalized word that's 3+ letters (last name)
        words = text.split()
        for j, word in enumerate(words):
            # Clean the word
            clean_word = re.sub(r"[^A-Za-z]", "", word)
            # If it's capitalized and 3+ letters, it's probably a name
            if clean_word and clean_word[0].isupper() and len(clean_word) >= 3:
                # If next word is also capitalized, this is first name, take next word (last name)
                if j + 1 < len(words):
                    next_word = re.sub(r"[^A-Za-z]", "", words[j + 1])
                    if next_word and next_word[0].isupper() and len(next_word) >= 3:
                        doc.close()
                        return next_word  # Return last name
                # Otherwise just return this word
                doc.close()
                return clean_word
    doc.close()
    return None

def extract_year_from_text(text: str) -> int:
    """
    Extract publication year.
    """
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

def extract_table_row_count(pdf_path: str, table_number: int) -> int:
    """
    Extract row count from a specific table number in the PDF.
    Returns the number of data rows (excluding header) if found, None otherwise.
    Uses Nougat OCR for table extraction.
    """
    print(f"\n    --- Extracting Table {table_number} ---")

    best_count = None
    best_method = None

    # Try NOUGAT OCR FIRST (processes whole PDF, finds tables from markdown)
    if NOUGAT_AVAILABLE:
        print(f"\n    Trying NOUGAT OCR (primary method)...")
        try:
            nougat_result = extract_table_with_nougat(pdf_path, table_number)
            if nougat_result:
                rows, table_data = nougat_result
                print(f"    Nougat: Found Table {table_number} with {rows} data rows")

                # Show FULL table with ACTUAL CELL DATA
                print(f"    Extracted table data (ALL ROWS):")
                for idx, row in enumerate(table_data):
                    row_str = " | ".join([str(cell)[:50] for cell in row])
                    print(f"      Row {idx}: {row_str}")

                if rows > 0:
                    best_count = rows
                    best_method = "Nougat OCR"
            else:
                print(f"    Nougat: Table {table_number} not found or no rows detected")
        except Exception as e:
            print(f"    Nougat ERROR: {e}")
    else:
        print(f"\n    Nougat OCR not available, skipping...")

    # AZURE - DISABLED (using Nougat only)
    # if not best_count:
    #     ... (commented out)

    # CAMELOT - DISABLED (using Nougat instead)
    # if not best_count:
    #     ... (commented out)

    if best_count:
        print(f"\n    ✓ Best result: {best_count} rows from {best_method}")
        return best_count
    else:
        print(f"\n    ✗ Could not extract Table {table_number} from any method")
        return None

# extract_table_with_azure - DISABLED (using Nougat only)
# def extract_table_with_azure(pdf_path: str, table_number: int):
#     ... (entire function commented out)

def extract_table_with_nougat(pdf_path: str, table_number: int):
    """
    Use Nougat OCR to extract tables from PDF.
    Uses Python API with PyMuPDF for rendering (bypasses pypdfium2 issues).
    Returns tuple of (row_count, table_data) if found, None otherwise.
    """
    try:
        print(f"    Extracting with Nougat OCR (Python API with GPU)...")
        import torch
        from nougat.utils.dataset import NougatDataset
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

        # Use PyMuPDF to render pages (more reliable than pypdfium2)
        doc = fitz.open(pdf_path)
        print(f"    Processing {len(doc)} pages with PyMuPDF...")

        # Nougat expects 896x672 images normalized
        transform = T.Compose([
            T.ToTensor(),
            T.Normalize(mean=[0.485, 0.456, 0.406], std=[0.229, 0.224, 0.225])
        ])

        predictions = []
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            # Render at 150 DPI then resize to Nougat's expected size
            pix = page.get_pixmap(matrix=fitz.Matrix(150/72, 150/72))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Resize to Nougat's expected input size (896 x 672)
            img = img.resize((672, 896), Image.Resampling.LANCZOS)

            # Transform to tensor
            image_tensor = transform(img).unsqueeze(0).to(device)

            with torch.no_grad():
                output = model.inference(image_tensors=image_tensor)

            # Handle different output formats from Nougat
            text = None
            if output is not None:
                if isinstance(output, str):
                    text = output
                elif isinstance(output, dict):
                    # Try common dict keys
                    text = output.get('predictions', output.get('text', output.get('generated_text', '')))
                    if isinstance(text, list) and len(text) > 0:
                        text = text[0]
                elif isinstance(output, (list, tuple)) and len(output) > 0:
                    text = output[0]
                else:
                    print(f"    DEBUG: Unknown output type: {type(output)}, value: {str(output)[:200]}")

            if text and len(str(text)) > 10:
                predictions.append(str(text))
                print(f"    Page {page_idx + 1}: {len(str(text))} chars extracted")
            else:
                print(f"    Page {page_idx + 1}: no output (got: {type(output)})")

        doc.close()

        if not predictions:
            print("    No text extracted from any page")
            return None

        # Join all page predictions into one document
        markdown_text = "\n\n".join(predictions)

        # DEBUG: Show first 2000 chars of Nougat output
        print(f"    Nougat output preview (first 2000 chars):")
        print(f"    {markdown_text[:2000]}")
        print(f"    ... (total {len(markdown_text)} chars)")

        # Try multiple table patterns - Nougat may output different formats
        tables = []

        # Pattern 1: Standard markdown tables
        md_table_pattern = r'\|[^\n]+\|\n\|[-:\s|]+\|\n(?:\|[^\n]+\|\n)+'
        md_tables = re.findall(md_table_pattern, markdown_text)
        if md_tables:
            print(f"    Found {len(md_tables)} markdown tables")
            tables.extend([('markdown', t) for t in md_tables])

        # Pattern 2: LaTeX tabular environments (common in scientific papers)
        latex_pattern = r'\\begin\{tabular\}.*?\\end\{tabular\}'
        latex_tables = re.findall(latex_pattern, markdown_text, re.DOTALL)
        if latex_tables:
            print(f"    Found {len(latex_tables)} LaTeX tables")
            tables.extend([('latex', t) for t in latex_tables])

        # Pattern 3: Nougat's special table format (may use \\ for rows)
        nougat_table_pattern = r'\\begin\{table\}.*?\\end\{table\}'
        nougat_tables = re.findall(nougat_table_pattern, markdown_text, re.DOTALL)
        if nougat_tables:
            print(f"    Found {len(nougat_tables)} Nougat-style tables")
            tables.extend([('nougat', t) for t in nougat_tables])

        if not tables:
            print(f"    ✗ No tables found in Nougat output")
            return None

        if table_number > len(tables):
            print(f"    ✗ Table {table_number} not found (Nougat found {len(tables)} tables)")
            return None

        # Get the requested table
        table_type, table_content = tables[table_number - 1]
        print(f"    Processing table {table_number} (type: {table_type})")
        print(f"    Table content: {table_content[:500]}...")

        # Parse based on table type
        table_data = []

        if table_type == 'markdown':
            # Parse markdown table into rows
            lines = [line.strip() for line in table_content.split('\n') if line.strip()]
            data_lines = [line for line in lines if not all(c in '-:|' for c in line.replace(' ', ''))]
            for line in data_lines:
                cells = [cell.strip() for cell in line.split('|')]
                cells = [c for c in cells if c]
                if cells:
                    table_data.append(cells)

        elif table_type in ('latex', 'nougat'):
            # Parse LaTeX table - rows separated by \\ or \\hline
            # Extract content between begin and end
            content_match = re.search(r'\\begin\{tabular\}\{[^}]*\}(.*?)\\end\{tabular\}', table_content, re.DOTALL)
            if content_match:
                content = content_match.group(1)
            else:
                content = table_content

            # Split by \\ (row separator in LaTeX)
            rows = re.split(r'\\\\|\\hline|\\tabularnewline', content)
            for row in rows:
                row = row.strip()
                if row and not row.startswith('\\'):
                    # Split by & (column separator in LaTeX)
                    cells = [cell.strip() for cell in row.split('&')]
                    cells = [c for c in cells if c]
                    if cells:
                        table_data.append(cells)

        if len(table_data) > 0:
            row_count = len(table_data) - 1  # Subtract header row
            print(f"    ✓ Nougat extracted {row_count} data rows from Table {table_number}")

            # Save debug image of the page containing this table
            try:
                doc = fitz.open(pdf_path)
                # Save first few pages as reference (tables usually in first pages)
                for page_idx in range(min(3, len(doc))):
                    page = doc[page_idx]
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    pdf_dir = Path(pdf_path).parent
                    study_match = re.search(r'(\d+)', Path(pdf_path).stem)
                    study_num = study_match.group(1) if study_match else "unknown"
                    img_path = pdf_dir / f"nougat_study{study_num}_page{page_idx+1}.png"
                    pix.save(str(img_path))
                    print(f"    Saved: {img_path}")
                doc.close()
            except Exception as e:
                print(f"    Could not save debug image: {e}")

            return (row_count, table_data)
        else:
            print("    ✗ No data rows found in table")
            return None

    except Exception as e:
        import traceback
        print(f"    Nougat ERROR: {e}")
        print(f"    Traceback: {traceback.format_exc()}")
        return None

def extract_sample_count_from_table(pdf_path: str, full_text: str) -> int:
    """
    Extract number of samples - TABLE EXTRACTION FIRST, then text detection as fallback.
    """
    print("\n    === SAMPLE DETECTION DIAGNOSTICS ===")
    search_text = full_text
    print(f"    Searching through {len(search_text)} characters of text")

    # Split sentences for later use
    sentences = re.split(r'[.!?](?=\s+[A-Z])', search_text)
    print(f"    Split text into {len(sentences)} sentences")

    # ==================================================================
    # PRIORITY 0: TABLE EXTRACTION (PRIMARY METHOD)
    # ==================================================================
    print("\n    ═══════════════════════════════════════════════")
    print("    PRIORITY 0: TABLE EXTRACTION (PRIMARY METHOD)")
    print("    ═══════════════════════════════════════════════")
    print("    Looking for table references near fabric/sample/material keywords...")

    group2_words = [
        "fabric", "fabrics",
        "material", "materials",
        "garment", "garments",
        "sample", "samples",
        "specimen", "specimens",
    ]

    table_count = None
    extracted_tables = set()

    for i, sentence in enumerate(sentences):
        sentence_lower = sentence.lower()

        # Check if sentence has a Group 2 term AND "table"
        group2_found_here = [word for word in group2_words if word in sentence_lower]
        has_group2 = len(group2_found_here) > 0
        table_match = re.search(r"table\s+(\d+)", sentence_lower, re.IGNORECASE)

        if has_group2 and table_match:
            table_num = int(table_match.group(1))

            # Skip if we've already extracted this table
            if table_num in extracted_tables:
                continue

            print(f"\n    ✓ Found Group 2 term + Table {table_num}")
            print(f"    Sentence: '{sentence.strip()[:200]}...'")
            group2_terms_list = ", ".join(group2_found_here)
            print(f"    Keywords found: {group2_terms_list}")
            print(f"    Extracting Table {table_num} with Nougat...")

            # Extract table with Nougat
            result = extract_table_row_count(pdf_path, table_num)
            extracted_tables.add(table_num)

            if result and result > 0:
                table_count = result
                print(f"\n    ═══════════════════════════════════════════════")
                print(f"    DECISION: TABLE EXTRACTION = {table_count} samples")
                print(f"    Source: Table {table_num} extracted by Nougat OCR")
                print("    ═══════════════════════════════════════════════\n")
                return table_count
            else:
                print(f"    ✗ Could not extract Table {table_num}")

    if not extracted_tables:
        print("    No table references found near fabric/sample keywords")
    else:
        print(f"    Tried {len(extracted_tables)} tables but none had valid data")

    # ==================================================================
    # PRIORITY 1: TEXT DETECTION (FALLBACK)
    # ==================================================================
    print("\n    ═══════════════════════════════════════════════")
    print("    PRIORITY 1: TEXT DETECTION (FALLBACK)")
    print("    ═══════════════════════════════════════════════")
    print("    Looking for explicit count statements in text...")
    print("    GROUP 1: 'total of' + number OR number before Group 2")
    print("    GROUP 2: fabrics/materials/samples/etc.")
    print("    GROUP 3: tested/produced/used/etc.")

    explicit_count = None
    found_best_combination = False


    # Number word to digit mapping - EVERY NUMBER 1-50 (lowercase only)
    # We convert sentences to lowercase before searching, so only lowercase keys needed
    word_to_num = {
        "one": 1,
        "two": 2,
        "three": 3,
        "four": 4,
        "five": 5,
        "six": 6,
        "seven": 7,
        "eight": 8,
        "nine": 9,
        "ten": 10,
        "eleven": 11,
        "twelve": 12,
        "thirteen": 13,
        "fourteen": 14,
        "fifteen": 15,
        "sixteen": 16,
        "seventeen": 17,
        "eighteen": 18,
        "nineteen": 19,
        "twenty": 20,
        "twenty-one": 21,
        "twenty-two": 22,
        "twenty-three": 23,
        "twenty-four": 24,
        "twenty-five": 25,
        "twenty-six": 26,
        "twenty-seven": 27,
        "twenty-eight": 28,
        "twenty-nine": 29,
        "thirty": 30,
        "thirty-one": 31,
        "thirty-two": 32,
        "thirty-three": 33,
        "thirty-four": 34,
        "thirty-five": 35,
        "thirty-six": 36,
        "thirty-seven": 37,
        "thirty-eight": 38,
        "thirty-nine": 39,
        "forty": 40,
        "forty-one": 41,
        "forty-two": 42,
        "forty-three": 43,
        "forty-four": 44,
        "forty-five": 45,
        "forty-six": 46,
        "forty-seven": 47,
        "forty-eight": 48,
        "forty-nine": 49,
        "fifty": 50,
    }
    
    # Roman numerals - EVERY NUMBER 1-50
    roman_to_num = {
        "i": 1,
        "ii": 2,
        "iii": 3,
        "iv": 4,
        "v": 5,
        "vi": 6,
        "vii": 7,
        "viii": 8,
        "ix": 9,
        "x": 10,
        "xi": 11,
        "xii": 12,
        "xiii": 13,
        "xiv": 14,
        "xv": 15,
        "xvi": 16,
        "xvii": 17,
        "xviii": 18,
        "xix": 19,
        "xx": 20,
        "xxi": 21,
        "xxii": 22,
        "xxiii": 23,
        "xxiv": 24,
        "xxv": 25,
        "xxvi": 26,
        "xxvii": 27,
        "xxviii": 28,
        "xxix": 29,
        "xxx": 30,
        "xxxi": 31,
        "xxxii": 32,
        "xxxiii": 33,
        "xxxiv": 34,
        "xxxv": 35,
        "xxxvi": 36,
        "xxxvii": 37,
        "xxxviii": 38,
        "xxxix": 39,
        "xl": 40,
        "xli": 41,
        "xlii": 42,
        "xliii": 43,
        "xliv": 44,
        "xlv": 45,
        "xlvi": 46,
        "xlvii": 47,
        "xlviii": 48,
        "xlix": 49,
        "l": 50,
    }

    # Build word number pattern for regex (used in multiple places)
    word_pattern_check = "|".join(word_to_num.keys())

    # Track if we find the BEST combination (Groups 1+2+3 together)
    found_best_combination = False

    # ===== SUPERSEDING CASE 2: "total of" + NUMBER immediately followed by Group 2 =====
    # This takes absolute priority - if found, use it and stop searching
    print("\n    SUPERSEDING CASE 2: Checking for 'total of' + number + Group 2 term...")

    for i, sentence in enumerate(sentences):
        sentence_lower = sentence.lower()
        
        # Check for "total of" + arabic number + group2
        # (Less likely to match sample IDs due to "total of", but check for consistency)
        total_of_arabic = re.search(
            r"total\s+of\s+(?<![a-zA-Z0-9.])(\d+)(?![a-zA-Z0-9.])\s+(?:different\s+)?(?:types?\s+of\s+)?(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)\b",
            sentence_lower
        )
        if total_of_arabic:
            num = int(total_of_arabic.group(1))
            if 1 <= num <= 100:
                print(f"\n    ✓✓✓ SUPERSEDING CASE FOUND - 'total of' + {num} ✓✓✓")
                print(f"    Sentence {i}: '{sentence.strip()[:200]}...'")
                print("\n    ═══════════════════════════════════════════════")
                print(f"    DECISION: EXPLICIT COUNT = {num}")
                print("    SUPERSEDES ALL OTHER METHODS - RETURNING NOW")
                print("    ═══════════════════════════════════════════════\n")
                return num
        
        # Check for "total of" + word number + group2 (CASE-INSENSITIVE)
        total_of_word = re.search(
            rf"total\s+of\s+({word_pattern_check})\s+(?:different\s+)?(?:types?\s+of\s+)?(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)\b",
            sentence_lower
        )
        if total_of_word:
            word_num = total_of_word.group(1)
            if word_num in word_to_num:
                num = word_to_num[word_num]
                print(f"\n    ✓✓✓ SUPERSEDING CASE FOUND - 'total of' + '{word_num}' → {num} ✓✓✓")
                print(f"    Sentence {i}: '{sentence.strip()[:200]}...'")
                print("\n    ═══════════════════════════════════════════════")
                print(f"    DECISION: EXPLICIT COUNT = {num}")
                print("    SUPERSEDES ALL OTHER METHODS - RETURNING NOW")
                print("    ═══════════════════════════════════════════════\n")
                return num
    
    print("    Result: NO 'total of' + number + Group 2 found")
    
    # ===== MAIN GROUP LOGIC =====
    print("\n    MAIN GROUP LOGIC: Checking all sentences for valid combinations...")

    for i, sentence in enumerate(sentences):
        sentence_lower = sentence.lower()

        # OPTIMIZATION: Only process sentences that contain numbers
        # Check for: arabic numerals (0-9), roman numerals, or word numbers
        has_any_number = (
            re.search(r'\d', sentence_lower) or  # Arabic numerals
            re.search(rf'\b({word_pattern_check})\b', sentence_lower) or  # Word numbers
            re.search(rf'\b({"|".join(roman_to_num.keys())})\b', sentence_lower)  # Roman numerals
        )
        if not has_any_number:
            continue  # Skip sentences without any numbers

        # CHECK GROUP 1: "total of" (1a) OR number immediately before group 2 (1b)
        # These are TWO SEPARATE cases
        has_total_of = "total of" in sentence_lower

        # Check if WHOLE, POSITIVE number IMMEDIATELY precedes group 2 words
        # IMMEDIATELY means the very next word, with NO words in between
        # Pattern: number + optional whitespace + Group 2 word (NO words in between)
        has_number_before_group2_check = False

        # STRICT pattern - number must be followed by Group 2 term with NO intervening words
        # Only whitespace allowed between number and Group 2 term
        # Exclude sample IDs like "S18", "F3" (number must be standalone)
        if re.search(
            r"(?<![a-zA-Z0-9.])(\d+)(?![a-zA-Z0-9.])\s+(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)\b",
            sentence_lower
        ):
            has_number_before_group2_check = True

        # Check for word numbers IMMEDIATELY before Group 2 (CASE-INSENSITIVE)
        if re.search(
            rf"\b({word_pattern_check})\s+(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)\b",
            sentence_lower
        ):
            has_number_before_group2_check = True

        # SPECIAL PATTERN: "X different types of [modifiers] [Group2]"
        # Allows: "nine different types of single jersey weft knitted fabrics"
        # Allows: "four types of tri-layer fabrics" (handles hyphens)
        # Modifiers between "types of" and Group 2 are allowed for this pattern ONLY
        if re.search(
            rf"\b({word_pattern_check})\s+(?:different\s+)?types?\s+of\s+(?:[\w-]+\s+){{0,6}}(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)\b",
            sentence_lower
        ):
            has_number_before_group2_check = True

        # SPECIAL PATTERN: digit + "different types of [modifiers] [Group2]"
        # Allows: "9 different types of single jersey weft knitted fabrics"
        # Allows: "4 types of tri-layer fabrics" (handles hyphens)
        # Modifiers between "types of" and Group 2 are allowed for this pattern ONLY
        # Exclude sample IDs like "S9", "F4" (number must be standalone)
        if re.search(
            r"(?<![a-zA-Z0-9.])(\d+)(?![a-zA-Z0-9.])\s+(?:different\s+)?types?\s+of\s+(?:[\w-]+\s+){{0,6}}(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)\b",
            sentence_lower
        ):
            has_number_before_group2_check = True

        has_group1 = has_total_of or has_number_before_group2_check

        # CHECK GROUP 2: BOTH SINGULAR AND PLURAL - track which terms found
        group2_words = [
            "fabric",
            "fabrics",
            "material",
            "materials",
            "variant",
            "variants",
            "garment",
            "garments",
            "sample",
            "samples",
            "textile",
            "textiles",
            "specimen",
            "specimens",
            "jersey",
            "jerseys",
            "structure",
            "structures",
        ]
        group2_found = [word for word in group2_words if word in sentence_lower]
        has_group2 = len(group2_found) > 0

        # CHECK GROUP 3: action words - track which terms found
        group3_words = [
            "tested",
            "produced",
            "used",
            "analyzed",
            "evaluated",
            "studied",
            "prepared",
            "examined",
            "knit",
            "knitted",
            "woven",
        ]
        # Use word boundaries to avoid false matches (e.g., "used" in "focused")
        group3_found = [word for word in group3_words if re.search(rf'\b{word}\b', sentence_lower)]
        has_group3 = len(group3_found) > 0
        
        # VALID COMBINATION CHECK
        # Group 1 (number immediately before Group 2) is REQUIRED for ALL combinations
        # Group 1 is satisfied by EITHER 1a (total of) OR 1b (number before Group 2)
        valid_combination = False
        combination_type = ""

        if has_group1 and has_group2 and has_group3:
            valid_combination = True
            combination_type = "Group 1 + 2 + 3 (BEST)"
            found_best_combination = True  # Mark that we found the best combination
        elif has_group1 and has_group2:
            valid_combination = True
            combination_type = "Group 1 + 2 (COMMON)"
        elif has_group1 and has_group3:
            valid_combination = True
            combination_type = "Group 1 + 3 (RARE)"
        # REMOVED: Group 2 + 3 without Group 1 - user requires number IMMEDIATELY before Group 2
        
        if not valid_combination:
            continue
        
        print(f"\n     Sentence {i}: VALID COMBINATION - {combination_type}")
        print(f"     Text: '{sentence.strip()[:250]}'")

        # Show what triggered Group 1
        # Group 1 = (1a: "total of") OR (1b: number immediately before Group 2)
        print(f"     GROUP 1 SATISFIED: {has_group1}")
        if has_total_of:
            total_of_pos = sentence_lower.find("total of")
            context = sentence_lower[max(0, total_of_pos-20):min(len(sentence_lower), total_of_pos+80)]
            print(f"       → Via 'total of': '...{context}...'")
        if has_number_before_group2_check:
            print(f"       → Via number IMMEDIATELY before Group 2")
        if not has_group1:
            print(f"       → NOT SATISFIED (need 'total of' OR number before Group 2)")

        group2_terms = ", ".join(group2_found) if group2_found else "none"
        print(f"     GROUP 2 (fabric terms): {has_group2}. Found: {group2_terms}")
        group3_terms = ", ".join(group3_found) if group3_found else "none"
        print(f"     GROUP 3 (action words): {has_group3}. Found: {group3_terms}")
        
        # ===== PRIORITY 1: ARABIC NUMERALS (most common) =====
        # Must be WHOLE, POSITIVE numbers only
        # Number must be IMMEDIATELY before Group 2 term (no words in between)
        # IMPORTANT: Find ALL numbers in sentence, take the highest
        # IMPORTANT: Exclude sample IDs like "S18", "F3", etc. (must be standalone)
        digit_patterns = [
            r"total\s+of\s+(?<![a-zA-Z0-9.])(\d+)(?![a-zA-Z0-9.])\s+(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)",
            r"(?<![a-zA-Z0-9.])(\d+)(?![a-zA-Z0-9.])\s+(?:different\s+)?types?\s+of\s+(?:[\w-]+\s+){0,6}(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)",
            r"(?<![a-zA-Z0-9.])(\d+)(?![a-zA-Z0-9.])\s+(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)",
            r"(?:tested|produced|used|analyzed|evaluated|studied|prepared|examined|knit|knitted|woven)\s+(?<![a-zA-Z0-9.])(\d+)(?![a-zA-Z0-9.])\s+(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)",
        ]

        # Find ALL matching numbers in this sentence (don't break early)
        # IMPORTANT: Only accept if Group 3 term is NEARBY (not at end of sentence)
        for pattern in digit_patterns:
            for match in re.finditer(pattern, sentence_lower):
                num = int(match.group(1))
                if 1 <= num <= 100:
                    # SKIP common section headers (appear in all studies)
                    match_context = sentence_lower[max(0, match.start()-5):min(len(sentence_lower), match.end()+40)]
                    section_headers = [
                        r'\d+\s+materials?\s+and\s+methods?',
                        r'\d+\s+introduction',
                        r'\d+\s+results?',
                        r'\d+\s+discussion',
                        r'\d+\s+conclusion',
                        r'\d+\s+experimental',
                        r'\d+\s+methodology',
                        r'\d+\s+background',
                        r'\d+\s+literature\s+review',
                    ]
                    is_section_header = False
                    for header_pattern in section_headers:
                        if re.search(header_pattern, match_context, re.IGNORECASE):
                            print(f"    ⚠ SKIPPING {num}: Section header detected")
                            is_section_header = True
                            break
                    if is_section_header:
                        continue
                    # Get position of this number match
                    match_pos = match.start()

                    # Find closest Group 3 term position
                    closest_group3_distance = float('inf')
                    for g3_word in group3_words:
                        for g3_match in re.finditer(rf'\b{g3_word}\b', sentence_lower):
                            distance = abs(g3_match.start() - match_pos)
                            closest_group3_distance = min(closest_group3_distance, distance)

                    # Only accept if Group 3 is within 100 characters (same clause/phrase)
                    # This filters out section headers like "2 Materials and Methods"
                    if closest_group3_distance > 100:
                        print(f"    ⚠ SKIPPING {num}: Group 3 term too far away ({closest_group3_distance} chars)")
                        continue

                    if not explicit_count or num > explicit_count:
                        explicit_count = num
                        print("\n    ***************************************************************************")
                        print("    ***************************************************************************")
                        print(f"    ✓ FOUND arabic numeral: {num} (Group 3 within {closest_group3_distance} chars)")
                        print("    ***************************************************************************")
                        print("    ***************************************************************************")
                        print()
                        print(f"     Sentence: '{sentence.strip()[:200]}...'")
        
        # ===== PRIORITY 2: WORD NUMBERS (one, two, three, etc.) - CASE-INSENSITIVE =====
        # Number word must be IMMEDIATELY before Group 2 term (no words in between)
        # ONLY exception: "X different types of [Group2]" or "X types of [Group2]"
        # IMPORTANT: Find ALL numbers in sentence, take the highest
        word_patterns = [
            rf"\b({word_pattern_check})\b\s+(?:different\s+)?types?\s+of\s+(?:[\w-]+\s+){{0,6}}(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)",
            rf"\b({word_pattern_check})\b\s+(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)",
            rf"\b({word_pattern_check})\b\s+(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)\s+of",
        ]

        # Find ALL matching numbers in this sentence (don't break early)
        # IMPORTANT: Only accept if Group 3 term is NEARBY (not at end of sentence)
        for pattern in word_patterns:
            for match in re.finditer(pattern, sentence_lower):
                word_num = match.group(1)
                if word_num in word_to_num:
                    found_count = word_to_num[word_num]

                    # DEBUG: Show exact match context
                    match_start = match.start()
                    match_end = match.end()
                    context_start = max(0, match_start - 20)
                    context_end = min(len(sentence_lower), match_end + 20)
                    match_context_full = sentence_lower[context_start:context_end]
                    print(f"    [DEBUG] Matched '{word_num}' at pos {match_start}: '...{match_context_full}...'")

                    # Get position of this number match
                    match_pos = match.start()

                    # Find closest Group 3 term position
                    closest_group3_distance = float('inf')
                    for g3_word in group3_words:
                        for g3_match in re.finditer(rf'\b{g3_word}\b', sentence_lower):
                            distance = abs(g3_match.start() - match_pos)
                            closest_group3_distance = min(closest_group3_distance, distance)

                    # Only accept if Group 3 is within 100 characters (same clause/phrase)
                    if closest_group3_distance > 100:
                        print(f"    ⚠ SKIPPING '{word_num}' ({found_count}): Group 3 term too far away ({closest_group3_distance} chars)")
                        continue

                    if not explicit_count or found_count > explicit_count:
                        explicit_count = found_count
                        print("\n    ***************************************************************************")
                        print("    ***************************************************************************")
                        print(f"    ✓ FOUND word number: '{word_num}' → {found_count} (Group 3 within {closest_group3_distance} chars)")
                        print("    ***************************************************************************")
                        print("    ***************************************************************************")
                        print()
                        print(f"     Sentence: '{sentence.strip()[:200]}...'")
        
        # ===== PRIORITY 3: ROMAN NUMERALS (rare but possible) =====
        # IMPORTANT: Find ALL numbers in sentence, take the highest
        roman_pattern = "|".join(roman_to_num.keys())
        roman_patterns = [
            rf"(?:^|\s)({roman_pattern})(?:\s+)(?:fabrics?|materials?|samples?|variants?|garments?|textiles?|specimens?|jerseys?|structures?)\b",
        ]

        # Find ALL matching numbers in this sentence (don't break early)
        # IMPORTANT: Only accept if Group 3 term is NEARBY (not at end of sentence)
        for pattern in roman_patterns:
            for match in re.finditer(pattern, sentence_lower):
                roman_num = match.group(1)
                if roman_num in roman_to_num:
                    found_count = roman_to_num[roman_num]

                    # Get position of this number match
                    match_pos = match.start()

                    # Find closest Group 3 term position
                    closest_group3_distance = float('inf')
                    for g3_word in group3_words:
                        for g3_match in re.finditer(rf'\b{g3_word}\b', sentence_lower):
                            distance = abs(g3_match.start() - match_pos)
                            closest_group3_distance = min(closest_group3_distance, distance)

                    # Only accept if Group 3 is within 100 characters (same clause/phrase)
                    if closest_group3_distance > 100:
                        print(f"    ⚠ SKIPPING '{roman_num.upper()}' ({found_count}): Group 3 term too far away ({closest_group3_distance} chars)")
                        continue

                    if not explicit_count or found_count > explicit_count:
                        explicit_count = found_count
                        print("\n    ***************************************************************************")
                        print("    ***************************************************************************")
                        print(f"    ✓ FOUND roman numeral: '{roman_num.upper()}' → {found_count} (Group 3 within {closest_group3_distance} chars)")
                        print("    ***************************************************************************")
                        print("    ***************************************************************************")
                        print()
                        print(f"     Sentence: '{sentence.strip()[:200]}...'")

        # After checking ALL patterns in this sentence, continue to next sentence
        # Don't return early - we want to find the HIGHEST number across ALL sentences
        if found_best_combination and explicit_count:
            print(f"\n    → Found Groups 1+2+3 in this sentence with count = {explicit_count}")
            print(f"    → Continuing to check remaining sentences for higher numbers...")
    
    if explicit_count:
        print("\n    ═══════════════════════════════════════════════")
        print(f"    DECISION: EXPLICIT COUNT = {explicit_count}")
        print("    HIGHEST NUMBER FOUND ACROSS ALL SENTENCES")
        print("    ═══════════════════════════════════════════════\n")
        return explicit_count

    print("\n    Result: NO explicit count found in text")

    # ==================================================================
    # PRIORITY 2: SINGLE SAMPLE
    # ==================================================================
    print("\n    PRIORITY 2: Checking for single-sample study...")
    single_sample_patterns = [
        r"\ba\s+(?:smart|novel|new|single)\s+fabrics?\b",
        r"\bone\s+fabrics?\b",
        r"\bsingle\s+(?:fabrics?|materials?|samples?)\b",
    ]
    
    for pattern in single_sample_patterns:
        if re.search(pattern, search_text, re.IGNORECASE):
            match = re.search(pattern, search_text, re.IGNORECASE)
            context_start = max(0, match.start() - 50)
            context_end = min(len(search_text), match.end() + 50)
            context = search_text[context_start:context_end]
            print(f"    ✓ Found: '{context}'")
            print("\n    ═══════════════════════════════════════════════")
            print("    DECISION: SINGLE SAMPLE STUDY = 1")
            print("    ═══════════════════════════════════════════════\n")
            return 1
    
    print("    Result: NOT a single sample study")

    # ==================================================================
    # FINAL FALLBACK: Default to 1
    # ==================================================================
    print("\n    ═══════════════════════════════════════════════")
    print("    DECISION: No indicators found - defaulting to 1")
    print("    ═══════════════════════════════════════════════\n")
    return 1

# ================== TEST FUNCTIONS ==================
def test_process_pdfs():
    """Test function to process all PDFs and extract metadata."""
    print(f"Loading input Excel: {INPUT_EXCEL}")
    
    if not INPUT_EXCEL.exists():
        print(f"ERROR: Input file not found -> {INPUT_EXCEL}")
        sys.exit(1)
    
    wb_input = openpyxl.load_workbook(INPUT_EXCEL)
    ws_input = wb_input.active
    
    study_ids = []
    for row in ws_input.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0]:
            study_ids.append(str(row[0]).strip())
    
    print(f"Found {len(study_ids)} study IDs in input Excel")
    
    # ================== PROCESS EACH PDF ==================
    output_columns = [
        "Study Number",
        "Study title",
        "Year of Publish",
        "Name of first-listed author",
        "Number of Sample Fabrics",
    ]
    output_rows = []
    
    for study_idx, study_id in enumerate(study_ids):
        # Extract study number from study_id (e.g., "Study1" -> 1)
        study_num_match = re.search(r'(\d+)', study_id)
        study_num = int(study_num_match.group(1)) if study_num_match else study_idx + 1

        # Stop after study 10 and print summary table
        if study_num > 10:
            print(f"\n{'=' * 60}")
            print("STOPPING AFTER STUDY 10 - SUMMARY TABLE")
            print(f"{'=' * 60}")
            print("\n" + "=" * 70)
            print(f"{'Study':<15} {'Samples':>10}")
            print("-" * 70)
            for row in output_rows:
                study = row.get("Study Number", "N/A")
                samples = row.get("Number of Sample Fabrics", "N/A")
                print(f"{study:<15} {samples:>10}")
            print("=" * 70)
            print(f"\nTotal studies processed: {len(output_rows)}")
            print("(Stopping here - first 10 studies finalized)")
            print("=" * 70 + "\n")
            break

        print(f"\n{'=' * 60}")
        print(f"Processing Study {study_id}")
        print(f"{'=' * 60}")
        
        # Find PDF
        pdf_path = PDF_FOLDER / f"{study_id}.pdf"
        if not os.path.exists(pdf_path):
            print(f"  PDF not found -> {pdf_path}")
            output_rows.append(
                {
                    "Study Number": study_id,
                    "Study title": "PDF not found",
                    "Year of Publish": "N/A",
                    "Name of first-listed author": "N/A",
                    "Number of Sample Fabrics": 0,
                }
            )
            continue
        
        # Extract full text
        doc = fitz.open(pdf_path)
        full_text = ""
        for page in doc:
            full_text += page.get_text()
        doc.close()
        print(f"  Extracted {len(full_text)} characters of text")
        
        # --------- EXTRACT TITLE ---------
        meta_title = extract_title_with_formatting(str(pdf_path))
        print(f"  Title: {meta_title}")
        
        # --------- EXTRACT YEAR ---------
        meta_year = extract_year_from_text(full_text)
        print(f"  Year: {meta_year}")
        
        # --------- EXTRACT AUTHOR ---------
        meta_first_last = extract_first_author_with_formatting(str(pdf_path))
        print(f"  Author: {meta_first_last}")
        
        # --------- EXTRACT NUMBER OF SAMPLES (REGEX) ---------
        sample_count = extract_sample_count_from_table(str(pdf_path), full_text)
        print(f"  Samples: {sample_count}")
        
        # --------- TRY LLM FOR MISSING DATA ---------
        missing = []
        if not meta_title:
            missing.append("title")
        if not meta_year:
            missing.append("year")
        if not meta_first_last:
            missing.append("author")
        
        # LLM FALLBACK - DISABLED
        # if missing:
        #     ... (commented out)
        
        # --------- SET DEFAULTS FOR MISSING VALUES ---------
        if not meta_title:
            meta_title = "Title not extracted"
        if not meta_year:
            meta_year = "Year not found"
        if not meta_first_last:
            meta_first_last = "Author not found"
        
        # --------- APPLY 100 CHARACTER LIMIT ---------
        if meta_title and len(str(meta_title)) > 100:
            meta_title = str(meta_title)[:100]
        if meta_first_last and len(str(meta_first_last)) > 100:
            meta_first_last = str(meta_first_last)[:100]
        
        print("\n  FINAL METADATA:")
        print(f"    Title: {meta_title}")
        print(f"    Year: {meta_year}")
        print(f"    Author: {meta_first_last}")
        print(f"    Samples: {sample_count}")
        
        # --------- ADD TO OUTPUT ---------
        output_rows.append(
            {
                "Study Number": study_id,
                "Study title": meta_title,
                "Year of Publish": meta_year,
                "Name of first-listed author": meta_first_last,
                "Number of Sample Fabrics": sample_count,
            }
        )
    
    # ================== WRITE OUTPUT EXCEL ==================
    print(f"\n{'=' * 60}")
    print(f"Writing output to: {OUTPUT_EXCEL}")
    print(f"{'=' * 60}")
    
    wb_output = Workbook()
    ws_output = wb_output.active
    ws_output.title = "Metadata"
    
    # Write header
    ws_output.append(output_columns)
    
    # Write data rows
    for row_dict in output_rows:
        row_values = [row_dict.get(col, "") for col in output_columns]
        ws_output.append(row_values)
    
    wb_output.save(OUTPUT_EXCEL)
    print(f"✓ Output saved: {OUTPUT_EXCEL}")
    print(f"✓ Processed {len(output_rows)} studies")

if __name__ == "__main__":
    # Run the test function directly when script is executed
    test_process_pdfs()
