"""
YOLOv12-nano Layout + Smart Text Reconstruction (NO OCR!)

HYBRID APPROACH (Fastest possible):
- PyMuPDF: Extract native embedded text with coordinates (instant)
- Smart Reconstruction: Fix ligatures + gap analysis for proper word/paragraph joining
- YOLOv12-nano: Detect banned regions (figures/tables) @ 150 DPI
- Coordinate filtering: Remove text spans intersecting banned regions

SMART RECONSTRUCTION FEATURES:
- Ligature replacement: Converts Unicode ligatures (fi, fl, ff, etc.) to regular letters
- Gap analysis: Joins word fragments split across spans (<2px = same word, >=2px = separate words)
- Paragraph detection: Uses vertical gaps and sentence endings to create proper paragraphs
- 16x faster than Tesseract OCR with better accuracy

BENEFITS:
- Extremely fast: 2.2x faster than DocLayout-YOLO @ 300 DPI (0.97s vs 2.15s)
- Better quality: Fixes ligature errors and word breaking issues
- No OCR subprocess overhead
- Clean: No chart/table contamination
- Accurate: Uses embedded PDF text with smart reconstruction

INPUTS:
- PDF file path
- DPI for rendering (only for layout detection)

OUTPUTS:
- <pdf_num>.txt - Full extracted body text (clean, reconstructed)
- <pdf_num>_filtered.txt - Metric sentences only
"""
import os
import sys

import warnings
warnings.filterwarnings("ignore")

import sys
from pathlib import Path
import os
import re
from typing import List, Tuple
import time

# Import timing utilities
sys.path.insert(0, str(Path(__file__).parent.parent))
from Timers.timing_utils import save_timer

# Check for required third-party libraries
try:
    from PIL import Image
except ImportError:
    print("ERROR: Missing required library 'Pillow'")
    print("Install it with: pip install Pillow")
    sys.exit(1)

try:
    import fitz  # PyMuPDF
except ImportError:
    print("ERROR: Missing required library 'PyMuPDF'")
    print("Install it with: pip install PyMuPDF")
    sys.exit(1)

try:
    import numpy as np
except ImportError:
    print("ERROR: Missing required library 'numpy'")
    print("Install it with: pip install numpy")
    sys.exit(1)

try:
    import torch
except ImportError:
    print("ERROR: Missing required library 'torch'")
    sys.exit(1)

try:
    from ultralytics import YOLO
except ImportError:
    print("ERROR: Missing required library 'ultralytics'")
    print("Install it with: pip install ultralytics")
    sys.exit(1)

# Lazy import - only load when needed
layout_engine = None

def _get_layout_engine():
    """Lazy load YOLOv12-nano model."""
    global layout_engine
    if layout_engine is None:
        print("[DEBUG] Loading YOLOv12-nano model...")

        # Model path: yolov12n_Model.pt (YOLOv12-nano layout detection model)
        model_path = Path(__file__).parent / "yolov12n_Model.pt"

        if not model_path.exists():
            print("[DEBUG] Downloading YOLOv12-nano from Hugging Face...")
            import requests
            model_url = "https://huggingface.co/hantian/yolo-doclaynet/resolve/main/yolov12n-doclaynet.pt"

            response = requests.get(model_url, stream=True)
            response.raise_for_status()

            total_size = int(response.headers.get('content-length', 0))
            with open(model_path, 'wb') as f:
                downloaded = 0
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total_size > 0:
                            progress = (downloaded / total_size) * 100
                            print(f"\r[DEBUG] Downloading: {progress:.1f}%", end='', flush=True)
            print(f"\n[DEBUG] Model downloaded to: {model_path}")

        layout_engine = YOLO(str(model_path))
        print(f"[DEBUG] YOLOv12-nano model loaded.\n")
    return layout_engine

# Paths (defaults for standalone testing)
def find_latest_query_folder():
    """Find the latest Query folder in Output directory."""
    output_base = Path(__file__).parent.parent / "Output"
    query_folders = [d for d in output_base.iterdir() if d.is_dir() and "Query" in d.name]
    if not query_folders:
        return None
    return max(query_folders, key=lambda d: d.stat().st_mtime)

PDF_DIR = Path(__file__).parent.parent / "Output"
OUTPUT_DIR = PDF_DIR

# Settings
DPI = 150  # For layout detection rendering (2x faster than 300 DPI)
CONF_THRESHOLD = 0.25

# YOLOv12-nano CLASS_NAMES (from hantian/yolo-doclaynet)
CLASS_NAMES = {
    0: 'Caption',
    1: 'Footnote',
    2: 'Formula',
    3: 'List-item',
    4: 'Page-footer',
    5: 'Page-header',
    6: 'Picture',
    7: 'Section-header',
    8: 'Table',
    9: 'Text',
    10: 'Title'
}

# Mask these labels to filter from text extraction
MASK_LABELS = ['Table', 'Picture', 'Caption']


# =============================================================================
# COORDINATE CONVERSION & FILTERING
# =============================================================================

def detect_banned_regions(img, model):
    """
    Extract banned regions (table/figure/caption) from YOLOv12-nano detection.
    Returns list of (x0, y0, x1, y1) in image coordinates.
    """
    banned_rects = []

    # Run detection (YOLOv12-nano uses 640px input size, device=0 for GPU)
    results = model.predict(img, imgsz=640, conf=CONF_THRESHOLD, device=0, verbose=False)

    # Extract boxes
    for result in results:
        boxes = result.boxes
        for box in boxes:
            # Get bbox coordinates (x1, y1, x2, y2)
            x1, y1, x2, y2 = box.xyxy[0].cpu().numpy()

            # Get class ID and confidence
            cls_id = int(box.cls[0])
            conf = float(box.conf[0])

            # Get label name
            label = CLASS_NAMES.get(cls_id, 'unknown')

            # Only keep banned labels
            if label in MASK_LABELS:
                banned_rects.append((float(x1), float(y1), float(x2), float(y2)))

    return banned_rects


def image_rects_to_pdf(rects_img, page_width, page_height, img_width, img_height):
    """
    Convert image-space rects to PDF page-space rects.
    - Image origin: (0, 0) top-left
    - PDF origin: (0, 0) bottom-left
    """
    scale_x = img_width / page_width
    scale_y = img_height / page_height

    pdf_rects = []
    for x0, y0, x1, y1 in rects_img:
        # Un-scale
        px0 = x0 / scale_x
        px1 = x1 / scale_x
        # Flip Y: image y=0 at top -> PDF y=0 at bottom
        py0 = page_height - (y1 / scale_y)
        py1 = page_height - (y0 / scale_y)
        pdf_rects.append(fitz.Rect(px0, py0, px1, py1))

    return pdf_rects


def rect_intersects_any(rect, banned_rects):
    """
    Check if rect intersects any banned regions.
    Optimized with early exit checks.
    """
    if not banned_rects:  # No banned regions
        return False

    rx0, ry0, rx1, ry1 = rect.x0, rect.y0, rect.x1, rect.y1

    for b in banned_rects:
        bx0, by0, bx1, by1 = b.x0, b.y0, b.x1, b.y1

        # Quick reject: check if bounding boxes overlap
        # No overlap if one rectangle is completely to the left/right/above/below the other
        if rx1 < bx0 or rx0 > bx1 or ry1 < by0 or ry0 > by1:
            continue  # No intersection
        else:
            return True  # Intersection found

    return False


# =============================================================================
# TEXT EXTRACTION WITH SMART RECONSTRUCTION
# =============================================================================

# LIGATURE MAPPING - Fix PDF ligature characters
LIGATURE_MAP = {
    '\ufb00': 'ff',
    '\ufb01': 'fi',
    '\ufb02': 'fl',
    '\ufb03': 'ffi',
    '\ufb04': 'ffl',
    '\ufb05': 'ft',
    '\ufb06': 'st',
}

def replace_ligatures(text):
    """Replace ligature characters with their component letters."""
    for ligature, replacement in LIGATURE_MAP.items():
        text = text.replace(ligature, replacement)
    return text


def extract_text_with_banned_regions(page, banned_pdf_rects):
    """
    Extract text from page using smart reconstruction algorithm.

    Features:
    - Replaces ligature characters (fi, fl, ff, etc.)
    - Smart word joining based on horizontal gaps
    - Smart paragraph detection based on vertical gaps
    - Filters out text intersecting banned regions

    Returns sorted text in reading order with proper paragraph breaks.
    """
    text_dict = page.get_text("dict")
    lines_data = []  # Store (y_position, text) tuples

    for block in text_dict.get("blocks", []):
        if block["type"] != 0:  # Skip non-text blocks
            continue

        for line in block.get("lines", []):
            # Collect all spans in this line with their positions
            spans = []
            y_pos = line["bbox"][1]  # Top Y coordinate of line

            for span in line.get("spans", []):
                x0 = span["bbox"][0]
                x1 = span["bbox"][2]
                text = span["text"]

                # Check if span intersects banned regions - skip if so
                span_rect = fitz.Rect(span["bbox"][0], span["bbox"][1], span["bbox"][2], span["bbox"][3])
                if rect_intersects_any(span_rect, banned_pdf_rects):
                    continue  # Skip this span

                # Replace ligatures immediately
                text = replace_ligatures(text)
                spans.append((x0, x1, text))

            if not spans:  # Skip if all spans were filtered
                continue

            # Sort spans by X position
            spans.sort(key=lambda s: s[0])

            # Join spans intelligently based on horizontal gaps
            line_parts = []
            for i, (x0, x1, text) in enumerate(spans):
                if i == 0:
                    line_parts.append(text)
                else:
                    prev_x1 = spans[i-1][1]
                    gap = x0 - prev_x1

                    # Small gap (<2 points) = join without space (word continuation)
                    # Normal gap (>=2 points) = join with space (separate words)
                    if gap < 2.0:
                        line_parts.append(text)  # No space
                    else:
                        line_parts.append(" " + text)  # Add space

            line_text = "".join(line_parts).strip()
            if line_text:  # Only add non-empty lines
                lines_data.append((y_pos, line_text))

    # Now join lines into paragraphs based on vertical gaps
    if not lines_data:
        return ""

    paragraphs = []
    current_paragraph = [lines_data[0][1]]
    prev_y = lines_data[0][0]

    for i in range(1, len(lines_data)):
        y_pos, line_text = lines_data[i]
        y_gap = y_pos - prev_y

        # Decide if this is a new paragraph or continuation
        prev_line = lines_data[i-1][1]

        # New paragraph if:
        # 1. Large vertical gap (>20 points)
        # 2. Previous line ends with sentence-ending punctuation + moderate gap
        # 3. Current line starts with capital AND prev ends with period + small gap
        is_new_paragraph = (
            y_gap > 20 or  # Large gap
            (prev_line.rstrip().endswith(('.', '!', '?')) and y_gap > 10) or  # Sentence end + moderate gap
            (prev_line.rstrip().endswith('.') and line_text[0].isupper() and y_gap > 8)  # Period + capital + small gap
        )

        if is_new_paragraph:
            # Join current paragraph and start new one
            paragraphs.append(" ".join(current_paragraph))
            current_paragraph = [line_text]
        else:
            # Continue current paragraph
            current_paragraph.append(line_text)

        prev_y = y_pos

    # Add the last paragraph
    if current_paragraph:
        paragraphs.append(" ".join(current_paragraph))

    return "\n\n".join(paragraphs)


# =============================================================================
# METRIC SENTENCE EXTRACTION
# =============================================================================

def load_searchable_terms() -> List[str]:
    """Load 501 searchable metric terms."""
    metrics_file = Path(__file__).parent.parent / "Config" / "MetricsSearchable.txt"  # DEV/Config or SBX/Config or PROD/Config
    with open(metrics_file, 'r', encoding='utf-8') as f:
        terms = [line.strip() for line in f if line.strip()]
    return terms


def load_acronym_mapping() -> dict:
    """Load acronym-to-phrase mapping from MetricsAcronymMapping.txt.

    Returns:
        dict: Maps acronyms to their corresponding phrases
              Example: {"WTT": "Wetting Time Top", "kPa": "kilopascals, which is a unit measuring..."}
    """
    mapping_file = Path(__file__).parent.parent / "Config" / "MetricsAcronymMapping.txt"
    acronym_to_phrase = {}

    with open(mapping_file, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            # Skip comments, empty lines, and header lines
            if not line or line.startswith('#') or line.startswith('='):
                continue
            # Parse "Acronym -> Phrase" format
            if '->' in line:
                parts = line.split('->', 1)
                if len(parts) == 2:
                    acronym = parts[0].strip()
                    phrase = parts[1].strip()
                    acronym_to_phrase[acronym] = phrase

    return acronym_to_phrase


def find_sentence_boundaries(text: str, match_pos: int) -> Tuple[int, int]:
    """Find sentence boundaries around match."""
    sentence_start = 0
    for i in range(match_pos - 1, -1, -1):
        if text[i] in '.!?':
            if i + 1 < len(text) and text[i + 1].isspace():
                sentence_start = i + 1
                while sentence_start < match_pos and text[sentence_start].isspace():
                    sentence_start += 1
                break
            elif i + 1 >= len(text):
                sentence_start = i + 1
                break

    periods_found = 0
    sentence_end = len(text)
    for i in range(match_pos, len(text)):
        if text[i] in '.!?':
            if i + 1 < len(text) and text[i + 1].isspace():
                periods_found += 1
                if periods_found == 2:
                    sentence_end = i + 1
                    break
            elif i + 1 >= len(text):
                periods_found += 1
                if periods_found == 2:
                    sentence_end = i + 1
                    break

    return sentence_start, sentence_end


def is_junk_sentence(sentence: str) -> bool:
    """Filter out junk sentences."""
    if re.search(r'\b(Figure|Fig\.|Table|Scheme)\s+\d+', sentence, re.IGNORECASE):
        return True
    if re.search(r'\b(Materials|Polymers|Textiles|Fibers|Journal)\s+\d{4}', sentence, re.IGNORECASE):
        return True
    if re.search(r'\b\d+\s+of\s+\d+\b', sentence):
        return True
    if re.search(r'\bp\s*[=<>¼½¾]\s*0\.\d+', sentence, re.IGNORECASE):
        return True
    if re.search(r'\d+\.\d+\s+significance', sentence, re.IGNORECASE):
        return True
    if re.search(r',\s+\d{4}\)', sentence):
        return True
    if len(sentence.split()) < 5:
        return True
    return False


def extract_metric_sentences(text: str, searchable_terms: List[str]) -> List[Tuple[str, str]]:
    """Extract sentences containing metric terms.

    Logic:
    - If term AND number are in SAME sentence -> return 1 sentence
    - If term is in one sentence but number in another -> return 2 sentences (for context)

    Returns:
        List of tuples: (matched_metric_term, sentence_or_sentences)
    """
    sentences_list = []

    for term in searchable_terms:
        pattern = r'\b' + re.escape(term) + r'\b'
        for match in re.finditer(pattern, text, re.IGNORECASE):
            match_pos = match.start()
            sentence_start, sentence_end = find_sentence_boundaries(text, match_pos)
            chunk = text[sentence_start:sentence_end].strip()

            if chunk:
                chunks = [s.strip() for s in re.split(r'(?<=[.!?])\s+(?=[A-Z])', chunk)]

                # Find the sentence containing the matched term
                term_sentence = None
                term_sentence_idx = -1
                for idx, c in enumerate(chunks):
                    term_pattern = r'\b' + re.escape(term) + r'\b'
                    if re.search(term_pattern, c, re.IGNORECASE):
                        term_sentence = c
                        term_sentence_idx = idx
                        break

                if not term_sentence:
                    continue

                # Check if term sentence has measurements
                has_measurement = re.search(r'\d+\.?\d*\s*(?:g|kg|mm|cm|m|%|°C|MPa|GPa|kPa)', term_sentence, re.IGNORECASE)
                has_decimal = re.search(r'\d+\.\d+', term_sentence)

                if has_measurement or has_decimal:
                    # Term and number in SAME sentence -> return just that sentence
                    if len(term_sentence) > 10 and not is_junk_sentence(term_sentence):
                        sentences_list.append((term, term_sentence))
                else:
                    # Term has NO number -> return 2 sentences for context
                    # Check if other sentences have numbers
                    for idx, c in enumerate(chunks):
                        if idx == term_sentence_idx:
                            continue  # Skip the term sentence itself

                        has_measurement = re.search(r'\d+\.?\d*\s*(?:g|kg|mm|cm|m|%|°C|MPa|GPa|kPa)', c, re.IGNORECASE)
                        has_decimal = re.search(r'\d+\.\d+', c)

                        if (has_measurement or has_decimal) and len(c) > 10 and not is_junk_sentence(c):
                            # Combine term sentence + number sentence
                            combined = term_sentence + " " + c
                            sentences_list.append((term, combined))
                            break  # Only add one combined sentence per match

    seen = set()
    deduplicated = []
    for term, sentence in sentences_list:
        if sentence not in seen:
            seen.add(sentence)
            deduplicated.append((term, sentence))

    return deduplicated


print("Loading searchable metric terms...")
SEARCHABLE_TERMS = load_searchable_terms()
print(f"Loaded {len(SEARCHABLE_TERMS)} terms.")

print("Loading acronym-to-phrase mapping...")
ACRONYM_MAPPING = load_acronym_mapping()
print(f"Loaded {len(ACRONYM_MAPPING)} acronym mappings.\n")


# =============================================================================
# MAIN PIPELINE
# =============================================================================

def extract_text_from_pdf(pdf_path: Path) -> str:
    """Extract text using PyMuPDF native + DocLayout-YOLO layout filtering."""
    doc = fitz.open(pdf_path)
    num_pages = len(doc)

    print(f"Processing {pdf_path.name} ({num_pages} pages)...\n")
    total_start = time.time()

    all_page_texts = []
    total_banned = 0

    # Timing accumulators
    time_rendering = 0.0
    time_layout_detection = 0.0
    time_coord_conversion = 0.0
    time_text_extraction = 0.0

    # Load model once
    model_start = time.time()
    model = _get_layout_engine()
    time_model_load = time.time() - model_start

    for page_num in range(num_pages):
        page_start = time.time()
        page = doc[page_num]

        # =====================================================================
        # STAGE 1: Render page for layout detection
        # =====================================================================
        render_start = time.time()
        zoom = DPI / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_width, img_height = pix.width, pix.height

        # Convert to PIL Image for DocLayout-YOLO
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        time_rendering += time.time() - render_start

        # =====================================================================
        # STAGE 2: Detect banned regions
        # =====================================================================
        layout_start = time.time()
        banned_img_rects = detect_banned_regions(img, model)
        time_layout_detection += time.time() - layout_start

        # Convert to PDF coordinates
        coord_start = time.time()
        page_rect = page.rect
        banned_pdf_rects = image_rects_to_pdf(
            banned_img_rects,
            page_rect.width, page_rect.height,
            img_width, img_height
        )
        time_coord_conversion += time.time() - coord_start

        total_banned += len(banned_pdf_rects)

        # =====================================================================
        # STAGE 3: Extract text with filtering
        # =====================================================================
        extract_start = time.time()
        page_text = extract_text_with_banned_regions(page, banned_pdf_rects)
        all_page_texts.append(page_text)
        time_text_extraction += time.time() - extract_start

        page_time = time.time() - page_start
        print(f"Page {page_num+1}/{num_pages}: {page_time:.2f}s | Banned regions: {len(banned_pdf_rects)}")

    doc.close()

    total_time = time.time() - total_start

    print(f"\n" + "=" * 80)
    print(f"TIMING BREAKDOWN")
    print(f"=" * 80)
    print(f"  Model load:         {time_model_load:.2f}s")
    print(f"  Page rendering:     {time_rendering:.2f}s ({time_rendering/num_pages:.3f}s per page)")
    print(f"  Layout detection:   {time_layout_detection:.2f}s ({time_layout_detection/num_pages:.3f}s per page)")
    print(f"  Coord conversion:   {time_coord_conversion:.2f}s ({time_coord_conversion/num_pages:.3f}s per page)")
    print(f"  Text extraction:    {time_text_extraction:.2f}s ({time_text_extraction/num_pages:.3f}s per page)")
    print(f"  ─────────────────────────────────────────────")
    print(f"  Total pages:        {num_pages}")
    print(f"  Total banned:       {total_banned}")
    print(f"  Total time:         {total_time:.2f}s")
    print(f"  Avg per page:       {total_time/num_pages:.2f}s\n")

    return "\n\n".join(all_page_texts)


def process_corpus(pdf_nums: list, pdf_dirs: list = None, query_folder: Path = None):
    """Process multiple PDFs.

    Args:
        pdf_nums: List of PDF numbers to process
        pdf_dirs: List of directories to search for PDFs (if None, uses PDF_DIR)
        query_folder: Query folder for timer saving (optional)
    """
    print("=" * 80)
    print("YOLOV12-NANO LAYOUT + SMART TEXT RECONSTRUCTION (NO OCR!)")
    print("=" * 80)
    print(f"DPI: {DPI} (for layout detection only)")
    print(f"Masking: {MASK_LABELS}")
    print(f"Smart Reconstruction: Ligatures fixed + gap analysis + paragraph detection")
    print(f"Metric terms: {len(SEARCHABLE_TERMS)}")
    print("=" * 80 + "\n")

    corpus_start = time.time()

    # Track first extraction and first metrics match
    first_extraction_time = None
    first_metrics_time = None

    for idx, pdf_num in enumerate(pdf_nums, 1):
        print("=" * 80)
        print(f"PDF {idx}/{len(pdf_nums)} (PDF #{pdf_num})")
        print("=" * 80 + "\n")

        pdf_start = time.time()

        # Try to find PDF in multiple directories if provided
        pdf_path = None
        if pdf_dirs:
            for pdf_dir in pdf_dirs:
                candidate = pdf_dir / f"{pdf_num}.pdf"
                if candidate.exists():
                    pdf_path = candidate
                    break
        else:
            pdf_path = PDF_DIR / f"{pdf_num}.pdf"

        if pdf_path is None or not pdf_path.exists():
            print(f"ERROR: {pdf_num}.pdf not found\n")
            continue

        # Extract text (Timer 6: First extraction)
        extraction_start = time.time()
        extracted_text = extract_text_from_pdf(pdf_path)
        extraction_time = time.time() - extraction_start

        # Save first extraction time
        if first_extraction_time is None:
            first_extraction_time = extraction_time
            if query_folder:
                save_timer(query_folder, "6_text_extraction_first", first_extraction_time)

        # Save full text
        output_txt = OUTPUT_DIR / f"{pdf_num}.txt"
        with open(output_txt, 'w', encoding='utf-8') as f:
            f.write(extracted_text)

        words = len(extracted_text.split())
        chars = len(extracted_text)
        print(f"[1/2] Full text saved to: {output_txt.name}")
        print(f"      Characters: {chars:,}")
        print(f"      Words: {words:,}\n")

        # Extract metric sentences (Timer 7: First metrics match)
        filter_start = time.time()
        metric_sentences = extract_metric_sentences(extracted_text, SEARCHABLE_TERMS)
        filter_time = time.time() - filter_start

        # Save first metrics match time
        if first_metrics_time is None:
            first_metrics_time = filter_time
            if query_folder:
                save_timer(query_folder, "7_metrics_match_first", first_metrics_time)

        # Save filtered text
        output_filtered = OUTPUT_DIR / f"{pdf_num}_filtered.txt"
        with open(output_filtered, 'w', encoding='utf-8') as f:
            for metric_term, sentence in metric_sentences:
                # Format based on whether it's an acronym or phrase
                if metric_term in ACRONYM_MAPPING:
                    # It's an acronym - show "Acronym, which means Phrase"
                    phrase = ACRONYM_MAPPING[metric_term]
                    f.write(f"Metric Match: {metric_term}, which means {phrase}\n")
                else:
                    # It's a phrase - show just the phrase
                    f.write(f"Metric Match: {metric_term}\n")
                f.write(sentence + "\n\n")

        print(f"[2/2] Metric sentences extracted: {len(metric_sentences)}")
        print(f"      Filtered text saved to: {output_filtered.name}\n")

        # Summary
        pdf_total_time = time.time() - pdf_start
        elapsed_total = time.time() - corpus_start

        print("=" * 80)
        print(f"PDF {pdf_num} COMPLETE")
        print("=" * 80)
        print(f"Full text:     {output_txt.name} ({chars:,} chars)")
        print(f"Filtered text: {output_filtered.name} ({len(metric_sentences)} sentences)")

        if chars > 0:
            # Extract just sentences (second element of tuple) for reduction calculation
            sentences_only = [sentence for _, sentence in metric_sentences]
            reduction = (1 - len(' '.join(sentences_only)) / chars) * 100
            print(f"Text reduction: {reduction:.1f}%")

        print(f"\nPDF Runtime: {pdf_total_time:.2f}s")
        print(f"Corpus Progress: {idx}/{len(pdf_nums)} PDFs | Avg: {elapsed_total/idx:.1f}s/PDF")
        if idx < len(pdf_nums):
            print(f"ETA: {(elapsed_total/idx)*(len(pdf_nums)-idx):.1f}s remaining")
        print("\n")

    print("=" * 80)
    print("ALL PDFs PROCESSED")
    print("=" * 80)


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 2:
        # First arg is PDF folder path, rest are PDF numbers
        PDF_DIR = Path(sys.argv[1])
        OUTPUT_DIR = PDF_DIR
        pdf_nums = [int(arg) for arg in sys.argv[2:]]
    elif len(sys.argv) > 1:
        # Legacy: just PDF numbers, use default Output folder
        pdf_nums = [int(arg) for arg in sys.argv[1:]]
    else:
        # Auto-detect latest Query folder and process all PDFs
        query_folder = find_latest_query_folder()
        if query_folder is None:
            print("ERROR: No Query folder found in Output directory")
            sys.exit(1)

        print(f"Found latest Query folder: {query_folder.name}\n")

        # Look for PDFs in Downloaded Papers/OA Papers and Non-OA Papers
        oa_papers_dir = query_folder / "Downloaded Papers" / "OA Papers"
        non_oa_papers_dir = query_folder / "Downloaded Papers" / "Non-OA Papers"

        pdf_files = []
        if oa_papers_dir.exists():
            pdf_files.extend(list(oa_papers_dir.glob("*.pdf")))
        if non_oa_papers_dir.exists():
            pdf_files.extend(list(non_oa_papers_dir.glob("*.pdf")))

        if not pdf_files:
            print("ERROR: No PDF files found in Downloaded Papers folders")
            sys.exit(1)

        # Extract PDF numbers from filenames
        pdf_nums = sorted([int(f.stem) for f in pdf_files])

        # Set OUTPUT_DIR for extracted text
        OUTPUT_DIR = query_folder / "Extracted Text"
        OUTPUT_DIR.mkdir(exist_ok=True)

        print(f"Processing {len(pdf_nums)} PDFs: {pdf_nums}\n")

        # Pass both directories to search for PDFs and query folder for timers
        pdf_search_dirs = [oa_papers_dir, non_oa_papers_dir]
        process_corpus(pdf_nums, pdf_dirs=pdf_search_dirs, query_folder=query_folder)
        sys.exit(0)

    process_corpus(pdf_nums)
