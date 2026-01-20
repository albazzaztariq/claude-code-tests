"""
PaddleOCR Layout + PyMuPDF Native Text Extraction (NO OCR!)

HYBRID APPROACH (Fastest possible):
- PyMuPDF: Extract native embedded text with coordinates (instant)
- PaddleOCR: Detect banned regions (figures/tables/charts)
- Coordinate filtering: Remove text spans intersecting banned regions

BENEFITS:
- Extremely fast: No OCR subprocess overhead
- Clean: No chart/table contamination
- Accurate: Uses embedded PDF text (no recognition errors)

INPUTS:
- PDF file path
- DPI for rendering (only for layout detection)

OUTPUTS:
- <pdf_num>.txt - Full extracted body text (clean)
- <pdf_num>_filtered.txt - Metric sentences only
"""
import warnings
warnings.filterwarnings("ignore")

import sys
from pathlib import Path
import os
import re
from typing import List, Tuple
import time

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

os.environ['DISABLE_MODEL_SOURCE_CHECK'] = 'True'

try:
    from paddleocr import LayoutDetection
except ImportError:
    print("ERROR: Missing required library 'paddleocr'")
    print("Install it with: pip install paddleocr")
    sys.exit(1)

# Paths
PDF_DIR = Path(r"C:\Users\azt12\OneDrive\Documents\Business\Textile\WrestlingRobe\Programming\FabricETL\GlobalDependencies\ResearchCorpus")
OUTPUT_DIR = PDF_DIR

# Settings
DPI = 300  # For layout detection rendering
MASK_LABELS = ["table", "figure", "reference", "equation"]

print("[DEBUG] Loading PaddleOCR Layout model...")
layout_engine = LayoutDetection()
print(f"[DEBUG] PaddleOCR Layout loaded.\n")

# GPU Check
import paddle
print("[DEBUG] GPU Check:")
print(f"  CUDA Available: {paddle.device.is_compiled_with_cuda()}")
if paddle.device.is_compiled_with_cuda():
    print(f"  GPU Count: {paddle.device.cuda.device_count()}")
    print(f"  GPU Name: {paddle.device.cuda.get_device_name(0)}")
    print(f"  Current Device: {paddle.get_device()}")
print()


# =============================================================================
# COORDINATE CONVERSION & FILTERING
# =============================================================================

def detect_banned_regions(img_array, layout_result):
    """
    Extract banned regions (table/figure/reference/equation) from layout detection.
    Returns list of (x0, y0, x1, y1) in image coordinates.
    """
    boxes = layout_result[0]["boxes"] if (isinstance(layout_result, list) and layout_result) else \
            (layout_result.get("boxes", []) if isinstance(layout_result, dict) else [])

    banned_rects = []
    for box in boxes:
        label = box.get("label", "unknown").lower()
        coord = box.get("coordinate", [])

        if len(coord) >= 4 and label in MASK_LABELS:
            x0, y0, x1, y1 = map(float, coord[:4])
            banned_rects.append((x0, y0, x1, y1))

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
# TEXT EXTRACTION
# =============================================================================

def extract_text_with_banned_regions(page, banned_pdf_rects):
    """
    Extract text from page, filtering out spans that intersect banned regions.
    Returns sorted text in reading order.
    """
    # Extract text blocks/spans with coordinates
    textpage = page.get_text("dict")
    kept_spans = []

    for block in textpage.get("blocks", []):
        if block["type"] != 0:  # 0 = text block
            continue

        for line in block.get("lines", []):
            for span in line.get("spans", []):
                x0, y0, x1, y1 = span["bbox"]
                span_rect = fitz.Rect(x0, y0, x1, y1)

                # Keep span only if it doesn't intersect banned regions
                if not rect_intersects_any(span_rect, banned_pdf_rects):
                    # Store (y, x, text) for sorting
                    kept_spans.append((y0, x0, span["text"]))

    # Sort by reading order (top-to-bottom, then left-to-right)
    kept_spans.sort(key=lambda t: (t[0], t[1]))

    # Join text with spaces (spans on same line get spaces between them)
    text_parts = []
    prev_y = None
    for y, x, text in kept_spans:
        if prev_y is not None and abs(y - prev_y) > 5:  # New line threshold
            text_parts.append("\n")
        elif text_parts:
            text_parts.append(" ")
        text_parts.append(text)
        prev_y = y

    return "".join(text_parts)


# =============================================================================
# METRIC SENTENCE EXTRACTION
# =============================================================================

def load_searchable_terms() -> List[str]:
    """Load 501 searchable metric terms."""
    metrics_file = Path(r"C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe\MetricsLists\MetricsSearchable.txt")
    with open(metrics_file, 'r', encoding='utf-8') as f:
        terms = [line.strip() for line in f if line.strip()]
    return terms


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


def extract_metric_sentences(text: str, searchable_terms: List[str]) -> List[str]:
    """Extract sentences containing metric terms."""
    sentences_list = []

    for term in searchable_terms:
        pattern = r'\b' + re.escape(term) + r'\b'
        for match in re.finditer(pattern, text, re.IGNORECASE):
            match_pos = match.start()
            sentence_start, sentence_end = find_sentence_boundaries(text, match_pos)
            chunk = text[sentence_start:sentence_end].strip()

            if chunk:
                chunks = [s.strip() for s in re.split(r'(?<=[.!?])\s+(?=[A-Z])', chunk)]
                for chunk in chunks:
                    has_measurement = re.search(r'\d+\.?\d*\s*(?:g|kg|mm|cm|m|%|°C|MPa|GPa|kPa)', chunk, re.IGNORECASE)
                    has_decimal = re.search(r'\d+\.\d+', chunk)

                    if chunk and len(chunk) > 10 and (has_measurement or has_decimal) and not is_junk_sentence(chunk):
                        sentences_list.append(chunk)

    seen = set()
    deduplicated = []
    for sentence in sentences_list:
        if sentence not in seen:
            seen.add(sentence)
            deduplicated.append(sentence)

    return deduplicated


print("Loading 501 searchable metric terms...")
SEARCHABLE_TERMS = load_searchable_terms()
print(f"Loaded {len(SEARCHABLE_TERMS)} terms.\n")


# =============================================================================
# MAIN PIPELINE
# =============================================================================

def extract_text_from_pdf(pdf_path: Path) -> str:
    """Extract text using PyMuPDF native + PaddleOCR layout filtering."""
    doc = fitz.open(pdf_path)
    num_pages = len(doc)

    print(f"Processing {pdf_path.name} ({num_pages} pages)...\n")
    total_start = time.time()

    all_page_texts = []
    total_banned = 0

    for page_num in range(num_pages):
        page_start = time.time()
        page = doc[page_num]

        # =====================================================================
        # STAGE 1: Render page for layout detection
        # =====================================================================
        zoom = DPI / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_width, img_height = pix.width, pix.height

        # Convert to numpy array for PaddleOCR
        img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
            pix.height, pix.width, pix.n
        )
        # Convert BGR to RGB if needed
        if pix.n == 3:
            img_array = img_array[:, :, ::-1]

        # =====================================================================
        # STAGE 2: Detect banned regions
        # =====================================================================
        layout_result = layout_engine.predict(img_array)
        banned_img_rects = detect_banned_regions(img_array, layout_result)

        # Convert to PDF coordinates
        page_rect = page.rect
        banned_pdf_rects = image_rects_to_pdf(
            banned_img_rects,
            page_rect.width, page_rect.height,
            img_width, img_height
        )

        total_banned += len(banned_pdf_rects)

        # =====================================================================
        # STAGE 3: Extract text with filtering
        # =====================================================================
        page_text = extract_text_with_banned_regions(page, banned_pdf_rects)
        all_page_texts.append(page_text)

        page_time = time.time() - page_start
        print(f"Page {page_num+1}/{num_pages}: {page_time:.2f}s | Banned regions: {len(banned_pdf_rects)}")

    doc.close()

    total_time = time.time() - total_start

    print(f"\n" + "=" * 80)
    print(f"TIMING SUMMARY")
    print(f"=" * 80)
    print(f"  Total pages: {num_pages}")
    print(f"  Total banned regions: {total_banned}")
    print(f"  Total time: {total_time:.2f}s")
    print(f"  Avg per page: {total_time/num_pages:.2f}s\n")

    return "\n\n".join(all_page_texts)


def process_corpus(pdf_nums: list):
    """Process multiple PDFs."""
    print("=" * 80)
    print("PADDLEOCR LAYOUT + PYMUPDF NATIVE TEXT EXTRACTION (NO OCR!)")
    print("=" * 80)
    print(f"DPI: {DPI} (for layout detection only)")
    print(f"Masking: {MASK_LABELS}")
    print(f"Metric terms: {len(SEARCHABLE_TERMS)}")
    print("=" * 80 + "\n")

    corpus_start = time.time()

    for idx, pdf_num in enumerate(pdf_nums, 1):
        print("=" * 80)
        print(f"PDF {idx}/{len(pdf_nums)} (PDF #{pdf_num})")
        print("=" * 80 + "\n")

        pdf_start = time.time()
        pdf_path = PDF_DIR / f"{pdf_num}.pdf"

        if not pdf_path.exists():
            print(f"ERROR: {pdf_path} not found\n")
            continue

        # Extract text
        extracted_text = extract_text_from_pdf(pdf_path)

        # Save full text
        output_txt = OUTPUT_DIR / f"{pdf_num}.txt"
        with open(output_txt, 'w', encoding='utf-8') as f:
            f.write(extracted_text)

        words = len(extracted_text.split())
        chars = len(extracted_text)
        print(f"[1/2] Full text saved to: {output_txt.name}")
        print(f"      Characters: {chars:,}")
        print(f"      Words: {words:,}\n")

        # Extract metric sentences
        filter_start = time.time()
        metric_sentences = extract_metric_sentences(extracted_text, SEARCHABLE_TERMS)
        filter_time = time.time() - filter_start

        # Save filtered text
        output_filtered = OUTPUT_DIR / f"{pdf_num}_filtered.txt"
        with open(output_filtered, 'w', encoding='utf-8') as f:
            for sentence in metric_sentences:
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
            reduction = (1 - len(' '.join(metric_sentences)) / chars) * 100
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
    if len(sys.argv) > 1:
        pdf_nums = [int(arg) for arg in sys.argv[1:]]
    else:
        pdf_nums = [1]

    process_corpus(pdf_nums)
