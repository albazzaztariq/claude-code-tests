"""
ChartOCRTester.py - Compare different layout parsers for academic document analysis

Test cases:
- 1.pdf page 9: 3 small XPS charts + images
- 2.pdf page 4: 2 figures (each with 3 sub-images) + 1 table
- 3.pdf page 4: Complex page with images and 1 real table + 1 graphical table
- 3.pdf page 8: Small wide table + images

Parsers to test:
1. PaddleOCR LayoutDetection (baseline)
2. Surya
3. Marker
4. deepdoctection
5. Unstructured.io
6. GROBID (if available)

Also includes:
- Relevance filter testing (checks if extracted data matches user's 693 metrics)
- DePlot chart-to-table extraction
- PaddleOCR text extraction (handles rotated tables, combo chart+table elements)
"""

import os
import sys
import time
import re
from typing import Dict, List, Any, Tuple

# Disable HuggingFace connectivity checks
os.environ['DISABLE_MODEL_SOURCE_CHECK'] = 'True'
os.environ['HF_HUB_OFFLINE'] = '0'

# Test configuration
TEST_CASES = [
    {"pdf": "1.pdf", "page": 9, "expected": "3 charts, several images, figure_title"},
    {"pdf": "2.pdf", "page": 4, "expected": "2 figure regions, 3 figure_titles, 1 table"},
    {"pdf": "3.pdf", "page": 4, "expected": "images + 1 real table + 1 graphical cam design table"},
    {"pdf": "3.pdf", "page": 8, "expected": "images + 1 small wide table"},
]


# ================== METRICS RELEVANCE FILTER (from ScrapeV2.py) ==================
_METRICS_CACHE = None
_METRICS_FILE = r"C:\Users\azt12\OneDrive\Documents\Wrestling Robe\Materials Science - Wickability\Studies for Analysis by LLM AI\Datafiles & Python Scripts\MetricsFullList.txt"

EXCLUDED_WORDS = {
    "title", "no", "no.", "number", "diagram", "design", "cam",
    "needle", "arrangement", "notation", "kd", "figure", "table",
    "for", "each", "layer", "or", "and", "the", "in", "of", "to",
}


def _extract_word_sequences(text: str, min_words: int = 2) -> set:
    """Extract all consecutive word sequences of length >= min_words."""
    text = re.sub(r'\([^)]*\)', '', text)
    text = re.sub(r'[^\w\s-]', ' ', text)
    text = text.lower().strip()
    words = text.split()
    sequences = set()
    for length in range(min_words, min(len(words) + 1, 5)):
        for i in range(len(words) - length + 1):
            seq = ' '.join(words[i:i + length])
            if len(seq) > 3:
                sequences.add(seq)
    return sequences


def _load_metrics_cache() -> set:
    """Load metrics list and pre-process into 2-word sequences."""
    global _METRICS_CACHE
    if _METRICS_CACHE is not None:
        return _METRICS_CACHE

    sequences = set()
    try:
        with open(_METRICS_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = re.sub(r'^\s*\d+[→.\s]+', '', line).strip()
                if not line or line.startswith('#'):
                    continue
                if 'LEAVE' in line.upper() and 'BLANK' in line.upper():
                    continue
                seqs = _extract_word_sequences(line, min_words=2)
                sequences.update(seqs)

        abbreviations = {
            'wtt', 'wtb', 'wt', 'art', 'arb', 'bar', 'tar',
            'ommc', 'owtc', 'aoti', 'mwrb', 'mwrt', 'ssb', 'sst',
            'rct', 'ret', 'wvtr', 'mvtr', 'wvp', 'qmax', 'gsm', 'denier', 'porosity',
        }
        sequences.update(abbreviations)
        _METRICS_CACHE = sequences
        print(f"    [CACHE] Loaded {len(sequences)} metric sequences")
    except FileNotFoundError:
        print(f"    [WARNING] Metrics file not found: {_METRICS_FILE}")
        _METRICS_CACHE = set()
    return _METRICS_CACHE


def check_table_relevance(headers: List[str], min_matches: int = 1) -> Tuple[bool, List[str], int]:
    """Check if headers contain relevant metrics using cached sequences."""
    metric_sequences = _load_metrics_cache()
    matching_headers = []

    for header in headers:
        header_clean = header.lower().strip().replace("<0x0a>", " ")
        words = re.sub(r'[^\w\s]', ' ', header_clean).split()
        non_excluded_words = [w for w in words if w not in EXCLUDED_WORDS]
        if not non_excluded_words:
            continue

        header_sequences = _extract_word_sequences(header_clean, min_words=2)

        for word in words:
            if word in metric_sequences:
                matching_headers.append(header)
                break
        else:
            if header_sequences & metric_sequences:
                matching_headers.append(header)

    is_relevant = len(matching_headers) >= min_matches
    return (is_relevant, matching_headers, len(matching_headers))


def filter_table_data(table_data: str) -> Tuple[bool, str, str]:
    """Filter extracted table data based on column header relevance."""
    if not table_data or not table_data.strip():
        return (False, "", "Empty data")

    lines = table_data.replace("<0x0A>", "\n").strip().split("\n")
    if len(lines) < 2:
        return (False, "", "Not enough rows")

    headers = []
    for line in lines[:3]:
        parts = [p.strip() for p in line.split("|") if p.strip()]
        if parts:
            headers.extend(parts)

    if not headers:
        return (False, "", "No headers found")

    is_relevant, matching, match_count = check_table_relevance(headers)

    if is_relevant:
        return (True, table_data, f"Found {match_count} relevant columns: {matching[:5]}")
    else:
        return (False, "", f"No relevant columns in: {headers[:10]}")


# ================== PADDLEOCR TEXT EXTRACTION ==================
# Handles rotated tables and combo chart+table elements
# DePlot fails on rotated text, but PaddleOCR handles it automatically

def extract_text_with_paddleocr(image_path: str) -> Dict[str, Any]:
    """
    Extract text from an image using PaddleOCR.
    Handles rotated tables automatically.

    Use for:
    - Rotated tables (DePlot fails on these)
    - Combo chart+table elements
    - Any image where you need raw text extraction

    Args:
        image_path: Path to image file

    Returns:
        Dict with 'status', 'texts' (list of extracted text), 'raw_result'
    """
    try:
        from paddleocr import PaddleOCR

        ocr = PaddleOCR(lang='en')
        result = ocr.predict(image_path)

        texts = []
        if result and len(result) > 0:
            for item in result:
                if 'rec_texts' in item:
                    for text in item['rec_texts']:
                        if len(text.strip()) > 1:  # Skip single chars
                            texts.append(text.strip())

        return {
            "status": "success",
            "texts": texts,
            "text_joined": " | ".join(texts),
            "raw_result": result
        }
    except ImportError as e:
        return {"status": "not_installed", "error": str(e)}
    except Exception as e:
        return {"status": "error", "error": str(e)}


def test_paddleocr(pdf_path: str, page_num: int) -> Dict[str, Any]:
    """Test PaddleOCR LayoutDetection"""
    try:
        import pypdfium2 as pdfium
        import numpy as np
        os.environ['DISABLE_MODEL_SOURCE_CHECK'] = 'True'
        from paddleocr import LayoutDetection

        pdf = pdfium.PdfDocument(pdf_path)
        page = pdf[page_num - 1]  # 0-indexed
        bitmap = page.render(scale=200/72)
        img = bitmap.to_pil()
        img_array = np.array(img)

        # Test with threshold=0.2 (lower to catch combo chart+table elements)
        ld = LayoutDetection(threshold=0.2, device="gpu")
        result = ld.predict(img_array)
        boxes = result[0]['boxes'] if result else []

        # Count by type
        type_counts = {}
        for box in boxes:
            label = box['label']
            type_counts[label] = type_counts.get(label, 0) + 1

        return {
            "status": "success",
            "counts": type_counts,
            "total": len(boxes),
            "details": [(b['label'], round(b['score'], 2)) for b in boxes]
        }
    except Exception as e:
        return {"status": "error", "error": str(e)}


def test_surya(pdf_path: str, page_num: int) -> Dict[str, Any]:
    """Test Surya layout detection"""
    try:
        from surya.foundation import FoundationPredictor
        from surya.layout import LayoutPredictor
        import pypdfium2 as pdfium

        # Load page
        pdf = pdfium.PdfDocument(pdf_path)
        page = pdf[page_num - 1]
        bitmap = page.render(scale=200/72)
        img = bitmap.to_pil()

        # Load models
        foundation = FoundationPredictor()
        predictor = LayoutPredictor(foundation)
        layout_predictions = predictor([img])

        # Count by type
        type_counts = {}
        layout = layout_predictions[0]
        for block in layout.bboxes:
            label = block.label
            type_counts[label] = type_counts.get(label, 0) + 1

        return {
            "status": "success",
            "counts": type_counts,
            "total": len(layout.bboxes),
            "details": [(b.label, round(b.confidence, 2) if hasattr(b, 'confidence') else 'N/A') for b in layout.bboxes]
        }
    except ImportError as e:
        return {"status": "not_installed", "error": f"surya not installed: {e}. Run: pip install surya-ocr"}
    except Exception as e:
        return {"status": "error", "error": str(e)}


def test_marker(pdf_path: str, page_num: int) -> Dict[str, Any]:
    """Test Marker PDF parser"""
    try:
        from marker.converters.pdf import PdfConverter
        from marker.models import create_model_dict

        # Load models
        model_dict = create_model_dict()
        converter = PdfConverter(artifact_dict=model_dict)

        # Convert specific page
        result = converter(pdf_path, page_range=[page_num - 1, page_num])

        # Analyze the output
        type_counts = {
            "images": len(result.images) if hasattr(result, 'images') else 0,
            "tables": result.markdown.count('|') > 10,  # Rough table detection
        }

        return {
            "status": "success",
            "counts": type_counts,
            "markdown_preview": result.markdown[:500] if hasattr(result, 'markdown') else "N/A"
        }
    except ImportError:
        return {"status": "not_installed", "error": "marker not installed. Run: pip install marker-pdf"}
    except Exception as e:
        return {"status": "error", "error": str(e)}


def test_deepdoctection(pdf_path: str, page_num: int) -> Dict[str, Any]:
    """Test deepdoctection layout detection"""
    try:
        import deepdoctection as dd
        import pypdfium2 as pdfium

        # Load page as image
        pdf = pdfium.PdfDocument(pdf_path)
        page = pdf[page_num - 1]
        bitmap = page.render(scale=200/72)
        img = bitmap.to_pil()

        # Save temp image
        temp_path = f"_temp_page_{page_num}.png"
        img.save(temp_path)

        # Create analyzer
        analyzer = dd.get_dd_analyzer()

        # Analyze
        df = analyzer.analyze(path=temp_path)
        df.reset_state()
        doc = iter(df)
        page_result = next(doc)

        # Count by type
        type_counts = {}
        for item in page_result.items:
            label = item.category_name
            type_counts[label] = type_counts.get(label, 0) + 1

        # Cleanup
        os.remove(temp_path)

        return {
            "status": "success",
            "counts": type_counts,
            "total": len(page_result.items),
        }
    except ImportError:
        return {"status": "not_installed", "error": "deepdoctection not installed. Run: pip install deepdoctection"}
    except Exception as e:
        return {"status": "error", "error": str(e)}


def test_unstructured(pdf_path: str, page_num: int) -> Dict[str, Any]:
    """Test Unstructured.io document parser"""
    try:
        from unstructured.partition.pdf import partition_pdf

        # Partition the PDF
        elements = partition_pdf(
            filename=pdf_path,
            strategy="hi_res",
            include_page_breaks=True,
        )

        # Filter to specific page
        page_elements = []
        current_page = 1
        for el in elements:
            if hasattr(el, 'metadata') and hasattr(el.metadata, 'page_number'):
                if el.metadata.page_number == page_num:
                    page_elements.append(el)

        # Count by type
        type_counts = {}
        for el in page_elements:
            label = type(el).__name__
            type_counts[label] = type_counts.get(label, 0) + 1

        return {
            "status": "success",
            "counts": type_counts,
            "total": len(page_elements),
        }
    except ImportError:
        return {"status": "not_installed", "error": "unstructured not installed. Run: pip install unstructured[pdf]"}
    except Exception as e:
        return {"status": "error", "error": str(e)}


def test_grobid(pdf_path: str, page_num: int) -> Dict[str, Any]:
    """Test GROBID (requires running GROBID server)"""
    try:
        from grobid_client.grobid_client import GrobidClient

        client = GrobidClient(config_path="./grobid_config.json")

        # Process PDF
        result = client.process(
            "processFulltextDocument",
            pdf_path,
            consolidate_citations=False,
            include_raw_citations=False,
        )

        # Parse XML result for figures/tables
        # This is simplified - full implementation would parse TEI XML
        type_counts = {
            "figures": result.count("<figure"),
            "tables": result.count("<table"),
            "formulas": result.count("<formula"),
        }

        return {
            "status": "success",
            "counts": type_counts,
        }
    except ImportError:
        return {"status": "not_installed", "error": "grobid_client not installed. Run: pip install grobid-client-python"}
    except Exception as e:
        return {"status": "error", "error": str(e)}


# ================== DEPLOT EXTRACTION (COMMENTED OUT - TOO SLOW ON CPU) ==================
# DePlot takes ~550 seconds per image on CPU. Uncomment if you have GPU with .to('cuda')
#
# def test_deplot_extraction(pdf_path: str, page_num: int) -> Dict[str, Any]:
#     """Test DePlot chart-to-table extraction with relevance filtering."""
#     try:
#         os.environ['DISABLE_MODEL_SOURCE_CHECK'] = 'True'
#         import pypdfium2 as pdfium
#         import numpy as np
#         from paddleocr import LayoutDetection
#         from transformers import Pix2StructProcessor, Pix2StructForConditionalGeneration
#
#         pdf = pdfium.PdfDocument(pdf_path)
#         page = pdf[page_num - 1]
#         bitmap = page.render(scale=200/72)
#         img = bitmap.to_pil()
#         img_array = np.array(img)
#
#         ld = LayoutDetection(threshold=0.2)
#         result = ld.predict(img_array)
#         boxes = result[0]['boxes'] if result else []
#         tables = [b for b in boxes if b['label'].lower() == 'table']
#
#         if not tables:
#             return {"status": "success", "message": "No tables detected", "tables": []}
#
#         print("      Loading DePlot model...")
#         processor = Pix2StructProcessor.from_pretrained('google/deplot')
#         model = Pix2StructForConditionalGeneration.from_pretrained('google/deplot')
#         # For GPU: model = model.to('cuda')
#
#         table_results = []
#         for i, box in enumerate(tables):
#             coord = box.get('coordinate', [])
#             if len(coord) < 4:
#                 continue
#             x1, y1, x2, y2 = int(coord[0]), int(coord[1]), int(coord[2]), int(coord[3])
#             cropped = img.crop((x1, y1, x2, y2))
#             cropped.save(f"_debug_table_{page_num}_{i+1}.png")
#
#             inputs = processor(images=cropped, text="Generate underlying data table of the figure below:", return_tensors="pt")
#             # For GPU: inputs = {k: v.to('cuda') for k, v in inputs.items()}
#             predictions = model.generate(**inputs, max_new_tokens=512)
#             chart_data = processor.decode(predictions[0], skip_special_tokens=True)
#
#             is_relevant, filtered_data, reason = filter_table_data(chart_data)
#             table_results.append({
#                 "index": i + 1, "raw_data": chart_data[:200],
#                 "is_relevant": is_relevant, "filter_reason": reason,
#             })
#
#         return {"status": "success", "tables_detected": len(tables), "table_results": table_results}
#     except ImportError as e:
#         return {"status": "not_installed", "error": str(e)}
#     except Exception as e:
#         import traceback
#         return {"status": "error", "error": str(e), "traceback": traceback.format_exc()}


# ================== COMPREHENSIVE TEST FUNCTION ==================
SKIP_OCR = True  # Skip OCR on elements, just count object types per page

COMPREHENSIVE_TEST_PAGES = {
    # "1.pdf": [7, 9, 11, 13, 14],  # Test pages with known chart/fig counts
    # "2.pdf": [4, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17],
    "3.pdf": [3, 4, 5, 6, 7, 8],
    # "4.pdf": [4, 6, 7, 8, 9],
}

# Crop folder base path
CROP_FOLDER = os.path.join(os.path.dirname(__file__), "OCR Crops")


def clear_crop_folder(pdf_name: str):
    """
    Clear all crop images for a specific PDF before running new tests.
    Clears both main folder and Text OCR crops subfolder.

    Args:
        pdf_name: e.g. "1.pdf" -> clears "OCR Crops/PDF 1/"
    """
    from pathlib import Path

    pdf_num = pdf_name.replace(".pdf", "")
    pdf_folder = Path(CROP_FOLDER) / f"PDF {pdf_num}"

    if not pdf_folder.exists():
        pdf_folder.mkdir(parents=True, exist_ok=True)
        (pdf_folder / "Text OCR crops").mkdir(exist_ok=True)
        print(f"    Created folder: {pdf_folder}")
        return 0

    # Clear main folder
    cleared = 0
    for f in pdf_folder.glob("*.png"):
        f.unlink()
        cleared += 1

    # Clear text subfolder
    text_folder = pdf_folder / "Text OCR crops"
    if text_folder.exists():
        for f in text_folder.glob("*.png"):
            f.unlink()
            cleared += 1
    else:
        text_folder.mkdir(exist_ok=True)

    if cleared > 0:
        print(f"    Cleared {cleared} old crops from PDF {pdf_num}")
    return cleared


def is_valid_figure_title(ocr_texts: list) -> tuple[bool, str]:
    """
    Check if OCR text indicates a valid figure title (starts with Figure/Table).

    Handles OCR truncation where first 1-4 chars are cut off:
    - "Figure 1." → "are 1.", "ure 1.", "gure 1.", "igure 1."
    - "Table 5." → "ble 5.", "able 5."

    Args:
        ocr_texts: List of OCR text strings from the element

    Returns:
        (is_valid, reason) - True if likely a real figure/table title
    """
    if not ocr_texts:
        return (False, "No OCR text")

    # Check first few text elements (title is usually at the start)
    check_texts = ocr_texts[:3]
    combined = " ".join(check_texts).lower().strip()

    # Patterns that indicate Figure/Table (accounting for OCR truncation)
    figure_patterns = ['figure', 'igure', 'gure', 'fig.', 'fig ']
    table_patterns = ['table', 'able', 'tab.', 'tab ']

    # Check first 20 characters for word patterns
    first_part = combined[:25]

    for pattern in figure_patterns:
        if pattern in first_part:
            return (True, f"Found '{pattern}' in: {first_part[:30]}")

    for pattern in table_patterns:
        if pattern in first_part:
            return (True, f"Found '{pattern}' in: {first_part[:30]}")

    # Check for heavily truncated patterns with number:
    # "Figure 1." -> "are 1." or "ure 1." (OCR cuts first 3-4 chars)
    # "Table 5." -> "ble 5." (OCR cuts first 2 chars)
    import re

    # Pattern: truncated_figure_word + space + number + punctuation
    # "are 1.", "ure 2,", "re 3.", "gre 3" etc.
    figure_truncated = re.match(r'^[a-z]{1,3}re?\s+\d+[.,:\s(]', first_part)
    if figure_truncated:
        return (True, f"Found truncated Figure pattern: {first_part[:15]}")

    # Extremely truncated Figure: "e 3 (a)" or "e 4." (only last letter visible)
    extreme_figure = re.match(r'^[egr]\s+\d+\s*[.,:(]', first_part)
    if extreme_figure:
        return (True, f"Found extreme truncated Figure: {first_part[:15]}")

    # Pattern: truncated_table_word + space + number + punctuation
    # "ble 5.", "le 6." etc.
    table_truncated = re.match(r'^[a-z]{0,2}ble\s+\d+[.,:\s(]', first_part)
    if table_truncated:
        return (True, f"Found truncated Table pattern: {first_part[:15]}")

    return (False, f"No Figure/Table pattern in: {first_part[:30]}")


def associate_titles_with_elements(boxes: list, max_distance: int = 100) -> dict:
    """
    Associate figure_titles with their parent tables/charts using spatial proximity.

    In academic papers:
    - Table captions are ABOVE the table
    - Figure/Chart captions are BELOW the figure/chart

    Args:
        boxes: List of detected boxes with 'label' and 'coordinate'
        max_distance: Maximum vertical distance (pixels) to consider association

    Returns:
        Dict mapping element index to associated title info:
        {element_idx: {'title_idx': int, 'title_text': str, 'distance': int}}
    """
    # Separate titles from tables/charts
    titles = []
    elements = []

    for i, box in enumerate(boxes):
        label = box['label'].lower()
        coord = box.get('coordinate', [])
        if len(coord) < 4:
            continue

        item = {
            'idx': i,
            'label': label,
            'x1': coord[0], 'y1': coord[1],
            'x2': coord[2], 'y2': coord[3],
            'center_x': (coord[0] + coord[2]) / 2,
            'texts': box.get('texts', [])
        }

        if label == 'figure_title':
            titles.append(item)
        elif label in ('table', 'chart', 'image'):
            elements.append(item)

    associations = {}

    for elem in elements:
        best_title = None
        best_distance = float('inf')

        for title in titles:
            # Check horizontal overlap (must be roughly aligned)
            x_overlap = min(elem['x2'], title['x2']) - max(elem['x1'], title['x1'])
            elem_width = elem['x2'] - elem['x1']
            if x_overlap < elem_width * 0.3:  # At least 30% horizontal overlap
                continue

            # For tables: title should be ABOVE (smaller Y)
            # For charts/images: title should be BELOW (larger Y)
            if elem['label'] == 'table':
                # Title above table: title.y2 < elem.y1
                if title['y2'] < elem['y1']:
                    distance = elem['y1'] - title['y2']
                    if distance < max_distance and distance < best_distance:
                        best_distance = distance
                        best_title = title
            else:
                # Title below chart/image: title.y1 > elem.y2
                if title['y1'] > elem['y2']:
                    distance = title['y1'] - elem['y2']
                    if distance < max_distance and distance < best_distance:
                        best_distance = distance
                        best_title = title

        if best_title:
            associations[elem['idx']] = {
                'title_idx': best_title['idx'],
                'title_label': best_title['label'],
                'distance': int(best_distance),
                'title_texts': best_title.get('texts', [])
            }

    return associations


def calculate_iou(box1: list, box2: list) -> float:
    """
    Calculate Intersection over Union (IoU) between two bounding boxes.

    Args:
        box1, box2: [x1, y1, x2, y2] coordinates

    Returns:
        IoU value between 0.0 and 1.0
    """
    x1 = max(box1[0], box2[0])
    y1 = max(box1[1], box2[1])
    x2 = min(box1[2], box2[2])
    y2 = min(box1[3], box2[3])

    # No intersection
    if x2 <= x1 or y2 <= y1:
        return 0.0

    intersection = (x2 - x1) * (y2 - y1)
    area1 = (box1[2] - box1[0]) * (box1[3] - box1[1])
    area2 = (box2[2] - box2[0]) * (box2[3] - box2[1])
    union = area1 + area2 - intersection

    return intersection / union if union > 0 else 0.0


def remove_duplicate_boxes(boxes: list, iou_threshold: float = 0.5) -> list:
    """
    Remove duplicate/overlapping bounding boxes using Non-Maximum Suppression (NMS).

    When two boxes of the same type overlap more than iou_threshold:
    - Keep the one with higher confidence score
    - If a larger box contains a smaller box (>70% overlap), keep only the smaller
      (more specific) detection

    Args:
        boxes: List of box dicts with 'coordinate', 'score', 'label'
        iou_threshold: IoU threshold above which boxes are considered duplicates

    Returns:
        Filtered list of boxes with duplicates removed
    """
    if not boxes:
        return []

    # Sort by score descending
    sorted_boxes = sorted(boxes, key=lambda x: x.get('score', 0), reverse=True)

    keep = []
    suppressed = set()

    for i, box in enumerate(sorted_boxes):
        if i in suppressed:
            continue

        coord1 = box.get('coordinate', [])
        if len(coord1) < 4:
            continue

        x1, y1, x2, y2 = coord1[0], coord1[1], coord1[2], coord1[3]
        area1 = (x2 - x1) * (y2 - y1)
        label1 = box['label'].lower()

        # Check against remaining boxes
        for j in range(i + 1, len(sorted_boxes)):
            if j in suppressed:
                continue

            other = sorted_boxes[j]
            coord2 = other.get('coordinate', [])
            if len(coord2) < 4:
                continue

            label2 = other['label'].lower()

            # Only compare same-type boxes for standard NMS
            # But also check if one contains the other for any type
            ox1, oy1, ox2, oy2 = coord2[0], coord2[1], coord2[2], coord2[3]
            area2 = (ox2 - ox1) * (oy2 - oy1)

            iou = calculate_iou([x1, y1, x2, y2], [ox1, oy1, ox2, oy2])

            # Calculate containment (what % of smaller box is inside larger)
            int_x1 = max(x1, ox1)
            int_y1 = max(y1, oy1)
            int_x2 = min(x2, ox2)
            int_y2 = min(y2, oy2)

            if int_x2 > int_x1 and int_y2 > int_y1:
                intersection = (int_x2 - int_x1) * (int_y2 - int_y1)
                smaller_area = min(area1, area2)
                containment = intersection / smaller_area if smaller_area > 0 else 0
            else:
                containment = 0

            # If same type and high IoU, suppress lower-scoring one
            if label1 == label2 and iou > iou_threshold:
                suppressed.add(j)
            # If same type and one box mostly contains the other (>70%), keep the smaller (more specific)
            elif label1 == label2 and containment > 0.7:
                if area1 > area2:
                    # Current box is larger, suppress it and keep smaller
                    suppressed.add(i)
                    break
                else:
                    suppressed.add(j)

        if i not in suppressed:
            keep.append(box)

    return keep


def run_comprehensive_paddleocr_test(save_images: bool = True, save_json: bool = True):
    """
    Run PaddleOCR layout detection + text extraction on all 21 test pages.

    Uses multi-pass detection:
    - Pass 1 at threshold 0.3 (high confidence)
    - Pass 2 at threshold 0.2 (rescue borderline elements)
    - Merge + dedup removes duplicates

    Outputs:
    - Layout detection results (tables, charts, images detected)
    - PaddleOCR text extraction for each detected element
    - Saves cropped images and JSON results

    Args:
        save_images: Save cropped element images
        save_json: Save JSON results file
    """
    import json
    import pypdfium2 as pdfium
    import numpy as np
    os.environ['DISABLE_MODEL_SOURCE_CHECK'] = 'True'
    from paddleocr import LayoutDetection, PaddleOCR

    # Calculate total pages being tested
    total_pages = sum(len(pages) for pages in COMPREHENSIVE_TEST_PAGES.values())
    pdf_list = ", ".join(COMPREHENSIVE_TEST_PAGES.keys())

    print("=" * 80)
    print(f"PADDLEOCR LAYOUT DETECTION - {total_pages} Pages")
    print(f"  PDFs: {pdf_list}")
    print("  Multi-pass detection (0.4 + 0.25 threshold) with duplicate removal")
    print("  Render scale: 900 DPI (3x)")
    print("=" * 80)

    # Initialize models once - two thresholds for multi-pass
    print("\nInitializing PaddleOCR models (GPU)...")
    start_init = time.time()
    ld_high = LayoutDetection(threshold=0.4, device="gpu")  # High confidence pass
    ld_low = LayoutDetection(threshold=0.25, device="gpu")  # Rescue pass
    ocr = PaddleOCR(lang='en', device="gpu")
    print(f"Models loaded in {time.time() - start_init:.1f}s (GPU enabled)")

    all_results = {}
    total_elements = {"table": 0, "chart": 0, "figure": 0, "image": 0}

    # Types to save in main folder vs text subfolder
    main_folder_types = {'image', 'chart', 'figure_title', 'formula', 'table'}

    for pdf_name, pages in COMPREHENSIVE_TEST_PAGES.items():
        pdf_path = pdf_name
        if not os.path.exists(pdf_path):
            print(f"\n[SKIP] {pdf_name} not found")
            continue

        # Clear old crops before processing this PDF
        clear_crop_folder(pdf_name)

        pdf = pdfium.PdfDocument(pdf_path)
        pdf_results = {}
        pdf_num = pdf_name.replace(".pdf", "")

        print(f"\n{'='*60}")
        print(f"PDF: {pdf_name}")
        print(f"{'='*60}")

        for page_num in pages:
            print(f"\n  Page {page_num}...")
            start_page = time.time()

            # Render page at 3x scale (900 DPI)
            page = pdf[page_num - 1]
            bitmap = page.render(scale=900/72)  # 900 DPI (3x)
            img = bitmap.to_pil()
            img_array = np.array(img)

            # Multi-pass layout detection
            result_high = ld_high.predict(img_array)
            result_low = ld_low.predict(img_array)
            boxes_high = result_high[0]['boxes'] if result_high else []
            boxes_low = result_low[0]['boxes'] if result_low else []

            # Only rescue tables and charts from low-confidence pass (avoid image/formula garbage)
            rescue_types = {'table', 'chart'}
            boxes_rescued = [b for b in boxes_low if b['label'].lower() in rescue_types]

            # Merge high-confidence boxes + rescued tables/charts
            raw_boxes = boxes_high + boxes_rescued

            # Apply duplicate removal (NMS)
            boxes = remove_duplicate_boxes(raw_boxes, iou_threshold=0.5)
            removed_count = len(raw_boxes) - len(boxes)
            if removed_count > 0:
                print(f"    [DEDUP] Removed {removed_count} duplicate boxes ({len(raw_boxes)} -> {len(boxes)})")

            # Count by type
            type_counts = {}
            for box in boxes:
                label = box['label'].lower()
                type_counts[label] = type_counts.get(label, 0) + 1
                if label in total_elements:
                    total_elements[label] += 1

            print(f"    Layout: {type_counts} ({time.time() - start_page:.1f}s)")

            # Process each element
            page_elements = []
            # Track saved element counts per type for naming
            saved_type_counts = {}

            # Progress tracking for ETA
            num_boxes = len(boxes)
            element_times = []
            ocr_start_page = time.time()

            # SKIP_OCR mode: just count elements, no OCR
            if SKIP_OCR:
                print(f"    [SKIP_OCR] Counting {num_boxes} elements (no OCR)")
                pdf_results[page_num] = {
                    "layout_counts": type_counts,
                    "elements": [],
                    "total_time": round(time.time() - start_page, 2),
                }
                continue

            for i, box in enumerate(boxes):
                element_start = time.time()
                label = box['label'].lower()
                score = box.get('score', 0)
                coord = box.get('coordinate', [])

                if len(coord) < 4:
                    continue

                x1, y1, x2, y2 = int(coord[0]), int(coord[1]), int(coord[2]), int(coord[3])
                cropped = img.crop((x1, y1, x2, y2))

                # Progress logging with ETA
                if element_times:
                    avg_time = sum(element_times) / len(element_times)
                    remaining = num_boxes - i
                    eta = avg_time * remaining
                    print(f"    [{i+1}/{num_boxes}] {label} (ETA: {eta:.0f}s)", end="\r")
                else:
                    print(f"    [{i+1}/{num_boxes}] {label}...", end="\r")

                # Run OCR on element FIRST (before saving)
                ocr_start = time.time()
                try:
                    # Save temp for OCR
                    temp_path = f"_temp_ocr_{i}.png"
                    cropped.save(temp_path)
                    ocr_result = ocr.predict(temp_path)
                    os.remove(temp_path)

                    # Extract texts
                    texts = []
                    if ocr_result and len(ocr_result) > 0:
                        for item in ocr_result:
                            if 'rec_texts' in item:
                                for text in item['rec_texts']:
                                    if len(text.strip()) > 1:
                                        texts.append(text.strip())

                    ocr_time = time.time() - ocr_start
                    text_preview = " | ".join(texts[:10])
                    if len(text_preview) > 100:
                        text_preview = text_preview[:100] + "..."

                    # Filter figure_titles: must contain Figure/Table pattern
                    if label == 'figure_title':
                        is_valid, filter_reason = is_valid_figure_title(texts)
                        if not is_valid:
                            element_times.append(time.time() - element_start)
                            print(" " * 60, end="\r")
                            print(f"      [{i+1}] {label} ({score:.2f}) - DISCARDED: {filter_reason}")
                            continue  # Skip this element - don't save or add to results

                    # Save cropped image ONLY for valid elements
                    if save_images:
                        from pathlib import Path
                        from datetime import datetime

                        # Naming: TIME_page#_type_#.png
                        time_str = datetime.now().strftime('%I%M%p').lower()

                        # Count elements of this type on this page (for naming)
                        type_key = f"{page_num}_{label}"
                        if type_key not in saved_type_counts:
                            saved_type_counts[type_key] = 0
                        saved_type_counts[type_key] += 1
                        type_num = saved_type_counts[type_key]

                        img_filename = f"{time_str}_page{page_num}_{label}_{type_num}.png"

                        # Save to main folder or text subfolder based on type
                        pdf_folder = Path(CROP_FOLDER) / f"PDF {pdf_num}"
                        if label in main_folder_types:
                            save_path = pdf_folder / img_filename
                        else:
                            save_path = pdf_folder / "Text OCR crops" / img_filename

                        cropped.save(save_path)

                    element_data = {
                        "index": i + 1,
                        "label": label,
                        "score": round(score, 3),
                        "bbox": [x1, y1, x2, y2],
                        "size": f"{x2-x1}x{y2-y1}",
                        "texts": texts,
                        "text_count": len(texts),
                        "ocr_time": round(ocr_time, 2),
                    }
                    page_elements.append(element_data)

                    # Track element time for ETA
                    element_times.append(time.time() - element_start)

                    # Clear progress line and print result
                    print(" " * 60, end="\r")
                    print(f"      [{i+1}] {label} ({score:.2f}) - {len(texts)} texts ({ocr_time:.1f}s): {text_preview}")

                except Exception as e:
                    element_times.append(time.time() - element_start)
                    print(" " * 60, end="\r")
                    print(f"      [{i+1}] {label} - OCR ERROR: {e}")
                    page_elements.append({
                        "index": i + 1,
                        "label": label,
                        "score": round(score, 3),
                        "error": str(e)
                    })

            pdf_results[page_num] = {
                "layout_counts": type_counts,
                "elements": page_elements,
                "total_time": round(time.time() - start_page, 2),
            }

        all_results[pdf_name] = pdf_results

    # Summary
    print("\n" + "=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"Total elements found: {sum(total_elements.values())}")
    for elem_type, count in total_elements.items():
        if count > 0:
            print(f"  - {elem_type}: {count}")

    # ASCII Table Summary - counts per page with Expected vs Actual
    print("\n" + "=" * 100)
    print("DETECTION RESULTS BY PAGE (E=Expected, A=Actual)")
    print("=" * 100)

    # Ground truth - expected counts per page: (pdf, page) -> {table, chart, fig_title}
    # Based on manual review
    EXPECTED_COUNTS = {
        # 1.pdf
        ("1.pdf", 7):  {"table": 0, "chart": 2, "figure_title": 1},
        ("1.pdf", 9):  {"table": 0, "chart": 3, "figure_title": 1},
        ("1.pdf", 11): {"table": 0, "chart": 1, "figure_title": 1},
        ("1.pdf", 13): {"table": 0, "chart": 0, "figure_title": 1},
        ("1.pdf", 14): {"table": 0, "chart": 3, "figure_title": 1},
        # 2.pdf
        ("2.pdf", 4):  {"table": 1, "chart": 0, "figure_title": 3},
        ("2.pdf", 7):  {"table": 0, "chart": 1, "figure_title": 1},
        ("2.pdf", 8):  {"table": 0, "chart": 0, "figure_title": 1},  # quasi-table discardable
        ("2.pdf", 9):  {"table": 0, "chart": 1, "figure_title": 1},
        ("2.pdf", 11): {"table": 0, "chart": 1, "figure_title": 1},
        ("2.pdf", 12): {"table": 0, "chart": 2, "figure_title": 2},
        ("2.pdf", 13): {"table": 0, "chart": 1, "figure_title": 1},
        ("2.pdf", 14): {"table": 0, "chart": 2, "figure_title": 2},
        ("2.pdf", 15): {"table": 0, "chart": 2, "figure_title": 2},
        ("2.pdf", 16): {"table": 0, "chart": 2, "figure_title": 2},
        ("2.pdf", 17): {"table": 0, "chart": 2, "figure_title": 2},
        # 3.pdf
        ("3.pdf", 3):  {"table": 1, "chart": 0, "figure_title": 1},
        ("3.pdf", 4):  {"table": 2, "chart": 0, "figure_title": 1},
        ("3.pdf", 5):  {"table": 2, "chart": 0, "figure_title": 2},
        ("3.pdf", 6):  {"table": 1, "chart": 2, "figure_title": 3},
        ("3.pdf", 7):  {"table": 1, "chart": 0, "figure_title": 1},
        ("3.pdf", 8):  {"table": 1, "chart": 0, "figure_title": 2},
    }

    # Primary columns with E/A pairs, then monitor columns
    primary_cols = [("table", "table"), ("chart", "chart"), ("fig_title", "figure_title")]
    monitor_cols = [("image", "image"), ("formula", "formula")]

    # Print header with E/A sub-columns for primary cols
    # Use abbreviations: tbl=table, cht=chart, fig=fig_title
    abbrev = {"table": "tbl", "chart": "cht", "fig_title": "fig"}
    header = f"{'Page':<10}"
    for col_name, _ in primary_cols:
        ab = abbrev[col_name]
        header += f"{ab+'E':<6}{ab+'A':<6}"
    header += " ||| "
    for col_name, _ in monitor_cols:
        header += f"{col_name:<10}"
    print(header)
    print("-" * len(header))

    # Track totals for E and A
    totals_e = {"table": 0, "chart": 0, "fig_title": 0}
    totals_a = {"table": 0, "chart": 0, "fig_title": 0}
    totals_monitor = {"image": 0, "formula": 0}

    # Print each PDF's pages
    for pdf_name, pdf_results in all_results.items():
        pdf_num = pdf_name.replace(".pdf", "")
        for page_num in sorted(pdf_results.keys()):
            page_data = pdf_results[page_num]
            counts = page_data.get("layout_counts", {})
            expected = EXPECTED_COUNTS.get((pdf_name, page_num), {})

            row = f"{pdf_num}-{page_num:<7}"
            for col_name, actual_key in primary_cols:
                exp = expected.get(actual_key, 0)
                act = counts.get(actual_key, 0)
                totals_e[col_name] += exp
                totals_a[col_name] += act
                # Mark mismatches with *
                marker = "*" if exp != act else ""
                row += f"{exp:<6}{act}{marker:<5}"
            row += " ||| "
            for col_name, actual_key in monitor_cols:
                act = counts.get(actual_key, 0)
                totals_monitor[col_name] += act
                row += f"{act:<10}"
            print(row)

    print("-" * len(header))

    # Print totals row
    total_row = f"{'TOTAL':<10}"
    for col_name, _ in primary_cols:
        total_row += f"{totals_e[col_name]:<6}{totals_a[col_name]:<6}"
    total_row += " ||| "
    for col_name, _ in monitor_cols:
        total_row += f"{totals_monitor[col_name]:<10}"
    print(total_row)
    print("=" * len(header))
    print("* = mismatch between expected and actual")

    # Note about text folder
    print("\nNote: 'text' type elements saved to 'Text OCR crops' subfolder")

    # Save JSON
    if save_json:
        json_path = "paddleocr_comprehensive_results.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(all_results, f, indent=2, ensure_ascii=False)
        print(f"\nResults saved to: {json_path}")

    return all_results


def run_relevance_filter_test():
    """Test the relevance filter on 3.pdf page 4 (cam design table should be filtered out)."""
    print("\n" + "=" * 80)
    print("RELEVANCE FILTER TEST - 3.pdf Page 4")
    print("Expected: Cam design table FILTERED OUT, real data table KEPT")
    print("=" * 80)

    result = test_deplot_extraction("3.pdf", 4)

    if result["status"] == "success":
        print(f"\n  Tables detected: {result.get('tables_detected', 0)}")
        for tbl in result.get("table_results", []):
            print(f"\n  --- Table {tbl['index']} ---")
            print(f"  Raw data: {tbl['raw_data']}...")
            print(f"  Is Relevant: {tbl['is_relevant']}")
            print(f"  Reason: {tbl['filter_reason']}")
    else:
        print(f"\n  Error: {result.get('error', 'Unknown')}")
        if 'traceback' in result:
            print(result['traceback'])


def run_all_tests():
    """Run all parser tests on all test cases"""

    parsers = [
        ("PaddleOCR", test_paddleocr),
        ("Surya", test_surya),
        ("Marker", test_marker),
        ("deepdoctection", test_deepdoctection),
        ("Unstructured", test_unstructured),
        ("GROBID", test_grobid),
    ]

    print("=" * 80)
    print("CHART/LAYOUT OCR TESTER - Comparing Parsers")
    print("=" * 80)

    for test_case in TEST_CASES:
        pdf = test_case["pdf"]
        page = test_case["page"]
        expected = test_case["expected"]

        print(f"\n{'='*80}")
        print(f"TEST: {pdf} - Page {page}")
        print(f"Expected: {expected}")
        print("=" * 80)

        for parser_name, parser_func in parsers:
            print(f"\n--- {parser_name} ---")
            start = time.time()

            try:
                result = parser_func(pdf, page)
                elapsed = time.time() - start

                if result["status"] == "success":
                    print(f"  Status: SUCCESS ({elapsed:.1f}s)")
                    print(f"  Counts: {result.get('counts', {})}")
                    if 'total' in result:
                        print(f"  Total elements: {result['total']}")
                elif result["status"] == "not_installed":
                    print(f"  Status: NOT INSTALLED")
                    print(f"  {result['error']}")
                else:
                    print(f"  Status: ERROR")
                    print(f"  {result.get('error', 'Unknown error')}")
            except Exception as e:
                print(f"  Status: EXCEPTION")
                print(f"  {str(e)}")

    print("\n" + "=" * 80)
    print("TESTING COMPLETE")
    print("=" * 80)


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        if sys.argv[1] == "--relevance":
            run_relevance_filter_test()
        elif sys.argv[1] == "--comprehensive":
            run_comprehensive_paddleocr_test()
        elif sys.argv[1] == "--help":
            print("Usage: python ChartOCRTester.py [option]")
            print("  (no args)       Run basic parser comparison tests")
            print("  --comprehensive Run PaddleOCR on all 21 test pages")
            print("  --relevance     Run relevance filter test")
        else:
            print(f"Unknown option: {sys.argv[1]}")
            print("Use --help for usage info")
    else:
        run_all_tests()
