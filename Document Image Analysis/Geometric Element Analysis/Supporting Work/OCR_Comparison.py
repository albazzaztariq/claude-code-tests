"""
OCR Tool Comparison - Accuracy by Object Type
==============================================
Run 2 tools at a time by commenting/uncommenting sections below.

Tools:
  1. PaddleOCR (baseline) - already tested via ChartOCRTester.py
  2. Marker (with Ollama LLM)
  3. Mistral OCR (cloud API - needs MISTRAL_API_KEY)
  4. PdfTable (GitHub install required)

Usage:
  - Uncomment the tools you want to test in TOOLS_TO_RUN below
  - Run: python OCR_Comparison.py
"""

import os
import json
import time
import subprocess
from pathlib import Path

BASE_DIR = Path(__file__).parent
PDF_DIR = BASE_DIR

# ============================================================
# CONFIGURATION - Uncomment tools to run (max 2 at a time)
# ============================================================
TOOLS_TO_RUN = [
    # "paddleocr",     # Baseline - uses existing ChartOCRTester results
    # "marker",        # PDF to markdown with Ollama LLM
    # "mistral_ocr", # Cloud API (needs MISTRAL_API_KEY)
    "pdftable",    # Deep learning table extraction
]

# Limit to specific PDFs for faster testing (set to None for all)
TEST_ONLY_PDFS = None  # Test all PDFs from ground truth

# ============================================================
# GROUND TRUTH
# ============================================================
def load_ground_truth():
    with open(BASE_DIR / "OCR_GroundTruth.json") as f:
        return json.load(f)

def get_expected_counts():
    """Return {(pdf, page): {table: N, chart: N, figure_title: N}}"""
    gt = load_ground_truth()
    counts = {}
    for pdf_name, pages in gt['pdfs'].items():
        for page_str, expected in pages.items():
            counts[(pdf_name, int(page_str))] = expected
    return counts


# ============================================================
# TOOL 1: PADDLEOCR (baseline)
# Uses existing results from ChartOCRTester.py
# ============================================================
def run_paddleocr():
    """Load PaddleOCR results from ChartOCRTester output."""
    print("\n" + "="*60)
    print("PADDLEOCR (baseline)")
    print("="*60)
    print("  [GPU] YES - uses device='gpu' in LayoutDetection and PaddleOCR")
    tool_start = time.time()

    results_file = BASE_DIR / "paddleocr_comprehensive_results.json"
    if not results_file.exists():
        print("  [PROGRESS] Results file not found, running ChartOCRTester.py...")
        print("  [PROGRESS] This may take several minutes...")
        subprocess.run(["python", "ChartOCRTester.py", "--comprehensive"], cwd=BASE_DIR)

    if results_file.exists():
        with open(results_file) as f:
            data = json.load(f)

        # Log what was loaded
        pdf_count = len([k for k in data.keys() if k.endswith('.pdf')])
        print(f"  [LOADED] {results_file.name}")
        print(f"  [STATS] {pdf_count} PDFs processed")
        for pdf_name, pdf_data in data.items():
            if pdf_name.endswith('.pdf') and isinstance(pdf_data, dict):
                page_count = len([k for k in pdf_data.keys() if k.isdigit()])
                print(f"    - {pdf_name}: {page_count} pages")

        print(f"  [TIME] Loaded in {time.time() - tool_start:.1f}s")
        return data
    else:
        print("  [ERROR] No results file found")
        return None


# ============================================================
# TOOL 2: MARKER (with Ollama LLM)
# ============================================================
def run_marker():
    """Run Marker PDF to markdown with Ollama."""
    print("\n" + "="*60)
    print("MARKER (with Ollama qwen2.5:14b)")
    print("="*60)
    print("  [GPU] YES - Marker uses PyTorch (GPU), Ollama uses GPU for LLM")
    tool_start = time.time()

    gt = load_ground_truth()
    results = {}
    total_pdfs = len(gt['pdfs'])
    pdf_idx = 0

    for pdf_name, pages in gt['pdfs'].items():
        pdf_idx += 1
        page_list = list(pages.keys())
        page_count = len(page_list)
        # 0-indexed for marker
        page_range = ",".join(str(int(p)-1) for p in page_list)

        pdf_path = PDF_DIR / pdf_name
        if not pdf_path.exists():
            print(f"  [{pdf_idx}/{total_pdfs}] {pdf_name}: [ERROR] FILE NOT FOUND")
            continue

        output_dir = BASE_DIR / "marker_output" / pdf_name.replace('.pdf', '')
        output_dir.mkdir(parents=True, exist_ok=True)

        print(f"  [{pdf_idx}/{total_pdfs}] {pdf_name} ({page_count} pages: {page_list})")
        print(f"    [PROGRESS] Starting Marker with Ollama LLM...")
        print(f"    [PROGRESS] Page range (0-indexed): {page_range}")
        start = time.time()

        # Run marker CLI
        cmd = [
            "marker_single", str(pdf_path),
            "--output_dir", str(output_dir),
            "--page_range", page_range,
            "--use_llm",
            "--llm_service", "marker.services.ollama.OllamaService",
            "--ollama_model", "qwen2.5:14b",
            "--output_format", "json",
        ]
        print(f"    [CMD] {' '.join(cmd[:3])}...")

        try:
            # Stream output for progress
            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)

            # Read output line by line
            line_count = 0
            for line in process.stdout:
                line_count += 1
                line = line.strip()
                if line and (line_count % 10 == 0 or "error" in line.lower() or "%" in line):
                    print(f"    [OUTPUT] {line[:80]}")

            process.wait(timeout=600)
            elapsed = time.time() - start

            # Parse output JSON
            json_files = list(output_dir.glob("*.json"))
            if json_files:
                with open(json_files[0]) as f:
                    marker_data = json.load(f)

                # Count detected elements from markdown
                md_files = list(output_dir.glob("*.md"))
                md_content = ""
                if md_files:
                    with open(md_files[0], encoding='utf-8') as f:
                        md_content = f.read()

                table_count = md_content.count('|---')
                image_count = md_content.count('![')

                results[pdf_name] = {
                    "status": "ok",
                    "time": round(elapsed, 1),
                    "tables_detected": table_count,
                    "images_detected": image_count,
                    "markdown_length": len(md_content)
                }
                print(f"    [DONE] {elapsed:.1f}s | Tables: ~{table_count}, Images: ~{image_count}")
            else:
                results[pdf_name] = {"status": "error", "error": "No JSON output"}
                print(f"    [ERROR] No JSON output generated")

        except subprocess.TimeoutExpired:
            process.kill()
            results[pdf_name] = {"status": "timeout"}
            print(f"    [TIMEOUT] Killed after 600s")
        except Exception as e:
            results[pdf_name] = {"status": "error", "error": str(e)}
            print(f"    [ERROR] {e}")

    total_time = time.time() - tool_start
    print(f"\n  [TOTAL TIME] Marker: {total_time:.1f}s")
    return results


# ============================================================
# TOOL 3: MISTRAL OCR (cloud API)
# ============================================================
def run_mistral_ocr():
    """Run Mistral OCR (requires MISTRAL_API_KEY)."""
    print("\n" + "="*60)
    print("MISTRAL OCR (Cloud API)")
    print("="*60)
    print("  [GPU] CLOUD - runs on Mistral's servers (GPU)")
    print("  [COST] ~$0.001 per page")
    tool_start = time.time()

    api_key = os.environ.get("MISTRAL_API_KEY")
    if not api_key:
        print("  [SKIPPED] Set MISTRAL_API_KEY environment variable")
        print("  [HINT] set MISTRAL_API_KEY=your_key_here")
        return None

    try:
        from mistralai import Mistral
        import base64
    except ImportError:
        print("  [ERROR] pip install mistralai")
        return None

    print("  [PROGRESS] Initializing Mistral client...")
    client = Mistral(api_key=api_key)
    gt = load_ground_truth()
    results = {}
    total_pdfs = len(gt['pdfs'])
    pdf_idx = 0

    for pdf_name, pages in gt['pdfs'].items():
        pdf_idx += 1
        pdf_path = PDF_DIR / pdf_name
        if not pdf_path.exists():
            print(f"  [{pdf_idx}/{total_pdfs}] {pdf_name}: [ERROR] FILE NOT FOUND")
            continue

        page_list = list(pages.keys())
        print(f"  [{pdf_idx}/{total_pdfs}] {pdf_name} ({len(page_list)} pages)")
        print(f"    [PROGRESS] Uploading PDF to Mistral API...")
        start = time.time()

        try:
            with open(pdf_path, 'rb') as f:
                pdf_data = base64.standard_b64encode(f.read()).decode('utf-8')

            print(f"    [PROGRESS] Processing with mistral-ocr-latest...")
            response = client.ocr.process(
                model="mistral-ocr-latest",
                document={
                    "type": "document_url",
                    "document_url": f"data:application/pdf;base64,{pdf_data}"
                }
            )

            elapsed = time.time() - start
            page_count = len(response.pages) if hasattr(response, 'pages') else 0

            results[pdf_name] = {
                "status": "ok",
                "time": round(elapsed, 1),
                "pages": page_count,
                "response": response
            }
            print(f"    [DONE] {elapsed:.1f}s | {page_count} pages processed")

        except Exception as e:
            results[pdf_name] = {"status": "error", "error": str(e)}
            print(f"    [ERROR] {e}")

    total_time = time.time() - tool_start
    print(f"\n  [TOTAL TIME] Mistral OCR: {total_time:.1f}s")
    return results


# ============================================================
# TOOL 4: PDFTABLE (deep learning)
# ============================================================
def run_pdftable():
    """Run PdfTable toolkit via CLI."""
    print("\n" + "="*60)
    print("PDFTABLE (Deep Learning Layout/Table Extraction)")
    print("="*60)
    print("  [GPU] YES - uses ONNX with CUDA (FP16)")
    print("  [LAYOUT] picodet (publaynet): text, title, list, table, figure")
    tool_start = time.time()

    # Check if pdftable CLI is available
    try:
        result = subprocess.run(["pdftable", "--help"], capture_output=True, text=True, timeout=30)
        if result.returncode != 0:
            print("  [ERROR] pdftable CLI not working")
            return None
    except FileNotFoundError:
        print("  [ERROR] pdftable not installed")
        print("  [INSTALL] git clone https://github.com/CycloneBoy/pdf_table")
        print("           cd pdf_table && python setup.py install")
        return None

    gt = load_ground_truth()
    results = {}
    total_pdfs = len(gt['pdfs'])
    pdf_idx = 0

    output_base = BASE_DIR / "pdftable_output"
    output_base.mkdir(exist_ok=True)

    for pdf_name, pages in gt['pdfs'].items():
        pdf_idx += 1
        pdf_path = PDF_DIR / pdf_name
        if not pdf_path.exists():
            print(f"  [{pdf_idx}/{total_pdfs}] {pdf_name}: [ERROR] FILE NOT FOUND")
            continue

        page_list = list(pages.keys())
        # pdftable uses 1-indexed pages, comma-separated
        page_range = ",".join(page_list)

        output_dir = output_base / pdf_name.replace('.pdf', '')
        output_dir.mkdir(parents=True, exist_ok=True)

        print(f"  [{pdf_idx}/{total_pdfs}] {pdf_name} ({len(page_list)} pages: {page_list})")
        print(f"    [PROGRESS] Running pdftable with layout detection...")
        start = time.time()

        cmd = [
            "pdftable",
            "--output_dir", str(output_dir),
            "--file_path_or_url", str(pdf_path),
            "--pages", page_range,
            "--lang", "en",  # publaynet labels: text, title, list, table, figure
            "--debug",
        ]

        try:
            # Run and capture output
            process = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding='utf-8', errors='replace'
            )

            output_lines = []
            for line in process.stdout:
                output_lines.append(line)
                # Show progress indicators
                if "Loading" in line or "开始提取" in line or "ERROR" in line.upper():
                    print(f"    [LOG] {line.strip()[:70]}")

            process.wait(timeout=300)
            elapsed = time.time() - start

            # Parse output for layout detection results
            # Look for detected elements in the output or HTML files
            full_output = "".join(output_lines)

            # Count elements from HTML output files
            html_files = list(output_dir.glob("*.html"))
            table_count = 0
            figure_count = 0

            for html_file in html_files:
                try:
                    content = html_file.read_text(encoding='utf-8', errors='replace')
                    # Count tables (HTML tables in output)
                    table_count += content.count('<table')
                    # Count figures (images referenced)
                    figure_count += content.count('<img')
                except:
                    pass

            results[pdf_name] = {
                "status": "ok",
                "time": round(elapsed, 1),
                "tables_detected": table_count,
                "figures_detected": figure_count,
                "html_files": len(html_files),
                "pages_processed": page_list
            }
            print(f"    [DONE] {elapsed:.1f}s | Tables: {table_count}, Figures: {figure_count}, HTML files: {len(html_files)}")

        except subprocess.TimeoutExpired:
            process.kill()
            results[pdf_name] = {"status": "timeout"}
            print(f"    [TIMEOUT] Killed after 300s")
        except Exception as e:
            results[pdf_name] = {"status": "error", "error": str(e)}
            print(f"    [ERROR] {e}")

    total_time = time.time() - tool_start
    print(f"\n  [TOTAL TIME] PdfTable: {total_time:.1f}s")
    return results


# ============================================================
# ACCURACY CALCULATION
# ============================================================
def calculate_accuracy(tool_results, tool_name):
    """Calculate accuracy per object type vs ground truth."""
    expected = get_expected_counts()

    accuracy = {
        "table": {"expected": 0, "detected": 0, "correct": 0},
        "chart": {"expected": 0, "detected": 0, "correct": 0},
        "figure_title": {"expected": 0, "detected": 0, "correct": 0},
    }

    if tool_results is None:
        return None

    # Tool-specific parsing
    if tool_name == "paddleocr":
        # Parse paddleocr_comprehensive_results.json format
        for pdf_name, pdf_data in tool_results.items():
            if not isinstance(pdf_data, dict):
                continue
            for page_str, page_data in pdf_data.items():
                try:
                    page_num = int(page_str)
                except:
                    continue

                key = (pdf_name, page_num)
                if key not in expected:
                    continue

                exp = expected[key]
                detected = page_data.get("layout_counts", {})

                for obj_type in ["table", "chart", "figure_title"]:
                    exp_count = exp.get(obj_type, 0)
                    det_count = detected.get(obj_type, 0)

                    accuracy[obj_type]["expected"] += exp_count
                    accuracy[obj_type]["detected"] += det_count
                    # Correct = min of expected and detected (can't be more correct than expected)
                    accuracy[obj_type]["correct"] += min(exp_count, det_count)

    elif tool_name == "pdftable":
        # Parse pdftable results - aggregated per PDF
        # pdftable detects: table, figure (mapped to chart)
        # Note: pdftable doesn't detect figure_title separately
        gt = load_ground_truth()
        for pdf_name, pdf_result in tool_results.items():
            if not isinstance(pdf_result, dict) or pdf_result.get("status") != "ok":
                continue

            # Get expected totals for this PDF
            if pdf_name in gt.get("totals", {}):
                exp_totals = gt["totals"][pdf_name]
                accuracy["table"]["expected"] += exp_totals.get("table", 0)
                accuracy["chart"]["expected"] += exp_totals.get("chart", 0)
                accuracy["figure_title"]["expected"] += exp_totals.get("figure_title", 0)

            # Get detected counts
            det_tables = pdf_result.get("tables_detected", 0)
            det_figures = pdf_result.get("figures_detected", 0)

            accuracy["table"]["detected"] += det_tables
            accuracy["chart"]["detected"] += det_figures  # figure → chart
            # figure_title not detected by pdftable

            # Calculate correct (min of expected and detected per PDF)
            if pdf_name in gt.get("totals", {}):
                exp_totals = gt["totals"][pdf_name]
                accuracy["table"]["correct"] += min(exp_totals.get("table", 0), det_tables)
                accuracy["chart"]["correct"] += min(exp_totals.get("chart", 0), det_figures)

    # Calculate percentages
    for obj_type in accuracy:
        exp = accuracy[obj_type]["expected"]
        det = accuracy[obj_type]["detected"]
        if exp > 0:
            accuracy[obj_type]["recall"] = round(accuracy[obj_type]["correct"] / exp * 100, 1)
        else:
            accuracy[obj_type]["recall"] = None
        if det > 0:
            accuracy[obj_type]["precision"] = round(accuracy[obj_type]["correct"] / det * 100, 1)
        else:
            accuracy[obj_type]["precision"] = None

    return accuracy


def print_accuracy_table(all_accuracy):
    """Print accuracy comparison table."""
    print("\n" + "="*80)
    print("ACCURACY BY OBJECT TYPE")
    print("="*80)
    print(f"{'Tool':<15} {'Tables':<20} {'Charts':<20} {'Fig Titles':<20}")
    print(f"{'':15} {'E/D/Rec%':<20} {'E/D/Rec%':<20} {'E/D/Rec%':<20}")
    print("-"*80)

    for tool_name, acc in all_accuracy.items():
        if acc is None:
            print(f"{tool_name:<15} SKIPPED")
            continue

        cols = []
        for obj_type in ["table", "chart", "figure_title"]:
            e = acc[obj_type]["expected"]
            d = acc[obj_type]["detected"]
            r = acc[obj_type]["recall"]
            r_str = f"{r}%" if r is not None else "N/A"
            cols.append(f"{e}/{d}/{r_str}")

        print(f"{tool_name:<15} {cols[0]:<20} {cols[1]:<20} {cols[2]:<20}")

    print("="*80)
    print("E=Expected, D=Detected, Rec%=Recall (correct/expected)")


# ============================================================
# MAIN
# ============================================================
def main():
    print("="*70)
    print("OCR TOOL COMPARISON - Accuracy by Object Type")
    print(f"Running: {TOOLS_TO_RUN}")
    print("="*70)

    all_results = {}
    all_accuracy = {}

    for tool in TOOLS_TO_RUN:
        if tool == "paddleocr":
            results = run_paddleocr()
        elif tool == "marker":
            results = run_marker()
        elif tool == "mistral_ocr":
            results = run_mistral_ocr()
        elif tool == "pdftable":
            results = run_pdftable()
        else:
            print(f"Unknown tool: {tool}")
            continue

        all_results[tool] = results
        all_accuracy[tool] = calculate_accuracy(results, tool)

    # Save results
    with open(BASE_DIR / "ocr_comparison_results.json", "w") as f:
        json.dump(all_results, f, indent=2, default=str)

    # Print accuracy table
    print_accuracy_table(all_accuracy)


if __name__ == "__main__":
    main()
