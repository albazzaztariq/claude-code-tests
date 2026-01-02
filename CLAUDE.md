# Project: Claude Code Tests

> **CLAUDE INSTRUCTIONS:** On session startup, read this file to understand current work status. Throughout the session, update the "Current Work" section below with major task progress. Keep updates high-level (don't log every edit, just major milestones and status changes).

---

## CRITICAL: Git Rules (DO NOT VIOLATE)

**NEVER run any git command in this working folder.** No `git init`, `git checkout`, `git restore`, `git pull`, or ANY git operation here. This folder must remain git-free.

**For GitHub backups:**
1. The user's GitHub repo is: `https://github.com/albazzaztariq/FabricETL`
2. Clone/use a SEPARATE folder for git operations (e.g., `C:\Users\azt12\claude-code-tests\`)
3. When pushing updates: COPY files from this working folder TO the repo folder, then commit/push from there
4. Periodically push updated files when significant changes are made (after major edits, end of session, or when user asks)

**Why:** Git previously overwrote local files via checkout. User's working files must NEVER be touched by git operations.

---

## Current Work (Updated: 2026-01-02 ~7:00AM)

### Git Workflow
- **Working folder:** `C:\Users\azt12\OneDrive\Documents\Wrestling Robe\Materials Science - Wickability\Studies for Analysis by LLM AI`
- **Git repo folder:** `C:\Users\azt12\OneDrive\Documents\Git Wrestling Robe`
- **Sync script:** `Sync-ToGit.ps1` - copies files from working folder to git repo, commits, pushes
- **NEVER run git in working folder** - use the separate git repo folder

### File Rename History
| Original | Renamed To | Purpose |
|----------|-----------|---------|
| `MasterPaperScrape.py` | `ScholarSweep.py` | API paper search (EXTRACT) |
| `SciMasterScrape#_test.py` | `TextileVision_v01.#.py` | Fabric metric extraction (TRANSFORM) |
| `Scrape#_test.py` | *(kept as-is)* | Legacy scripts with reusable OCR code |
| `compare_extractors.py` | `TableBench_NougatVsPyMuPDF.py` | Table extraction benchmark |
| `ChartOCRTester.py` | `ChartF2TBench.py` | Chart detection benchmark |
| `OCR Testing.xlsx` | `F2T Testing.xlsx` | F2T = Figure-to-Text |
| `PDF OCR Object Counts.xlsx` | `PDF Element Counts.xlsx` | Element counts per PDF |

### DIA Branch Structure (2026-01-02)
```
Document-Image-Analysis/
├── Figure-to-Text Extraction/
│   ├── Supporting Work/
│   │   ├── ChartF2TBench.py
│   │   ├── F2TBench.py
│   │   ├── F2T Testing.xlsx
│   │   └── PDF Element Counts.xlsx
│   ├── TableBench_NougatVsPyMuPDF.py
│   ├── TextileVision_v01.py
│   ├── TextileVision_v02.py
│   └── TextileVision_v03.py
├── OCR/
│   ├── TextileParse_v01.py
│   ├── TextileParse_v02.py
│   └── TextileParse_v03.py
├── Supporting Work/
│   ├── Scrape#_test README.txt   <- package dependencies
│   ├── Scrape1_test.py           <- pdfplumber, pandas
│   ├── Scrape2_test.py           <- PyMuPDF, camelot, tabula
│   ├── Scrape3_test.py           <- Google Cloud Vision
│   └── Scrape4_test.py           <- Azure Doc Intelligence, Nougat
└── TextileVision_INSTALL.md
```

### FabricETL Pipeline Architecture
```
FabricETL.py (orchestrator)
    │
    ├── ScholarSweep.py (EXTRACT)
    │   └── API search → filter → download → metadata.xlsx
    │
    └── TextileVision_v01.3.py (TRANSFORM)
        └── OCR/F2T → fabric_metrics.xlsx
```

**Two Excel outputs:**
- `{timestamp} Paper API Query.xlsx` - paper metadata (ScholarSweep)
- `{folder}_metrics.xlsx` - extracted fabric properties (TextileVision)

### GitHub Repo Structure
- **Repo:** https://github.com/albazzaztariq/FabricETL
- **Branches:**
  - `Main` - Documentation only (CLAUDE.md)
  - `Document-Image-Analysis` - OCR/vision tools (TextileVision, F2TBench)
  - `Research-Corpus-Generation` - ScholarSweep.py

### File Locations
- **Working folder:** `C:\Users\azt12\OneDrive\Documents\Wrestling Robe\Materials Science - Wickability\Studies for Analysis by LLM AI`
- **Test PDFs:** `1.pdf` through `57.pdf` in working folder (ground truth covers 1.pdf, 2.pdf, 3.pdf)
- **pdftable repo:** `pdf_table_repo/` subfolder (installed via `python setup.py install`)

### Active Tasks

1. **ChartOCRTester.py** - PDF layout detection validation
   - Status: SKIP_OCR mode enabled for fast testing
   - Detection settings: 900 DPI (3x), multi-pass thresholds 0.4 + 0.25
   - Results:
     - 1.pdf: 8/9 charts
     - 2.pdf: 14/14 charts perfect!
     - 3.pdf: 8/8 tables ✓, 2/2 charts ✓, 12/10 fig_titles (2 extra without OCR filter)
   - Outputs ASCII table with E/A (Expected vs Actual) columns

2. **ScholarSweep.py** (was MasterPaperScrape.py) - Multi-API academic paper search tool
   - Status: **THRESHOLDS DEFINED** (2026-01-01)
   - Location: `Research Corpus Generation/ScholarSweep.py` in GitHub repo
   - **528 TEXTILE JOURNALS EMBEDDED** with acronym lookup (TEXTILE_JOURNALS dict)
   - **PROCESSING THRESHOLDS (constants for future use):**
     - `THRESHOLD_END_BATCH = 50,000` - below: hold all in RAM
     - `THRESHOLD_CHUNKED = 500,000` - below: chunked processing
     - Above 500K: streaming mode (not yet implemented)
     - `decide_processing_mode(count)` function ready for future use
   - **JOURNAL FILTER FUNCTIONS:**
     - `interactive_journal_selection()` - User prompt for journal input
     - `resolve_journal_input()` - Matches name, acronym, or partial
     - CrossRef/OpenAlex: server-side filtering
     - PubMed: client-side filtering (textile journals not in PubMed)
   - **TERMINAL COLORS:** GREEN, DIM_RED, RESET constants added
   - **NEW INTERACTIVE QUERY BUILDER:**
     1. Select APIs by ID (1=OpenAlex, 2=PubMed, 3=CrossRef)
     2. For each API, select search regions (1=Title, 2=Abstract, 3=Full-text for OpenAlex)
     3. For each region, enter comma-separated search terms (multi-word = exact phrase)
   - **PARALLEL EXECUTION:** APIs now run simultaneously using ThreadPoolExecutor
   - **OPTIMIZED BATCH SIZES:**
     - CrossRef: 1000 (was 500) - max allowed
     - OpenAlex: 200 - max allowed
     - PubMed: 10000 (was 200) - max allowed
   - **API-SPECIFIC NOTES:** CrossRef prompts now show "best match" warning
   - **TITLE-SPECIFIC SEARCH:** Now supports title-only, abstract-only, or both per API
   - **DIRECTORY STRUCTURE:**
     ```
     ScrapedResearch/
     └── MM-DD-YY-HHmm Query/     (e.g., "01-01-26-0439 Query")
         ├── Query Parameters/
         ├── Downloaded Papers/
         │   ├── OA Papers/
         │   └── Non-OA Papers/
         ├── 01-01-26-0439 Paper API Query.xlsx
         ├── 01-01-26-0439 BACKUP.json
         └── 01-01-26-0439 Full-Text Search.xlsx
     ```
   - **OUTPUT FORMAT:** Excel only (no CSV) - openpyxl for colored row status (green=downloaded, red=searched)
   - **DEPENDENCIES:** `requests`, `openpyxl`, `pdfplumber`
   - **MODULE API (2026-01-02):** Refactored for import as module:
     ```python
     from ScholarSweep import api_search, filter_and_save, download_papers, search_fulltext
     papers = api_search("textile wicking", apis=["openalex"], max_results=500)
     excel_path, oa, non_oa = filter_and_save(papers)
     pdf_folder = download_papers(excel_path, max_downloads=100)
     results = search_fulltext(pdf_folder, excel_path, ["absorption", "wicking"])
     ```

8. **Textile/Materials Journal List** - COMPLETED
   - File: `List of Journals.xlsx` (column B) - user moved file location
   - **452 unique journals** collected from OpenAlex + CrossRef
   - Search terms: textile, fabric, fiber, fibre, apparel, woven, knit, cloth, garment, clothing, **materials science**
   - Purpose: Filter future paper searches by journal name server-side

9. **FabricETL.py** - NEW (2026-01-02)
   - Main orchestrator for Extract-Transform-Load pipeline
   - Imports ScholarSweep (EXTRACT) and TextileVision (TRANSFORM)
   - CLI interface with 4 options: full pipeline, extract only, transform only, exit

10. **TextileVision** - Renamed from SciMasterScrape (2026-01-02)
    - `TextileVision_v01.1.py` - (was SciMasterScrape1_test.py)
    - `TextileVision_v01.2.py` - (was SciMasterScrape2_test.py)
    - `TextileVision_v01.3.py` - (was SciMasterScrape3_test.py) **← MAIN VERSION**
    - Uses Nougat OCR + GPT-4 for structured fabric data extraction
    - **MODULE API:**
      ```python
      from TextileVision_v01_3 import extract_metrics, process_pdf_batch
      data, df = extract_metrics("paper.pdf")
      ```

---

## API Journal Filtering Capabilities

| API | Filter Method | Exact Match? | Multiple Journals? | Notes |
|-----|---------------|--------------|-------------------|-------|
| **CrossRef** | `container-title:Name` | YES | YES (comma) | Direct name, no lookup |
| **OpenAlex** | `source.id:ID` | YES (by ID) | YES (pipe `\|`) | Requires name→ID lookup |
| **PubMed** | `"Name"[Journal]` | YES | YES (OR) | But textile journals not indexed |

**CrossRef example:**
```
filter=container-title:Textile Research Journal,container-title:Journal of the Textile Institute
```

**OpenAlex example:**
```
filter=primary_location.source.id:S70817854|S105605216
```

**PubMed note:** Supports exact journal filtering, but dedicated textile journals aren't in PubMed. Textile papers exist (16,221) but in general science journals (Nature Comms, ACS Applied Materials, etc.)

3. **PaddleOCR** - PDF document structure extraction
   - Ground truth: `OCR_GroundTruth.json` (machine-readable counts per page)
   - Also: `OCR Testing.xlsx` (sheets 1-3 have page-by-page counts)
   - Uses `device="gpu"` for GPU acceleration
   - PaddleOCR version: 3.3.2 installed

4. **F2TBench.py** (was OCR_Comparison.py) - Figure-to-Text Benchmark
   - Location: `Document Image Analysis/Figure-to-Text Extraction/Supporting Work/`
   - Compares layout detection accuracy by object type (tables, charts, fig_titles)
   - Tools configured (uncomment in TOOLS_TO_RUN to enable):
     - `paddleocr` - baseline, 100% tables/charts on 3.pdf
     - `marker` - PDF→markdown with Ollama qwen2.5:14b
     - `mistral_ocr` - cloud API (needs MISTRAL_API_KEY)
     - `pdftable` - deep learning, NOW INSTALLED with CUDA/cuDNN
   - Ground truth: `OCR_GroundTruth.json`
   - Outputs accuracy table: Expected/Detected/Recall% per object type
   - NOT layout tools (removed): LaTeX-OCR (equations only), Camelot (tables only)
   - MassivePix is commercial (web service only), NeMo needs NVIDIA NIM containers

5. **pdftable** - TESTED, NOT SUITABLE
   - Installed from: `pdf_table_repo/` via `python setup.py install`
   - CUDA 12.4 + cuDNN, Ghostscript 10.03.0 installed
   - Fixed bugs: Windows path (file_utils.py:466), gswin64c (ghostscript_backend.py:76)
   - **PROBLEM:** Layout model (picodet/publaynet) detects whole figure regions, NOT individual charts
   - publaynet categories: text, title, list, table, figure (no "chart" category)
   - Detects one giant bbox for multi-chart panels instead of separate charts
   - **VERDICT:** Wrong tool - designed for document structure, not chart detection

6. **detectron2** - INSTALLED (WSL2 Ubuntu)
   - Windows native install failed (C++ compilation issues)
   - **Successfully installed in WSL2 Ubuntu** with GPU access
   - Location: `~/detectron2_env/` virtual environment in WSL
   - To use: `wsl -d Ubuntu -e bash -c "source ~/detectron2_env/bin/activate && python ..."`
   - detectron2 v0.6, PyTorch 2.6.0+cu124, RTX 4070 GPU working

7. **ChartDete/CACHED** - FULLY INSTALLED & WORKING (WSL2 Ubuntu)
   - ICDAR 2023 winner for chart element detection
   - **18 chart element classes** (bars, axes, labels, legends, data points, etc.)
   - Location: `~/ChartDete/` in WSL2
   - Conda env: `ChartDete` (Python 3.8, PyTorch 1.13.1, CUDA 11.7)
   - MMDetection: mmcv-full 1.7.2, mmdet 2.28.1
   - Models: `~/ChartDete/work_dirs/` (5 model variants)
   - **TESTED: Detected 200 chart elements on sample.jpg**

   **Available models:**
   - `cascade_rcnn_swin-t_fpn_LGF_VCE_PCE_coco_focalsmoothloss` (CACHED - best)
   - `cascade_rcnn_swin-t-p4-w7_fpn_1x_coco`
   - `cascade_rcnn_x101_64x4d_fpn_20e_coco`
   - `faster_rcnn_x101_64x4d_fpn_2x_coco`
   - `detr_finetune`

   **To run inference from Windows:**
   ```bash
   wsl -d Ubuntu -e bash -c "source ~/miniconda3/bin/activate && conda activate ChartDete && export LD_LIBRARY_PATH=/usr/lib/wsl/lib:\$LD_LIBRARY_PATH && cd ~/ChartDete && python << 'EOF'
   from mmdet.apis import init_detector, inference_detector
   config = './work_dirs/cascade_rcnn_swin-t_fpn_LGF_VCE_PCE_coco_focalsmoothloss/cascade_rcnn_swin-t_fpn_LGF_VCE_PCE_coco_focalsmoothloss.py'
   checkpoint = './work_dirs/cascade_rcnn_swin-t_fpn_LGF_VCE_PCE_coco_focalsmoothloss/checkpoint.pth'
   model = init_detector(config, checkpoint, device='cuda:0')
   result = inference_detector(model, '/mnt/c/path/to/image.jpg')
   model.show_result('/mnt/c/path/to/image.jpg', result, out_file='/mnt/c/path/to/result.jpg', score_thr=0.3)
   EOF"
   ```

   **IMPORTANT:** Must set `export LD_LIBRARY_PATH=/usr/lib/wsl/lib:$LD_LIBRARY_PATH` for CUDA to work

---

## Chart Detection Model Options

**Goal:** Find a model that detects individual charts (not just "figure" regions)

| Model | Framework | Chart Detection | Pretrained | Notes |
|-------|-----------|-----------------|------------|-------|
| **ChartDete/CACHED** | MMDetection | ✅ 18 chart element classes | [Google Drive](https://drive.google.com/file/d/1n9UtHgfOA6H8cxp4Y44fG7OdXbVJzMnJ) | ICDAR 2023 winner, best option |
| **Detectron2_DocLayNet** | Detectron2 | ⚠️ "Picture" class only | [HuggingFace](https://huggingface.co/AlexShmak/Detectron2_DocLayNet) | 11 classes, not chart-specific |
| **DocLayout-YOLO** | YOLO | ⚠️ "Figure" class only | [arXiv](https://arxiv.org/html/2410.12628v1) | 79.7% mAP, 85 FPS, fastest |
| **detectron2-publaynet** | Detectron2 | ⚠️ "Figure" class only | [GitHub](https://github.com/JPLeoRX/detectron2-publaynet) | Same as pdftable |
| **ChartEye** | YOLOv7 | ✅ 15 chart types | ❌ Not released | F1=0.97 classification |
| **CHART-Info dataset** | Any | ✅ Classification | [chartinfo.github.io](https://chartinfo.github.io/) | PMC dataset, no pretrained |

**Recommendation:** ChartDete/CACHED - uses MMDetection but has 18 chart element classes with pretrained weights

**User-noted models (not yet tested):**
- **ChartVLM-L/B** - Best for charts (per user)
- **Table-LLaVA** - Best for tables (per user)

**ChartDete Install (WSL2):**
```bash
conda create -n ChartDete python=3.8
conda activate ChartDete
conda install pytorch==1.13.1 torchvision==0.14.1 pytorch-cuda=11.7 -c pytorch -c nvidia
pip install -U openmim
mim install mmcv-full
git clone https://github.com/pengyu965/ChartDete
cd ChartDete && pip install -v -e .
```

---

## OCR Runtime Benchmark

**IMPORTANT: Benchmark uses 1.pdf PAGES 1 AND 7 ONLY**

Results file: `ocr_runtime_results.csv`

| Tool | GPU | Page 1 (tbl/chart/fig) | Page 7 (tbl/chart/fig) | Runtime |
|------|-----|------------------------|------------------------|---------|
| **paddleocr** | YES | 0/0/0 | 0/3/3 | **0.6s** |
| mistral_ocr | CLOUD | 0/0/0 | 0/2/3 | 3.8s |
| pdftable | YES | 0/1/0 | 0/1/1 | ~9s* |
| marker | YES + Ollama | 0/0/1 | 0/0/1 | 98.4s |

*pdftable ~9s after model cached (first run 14.8s includes model loading)

**Winner: paddleocr** - Fastest AND most accurate for layout detection.

**ChartDete** (separate tool - different purpose): Detects 18 chart ELEMENT classes (axes, labels, value_label, plot_area, etc.) with bounding boxes. Use for chart VALUE EXTRACTION, not region counting. Gives element locations but requires OCR to read actual text/numbers. Runtime: 2.27s. Visualizations: `C:/temp/ocr_test/page*_chartdete_result.png`.

**To run pdftable benchmark:**
1. Install Ghostscript: https://www.ghostscript.com/releases/gsdnld.html
2. Copy 1.pdf to `C:/temp/ocr_test/` (no spaces in path)
3. Run: `pdftable --output_dir "C:/temp/ocr_test/output" --file_path_or_url "C:/temp/ocr_test/1.pdf" --pages "1,7" --lang en`

---

## Session Continuity

This project uses session summaries to maintain context across Claude CLI sessions.

### How it works:
1. Before ending a session, run `/save-session` to save a summary
2. The summary is written to `SESSION_HISTORY.md` in the project root
3. On new sessions, Claude reads this file to understand previous work

### Session History Location:
- **File:** `SESSION_HISTORY.md` (project root)
- **Command:** `/save-session`

## Important Files
- `.claude/settings.local.json` - Local Claude settings with BurntToast notification hooks
- `.claude/commands/save-session.md` - Session save command
- `MasterPaperScrape.py` - Multi-API paper search (CrossRef, PubMed, OpenAlex)

## Notes
- BurntToast notifications configured for Stop, idle_prompt, and permission_prompt hooks
- **FIXED:** Moved C# foreground-check code to external script  to fix here-string JSON escaping error
- No built-in command exists to refresh settings mid-session; restart Claude CLI to apply changes

## Technical Concepts Discussed (2026-01-01)
- **Syscalls:** User mode to kernel mode transitions have fixed overhead per call
- **Disk I/O batching:** Fewer large writes faster than many small writes (transaction cost, not context-switching)
- **CSV vs RAM:** Variable-length rows prevent direct byte addressing; RAM uses memory addresses for O(1) lookup
- **Streaming vs End-batch:** Both just append to files; difference is write frequency (buffer size)
