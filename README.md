# FabricETL Production

**Last Updated:** 2026-01-17

This folder contains the **production-ready** FabricETL pipeline with only the essential files needed to run.

## Folder Structure

```
PROD/
├── FabricETL/
│   ├── FabricETL.py                          # Main orchestrator
│   └── Dependencies/
│       ├── ScholarSweep/
│       │   ├── ScholarSweep.py               # EXTRACT: API search & download
│       │   ├── query.txt                     # Search configuration
│       │   └── JournalList.py                # Journal mappings (optional)
│       └── TextOCR/
│           ├── text_pipeline.py              # TRANSFORM: Text extraction (v01-6)
│           └── MetricsSearchable.txt         # Metric keywords for filtering
└── Output/                                    # Pipeline outputs (Excel, PDFs, text files)
```

## What's Included

### 1. FabricETL.py
- Main orchestrator that runs the full pipeline
- Coordinates EXTRACT and TRANSFORM phases
- CLI interface for user interaction

### 2. ScholarSweep v01-1 (EXTRACT)
- **ScholarSweep.py**: API search, metadata Excel, PDF downloads
- **query.txt**: Configuration file for search terms
- **JournalList.py**: Optional journal name mappings
- **Outputs:**
  - Excel file with OA and Non-OA sheets
  - Green rows: Downloaded successfully
  - Red rows: Download failed
  - PDFs in `OA Papers/` and `Non-OA Papers/` folders

### 3. Text Pipeline v01-6 (TRANSFORM)
- **text_pipeline.py**: PaddleOCR Layout + PyMuPDF native text extraction
- **MetricsSearchable.txt**: 501 searchable metric terms
- **Method:**
  - PaddleOCR: Detect tables/figures (what NOT to extract)
  - PyMuPDF: Extract native PDF text (NO OCR - 16x faster!)
  - Coordinate filtering: Remove text in banned regions
  - Metric filtering: Extract sentences with fabric metrics
- **Outputs per PDF:**
  - `<pdf_num>.txt`: Full body text (clean, no tables/figures)
  - `<pdf_num>_filtered.txt`: Metric sentences only (LLM-ready)
  - `<pdf_num>_table_bboxes.json`: Table coordinates for table extraction

## How to Run

```bash
cd C:\Users\azt12\OneDrive\Documents\Business\Textile\WrestlingRobe\Programming\FabricETL\PROD\FabricETL

python FabricETL.py
```

**Options:**
1. Run full pipeline (Extract → Transform)
2. Extract only (ScholarSweep)
3. Transform only (Text OCR on existing PDFs)
4. Exit

## Dependencies Required

**Python Packages:**
- ScholarSweep: `requests`, `openpyxl`
- Text Pipeline: `paddleocr`, `fitz` (PyMuPDF), `pillow`, `paddle`

## NOT Included (Not Hooked Up Yet)

- Table Extraction Pipeline (ready, tested, not integrated)
- LLM Processing (not implemented)

## Production Status

✅ **ScholarSweep → Text Pipeline** hooked up and ready to use
❌ Table Extraction - ready but not integrated
❌ LLM Processing - not implemented

**Output Location:** `PROD/Output/` (all query results, Excel files, PDFs, and text files)

See `High-Level-Design/pipeline_hookupStatus.md` for current status.
