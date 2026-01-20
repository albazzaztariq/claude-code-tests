"""
FabricETL - Extract-Transform-Load Pipeline for Textile Research

Main orchestrator that coordinates:
1. EXTRACT: ScholarSweep v01-1 - API search, paper metadata, PDF downloads
2. TRANSFORM: PaddleOCR Layout + PyMuPDF - Text extraction from PDFs (v01-6)
3. LOAD: Final Excel output with all extracted data

USAGE:
    python FabricETL.py

    Or import as module:
        from FabricETL import run_pipeline, extract_papers, transform_pdfs

DEPENDENCIES:
    - ScholarSweep v01-1 (API search & download)
    - text_pipeline.py (PaddleOCR Layout + PyMuPDF native text extraction with bbox caching)
"""

import sys
import subprocess
from pathlib import Path
from datetime import datetime

# Paths to pipeline components (PROD structure)
SCHOLARSWEEP_PATH = Path(__file__).parent / "Dependencies" / "ScholarSweep" / "ScholarSweep.py"
TEXTILEVISION_PATH = Path(__file__).parent / "Dependencies" / "TextOCR" / "text_pipeline.py"

# Check availability
SCHOLAR_SWEEP_AVAILABLE = SCHOLARSWEEP_PATH.exists()
TEXTILE_VISION_AVAILABLE = TEXTILEVISION_PATH.exists()

if not SCHOLAR_SWEEP_AVAILABLE:
    print(f"WARNING: ScholarSweep not found at {SCHOLARSWEEP_PATH}")
if not TEXTILE_VISION_AVAILABLE:
    print(f"WARNING: TextileVision not found at {TEXTILEVISION_PATH}")

# Import TextileVision module for direct function calls
TEXTILEVISION_OCR = None
if TEXTILE_VISION_AVAILABLE:
    try:
        sys.path.insert(0, str(TEXTILEVISION_PATH.parent))
        from text_pipeline import extract_text_from_pdf
        TEXTILEVISION_OCR = extract_text_from_pdf
    except ImportError as e:
        print(f"WARNING: Could not import TextileVision OCR functions: {e}")


# =============================================================================
# PUBLIC API
# =============================================================================

__all__ = [
    "run_pipeline",
    "extract_papers",
    "transform_pdfs",
]


def extract_papers(config_file: Path = None) -> tuple:
    """
    EXTRACT phase: Run ScholarSweep v01-1 to search APIs and download PDFs.

    Args:
        config_file: Path to query.txt config file (default: ScholarSweep's default)

    Returns:
        Tuple of (excel_path, pdf_folder, query_folder)
    """
    if not SCHOLAR_SWEEP_AVAILABLE:
        raise ImportError(f"ScholarSweep not found at {SCHOLARSWEEP_PATH}")

    print("=" * 70)
    print("FABRICETL - EXTRACT PHASE (ScholarSweep v01-1)")
    print("=" * 70)

    # Run ScholarSweep as subprocess
    cmd = [sys.executable, str(SCHOLARSWEEP_PATH)]
    if config_file:
        cmd.append(str(config_file))

    print(f"\nRunning: {' '.join(cmd)}\n")

    result = subprocess.run(cmd, capture_output=False, text=True)

    if result.returncode != 0:
        raise RuntimeError(f"ScholarSweep failed with code {result.returncode}")

    # Find the most recent Output folder (at PROD level)
    output_dir = Path(__file__).parent.parent / "Output"
    query_folders = sorted(output_dir.glob("*Query"), key=lambda p: p.stat().st_mtime, reverse=True)

    if not query_folders:
        raise FileNotFoundError("No query output folders found")

    query_folder = query_folders[0]
    excel_files = list(query_folder.glob("*.xlsx"))
    pdf_folder = query_folder / "Downloaded Papers"

    excel_path = excel_files[0] if excel_files else None

    print(f"\n{'='*70}")
    print("EXTRACT COMPLETE")
    print(f"  Query folder: {query_folder}")
    print(f"  Excel: {excel_path}")
    print(f"  PDFs: {pdf_folder}")
    print(f"{'='*70}")

    return excel_path, pdf_folder, query_folder


def transform_pdfs(pdf_folder: Path) -> Path:
    """
    TRANSFORM phase: Extract text from PDFs using PaddleOCR Layout + PyMuPDF + Metric Filtering.

    Processes all PDFs and creates TWO text files per PDF:
    - <pdf_num>.txt - Full extracted text (clean, no tables/figures)
    - <pdf_num>_filtered.txt - Metric sentences only (ready for LLM)

    Also creates table bbox JSON files for table extraction pipeline.

    Args:
        pdf_folder: Folder containing downloaded PDFs (OA Papers/ and Non-OA Papers/)

    Returns:
        Path to the folder with extracted text files
    """
    if not TEXTILE_VISION_AVAILABLE or not TEXTILEVISION_OCR:
        raise ImportError(f"TextileVision not found at {TEXTILEVISION_PATH}")

    pdf_folder = Path(pdf_folder)

    print("=" * 70)
    print("FABRICETL - TRANSFORM PHASE")
    print("PaddleOCR Layout + PyMuPDF Native Text + Metric Filtering")
    print("=" * 70)

    # Get all PDFs from both OA and Non-OA subfolders
    pdf_files = []
    for subfolder in ["OA Papers", "Non-OA Papers"]:
        subfolder_path = pdf_folder / subfolder
        if subfolder_path.exists():
            pdf_files.extend(list(subfolder_path.glob("*.pdf")))

    print(f"\n  Found {len(pdf_files)} PDFs in {pdf_folder}")

    if not pdf_files:
        print("  No PDFs found. Skipping transform phase.")
        return None

    # TRANSFORM: Process each PDF with PaddleOCR Layout + PyMuPDF + Filtering
    print(f"\n  Processing PDFs with PaddleOCR Layout + PyMuPDF + Metric Filtering...\n")

    for i, pdf_path in enumerate(pdf_files, 1):
        print(f"[{i}/{len(pdf_files)}] Processing {pdf_path.name}...")

        try:
            # Extract text using PaddleOCR Layout + PyMuPDF (also creates _filtered.txt and bbox JSON)
            extracted_text = TEXTILEVISION_OCR(pdf_path)

            # The OCR stack now creates THREE outputs per PDF:
            # - <pdf_num>.txt (full text) - created by OCR stack
            # - <pdf_num>_filtered.txt (metric sentences) - created by OCR stack
            # - <pdf_num>_table_bboxes.json (table bboxes for table extraction) - created by OCR stack

            txt_path = pdf_path.with_suffix('.txt')
            filtered_path = pdf_path.with_suffix('').with_suffix('').with_name(f"{pdf_path.stem}_filtered.txt")

            print(f"  Full text:     {txt_path.name}")
            print(f"  Filtered text: {filtered_path.name}")
            print(f"  Both files ready for LLM processing\n")

        except Exception as e:
            print(f"  Error: {e}\n")

    print(f"\n{'='*70}")
    print("TRANSFORM COMPLETE")
    print(f"  Output files in: {pdf_folder}")
    print(f"  - Full text files: <pdf_num>.txt")
    print(f"  - Filtered files: <pdf_num>_filtered.txt")
    print(f"  - Table bbox JSON: <pdf_num>_table_bboxes.json")
    print(f"\n  NOTE: Filtered files are ready for LLM API (not yet implemented)")
    print(f"{'='*70}")

    return pdf_folder


def run_pipeline(config_file: Path = None) -> dict:
    """
    Run complete ETL pipeline: Extract → Transform.

    Args:
        config_file: Path to ScholarSweep query.txt config file

    Returns:
        Dict with paths to all outputs:
        {
            "metadata_excel": Path,
            "pdf_folder": Path,
            "query_folder": Path,
            "text_files_folder": Path,
        }
    """
    print("\n" + "=" * 70)
    print("FABRICETL - COMPLETE PIPELINE")
    print("=" * 70)

    results = {}

    # EXTRACT: ScholarSweep v01-1 (API search + downloads)
    excel_path, pdf_folder, query_folder = extract_papers(config_file)
    results["metadata_excel"] = excel_path
    results["pdf_folder"] = pdf_folder
    results["query_folder"] = query_folder

    # TRANSFORM: PaddleOCR Layout + PyMuPDF (text extraction)
    text_folder = transform_pdfs(pdf_folder)
    results["text_files_folder"] = text_folder

    # Summary
    print("\n" + "=" * 70)
    print("PIPELINE COMPLETE")
    print("=" * 70)
    print(f"  Metadata Excel:   {results['metadata_excel']}")
    print(f"  PDF Folder:       {results['pdf_folder']}")
    print(f"  Text files:       {results['text_files_folder']}")
    print("=" * 70)

    return results


# =============================================================================
# CLI INTERFACE
# =============================================================================

def main():
    """Interactive CLI for FabricETL pipeline."""
    print("=" * 70)
    print("FABRICETL - Textile Research ETL Pipeline")
    print("ScholarSweep v01-1 -> PaddleOCR Layout + PyMuPDF (v01-6)")
    print("=" * 70)

    # Check available components
    print("\nComponent Status:")
    print(f"  ScholarSweep:              {'✓ Available' if SCHOLAR_SWEEP_AVAILABLE else '✗ Not available'}")
    print(f"  PaddleOCR+PyMuPDF (v01-6): {'✓ Available' if TEXTILE_VISION_AVAILABLE else '✗ Not available'}")

    if not SCHOLAR_SWEEP_AVAILABLE and not TEXTILE_VISION_AVAILABLE:
        print("\nERROR: No pipeline components available.")
        return

    print("\nOptions:")
    print("  1: Run full pipeline (Extract → Transform)")
    print("  2: Extract only (ScholarSweep: API search + download)")
    print("  3: Transform only (PaddleOCR+PyMuPDF on existing PDFs)")
    print("  4: Exit")

    choice = input("\nSelect option (1-4): ").strip()

    if choice == "1":
        # Full pipeline
        config_file = input("\nPath to query.txt config file (Enter for default): ").strip()
        config_file = Path(config_file) if config_file else None

        run_pipeline(config_file)

    elif choice == "2":
        # Extract only
        if not SCHOLAR_SWEEP_AVAILABLE:
            print("ScholarSweep not available.")
            return

        config_file = input("\nPath to query.txt config file (Enter for default): ").strip()
        config_file = Path(config_file) if config_file else None

        extract_papers(config_file)

    elif choice == "3":
        # Transform only
        if not TEXTILE_VISION_AVAILABLE:
            print("PaddleOCR+PyMuPDF OCR not available.")
            return

        pdf_folder = input("\nEnter PDF folder path: ").strip()
        if not pdf_folder or not Path(pdf_folder).exists():
            print("Invalid folder path. Exiting.")
            return

        transform_pdfs(pdf_folder=Path(pdf_folder))

    elif choice == "4":
        print("Exiting.")
        return

    else:
        print("Invalid option. Exiting.")


if __name__ == "__main__":
    main()
