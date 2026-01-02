"""
Compare PyMuPDF vs Nougat extraction quality
Extracts first 8 pages from 1.pdf using PyMuPDF for comparison
"""

import fitz  # PyMuPDF
from pathlib import Path

def extract_with_pymupdf(pdf_path: str, max_pages: int = 8):
    """Extract text using PyMuPDF (simple text extraction)"""

    print(f"\n{'='*60}")
    print(f"PyMuPDF Extraction - {Path(pdf_path).name}")
    print(f"{'='*60}\n")

    try:
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        pages_to_process = min(max_pages, total_pages)

        print(f"Total pages: {total_pages}")
        print(f"Processing: First {pages_to_process} pages\n")

        all_text = []

        for page_num in range(pages_to_process):
            print(f"Extracting page {page_num + 1}/{pages_to_process}...", end=" ")
            page = doc[page_num]
            text = page.get_text()

            print(f"✓ {len(text)} characters")

            # Add page break marker
            all_text.append(f"<!-- PAGE {page_num + 1} -->")
            all_text.append(text)
            all_text.append("\n<!-- PAGE BREAK -->\n")

        doc.close()

        full_text = "\n".join(all_text)

        # Save to file
        output_path = Path(pdf_path).stem + "_pymupdf.txt"
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(full_text)

        print(f"\n✓ Saved PyMuPDF output to {output_path}")
        print(f"✓ Total characters extracted: {len(full_text)}")

        return full_text

    except Exception as e:
        print(f"\n✗ PyMuPDF extraction failed: {e}")
        return None


if __name__ == "__main__":
    # Extract from 1.pdf using PyMuPDF
    extract_with_pymupdf("1.pdf", max_pages=8)

    print("\n" + "="*60)
    print("COMPARISON READY")
    print("="*60)
    print("\nFiles created:")
    print("  • 1_pymupdf.txt  - PyMuPDF extraction (simple text)")
    print("  • 1_nougat.txt   - Nougat extraction (markdown structure)")
    print("\nYou can now compare:")
    print("  - Text quality and completeness")
    print("  - Table preservation")
    print("  - Structure and formatting")
    print("  - Mathematical equations")
