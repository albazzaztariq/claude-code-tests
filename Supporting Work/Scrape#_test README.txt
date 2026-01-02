SUPPORTING WORK - PACKAGE REFERENCE
====================================
Lists which packages/libraries are used in each legacy Scrape script.
Use this to identify reusable code for specific OCR/extraction packages.


Scrape1_test.py
---------------
  Standard Library:
    - os
    - re
    - json

  Third-Party:
    - requests        (HTTP requests to Ollama)
    - pandas          (DataFrame operations)
    - pdfplumber      (PDF text/table extraction)

  Services:
    - Ollama (local LLM - gemma3:1b)


Scrape2_test.py
---------------
  Standard Library:
    - os
    - re
    - sys
    - pathlib

  Third-Party:
    - fitz (PyMuPDF)  (PDF parsing, font detection)
    - requests        (HTTP requests to Ollama)
    - openpyxl        (Excel read/write)
    - pdfplumber      (PDF text/table extraction)
    - camelot         (PDF table extraction - lattice/stream)
    - tabula          (PDF table extraction - Java-based)
    - pytest          (testing framework)

  Services:
    - Ollama (local LLM - gemma2:2b)


Scrape3_test.py
---------------
  Standard Library:
    - os
    - re
    - sys
    - pathlib
    - io

  Third-Party:
    - fitz (PyMuPDF)  (PDF parsing, font detection)
    - requests        (HTTP requests to Ollama)
    - openpyxl        (Excel read/write)
    - pdfplumber      (PDF text/table extraction)
    - camelot         (PDF table extraction)
    - pytest          (testing framework)
    - google.cloud.vision (Google Cloud Vision API)

  Services:
    - Ollama (local LLM - gemma2:2b)
    - Google Cloud Vision API


Scrape4_test.py
---------------
  Standard Library:
    - os
    - re
    - sys
    - pathlib
    - io

  Third-Party:
    - fitz (PyMuPDF)  (PDF parsing, font detection)
    - requests        (HTTP requests to Ollama)
    - openpyxl        (Excel read/write)
    - pdfplumber      (PDF text/table extraction)
    - camelot         (PDF table extraction)
    - pytest          (testing framework)
    - PIL (Pillow)    (image processing)
    - azure.ai.formrecognizer (Azure Document Intelligence)
    - azure.core.credentials  (Azure authentication)
    - albumentations  (image augmentation - for Nougat compat)
    - nougat          (Neural OCR for academic documents)

  Services:
    - Ollama (local LLM - gemma2:2b)
    - Azure Document Intelligence API
    - Nougat OCR (local neural model)


INSTALL COMMANDS
================
# Core packages (all scripts)
pip install requests pdfplumber openpyxl PyMuPDF

# Scrape1 only
pip install pandas

# Scrape2 additions
pip install camelot-py[cv] tabula-py pytest

# Scrape3 additions
pip install google-cloud-vision

# Scrape4 additions
pip install azure-ai-formrecognizer Pillow nougat-ocr albumentations
