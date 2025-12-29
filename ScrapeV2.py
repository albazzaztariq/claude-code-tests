#!/usr/bin/env python3
"""
ScrapeV2.py - Comprehensive Textile Specification Extractor

Uses Nougat OCR for table extraction and LLM (Ollama/OpenAI) for semantic
understanding to extract all yarn, fabric, and transfer metrics from PDFs.

Focused on 1.pdf for initial development.
"""

import os
import re
import json
import sys
from pathlib import Path
from typing import Optional, Dict, List, Any, Tuple
import fitz  # PyMuPDF
import requests
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ================== CONFIGURATION ==================
# LLM Configuration - Choose one
LLM_PROVIDER = os.getenv("LLM_PROVIDER", "ollama")  # "ollama" or "openai"

# Ollama settings
OLLAMA_URL = os.getenv("OLLAMA_URL", "http://localhost:11434/api/generate")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "llama3.1:8b")  # Recommend larger model for this task

# OpenAI settings (if using)
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4-turbo-preview")

# Paths
BASE_DIR = Path(
    r"C:\Users\azt12\OneDrive\Documents\Wrestling Robe\Materials Science - Wickability\Studies for Analysis by LLM AI\Datafiles & Python Scripts"
)
PDF_FOLDER = Path(
    r"C:\Users\azt12\OneDrive\Documents\Wrestling Robe\Materials Science - Wickability\Studies for Analysis by LLM AI"
)
OUTPUT_EXCEL = BASE_DIR / "Textile-Specifications-Extracted.xlsx"

# Focus on 1.pdf only for now
FOCUS_PDF = "1.pdf"

# ================== METRICS SCHEMA ==================
# This defines ALL the metrics we want to extract
# Each metric has: name, layer_types (which layer columns it applies to), category

LAYER_TYPES = ["Inner Layer", "Middle Layer", "Outer Layer", "Monolayer"]
SIDES = ["Bottom Side", "Top Side", ""]  # "" means general/both sides

# Study-level metadata (one per study)
STUDY_METADATA = [
    "Study Number",
    "Study Title",
    "Year of Publish",
    "Name of First-Listed Author",
    "Number of Sample Fabrics",
    "Testing/Standards Bodies Methods Used",
]

# Sample-level metadata (one per sample)
SAMPLE_METADATA = [
    "Sample ID/Name",
    "Number of Fabric Layers",
]

# Material, Structure, Treatment fields - per layer
MATERIAL_STRUCTURE_FIELDS = [
    {"name": "Material", "layers": True},
    {"name": "Structure", "layers": True, "note": "Knit or Woven or both"},
]

# Treatment fields - per layer and per side
TREATMENT_FIELDS = [
    {"name": "Treatment", "layers": True, "sides": True},
    {"name": "Change in Water Affinity", "layers": True, "sides": True,
     "note": "e.g. hydrophobic -> superhydrophobic"},
]

# Yarn specification fields - per layer
YARN_SPEC_FIELDS = [
    {"name": "Yarn Fiber Type", "layers": True, "note": "staple, filament, or composite"},
    {"name": "Filament Yarn Texture", "layers": True, "note": "Spun or Filament"},
    {"name": "Filament Yarn Structure", "layers": True, "note": "Textured or Flat (filament only)"},
    {"name": "Filament Yarn Cross-section Type", "layers": True, "note": "round, trilobal, etc."},
    {"name": "Yarn Denier/Linear Density", "layers": True},
    {"name": "Filament Count", "layers": True, "note": "number of filaments per yarn"},
    {"name": "Yarn Twist", "layers": True, "note": "direction and turns per inch/cm"},
]

# Porosity fields - per layer
POROSITY_FIELDS = [
    {"name": "Pore Size", "layers": True},
    {"name": "Pore Size Distribution", "layers": True},
    {"name": "Open Area Fraction", "layers": True},
    {"name": "Hydraulic Diameter", "layers": True},
]

# Structure-specific fields
STRUCTURE_SPECIFIC_FIELDS = [
    {"name": "Crimp (Wovens)", "layers": True, "note": "wovens only"},
    {"name": "Loop Length (Knits)", "layers": True, "note": "knits only"},
    {"name": "Float Length (Knits)", "layers": True, "note": "knits only"},
    {"name": "Proportion of Tucks/Misses (Knits)", "layers": True, "note": "knits only"},
    {"name": "Surface Roughness", "layers": True, "note": "in micrometers or microinches"},
]

# Physical properties - per layer
PHYSICAL_FIELDS = [
    {"name": "Fabric Weight/GSM", "layers": True, "note": "g/m² or oz/yd²"},
    {"name": "Fabric Thickness", "layers": True, "note": "mm or mil"},
    {"name": "Fabric Density", "layers": True, "note": "threads/inch, picks/inch, loops/inch"},
    {"name": "Tensile Strength", "layers": True},
    {"name": "Elongation at Break", "layers": True},
    {"name": "Tear Strength", "layers": True},
    {"name": "Fabric Hand/Feel", "layers": True},
]

# AATCC Moisture metrics - per layer
AATCC_MOISTURE_FIELDS = [
    {"name": "Wetting Time Bottom (WTB)", "layers": True, "standard": "AATCC"},
    {"name": "Wetting Time Top (WTT)", "layers": True, "standard": "AATCC"},
    {"name": "Wetting Time (WT)", "layers": True, "standard": "AATCC"},
    {"name": "Absorption Rate Bottom (ARB)", "layers": True, "standard": "AATCC"},
    {"name": "Absorption Rate Top (ART)", "layers": True, "standard": "AATCC"},
    {"name": "Bottom Absorption Rate (BAR)", "layers": True, "standard": "AATCC"},
    {"name": "Top Absorption Rate (TAR)", "layers": True, "standard": "AATCC"},
    {"name": "Wicking Distance/Rate", "layers": True, "standard": "AATCC"},
    {"name": "Vertical Wicking", "layers": True, "standard": "AATCC"},
    {"name": "Spreading Speed Bottom (SSb)", "layers": True, "standard": "AATCC"},
    {"name": "Spreading Speed Top (SSt)", "layers": True, "standard": "AATCC"},
    {"name": "Spreading Speed (SS)", "layers": True, "standard": "AATCC"},
    {"name": "Max Wetted Radius Bottom (MWRb)", "layers": True, "standard": "AATCC"},
    {"name": "Max Wetted Radius Top (MWRt)", "layers": True, "standard": "AATCC"},
    {"name": "Accumulative One-Way Transport Index (AOTI)", "layers": True, "standard": "AATCC"},
    {"name": "One Way Transport Capability (OWTC)", "layers": True, "standard": "AATCC"},
    {"name": "Overall Moisture Management Capacity (OMMC)", "layers": True, "standard": "AATCC"},
    {"name": "Drying Rate", "layers": True, "standard": "AATCC"},
    {"name": "Drying Time", "layers": True, "standard": "AATCC"},
    {"name": "Moisture Retention", "layers": True, "standard": "AATCC"},
    {"name": "Capillary Rise Height", "layers": True, "standard": "AATCC"},
    {"name": "Capillary Rise Rate", "layers": True, "standard": "AATCC"},
    {"name": "Water Retention", "layers": True, "standard": "AATCC"},
    {"name": "Maximum Water Uptake", "layers": True, "standard": "AATCC"},
]

# Wettability fields - per layer
WETTABILITY_FIELDS = [
    {"name": "Contact Angle", "layers": True},
    {"name": "Contact Angle - Advancing", "layers": True},
    {"name": "Contact Angle - Receding", "layers": True},
    {"name": "Wettability Classification", "layers": True, "note": "hydrophobic, hydrophilic, etc."},
    {"name": "Wettability Gradient", "layers": True},
    {"name": "Surface Energy/Tension", "layers": True},
    {"name": "Interfacial Tension", "layers": True},
]

# Transport metrics - per layer
TRANSPORT_FIELDS = [
    {"name": "Moisture Transport (Transplanar)", "layers": True},
    {"name": "Moisture Transport (Horizontal)", "layers": True},
    {"name": "Directional Moisture Transport", "layers": True},
    {"name": "Unidirectional Transport Direction", "layers": True},
    {"name": "Air Permeability", "layers": True},
    {"name": "Liquid Permeability", "layers": True},
    {"name": "Water Vapor Transmission Rate (WVTR)", "layers": True},
    {"name": "Moisture Vapor Transmission Rate (MVTR)", "layers": True},
]

# Thermal metrics - per layer
THERMAL_FIELDS = [
    {"name": "Thermal Conductivity", "layers": True},
    {"name": "Thermal Resistance (Rct)", "layers": True},
    {"name": "Evaporative Resistance (Ret)", "layers": True},
    {"name": "Thermal Diffusivity", "layers": True},
    {"name": "Heat Flux", "layers": True},
    {"name": "Thermal Absorptivity (Qmax)", "layers": True},
    {"name": "Heat Dissipation/Loss", "layers": True},
    {"name": "Breathability", "layers": True},
    {"name": "Comfort Index", "layers": True},
    {"name": "Moisture Management Capacity (MMC)", "layers": True},
]

# JIS-specific metrics - per layer
JIS_FIELDS = [
    {"name": "Wetting Time (JIS)", "layers": True, "standard": "JIS"},
    {"name": "Absorption Rate (JIS)", "layers": True, "standard": "JIS"},
    {"name": "Spreading Speed (JIS)", "layers": True, "standard": "JIS"},
    {"name": "Vertical Wicking (JIS)", "layers": True, "standard": "JIS"},
    {"name": "Horizontal Wicking (JIS)", "layers": True, "standard": "JIS"},
    {"name": "Drying Time (JIS)", "layers": True, "standard": "JIS"},
    {"name": "Drying Rate (JIS)", "layers": True, "standard": "JIS"},
    {"name": "Drying Index (JIS)", "layers": True, "standard": "JIS"},
    {"name": "Moisture Regain (JIS)", "layers": True, "standard": "JIS"},
    {"name": "Water Vapor Resistance (JIS)", "layers": True, "standard": "JIS"},
    {"name": "WVTR (JIS)", "layers": True, "standard": "JIS"},
    {"name": "Evaporative Heat Flux (JIS)", "layers": True, "standard": "JIS"},
]

# ISO/EN-specific metrics - per layer
ISO_EN_FIELDS = [
    {"name": "Thermal Conductivity (ISO)", "layers": True, "standard": "ISO/EN"},
    {"name": "Thermal Resistance (ISO)", "layers": True, "standard": "ISO/EN"},
    {"name": "Thermal Diffusivity (ISO)", "layers": True, "standard": "ISO/EN"},
    {"name": "Heat Flux (ISO)", "layers": True, "standard": "ISO/EN"},
    {"name": "Hydrostatic Head (ISO/EN)", "layers": True, "standard": "ISO/EN"},
    {"name": "Water Resistance (ISO/EN)", "layers": True, "standard": "ISO/EN"},
    {"name": "Air Permeability (ISO)", "layers": True, "standard": "ISO/EN"},
    {"name": "Water Vapor Resistance (ISO)", "layers": True, "standard": "ISO/EN"},
    {"name": "Water Vapor Permeability (ISO)", "layers": True, "standard": "ISO/EN"},
    {"name": "Evaporative Heat Flux (ISO)", "layers": True, "standard": "ISO/EN"},
    {"name": "Drying Rate (ISO)", "layers": True, "standard": "ISO/EN"},
]

# ASTM-specific metrics - per layer
ASTM_FIELDS = [
    {"name": "Qualitative Absorbency Rating (ASTM)", "layers": True, "standard": "ASTM"},
    {"name": "Hydrostatic Head (ASTM)", "layers": True, "standard": "ASTM"},
    {"name": "Water Resistance (ASTM)", "layers": True, "standard": "ASTM"},
    {"name": "Air Permeability (ASTM)", "layers": True, "standard": "ASTM"},
    {"name": "WVTR (ASTM)", "layers": True, "standard": "ASTM"},
    {"name": "Water Vapor Permeability (ASTM)", "layers": True, "standard": "ASTM"},
    {"name": "Water Absorbency (ASTM)", "layers": True, "standard": "ASTM"},
    {"name": "Absorption Time (ASTM)", "layers": True, "standard": "ASTM"},
    {"name": "Liquid Absorption Capacity (ASTM)", "layers": True, "standard": "ASTM"},
    {"name": "Thermal Resistance Rct (ASTM)", "layers": True, "standard": "ASTM"},
    {"name": "Evaporative Resistance Ret (ASTM)", "layers": True, "standard": "ASTM"},
]

# Other/unspecified standard metrics
OTHER_FIELDS = [
    {"name": "Capillary Pressure", "layers": True, "standard": "Other"},
    {"name": "SMF", "layers": True, "standard": "Other"},
    {"name": "PMI", "layers": True, "standard": "Other"},
]


def get_all_column_headers() -> List[str]:
    """
    Generate all column headers based on the schema.
    Returns a flat list of all column names.
    """
    headers = []

    # Study metadata
    headers.extend(STUDY_METADATA)

    # Sample metadata
    headers.extend(SAMPLE_METADATA)

    # Build layer-specific columns
    all_field_groups = [
        ("Material/Structure", MATERIAL_STRUCTURE_FIELDS),
        ("Yarn Specs", YARN_SPEC_FIELDS),
        ("Porosity", POROSITY_FIELDS),
        ("Structure-Specific", STRUCTURE_SPECIFIC_FIELDS),
        ("Physical Properties", PHYSICAL_FIELDS),
        ("Wettability", WETTABILITY_FIELDS),
        ("Transport", TRANSPORT_FIELDS),
        ("Thermal", THERMAL_FIELDS),
        ("AATCC Moisture", AATCC_MOISTURE_FIELDS),
        ("JIS", JIS_FIELDS),
        ("ISO/EN", ISO_EN_FIELDS),
        ("ASTM", ASTM_FIELDS),
        ("Other", OTHER_FIELDS),
    ]

    # Treatment fields with sides
    for field in TREATMENT_FIELDS:
        for layer in LAYER_TYPES:
            if field.get("sides"):
                for side in SIDES:
                    if side:
                        headers.append(f"{field['name']}: {layer}, {side}")
                    else:
                        headers.append(f"{field['name']}: {layer}")
            else:
                headers.append(f"{field['name']}: {layer}")

    # Regular layer fields (no sides)
    for group_name, fields in all_field_groups:
        for field in fields:
            if field.get("layers"):
                for layer in LAYER_TYPES:
                    headers.append(f"{field['name']}: {layer}")
            else:
                headers.append(field['name'])

    return headers


# ================== LLM FUNCTIONS ==================

def call_ollama(prompt: str, system_prompt: str = "") -> str:
    """Call local Ollama LLM."""
    full_prompt = f"{system_prompt}\n\n{prompt}" if system_prompt else prompt

    payload = {
        "model": OLLAMA_MODEL,
        "prompt": full_prompt,
        "stream": False,
        "options": {
            "temperature": 0.1,  # Low temperature for more deterministic extraction
            "num_predict": 8000,  # Allow longer responses for JSON
        }
    }

    try:
        resp = requests.post(OLLAMA_URL, json=payload, timeout=300)
        resp.raise_for_status()
        data = resp.json()
        return data.get("response", "")
    except requests.exceptions.ConnectionError:
        print("ERROR: Cannot connect to Ollama. Is it running?")
        print("Start Ollama with: ollama serve")
        return ""
    except Exception as e:
        print(f"Ollama error: {e}")
        return ""


def call_openai(prompt: str, system_prompt: str = "") -> str:
    """Call OpenAI API."""
    if not OPENAI_API_KEY:
        print("ERROR: OPENAI_API_KEY not set")
        return ""

    import openai
    client = openai.OpenAI(api_key=OPENAI_API_KEY)

    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})
    messages.append({"role": "user", "content": prompt})

    try:
        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=messages,
            temperature=0.1,
            max_tokens=8000,
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"OpenAI error: {e}")
        return ""


def call_llm(prompt: str, system_prompt: str = "") -> str:
    """Call the configured LLM provider."""
    if LLM_PROVIDER == "openai":
        return call_openai(prompt, system_prompt)
    else:
        return call_ollama(prompt, system_prompt)


# ================== EXTRACTION FUNCTIONS ==================

def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract all text from PDF using PyMuPDF."""
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc:
        full_text += page.get_text()
    doc.close()
    return full_text


def extract_study_metadata(pdf_path: str, full_text: str) -> Dict[str, Any]:
    """Extract study-level metadata using regex and LLM."""
    metadata = {}

    # Use existing functions for basic metadata
    from Scrape4_test import (
        extract_title_with_formatting,
        extract_year_from_text,
        extract_first_author_with_formatting,
    )

    metadata["Study Title"] = extract_title_with_formatting(pdf_path) or "Not found"
    metadata["Year of Publish"] = extract_year_from_text(full_text) or "Not found"
    metadata["Name of First-Listed Author"] = extract_first_author_with_formatting(pdf_path) or "Not found"

    # Use LLM to extract testing standards used
    standards_prompt = f"""
From this textile research paper text, identify ALL testing/standards bodies mentioned.
Look for: AATCC, ISO, JIS, EN, ASTM, GB/T, MMT, or any other testing standards.

Return ONLY a comma-separated list of the standards found, e.g.: "AATCC, ISO 9237, ASTM D737"
If none found, return "Not specified"

TEXT (first 8000 chars):
{full_text[:8000]}
"""

    standards_response = call_llm(standards_prompt)
    metadata["Testing/Standards Bodies Methods Used"] = standards_response.strip() if standards_response else "Not specified"

    return metadata


def extract_samples_info(full_text: str) -> List[Dict[str, Any]]:
    """
    Use LLM to identify all samples in the study and their basic properties.
    Returns a list of sample dictionaries.
    """

    system_prompt = """You are a textile research data extractor. Your task is to identify all fabric samples
mentioned in a research paper and extract their basic properties.

IMPORTANT RULES:
1. Return ONLY valid JSON, no other text
2. Each sample should have a unique ID from the paper (e.g., "S1", "Sample A", "Fabric 1")
3. If samples aren't explicitly named, use "Sample 1", "Sample 2", etc.
4. Determine the number of layers for each sample (1 = monolayer, 2 = bilayer, 3 = trilayer)
5. Extract material composition for each layer"""

    extraction_prompt = f"""
Analyze this textile research paper and identify ALL fabric samples tested.

For each sample, extract:
1. sample_id: The name/ID used in the paper
2. num_layers: Number of fabric layers (1, 2, or 3)
3. materials: Dictionary with layer materials, e.g., {{"inner": "cotton", "outer": "polyester"}} or {{"monolayer": "cotton/polyester blend"}}
4. structure: Dictionary with layer structures, e.g., {{"inner": "woven", "outer": "knit"}} or {{"monolayer": "woven"}}

Return as a JSON array. Example:
[
  {{
    "sample_id": "S1",
    "num_layers": 1,
    "materials": {{"monolayer": "100% cotton"}},
    "structure": {{"monolayer": "plain weave"}}
  }},
  {{
    "sample_id": "S2",
    "num_layers": 2,
    "materials": {{"inner": "polyester", "outer": "cotton"}},
    "structure": {{"inner": "knit", "outer": "woven"}}
  }}
]

PAPER TEXT:
{full_text[:15000]}

Return ONLY the JSON array, no explanation:"""

    response = call_llm(extraction_prompt, system_prompt)

    # Parse JSON response
    try:
        # Try to extract JSON from response
        json_match = re.search(r'\[[\s\S]*\]', response)
        if json_match:
            samples = json.loads(json_match.group())
            return samples
    except json.JSONDecodeError as e:
        print(f"Failed to parse samples JSON: {e}")
        print(f"Response was: {response[:500]}")

    # Fallback: return single unknown sample
    return [{"sample_id": "Sample 1", "num_layers": 1, "materials": {}, "structure": {}}]


def extract_metrics_for_sample(full_text: str, sample: Dict[str, Any], all_headers: List[str]) -> Dict[str, Any]:
    """
    Use LLM to extract all available metrics for a specific sample.
    """
    sample_id = sample.get("sample_id", "Unknown")
    num_layers = sample.get("num_layers", 1)

    # Determine which layer columns to use based on num_layers
    if num_layers == 1:
        relevant_layers = ["Monolayer"]
    elif num_layers == 2:
        relevant_layers = ["Inner Layer", "Outer Layer"]
    else:  # 3 layers
        relevant_layers = ["Inner Layer", "Middle Layer", "Outer Layer"]

    # Filter headers to only include relevant layers
    relevant_headers = []
    for header in all_headers:
        # Skip study/sample metadata
        if header in STUDY_METADATA or header in SAMPLE_METADATA:
            continue

        # Check if header contains a layer specification
        has_layer = any(layer in header for layer in LAYER_TYPES)

        if has_layer:
            # Only include if it's a relevant layer for this sample
            if any(layer in header for layer in relevant_layers):
                relevant_headers.append(header)
        else:
            # Non-layer-specific headers
            relevant_headers.append(header)

    # Create a focused extraction prompt
    system_prompt = """You are a precise textile data extractor. Extract numerical values and specifications
from research papers into a structured JSON format.

RULES:
1. Return ONLY valid JSON
2. Use null for values not found in the text
3. Include units when available (e.g., "25.3 g/m²" not just "25.3")
4. Be precise - only extract values that are clearly stated
5. For the specific sample requested, find its data in tables and text"""

    # Break headers into chunks to avoid overwhelming the LLM
    chunk_size = 50
    all_extracted = {}

    for i in range(0, len(relevant_headers), chunk_size):
        chunk_headers = relevant_headers[i:i+chunk_size]
        headers_list = "\n".join([f"- {h}" for h in chunk_headers])

        extraction_prompt = f"""
Extract data for SAMPLE "{sample_id}" from this textile research paper.

This sample has {num_layers} layer(s). Only extract values for the relevant layer type(s).

Extract these specific metrics (return null if not found):
{headers_list}

PAPER TEXT:
{full_text[:12000]}

Return as JSON object with metric names as keys. Example:
{{
  "Fabric Weight/GSM: Monolayer": "185 g/m²",
  "Thermal Conductivity: Monolayer": "0.045 W/m·K",
  "Wetting Time (WT): Monolayer": "3.2 seconds"
}}

JSON output for sample "{sample_id}":"""

        response = call_llm(extraction_prompt, system_prompt)

        # Parse response
        try:
            json_match = re.search(r'\{[\s\S]*\}', response)
            if json_match:
                chunk_data = json.loads(json_match.group())
                all_extracted.update(chunk_data)
        except json.JSONDecodeError:
            print(f"  Warning: Failed to parse chunk {i//chunk_size + 1}")
            continue

    return all_extracted


def process_single_pdf(pdf_path: str) -> List[Dict[str, Any]]:
    """
    Process a single PDF and extract all specifications.
    Returns a list of rows (one per sample).
    """
    print(f"\n{'='*60}")
    print(f"Processing: {pdf_path}")
    print(f"{'='*60}")

    # Extract text
    full_text = extract_text_from_pdf(pdf_path)
    print(f"  Extracted {len(full_text)} characters of text")

    # Get study metadata
    print("\n  Extracting study metadata...")
    study_metadata = extract_study_metadata(pdf_path, full_text)
    for key, value in study_metadata.items():
        print(f"    {key}: {value}")

    # Identify samples
    print("\n  Identifying samples...")
    samples = extract_samples_info(full_text)
    print(f"    Found {len(samples)} samples")
    for sample in samples:
        print(f"      - {sample.get('sample_id')}: {sample.get('num_layers')} layer(s)")

    # Get all headers
    all_headers = get_all_column_headers()

    # Extract metrics for each sample
    output_rows = []
    for idx, sample in enumerate(samples):
        print(f"\n  Extracting metrics for sample {idx+1}/{len(samples)}: {sample.get('sample_id')}...")

        # Start with study metadata
        row = {
            "Study Number": Path(pdf_path).stem,
            **study_metadata,
            "Sample ID/Name": sample.get("sample_id"),
            "Number of Fabric Layers": sample.get("num_layers"),
        }

        # Add material and structure from sample identification
        materials = sample.get("materials", {})
        structure = sample.get("structure", {})

        for layer_key, layer_name in [("inner", "Inner Layer"), ("middle", "Middle Layer"),
                                       ("outer", "Outer Layer"), ("monolayer", "Monolayer")]:
            if layer_key in materials:
                row[f"Material: {layer_name}"] = materials[layer_key]
            if layer_key in structure:
                row[f"Structure: {layer_name}"] = structure[layer_key]

        # Extract all metrics
        metrics = extract_metrics_for_sample(full_text, sample, all_headers)
        row.update(metrics)

        # Count non-null metrics
        non_null = sum(1 for v in metrics.values() if v and v != "null")
        print(f"    Extracted {non_null} non-null metrics")

        output_rows.append(row)

    return output_rows


def write_to_excel(rows: List[Dict[str, Any]], output_path: Path):
    """Write extracted data to Excel file."""
    print(f"\n{'='*60}")
    print(f"Writing to Excel: {output_path}")
    print(f"{'='*60}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Textile Specifications"

    # Get all headers
    all_headers = get_all_column_headers()

    # Write header row with formatting
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for col_idx, header in enumerate(all_headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Write data rows
    for row_idx, row_data in enumerate(rows, 2):
        for col_idx, header in enumerate(all_headers, 1):
            value = row_data.get(header, "")
            if value == "null" or value is None:
                value = ""
            ws.cell(row=row_idx, column=col_idx, value=value)

    # Adjust column widths
    for col_idx in range(1, len(all_headers) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 20

    # Freeze header row
    ws.freeze_panes = "A2"

    wb.save(output_path)
    print(f"  ✓ Saved {len(rows)} rows to {output_path}")


def main():
    """Main entry point."""
    print("\n" + "="*70)
    print("ScrapeV2 - Comprehensive Textile Specification Extractor")
    print("="*70)
    print(f"LLM Provider: {LLM_PROVIDER}")
    if LLM_PROVIDER == "ollama":
        print(f"Ollama Model: {OLLAMA_MODEL}")
    print(f"Focus PDF: {FOCUS_PDF}")
    print("="*70)

    # Check PDF exists
    pdf_path = PDF_FOLDER / FOCUS_PDF
    if not pdf_path.exists():
        print(f"\nERROR: PDF not found: {pdf_path}")
        print("Please check the path and try again.")
        return 1

    # Process the PDF
    all_rows = process_single_pdf(str(pdf_path))

    if not all_rows:
        print("\nNo data extracted. Check the PDF content and LLM connection.")
        return 1

    # Write to Excel
    write_to_excel(all_rows, OUTPUT_EXCEL)

    print("\n" + "="*70)
    print("EXTRACTION COMPLETE")
    print(f"Total samples extracted: {len(all_rows)}")
    print(f"Output file: {OUTPUT_EXCEL}")
    print("="*70 + "\n")

    return 0


if __name__ == "__main__":
    sys.exit(main())
