"""
SciMasterScrape3_test.py

NOUGAT-ONLY VERSION: Debug and optimize Nougat OCR for research papers.
NO FALLBACK - Forces Nougat usage to troubleshoot issues.

FOCUS:
- Get Nougat working properly (no PyMuPDF fallback)
- Better table and structure preservation
- Verbose debugging to see exactly what's failing
- Test on 1 PDF first

Architecture:
- Stage 1: PDF preprocessing with Nougat ONLY (verbose debugging)
- Stage 2: Multi-stage LLM extraction (metadata → samples → specs)
- Stage 3: Data validation and normalization

Output: Each layer of each sample gets its own row in structured format.
"""

import os
import re
import sys
import io  # ADDED: Missing import for BytesIO
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any
from dataclasses import dataclass, field, asdict
from enum import Enum
import json

# PDF Processing
import fitz  # PyMuPDF (fallback)

# ==================== ALBUMENTATIONS COMPATIBILITY FIX ====================
# Fix for albumentations v1.4.0+ compatibility with nougat-ocr
# Nougat uses old albumentations API - need to monkeypatch before importing
# This is the most likely cause of Nougat failures!

print("=" * 60)
print("CHECKING ALBUMENTATIONS COMPATIBILITY")
print("=" * 60)

try:
    import albumentations as alb
    from albumentations.core.transforms_interface import ImageOnlyTransform

    print(f"✓ Albumentations version: {alb.__version__}")
    print(f"✓ Applying compatibility patches...")

    # Monkeypatch ImageCompression - just make it a no-op for nougat
    class ImageCompression(ImageOnlyTransform):
        def __init__(self, *args, **kwargs):
            super(ImageOnlyTransform, self).__init__(always_apply=kwargs.get('always_apply', False),
                                                     p=kwargs.get('p', 0.5))
        def apply(self, img, **params):
            return img  # Return image unchanged
    alb.ImageCompression = ImageCompression
    print("  ✓ ImageCompression patched (no-op)")

    # Monkeypatch GaussNoise - normalize std_range to 0-1
    _original_GaussNoise = alb.GaussNoise
    class GaussNoise(_original_GaussNoise):
        def __init__(self, std_or_limit=20, p=0.5, **kwargs):
            # Normalize to 0-1 range (20 -> 0.2, etc.)
            if isinstance(std_or_limit, (int, float)):
                normalized = std_or_limit / 100.0  # Convert to 0-1 range
                super().__init__(var_limit=(0, normalized), p=p, **kwargs)
            else:
                super().__init__(var_limit=std_or_limit, p=p, **kwargs)
    alb.GaussNoise = GaussNoise
    print("  ✓ GaussNoise patched (normalized to 0-1)")

    print("✓ All patches applied successfully!")
    print("=" * 60 + "\n")

except ImportError as e:
    print(f"⚠ Albumentations not installed: {e}")
    print("=" * 60 + "\n")

try:
    from nougat import NougatModel
    from nougat.utils.checkpoint import get_checkpoint
    NOUGAT_AVAILABLE = True
except (ImportError, Exception) as e:
    NOUGAT_AVAILABLE = False
    print(f"⚠ Nougat not available: {e}")

# LLM and Workflow
try:
    # Try new langchain imports first (v0.1.0+)
    try:
        from langchain_openai import ChatOpenAI
    except ImportError:
        from langchain.chat_models import ChatOpenAI

    # ChatPromptTemplate moved to langchain_core in newer versions
    try:
        from langchain.prompts import ChatPromptTemplate
    except ImportError:
        from langchain_core.prompts import ChatPromptTemplate

    # FewShotChatMessagePromptTemplate moved in newer versions
    try:
        from langchain.prompts import FewShotChatMessagePromptTemplate
    except ImportError:
        try:
            from langchain_core.prompts import FewShotChatMessagePromptTemplate
        except ImportError:
            # If not available, we'll create a simple alternative
            FewShotChatMessagePromptTemplate = None

    # PydanticOutputParser moved to langchain_core in newer versions
    try:
        from langchain.output_parsers import PydanticOutputParser
    except ImportError:
        from langchain_core.output_parsers import PydanticOutputParser

    # LLMChain deprecated in newer versions
    try:
        from langchain.chains import LLMChain
    except ImportError:
        LLMChain = None

    LANGCHAIN_AVAILABLE = True
except ImportError as e:
    LANGCHAIN_AVAILABLE = False
    print(f"⚠ LangChain not available: {e}")
    print("Install with: pip install langchain langchain-openai langchain-core")

# Structured Data Validation
try:
    from pydantic import BaseModel, Field, validator
    PYDANTIC_AVAILABLE = True
except ImportError:
    PYDANTIC_AVAILABLE = False
    print("⚠ Pydantic not available. Install with: pip install pydantic")

# Data Processing
import pandas as pd
import numpy as np

# Unit Conversion
try:
    from pint import UnitRegistry
    ureg = UnitRegistry()
    PINT_AVAILABLE = True
except ImportError:
    PINT_AVAILABLE = False
    print("⚠ Pint not available for unit conversion. Install with: pip install pint")

# OpenAI Integration
try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False
    print("⚠ OpenAI not available. Install with: pip install openai")


# ==================== CONFIGURATION ====================

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
NOUGAT_MODEL = "facebook/nougat-base"
GPT_MODEL = "gpt-4-turbo-preview"  # 128k context (or use "gpt-3.5-turbo-16k" for cheaper)
TEMPERATURE = 0.1  # Low for deterministic extraction

# Output paths
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "extracted_data"
OUTPUT_DIR.mkdir(exist_ok=True)


# ==================== PYDANTIC DATA MODELS ====================

class YarnType(str, Enum):
    """Common yarn types in textile research"""
    SPUN = "spun"
    FILAMENT = "filament"
    TEXTURED = "textured"
    CORE_SPUN = "core-spun"
    AIR_JET = "air-jet"
    UNKNOWN = "unknown"


class FabricPattern(str, Enum):
    """Common fabric patterns"""
    PLAIN = "plain"
    TWILL = "twill"
    SATIN = "satin"
    JERSEY = "jersey"
    RIB = "rib"
    INTERLOCK = "interlock"
    PIQUE = "pique"
    UNKNOWN = "unknown"


class YarnSpecification(BaseModel):
    """Yarn specifications for a layer"""
    yarn_type: Optional[YarnType] = None
    filament_count: Optional[int] = Field(None, description="Number of filaments")
    tpi: Optional[float] = Field(None, description="Twists per inch")
    linear_density: Optional[float] = Field(None, description="Denier or Tex")
    linear_density_unit: Optional[str] = Field(None, description="'denier' or 'tex'")
    material: Optional[str] = None
    blend_ratio: Optional[str] = Field(None, description="e.g., '65/35'")

    class Config:
        use_enum_values = True


class FabricSpecification(BaseModel):
    """Fabric specifications for a layer"""
    density: Optional[float] = Field(None, description="Fabric density (g/m²)")
    thickness: Optional[float] = Field(None, description="Thickness (mm)")
    thickness_unit: Optional[str] = "mm"
    pattern: Optional[FabricPattern] = None
    stitch_density: Optional[float] = Field(None, description="Stitches per unit area")
    epi: Optional[int] = Field(None, description="Ends per inch (warp)")
    ppi: Optional[int] = Field(None, description="Picks per inch (weft)")
    cpi: Optional[int] = Field(None, description="Courses per inch")
    wpi: Optional[int] = Field(None, description="Wales per inch")
    loop_length: Optional[float] = Field(None, description="Loop length (mm)")

    class Config:
        use_enum_values = True


class ThermalMetrics(BaseModel):
    """Thermal and heat-related performance metrics"""
    evaporative_resistance: Optional[float] = Field(None, description="Ret (m²·Pa/W)")
    thermal_resistance: Optional[float] = Field(None, description="Rct (m²·K/W)")
    total_heat_loss: Optional[float] = Field(None, description="W/m² or similar")
    dry_heat_loss: Optional[float] = None
    evaporative_heat_loss: Optional[float] = None
    clo_value: Optional[float] = Field(None, description="Thermal insulation")


class LiquidTransportMetrics(BaseModel):
    """Liquid transport and moisture management metrics"""
    one_way_transport_index: Optional[float] = Field(None, description="R value")
    wetting_time_top: Optional[float] = Field(None, description="Seconds")
    wetting_time_bottom: Optional[float] = Field(None, description="Seconds")
    absorption_rate_top: Optional[float] = Field(None, description="%/s")
    absorption_rate_bottom: Optional[float] = Field(None, description="%/s")
    max_wetted_radius_top: Optional[float] = Field(None, description="mm")
    max_wetted_radius_bottom: Optional[float] = Field(None, description="mm")
    spreading_speed_top: Optional[float] = Field(None, description="mm/s")
    spreading_speed_bottom: Optional[float] = Field(None, description="mm/s")
    overall_moisture_management_capacity: Optional[float] = Field(None, description="OMMC")


class TreatmentInfo(BaseModel):
    """Surface treatments and finishes"""
    finish_type: Optional[str] = Field(None, description="e.g., 'hydrophilic', 'antimicrobial'")
    application_side: Optional[str] = Field(None, description="'top', 'bottom', 'both'")
    wettability_change: Optional[str] = Field(None, description="Description of wettability change")
    contact_angle_before: Optional[float] = Field(None, description="Degrees")
    contact_angle_after: Optional[float] = Field(None, description="Degrees")


class SampleLayer(BaseModel):
    """Single layer within a sample"""
    layer_number: int = Field(..., description="Layer position (1=innermost, N=outermost)")
    yarn_specs: Optional[YarnSpecification] = None
    fabric_specs: Optional[FabricSpecification] = None
    material: Optional[str] = Field(None, description="Primary material")
    blend_components: Optional[List[str]] = Field(None, description="For blends")
    treatment: Optional[TreatmentInfo] = None


class Sample(BaseModel):
    """Complete sample with all layers"""
    sample_id: str = Field(..., description="Sample identifier (e.g., 'S1', 'Sample A')")
    study_id: str = Field(..., description="Study identifier")
    num_layers: int = Field(..., description="Total number of layers")
    layers: List[SampleLayer] = Field(..., description="All layers, ordered inside-out")
    thermal_metrics: Optional[ThermalMetrics] = None
    liquid_transport_metrics: Optional[LiquidTransportMetrics] = None
    overall_thickness: Optional[float] = Field(None, description="Total thickness (mm)")
    overall_weight: Optional[float] = Field(None, description="Total weight (g/m²)")


class StudyMetadata(BaseModel):
    """Metadata about the research study"""
    study_id: str
    title: str
    first_author: str
    year: Optional[int] = None
    doi: Optional[str] = None
    total_samples: Optional[int] = None


class ExtractedData(BaseModel):
    """Complete extracted data from a study"""
    metadata: StudyMetadata
    samples: List[Sample]


# ==================== STAGE 1: PDF PREPROCESSING ====================

class PDFPreprocessor:
    """Handles PDF parsing and text extraction"""

    def __init__(self):
        self.nougat_model = None

    def load_nougat_model(self):
        """Load Nougat model (cached)"""
        if not NOUGAT_AVAILABLE:
            raise RuntimeError("Nougat not installed")

        if self.nougat_model is None:
            print("Loading Nougat model...")
            checkpoint = get_checkpoint(NOUGAT_MODEL)
            self.nougat_model = NougatModel.from_pretrained(checkpoint)
            print("✓ Nougat model loaded")

        return self.nougat_model

    def extract_with_nougat(self, pdf_path: str) -> str:
        """
        Extract text using Nougat ONLY (no fallback).
        Returns markdown-formatted text with tables and structure preserved.
        """
        if not NOUGAT_AVAILABLE:
            raise RuntimeError(
                "❌ Nougat not available! Install with:\n"
                "   pip install nougat-ocr\n"
                "This version requires Nougat - no fallback to PyMuPDF."
            )

        print(f"\n{'='*60}")
        print("NOUGAT EXTRACTION - VERBOSE MODE")
        print(f"{'='*60}")

        try:
            from PIL import Image
            import fitz  # PyMuPDF for PDF to image conversion

            print("✓ Imports successful (PIL, fitz)")

            # Load model
            print("\n1. Loading Nougat model...")
            model = self.load_nougat_model()
            print(f"   ✓ Model loaded: {type(model)}")

            # Open PDF
            print(f"\n2. Opening PDF: {Path(pdf_path).name}")
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            print(f"   ✓ PDF opened: {total_pages} pages")

            all_text = []

            # Process only first 8 pages (or all pages if fewer than 8)
            pages_to_process = min(8, total_pages)
            print(f"   → Processing first {pages_to_process} pages only")

            for page_num in range(pages_to_process):
                print(f"\n3. Processing page {page_num + 1}/{pages_to_process}...")

                # Render page to image
                page = doc[page_num]
                pix = page.get_pixmap(dpi=150)
                print(f"   ✓ Rendered to image: {pix.width}x{pix.height} pixels")

                # Convert to PIL Image
                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data))
                print(f"   ✓ Converted to PIL Image: {img.size}")

                # Run Nougat inference
                print(f"   → Running Nougat inference...")
                predictions = model.inference(image=img)  # FIXED: removed batch_size parameter

                # Extract clean text from predictions dict
                if predictions and isinstance(predictions, dict) and 'predictions' in predictions:
                    # predictions is a dict with key 'predictions' containing a list of strings
                    page_text = predictions['predictions'][0] if predictions['predictions'] else ""
                    print(f"   ✓ Extracted {len(page_text)} characters from page {page_num + 1}")
                    all_text.append(page_text)
                else:
                    print(f"   ⚠ WARNING: Page {page_num + 1} returned empty or unexpected format!")
                    print(f"   → Predictions: {predictions}")

            doc.close()

            # Combine all pages
            if all_text:
                markdown_text = "\n\n<!-- PAGE BREAK -->\n\n".join(all_text)
                print(f"\n{'='*60}")
                print(f"✅ NOUGAT SUCCESS!")
                print(f"   Total characters: {len(markdown_text)}")
                print(f"   Pages processed: {len(all_text)}/{pages_to_process} (first 8 of {total_pages} total)")
                print(f"{'='*60}\n")
                return markdown_text
            else:
                raise RuntimeError(
                    f"❌ Nougat extracted 0 text from {total_pages} pages!\n"
                    f"   This usually means:\n"
                    f"   1. Albumentations compatibility issue\n"
                    f"   2. Model inference failing silently\n"
                    f"   3. PDF rendering issue"
                )

        except Exception as e:
            print(f"\n{'='*60}")
            print(f"❌ NOUGAT FAILED!")
            print(f"   Error: {e}")
            print(f"   Type: {type(e).__name__}")
            print(f"{'='*60}\n")
            import traceback
            traceback.print_exc()
            raise  # NO FALLBACK - let it fail

    def extract_with_pymupdf(self, pdf_path: str) -> str:
        """Fallback: Extract text using PyMuPDF"""
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        print(f"✓ PyMuPDF extracted {len(text)} characters")
        return text

    def preprocess(self, pdf_path: str) -> str:
        """
        Main preprocessing entry point - NOUGAT ONLY.
        Returns markdown-formatted text with tables and structure preserved.
        """
        print(f"\n=== Preprocessing: {Path(pdf_path).name} ===")
        print("Using: NOUGAT ONLY (no fallback)")

        return self.extract_with_nougat(pdf_path)


# ==================== STAGE 2: LLM-BASED EXTRACTION ====================

class LLMExtractor:
    """Handles structured data extraction using LLMs"""

    def __init__(self, api_key: str = OPENAI_API_KEY, model: str = GPT_MODEL):
        if not OPENAI_AVAILABLE or not LANGCHAIN_AVAILABLE:
            raise RuntimeError("OpenAI and LangChain required for extraction")

        self.api_key = api_key
        self.model = model
        openai.api_key = api_key

        # Initialize LangChain LLM
        self.llm = ChatOpenAI(
            model=model,
            temperature=TEMPERATURE,
            openai_api_key=api_key
        )

    def create_few_shot_prompt(self, examples: List[Dict], example_template: str):
        """Create few-shot prompt from examples"""
        if FewShotChatMessagePromptTemplate is not None:
            return FewShotChatMessagePromptTemplate(
                examples=examples,
                example_prompt=ChatPromptTemplate.from_template(example_template)
            )
        else:
            # Fallback: Create a simple prompt with examples as text
            examples_text = "\n\n".join([
                example_template.format(**ex) for ex in examples
            ])
            return ChatPromptTemplate.from_template(
                f"Here are some examples:\n\n{examples_text}\n\n" + "{input}"
            )

    def extract_metadata(self, text: str) -> StudyMetadata:
        """Extract study metadata (title, author, year, DOI) - IMPROVED VERSION"""
        print("\n--- Extracting Metadata ---")

        # Use more text - first 10000 chars to capture full introduction/abstract
        text_chunk = text[:10000]

        prompt = ChatPromptTemplate.from_template("""
You are an expert at extracting metadata from scientific research papers.

Extract the following information from this research paper:

1. TITLE: The full title of the paper (usually at the top of page 1, often in larger font or bold)
2. FIRST AUTHOR: The family name and initials of the first listed author (e.g., "Smith J.R.")
3. YEAR: The publication year (look for copyright, dates, journal info)
4. DOI: The Digital Object Identifier if present (format: 10.xxxx/xxxxx)
5. TOTAL SAMPLES: How many fabric samples/specimens were tested in this study

IMPORTANT:
- Look carefully at the beginning of the paper for title, authors, journal info
- For sample count, look in the abstract, introduction, or methods section
- Be precise - extract exactly what you see
- If you cannot find a field, use null

Text from paper:
{text}

Return ONLY a JSON object with this exact structure:
{{
  "study_id": "UNKNOWN",
  "title": "extracted title here",
  "first_author": "LastName A.B.",
  "year": 2023,
  "doi": "10.xxxx/xxxxx",
  "total_samples": 5
}}
""")

        chain = prompt | self.llm
        response = chain.invoke({"text": text_chunk})

        try:
            # Parse JSON from response - more robust parsing
            content = response.content

            # Try to find JSON block (handle markdown code blocks too)
            json_match = re.search(r'```(?:json)?\s*(\{[^`]+\})\s*```', content, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                # Try without code block
                json_match = re.search(r'\{[^}]+\}', content, re.DOTALL)
                if json_match:
                    json_str = json_match.group(0)
                else:
                    print(f"⚠ Could not find JSON in response: {content[:200]}")
                    return StudyMetadata(study_id="UNKNOWN", title="UNKNOWN", first_author="UNKNOWN")

            metadata_dict = json.loads(json_str)
            print(f"✓ Metadata: {metadata_dict.get('title', 'UNKNOWN')[:60]}...")
            return StudyMetadata(**metadata_dict)

        except json.JSONDecodeError as e:
            print(f"⚠ JSON parse error: {e}")
            print(f"   Response: {content[:500]}")
            return StudyMetadata(study_id="UNKNOWN", title="UNKNOWN", first_author="UNKNOWN")
        except Exception as e:
            print(f"⚠ Metadata extraction error: {e}")
            return StudyMetadata(study_id="UNKNOWN", title="UNKNOWN", first_author="UNKNOWN")

    def extract_sample_list(self, text: str, metadata: StudyMetadata) -> List[Dict]:
        """
        STAGE 1: Extract just sample IDs and layer counts.
        Returns basic info: [{"sample_id": "S1", "num_layers": 3}, ...]
        """
        print("\n--- Stage 1: Extracting Sample List ---")

        # Focus on Methods section where samples are described
        methods_section = self._find_section(text, ["materials and methods", "methods", "experimental"])

        prompt = ChatPromptTemplate.from_template("""
You are an expert at analyzing textile research papers.

Your task: Identify ALL fabric samples tested in this study.

For each sample, extract:
1. Sample ID (e.g., "S1", "Sample A", "Fabric 1", "F-1")
2. Number of layers (single-layer = 1, bi-layer = 2, tri-layer = 3, etc.)

WHERE TO LOOK:
- Methods section (usually describes sample construction)
- Tables showing sample properties
- Results section mentioning samples

Text:
{text}

Return ONLY a JSON array like this:
[
  {{"sample_id": "S1", "num_layers": 2}},
  {{"sample_id": "S2", "num_layers": 3}},
  {{"sample_id": "S3", "num_layers": 1}}
]

IMPORTANT:
- Extract EVERY sample mentioned
- If layer count is unclear, use 1 as default
- Sample IDs should match how they're referenced in the paper
""")

        chain = prompt | self.llm
        response = chain.invoke({"text": methods_section or text[:15000]})

        try:
            content = response.content
            # Try markdown code block first
            json_match = re.search(r'```(?:json)?\s*(\[.*?\])\s*```', content, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                # Try without code block
                json_match = re.search(r'\[.*?\]', content, re.DOTALL)
                if json_match:
                    json_str = json_match.group(0)
                else:
                    print(f"⚠ Could not find sample list JSON")
                    return []

            sample_list = json.loads(json_str)
            print(f"✓ Found {len(sample_list)} samples: {[s['sample_id'] for s in sample_list]}")
            return sample_list

        except Exception as e:
            print(f"⚠ Sample list extraction error: {e}")
            return []

    def extract_sample_details(self, text: str, sample_id: str, num_layers: int, metadata: StudyMetadata) -> Optional[Sample]:
        """
        STAGE 2: Extract detailed specifications for a single sample.
        Much more focused and accurate than trying to extract all at once.
        """
        print(f"\n--- Stage 2: Extracting Details for {sample_id} ---")

        # Find relevant sections mentioning this sample
        sample_context = self._find_sample_context(text, sample_id)

        prompt = ChatPromptTemplate.from_template("""
You are an expert at extracting detailed textile specifications from research papers.

Extract ALL available information for sample "{sample_id}" which has {num_layers} layer(s).

For EACH layer (1 to {num_layers}), extract:

YARN SPECIFICATIONS:
- yarn_type: "spun", "filament", "textured", "core-spun", "air-jet", or "unknown"
- filament_count: number (if filament yarn)
- tpi: twists per inch (number)
- linear_density: denier or tex value (number)
- linear_density_unit: "denier" or "tex"
- material: e.g., "polyester", "cotton", "nylon"
- blend_ratio: e.g., "65/35" (if blend)

FABRIC SPECIFICATIONS:
- density: fabric weight in g/m² (number)
- thickness: in mm (number)
- pattern: "plain", "twill", "satin", "jersey", "rib", "interlock", "pique", or "unknown"
- stitch_density: stitches per unit area (number)
- epi: ends per inch for woven (number)
- ppi: picks per inch for woven (number)
- cpi: courses per inch for knitted (number)
- wpi: wales per inch for knitted (number)
- loop_length: in mm (number)

MATERIAL & TREATMENT:
- material: primary material name
- blend_components: array of materials if blend
- finish_type: e.g., "hydrophilic", "antimicrobial", "water-repellent"
- application_side: "top", "bottom", or "both"

PERFORMANCE METRICS (for whole sample):
THERMAL:
- evaporative_resistance: Ret in m²·Pa/W
- thermal_resistance: Rct in m²·K/W
- total_heat_loss: W/m²
- clo_value: thermal insulation

LIQUID TRANSPORT:
- one_way_transport_index: R value
- wetting_time_top: seconds
- wetting_time_bottom: seconds
- absorption_rate_top: %/s
- absorption_rate_bottom: %/s
- max_wetted_radius_top: mm
- max_wetted_radius_bottom: mm
- spreading_speed_top: mm/s
- spreading_speed_bottom: mm/s
- overall_moisture_management_capacity: OMMC value

Text about {sample_id}:
{text}

Return ONLY a JSON object with this structure:
{{
  "sample_id": "{sample_id}",
  "study_id": "{study_id}",
  "num_layers": {num_layers},
  "layers": [
    {{
      "layer_number": 1,
      "yarn_specs": {{...}},
      "fabric_specs": {{...}},
      "material": "...",
      "blend_components": null,
      "treatment": {{...}}
    }}
  ],
  "thermal_metrics": {{...}},
  "liquid_transport_metrics": {{...}},
  "overall_thickness": null,
  "overall_weight": null
}}

Use null for any field you cannot find. Extract exact numbers with proper units.
""")

        chain = prompt | self.llm
        response = chain.invoke({
            "text": sample_context or text[:20000],
            "sample_id": sample_id,
            "num_layers": num_layers,
            "study_id": metadata.study_id
        })

        try:
            content = response.content
            # Try markdown code block first
            json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                # Try without code block
                json_match = re.search(r'\{.*\}', content, re.DOTALL)
                if json_match:
                    json_str = json_match.group(0)
                else:
                    print(f"⚠ Could not find sample JSON for {sample_id}")
                    return None

            sample_data = json.loads(json_str)
            sample = Sample(**sample_data)
            print(f"✓ Extracted {len(sample.layers)} layer(s) for {sample_id}")
            return sample

        except Exception as e:
            print(f"⚠ Error extracting {sample_id}: {e}")
            return None

    def _find_section(self, text: str, section_names: List[str]) -> Optional[str]:
        """Find and extract a specific section from paper (e.g., Methods, Results)"""
        text_lower = text.lower()

        for section_name in section_names:
            # Look for section headers like "2. Materials and Methods" or "Methods"
            pattern = rf'\b\d*\.?\s*{re.escape(section_name)}\b'
            match = re.search(pattern, text_lower)

            if match:
                start = match.start()
                # Find next section or take next 5000 chars
                next_section = re.search(r'\n\d+\.\s+[A-Z]', text[start+100:])
                if next_section:
                    end = start + 100 + next_section.start()
                else:
                    end = start + 5000

                section_text = text[start:end]
                print(f"✓ Found '{section_name}' section ({len(section_text)} chars)")
                return section_text

        return None

    def _find_sample_context(self, text: str, sample_id: str) -> Optional[str]:
        """Find all text mentioning a specific sample ID"""
        # Look for mentions of this sample ID
        pattern = rf'\b{re.escape(sample_id)}\b'
        matches = list(re.finditer(pattern, text, re.IGNORECASE))

        if not matches:
            print(f"⚠ Sample {sample_id} not found in text")
            return None

        # Collect context around each mention (500 chars before and after)
        contexts = []
        for match in matches[:10]:  # Limit to first 10 mentions
            start = max(0, match.start() - 500)
            end = min(len(text), match.end() + 500)
            contexts.append(text[start:end])

        combined = "\n\n....\n\n".join(contexts)
        print(f"✓ Found {len(matches)} mentions of {sample_id}")
        return combined

    def extract_samples(self, text: str, metadata: StudyMetadata) -> List[Sample]:
        """
        IMPROVED MULTI-STAGE EXTRACTION:
        Stage 1: Get list of all samples and layer counts
        Stage 2: Extract detailed specs for each sample individually
        """
        print("\n--- Extracting Samples (Multi-Stage) ---")

        # Stage 1: Get sample list
        sample_list = self.extract_sample_list(text, metadata)

        if not sample_list:
            print("⚠ No samples found")
            return []

        # Stage 2: Extract details for each sample
        samples = []
        for sample_info in sample_list:
            sample = self.extract_sample_details(
                text,
                sample_info["sample_id"],
                sample_info["num_layers"],
                metadata
            )
            if sample:
                samples.append(sample)

        print(f"\n✓ Total samples extracted: {len(samples)}/{len(sample_list)}")
        return samples

    def extract_all(self, text: str) -> ExtractedData:
        """Main extraction pipeline"""
        # Step 1: Extract metadata
        metadata = self.extract_metadata(text)
        print(f"✓ Study: {metadata.title[:60]}...")

        # Step 2: Extract samples
        samples = self.extract_samples(text, metadata)

        return ExtractedData(metadata=metadata, samples=samples)


# ==================== STAGE 3: DATA NORMALIZATION ====================

class DataNormalizer:
    """Normalizes units, materials, and validates data"""

    def normalize_units(self, sample: Sample) -> Sample:
        """Convert all units to standard formats"""
        if not PINT_AVAILABLE:
            return sample

        # Normalize thickness to mm
        for layer in sample.layers:
            if layer.fabric_specs and layer.fabric_specs.thickness:
                if layer.fabric_specs.thickness_unit == "cm":
                    layer.fabric_specs.thickness *= 10
                    layer.fabric_specs.thickness_unit = "mm"

        return sample

    def validate_data(self, data: ExtractedData) -> ExtractedData:
        """Validate extracted data for consistency"""
        # Check for missing critical fields
        for sample in data.samples:
            if sample.num_layers != len(sample.layers):
                print(f"⚠ Sample {sample.sample_id}: num_layers mismatch")

        return data

    def normalize(self, data: ExtractedData) -> ExtractedData:
        """Main normalization pipeline"""
        print("\n--- Normalizing Data ---")

        # Normalize units for each sample
        for i, sample in enumerate(data.samples):
            data.samples[i] = self.normalize_units(sample)

        # Validate data
        data = self.validate_data(data)

        return data


# ==================== OUTPUT FORMATTING ====================

def flatten_to_dataframe(data: ExtractedData) -> pd.DataFrame:
    """
    Convert extracted data to flat DataFrame where each layer gets its own row.
    """
    rows = []

    for sample in data.samples:
        for layer in sample.layers:
            row = {
                # Study metadata
                'study_id': data.metadata.study_id,
                'title': data.metadata.title,
                'first_author': data.metadata.first_author,
                'year': data.metadata.year,
                'doi': data.metadata.doi,

                # Sample info
                'sample_id': sample.sample_id,
                'total_layers': sample.num_layers,
                'layer_number': layer.layer_number,

                # Material
                'material': layer.material,
                'blend_components': ','.join(layer.blend_components) if layer.blend_components else None,

                # Yarn specs
                'yarn_type': layer.yarn_specs.yarn_type if layer.yarn_specs else None,
                'filament_count': layer.yarn_specs.filament_count if layer.yarn_specs else None,
                'tpi': layer.yarn_specs.tpi if layer.yarn_specs else None,
                'linear_density': layer.yarn_specs.linear_density if layer.yarn_specs else None,
                'linear_density_unit': layer.yarn_specs.linear_density_unit if layer.yarn_specs else None,

                # Fabric specs
                'fabric_density_gsm': layer.fabric_specs.density if layer.fabric_specs else None,
                'thickness_mm': layer.fabric_specs.thickness if layer.fabric_specs else None,
                'pattern': layer.fabric_specs.pattern if layer.fabric_specs else None,
                'stitch_density': layer.fabric_specs.stitch_density if layer.fabric_specs else None,
                'epi': layer.fabric_specs.epi if layer.fabric_specs else None,
                'ppi': layer.fabric_specs.ppi if layer.fabric_specs else None,
                'cpi': layer.fabric_specs.cpi if layer.fabric_specs else None,
                'wpi': layer.fabric_specs.wpi if layer.fabric_specs else None,
                'loop_length_mm': layer.fabric_specs.loop_length if layer.fabric_specs else None,

                # Treatment
                'finish_type': layer.treatment.finish_type if layer.treatment else None,
                'application_side': layer.treatment.application_side if layer.treatment else None,
                'wettability_change': layer.treatment.wettability_change if layer.treatment else None,
                'contact_angle_before': layer.treatment.contact_angle_before if layer.treatment else None,
                'contact_angle_after': layer.treatment.contact_angle_after if layer.treatment else None,

                # Thermal metrics (sample-level)
                'evaporative_resistance': sample.thermal_metrics.evaporative_resistance if sample.thermal_metrics else None,
                'thermal_resistance': sample.thermal_metrics.thermal_resistance if sample.thermal_metrics else None,
                'total_heat_loss': sample.thermal_metrics.total_heat_loss if sample.thermal_metrics else None,
                'clo_value': sample.thermal_metrics.clo_value if sample.thermal_metrics else None,

                # Liquid transport metrics (sample-level)
                'one_way_transport_index': sample.liquid_transport_metrics.one_way_transport_index if sample.liquid_transport_metrics else None,
                'wetting_time_top': sample.liquid_transport_metrics.wetting_time_top if sample.liquid_transport_metrics else None,
                'wetting_time_bottom': sample.liquid_transport_metrics.wetting_time_bottom if sample.liquid_transport_metrics else None,
                'absorption_rate_top': sample.liquid_transport_metrics.absorption_rate_top if sample.liquid_transport_metrics else None,
                'absorption_rate_bottom': sample.liquid_transport_metrics.absorption_rate_bottom if sample.liquid_transport_metrics else None,
                'ommc': sample.liquid_transport_metrics.overall_moisture_management_capacity if sample.liquid_transport_metrics else None,

                # Sample-level overall
                'overall_thickness_mm': sample.overall_thickness,
                'overall_weight_gsm': sample.overall_weight,
            }
            rows.append(row)

    df = pd.DataFrame(rows)
    print(f"\n✓ Created DataFrame with {len(df)} rows (one per layer)")
    return df


# ==================== MAIN PIPELINE ====================

class SciMasterScraper:
    """Main pipeline orchestrator - Nougat-only version"""

    def __init__(self, openai_api_key: str = None):
        self.preprocessor = PDFPreprocessor()
        # Nougat-only mode: no LLM extraction needed
        self.extractor = None
        self.normalizer = DataNormalizer()

    def process_pdf(self, pdf_path: str) -> Tuple[ExtractedData, pd.DataFrame]:
        """
        Complete pipeline: PDF → Extracted Data → DataFrame (NOUGAT ONLY)

        Args:
            pdf_path: Path to PDF file

        Returns:
            (ExtractedData, DataFrame) - Structured data and flattened table
        """
        print(f"\n{'='*60}")
        print(f"Processing: {Path(pdf_path).name} (NOUGAT ONLY)")
        print(f"{'='*60}")

        # Stage 1: Preprocess with Nougat
        text = self.preprocessor.preprocess(pdf_path)

        # Nougat-only mode: Save extracted text and skip LLM extraction
        if self.extractor is None:
            # Save the raw Nougat markdown output
            output_txt = Path(pdf_path).stem + "_nougat.txt"
            with open(output_txt, 'w', encoding='utf-8') as f:
                f.write(text)
            print(f"✓ Saved Nougat output to {output_txt}")

            # Return empty data structures (no LLM extraction)
            return None, None

        # Stage 2: Extract (if extractor is available)
        extracted_data = self.extractor.extract_all(text)

        # Stage 3: Normalize
        normalized_data = self.normalizer.normalize(extracted_data)

        # Format output
        df = flatten_to_dataframe(normalized_data)

        return normalized_data, df

    def process_batch(self, pdf_paths: List[str], output_excel: str) -> pd.DataFrame:
        """
        Process multiple PDFs and combine into single Excel file.

        Args:
            pdf_paths: List of PDF file paths
            output_excel: Output Excel file path

        Returns:
            Combined DataFrame
        """
        all_dfs = []

        for pdf_path in pdf_paths:
            try:
                _, df = self.process_pdf(pdf_path)
                all_dfs.append(df)
            except Exception as e:
                print(f"✗ Error processing {pdf_path}: {e}")
                continue

        # Combine all DataFrames
        combined_df = pd.concat(all_dfs, ignore_index=True)

        # Save to Excel
        combined_df.to_excel(output_excel, index=False)
        print(f"\n✓ Saved {len(combined_df)} rows to {output_excel}")

        return combined_df


# ==================== EXAMPLE USAGE ====================

def main():
    """Example usage - Nougat-only extraction (no API key needed)"""

    # Initialize scraper (Nougat-only, no API key needed)
    scraper = SciMasterScraper(openai_api_key=None)

    # Process PDFs 1.pdf through 10.pdf in current directory
    current_dir = Path.cwd()
    all_results = []

    # TESTING: Start with just 1.pdf first
    test_range = range(1, 2)  # Change to range(1, 11) to process all PDFs

    for i in test_range:
        pdf_path = current_dir / f"{i}.pdf"

        if not pdf_path.exists():
            print(f"\n⚠ Skipping: {pdf_path.name} (not found)")
            continue

        print(f"\n{'='*60}")
        print(f"Processing: {pdf_path.name}")
        print(f"{'='*60}")

        try:
            extracted_data, df = scraper.process_pdf(str(pdf_path))

            # Nougat-only mode: text is already saved, skip JSON/Excel
            if extracted_data is None:
                print(f"✓ Nougat extraction complete for {pdf_path.name}")
                continue

            # Save individual outputs (only if LLM extraction was done)
            OUTPUT_DIR.mkdir(exist_ok=True)
            output_json = OUTPUT_DIR / f"{i}_extracted.json"
            output_excel = OUTPUT_DIR / f"{i}_data.xlsx"

            # Save JSON
            with open(output_json, 'w') as f:
                json.dump(extracted_data.dict(), f, indent=2)
            print(f"✓ Saved JSON to {output_json}")

            # Save Excel
            df.to_excel(output_excel, index=False)
            print(f"✓ Saved Excel to {output_excel}")

            all_results.append(df)

        except Exception as e:
            print(f"✗ Error processing {pdf_path.name}: {e}")
            continue

    # Combine all results into single Excel file
    if all_results:
        import pandas as pd
        combined_df = pd.concat(all_results, ignore_index=True)
        combined_excel = OUTPUT_DIR / "all_studies_combined.xlsx"
        combined_df.to_excel(combined_excel, index=False)
        print(f"\n{'='*60}")
        print(f"✓ Combined results: {combined_excel}")
        print(f"✓ Total rows: {len(combined_df)}")
        print(f"✓ Total PDFs processed: {len(all_results)}")
        print(f"{'='*60}")
    else:
        print("\n⚠ No PDFs were successfully processed")


if __name__ == "__main__":
    main()