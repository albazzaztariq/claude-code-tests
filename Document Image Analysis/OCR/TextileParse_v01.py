"""
SciMasterScrape1_test.py

Comprehensive textile engineering data extraction from research PDFs.
Extracts detailed yarn, fabric, material, and performance metrics with minimal training.

Architecture:
- Stage 1: PDF preprocessing with Nougat (Neural Optical Understanding)
- Stage 2: Structured data extraction with LLM-assisted methods
- Stage 3: Data validation and normalization

Output: Each layer of each sample gets its own row in structured format.
"""

import os
import re
import sys
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any
from dataclasses import dataclass, field, asdict
from enum import Enum
import json

# PDF Processing
import fitz  # PyMuPDF (fallback)
try:
    from nougat import NougatModel
    from nougat.utils.checkpoint import get_checkpoint
    NOUGAT_AVAILABLE = True
except ImportError:
    NOUGAT_AVAILABLE = False
    print("⚠ Nougat not available. Install with: pip install nougat-ocr")

# LLM and Workflow
try:
    from langchain.chat_models import ChatOpenAI
    from langchain.prompts import ChatPromptTemplate, FewShotChatMessagePromptTemplate
    from langchain.output_parsers import PydanticOutputParser
    from langchain.chains import LLMChain
    LANGCHAIN_AVAILABLE = True
except ImportError:
    LANGCHAIN_AVAILABLE = False
    print("⚠ LangChain not available. Install with: pip install langchain langchain-openai")

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
GPT_MODEL = "gpt-4"  # or "gpt-3.5-turbo" for faster/cheaper
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
        Extract text using Nougat (preserves structure, tables, math).
        Returns markdown-formatted text.
        """
        if not NOUGAT_AVAILABLE:
            print("⚠ Falling back to PyMuPDF")
            return self.extract_with_pymupdf(pdf_path)

        try:
            model = self.load_nougat_model()
            predictions = model.inference(pdf_path=pdf_path, batch_size=1)

            if predictions and len(predictions) > 0:
                markdown_text = predictions[0]
                print(f"✓ Nougat extracted {len(markdown_text)} characters")
                return markdown_text
            else:
                print("⚠ Nougat returned empty, falling back to PyMuPDF")
                return self.extract_with_pymupdf(pdf_path)

        except Exception as e:
            print(f"⚠ Nougat error: {e}. Falling back to PyMuPDF")
            return self.extract_with_pymupdf(pdf_path)

    def extract_with_pymupdf(self, pdf_path: str) -> str:
        """Fallback: Extract text using PyMuPDF"""
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        print(f"✓ PyMuPDF extracted {len(text)} characters")
        return text

    def preprocess(self, pdf_path: str, use_nougat: bool = True) -> str:
        """
        Main preprocessing entry point.
        Returns structured text (markdown if Nougat, plain text otherwise).
        """
        print(f"\n=== Preprocessing: {Path(pdf_path).name} ===")

        if use_nougat and NOUGAT_AVAILABLE:
            return self.extract_with_nougat(pdf_path)
        else:
            return self.extract_with_pymupdf(pdf_path)


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

    def create_few_shot_prompt(self, examples: List[Dict], example_template: str) -> FewShotChatMessagePromptTemplate:
        """Create few-shot prompt from examples"""
        return FewShotChatMessagePromptTemplate(
            examples=examples,
            example_prompt=ChatPromptTemplate.from_template(example_template)
        )

    def extract_metadata(self, text: str) -> StudyMetadata:
        """Extract study metadata (title, author, year, DOI)"""
        print("\n--- Extracting Metadata ---")

        prompt = ChatPromptTemplate.from_template("""
You are an expert at extracting metadata from scientific papers.

Extract the following information from the text:
- Title (full title of the paper)
- First author (family name and initials)
- Year (publication year)
- DOI (if available)
- Total samples (number of samples/specimens tested, if mentioned)

Text:
{text}

Return a JSON object with keys: study_id, title, first_author, year, doi, total_samples.
Use "UNKNOWN" for study_id. If a field is not found, use null.
""")

        chain = prompt | self.llm
        response = chain.invoke({"text": text[:3000]})  # Use first 3000 chars

        try:
            # Parse JSON from response
            content = response.content
            # Find JSON block
            json_match = re.search(r'\{[^}]+\}', content, re.DOTALL)
            if json_match:
                metadata_dict = json.loads(json_match.group(0))
                return StudyMetadata(**metadata_dict)
            else:
                print(f"⚠ Could not parse metadata JSON from: {content}")
                return StudyMetadata(study_id="UNKNOWN", title="UNKNOWN", first_author="UNKNOWN")
        except Exception as e:
            print(f"⚠ Metadata extraction error: {e}")
            return StudyMetadata(study_id="UNKNOWN", title="UNKNOWN", first_author="UNKNOWN")

    def extract_samples(self, text: str, metadata: StudyMetadata) -> List[Sample]:
        """
        Extract all samples with their layers and specifications.
        This is the main extraction logic.
        """
        print("\n--- Extracting Samples ---")

        # Few-shot examples for yarn extraction
        yarn_examples = [
            {
                "input": "The inner layer was made of 100% polyester filament yarn (150 denier, 48 filaments).",
                "output": json.dumps({
                    "yarn_type": "filament",
                    "filament_count": 48,
                    "linear_density": 150,
                    "linear_density_unit": "denier",
                    "material": "polyester"
                })
            },
            {
                "input": "Cotton spun yarn (30s Ne, 15 TPI) was used for the outer layer.",
                "output": json.dumps({
                    "yarn_type": "spun",
                    "tpi": 15,
                    "material": "cotton"
                })
            }
        ]

        # Create comprehensive extraction prompt
        prompt = ChatPromptTemplate.from_template("""
You are an expert at extracting textile engineering data from research papers.

Extract ALL samples mentioned in the text. For each sample:
1. Identify sample ID (e.g., "S1", "Sample A", "Fabric 1")
2. Determine number of layers
3. For EACH layer, extract:
   - Yarn specifications (type, filament count, TPI, linear density, material)
   - Fabric specifications (density, thickness, pattern, stitch density, EPI/PPI, CPI/WPI, loop length)
   - Materials and blends
   - Treatments/finishes (type, application side, wettability changes, contact angles)
4. Extract performance metrics:
   - Thermal: evaporative resistance, thermal resistance, heat loss, CLO value
   - Liquid transport: one-way transport index (R), wetting times, absorption rates, spreading speeds, OMMC

IMPORTANT:
- If a sample has 3 layers, create 3 separate entries (one per layer)
- Number layers from inside (1) to outside (N)
- Extract exact numerical values with units
- Use null for missing data

Text:
{text}

Return a JSON array of samples. Each sample should follow this structure:
{{
  "sample_id": "S1",
  "study_id": "{study_id}",
  "num_layers": 2,
  "layers": [
    {{
      "layer_number": 1,
      "yarn_specs": {{}},
      "fabric_specs": {{}},
      "material": "polyester",
      "treatment": {{}}
    }}
  ],
  "thermal_metrics": {{}},
  "liquid_transport_metrics": {{}}
}}
""")

        chain = prompt | self.llm
        response = chain.invoke({
            "text": text,
            "study_id": metadata.study_id
        })

        try:
            content = response.content
            # Extract JSON array
            json_match = re.search(r'\[.*\]', content, re.DOTALL)
            if json_match:
                samples_data = json.loads(json_match.group(0))
                samples = [Sample(**s) for s in samples_data]
                print(f"✓ Extracted {len(samples)} samples")
                return samples
            else:
                print(f"⚠ Could not parse samples JSON from response")
                return []
        except Exception as e:
            print(f"⚠ Sample extraction error: {e}")
            return []

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
    """Main pipeline orchestrator"""

    def __init__(self, openai_api_key: str = OPENAI_API_KEY):
        self.preprocessor = PDFPreprocessor()
        self.extractor = LLMExtractor(api_key=openai_api_key)
        self.normalizer = DataNormalizer()

    def process_pdf(self, pdf_path: str, use_nougat: bool = True) -> Tuple[ExtractedData, pd.DataFrame]:
        """
        Complete pipeline: PDF → Extracted Data → DataFrame

        Args:
            pdf_path: Path to PDF file
            use_nougat: Whether to use Nougat for parsing (vs PyMuPDF)

        Returns:
            (ExtractedData, DataFrame) - Structured data and flattened table
        """
        print(f"\n{'='*60}")
        print(f"Processing: {Path(pdf_path).name}")
        print(f"{'='*60}")

        # Stage 1: Preprocess
        text = self.preprocessor.preprocess(pdf_path, use_nougat=use_nougat)

        # Stage 2: Extract
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
    """Example usage"""

    # Check dependencies
    if not OPENAI_API_KEY:
        print("⚠ Warning: OPENAI_API_KEY not set. Set it via environment variable.")
        print("   export OPENAI_API_KEY='your-key-here'")
        return

    # Initialize scraper
    scraper = SciMasterScraper(openai_api_key=OPENAI_API_KEY)

    # Example: Process single PDF
    pdf_path = "example_textile_paper.pdf"
    if Path(pdf_path).exists():
        extracted_data, df = scraper.process_pdf(pdf_path)

        # Save outputs
        output_json = OUTPUT_DIR / f"{Path(pdf_path).stem}_extracted.json"
        output_excel = OUTPUT_DIR / f"{Path(pdf_path).stem}_data.xlsx"

        # Save JSON
        with open(output_json, 'w') as f:
            json.dump(extracted_data.dict(), f, indent=2)
        print(f"✓ Saved JSON to {output_json}")

        # Save Excel
        df.to_excel(output_excel, index=False)
        print(f"✓ Saved Excel to {output_excel}")

        # Display sample
        print("\nSample output (first 5 rows):")
        print(df.head())

    # Example: Batch processing
    pdf_folder = Path("textile_papers")
    if pdf_folder.exists():
        pdf_files = list(pdf_folder.glob("*.pdf"))
        output_excel = OUTPUT_DIR / "all_textile_data.xlsx"

        combined_df = scraper.process_batch(pdf_files, output_excel)
        print(f"\n✓ Processed {len(pdf_files)} PDFs")
        print(f"✓ Total rows: {len(combined_df)}")


if __name__ == "__main__":
    main()
