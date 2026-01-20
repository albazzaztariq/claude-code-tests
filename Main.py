"""
ScholarSweep Main Pipeline Orchestrator

Simple orchestrator that calls pipelines in sequence:
1. QueueManager - Create all queues
2. Pipeline 1 - GenerateCorpus
3. Pipeline 2 - NativeTextExtract
4. Pipeline 3 - LLMFilter
5. Pipeline 4a & 4b - Run in parallel (LLM + Table extraction)
6. Pipeline 5 - GenerateOutputs (combine results)
"""

import subprocess
import sys
from pathlib import Path

# Pipeline paths
QUEUE_MANAGER = Path(__file__).parent / "QueueManager" / "QueueManager.py"
PIPELINE_1 = Path(__file__).parent / "1_GenerateCorpus" / "GenerateCorpus.py"
PIPELINE_2 = Path(__file__).parent / "2_NativeTextExtraction" / "NativeTextExtract_pipeline.py"
PIPELINE_3 = Path(__file__).parent / "3_LLMFilter" / "llm_filter.py"
PIPELINE_4A = Path(__file__).parent / "4a_LLMTextExtract" / "llm_extract_values.py"
PIPELINE_4B = Path(__file__).parent / "4b_TableOCR" / "pipeline_4b.py"
PIPELINE_5 = Path(__file__).parent / "5_GenerateOutputs" / "GenerateOutputs.py"

def main():
    print("=" * 80)
    print("SCHOLARSWEEP PIPELINE ORCHESTRATOR")
    print("=" * 80)

    # QueueManager: Create queues
    print("\n[QUEUE MANAGER] Creating queues...")
    subprocess.run([sys.executable, str(QUEUE_MANAGER)], check=True)

    # Pipeline 1: GenerateCorpus
    print("\n[PIPELINE 1] GenerateCorpus - API search and download...")
    subprocess.run([sys.executable, str(PIPELINE_1)], check=True)

    # Pipeline 2: NativeTextExtract
    print("\n[PIPELINE 2] NativeTextExtract - Extract text from PDFs...")
    subprocess.run([sys.executable, str(PIPELINE_2)], check=True)

    # Pipeline 3: LLMFilter
    print("\n[PIPELINE 3] LLMFilter - Filter irrelevant papers...")
    subprocess.run([sys.executable, str(PIPELINE_3)], check=True)

    # Pipeline 4a & 4b: Run in parallel
    print("\n[PIPELINE 4a & 4b] Running LLM + Table extraction in parallel...")
    import concurrent.futures
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        future_4a = executor.submit(subprocess.run, [sys.executable, str(PIPELINE_4A)], check=True)
        future_4b = executor.submit(subprocess.run, [sys.executable, str(PIPELINE_4B)], check=True)
        future_4a.result()
        future_4b.result()

    # Pipeline 5: GenerateOutputs
    print("\n[PIPELINE 5] GenerateOutputs - Combine results...")
    subprocess.run([sys.executable, str(PIPELINE_5)], check=True)

    print("\n" + "=" * 80)
    print("PIPELINE COMPLETE")
    print("=" * 80)

if __name__ == "__main__":
    main()
