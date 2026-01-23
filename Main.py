#!/usr/bin/env python
"""
ScholarSweep Main Pipeline Orchestrator - SEQUENTIAL EXECUTION (WINDOWS-SAFE)

EXECUTION STRATEGY:
Pipelines run SEQUENTIALLY (1 → 2 → 3) to avoid Windows spawn deadlocks.
- Pipeline 1 MUST complete before Pipeline 2 spawns multiprocessing.Process
- No ThreadPoolExecutor at top level (would deadlock with Process.start())

PARALLELISM STRATEGY:
Each pipeline has INTERNAL parallelism (no cross-pipeline parallelism needed):
- Pipeline 1: Threading for concurrent downloads
- Pipeline 2: Multiprocessing (GPU owner + 4 CPU workers) + GPU parallelism
- Pipeline 3: Threading for concurrent LLM calls

PIPELINE 1 (Download):
- Pipeline-level: 30 download workers + 10 validation workers (parallel downloading)
- Inter-pipeline: Push to Pipeline 2 IMMEDIATELY when each PDF is validated

PIPELINE 2 (Text Extraction - TWO-TIER ARCHITECTURE):
- RUNS IN MAIN THREAD (cannot use ThreadPoolExecutor - Windows spawn deadlock)
- GPU owner process: ONE process runs all YOLO inference (single CUDA context, no memory leak)
- CPU workers: 4 processes for PDF rendering + text extraction (true parallelism, no GIL)
- Single shared gpu_out queue with routing by worker_id (GPU sends GPU_DONE sentinels)
- Non-daemon processes with proper shutdown protocol (no early GPU exit, no hangs)
- Inter-pipeline: Push to Pipeline 3 IMMEDIATELY after each PDF extraction completes
- Architecture: o1-recommended design for Windows spawn safety + max performance

PIPELINE 3 (LLM Filter):
- Pipeline-level: 4 workers filter multiple papers simultaneously
- Inter-pipeline: N/A (last pipeline in chain)

BENEFITS:
✅ No Windows spawn deadlocks (sequential execution)
✅ Maximum GPU/CPU utilization within each pipeline
✅ Internal parallelism preserved (threading + multiprocessing)
✅ Models loaded once and reused
✅ Simple, debuggable control flow

EXECUTION ORDER:
- 1 → 2 → 3 → (4a || 4b) → 5
- Pipelines 1-3 run sequentially (Windows requirement)
- Pipelines 4a-4b still run in parallel
"""

import os
# CRITICAL: Enable expandable segments to prevent memory fragmentation
os.environ["PYTORCH_CUDA_ALLOC_CONF"] = "expandable_segments:True"

# Suppress warnings
import warnings
warnings.filterwarnings("ignore")
os.environ["PYTHONWARNINGS"] = "ignore"

import sys
import time
import threading
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

# Add pipeline directories to path
sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent / "QueueManager"))

from QueueManager import PipelineQueueManager


def run_pipeline1_with_queue(queue_mgr: PipelineQueueManager, output_base: Path):
    """
    Run Pipeline 1 (Download) with TRUE inter-pipeline streaming.

    Uses validation callback to push papers to Pipeline 2 IMMEDIATELY when validated,
    not after all downloads complete.
    """
    sys.path.insert(0, str(Path(__file__).parent / "1_GenerateCorpus"))

    # Define callback that pushes papers to queue immediately when validated
    def on_pdf_validated(pdf_num, pdf_path):
        """Called by DownloadValidationManager when a PDF is validated."""
        queue_mgr.pipeline1_put(pdf_num, pdf_path)

    # Define callback to set query folder immediately (allows Pipeline 2 to start)
    def on_query_folder_created(query_folder):
        """Called by GenerateCorpus immediately after creating query folder."""
        queue_mgr.set_query_folder(query_folder)

    # Import and run Pipeline 1 with callbacks
    import GenerateCorpus as p1
    query_folder = p1.main(
        on_pdf_validated=on_pdf_validated,
        on_query_folder_created=on_query_folder_created
    )

    # Verify query folder was set
    if query_folder is None:
        raise ValueError("Pipeline 1 did not return query folder")

    # Signal Pipeline 1 is done (all downloads and validations complete)
    queue_mgr.pipeline1_complete()


def run_pipeline2_with_queue(queue_mgr: PipelineQueueManager, output_base: Path):
    """
    Run Pipeline 2 (Text Extraction) with TWO-TIER parallelization:
    - ONE GPU owner process: Runs all YOLO inference (single CUDA context)
    - N CPU worker processes: PDF rendering and text extraction (true parallelism, no GIL)
    - Inter-pipeline: Push each finished PDF to Pipeline 3 immediately
    """
    # Get query folder from Pipeline 1 (waits until it's set)
    query_folder = queue_mgr.get_query_folder(timeout=60)

    sys.path.insert(0, str(Path(__file__).parent / "2_NativeTextExtraction"))

    # Import two-tier architecture modules
    from multiprocessing import Process, Queue
    from gpu_worker import gpu_worker
    from cpu_worker import cpu_worker

    # Track timing and results
    extraction_start = time.time()
    all_table_detections = []
    total_metrics_time = 0.0

    # Setup
    model_path = Path(__file__).parent / "2_NativeTextExtraction" / "yolov12n_Model.pt"
    output_dir = query_folder / "Extracted Text"
    output_dir.mkdir(exist_ok=True)

    # Create queues for two-tier architecture with PER-WORKER output queues
    NUM_CPU_WORKERS = 4

    # Use regular Queue objects (passed as Process constructor args - no Manager overhead)
    gpu_in = Queue(maxsize=16)     # PageBatch objects to GPU (increased for better flow)
    gpu_out_queues = [Queue(maxsize=32) for _ in range(NUM_CPU_WORKERS)]  # Per-worker output queues
    result_queue = Queue()         # Results from CPU workers to orchestrator

    # Input queues for each worker
    pdf_queues = [Queue(maxsize=8) for _ in range(NUM_CPU_WORKERS)]

    cpu_workers = []
    STOP = "__STOP__"

    # Start GPU owner process with dict of per-worker output queues
    gpu_out_dict = {i: gpu_out_queues[i] for i in range(NUM_CPU_WORKERS)}
    gpu_proc = Process(
        target=gpu_worker,
        args=(str(model_path), gpu_in, gpu_out_dict, STOP),
    )
    gpu_proc.start()

    # Start CPU worker processes (each gets its own dedicated output queue)
    for worker_id in range(NUM_CPU_WORKERS):
        p = Process(
            target=cpu_worker,
            args=(worker_id,
                  pdf_queues[worker_id],
                  gpu_in,
                  gpu_out_queues[worker_id],  # dedicated output queue
                  result_queue,
                  STOP),
        )
        p.start()
        cpu_workers.append(p)

    papers_submitted = 0
    results_received = 0
    next_worker = 0  # Round-robin worker assignment

    def handle_result(result):
        """Process a result from a CPU worker."""
        nonlocal total_metrics_time, all_table_detections

        if result["error"]:
            return

        # Accumulate metrics match time
        total_metrics_time += result.get("metrics_time", 0.0)

        # Collect table detections
        if result["table_detections"]:
            all_table_detections.extend(result["table_detections"])

        # Push to Pipeline 3 immediately
        text_path = output_dir / f"{result['pdf_num']}.txt"
        filtered_path = output_dir / f"{result['pdf_num']}_filtered.txt"

        queue_mgr.pipeline2_put(
            pdf_num=result['pdf_num'],
            text_path=text_path,
            filtered_path=filtered_path
        )

    try:
        while True:
            # Get paper from Pipeline 1 queue (non-blocking check)
            paper = queue_mgr.pipeline2_get(timeout=0.5)

            if paper is not None:
                # Submit PDF to CPU worker (round-robin)
                pdf_queues[next_worker].put((
                    paper['pdf_num'],
                    str(paper['pdf_path']),
                    str(output_dir)
                ))
                papers_submitted += 1
                next_worker = (next_worker + 1) % NUM_CPU_WORKERS

            # Check for results from CPU workers (non-blocking)
            while not result_queue.empty():
                result = result_queue.get_nowait()
                results_received += 1
                handle_result(result)

            # Exit condition: Pipeline 1 done AND no more papers in queue
            if queue_mgr.pipeline1_done.is_set() and paper is None:
                break

        # CRITICAL: Wait for ALL results BEFORE sending any STOP tokens
        # Workers are still processing - GPU is still needed!
        while results_received < papers_submitted:
            try:
                result = result_queue.get(timeout=30.0)  # BLOCKING wait
                results_received += 1
                handle_result(result)
            except Exception:
                pass

        # NOW send stop tokens (workers are idle, waiting for next PDF)
        for worker_id in range(NUM_CPU_WORKERS):
            pdf_queues[worker_id].put(STOP)

        # Stop GPU owner process (it will send GPU_DONE to all workers)
        gpu_in.put(STOP)
        gpu_proc.join()

        # Wait for CPU workers to exit
        for p in cpu_workers:
            p.join()

    finally:
        # Ensure all processes are terminated
        if gpu_proc.is_alive():
            gpu_proc.terminate()
            gpu_proc.join()
        for p in cpu_workers:
            if p.is_alive():
                p.terminate()
                p.join()

    extraction_time = time.time() - extraction_start

    # Save timers
    from Timers.timing_utils import save_timer
    if extraction_time > 0:
        save_timer(query_folder, "5_text_extraction_total", extraction_time)

    if total_metrics_time > 0:
        save_timer(query_folder, "6_metrics_match_total", total_metrics_time)

    # Save table detections JSON
    if all_table_detections:
        table_detections_json = query_folder / "yolo_table_detections.json"
        import json
        with open(table_detections_json, 'w', encoding='utf-8') as f:
            json.dump(all_table_detections, f, indent=2)

    # Signal Pipeline 2 is done
    queue_mgr.pipeline2_complete()


def run_pipeline3_with_queue(queue_mgr: PipelineQueueManager, output_base: Path):
    """
    Run Pipeline 3 (LLM Filter) with DUAL parallelization:
    - Pipeline-level: Filter 4 papers simultaneously using ThreadPoolExecutor
    - Inter-pipeline: Process papers as they arrive from Pipeline 2
    """
    # Get query folder from Pipeline 1 (waits until it's set)
    query_folder = queue_mgr.get_query_folder(timeout=60)

    sys.path.insert(0, str(Path(__file__).parent / "3_LLMFilter"))

    # Import filtering function and utilities
    from llm_filter import filter_single_paper_for_queue
    from Utils.corpus_utils import load_corpus_json, save_corpus_json
    from concurrent.futures import ThreadPoolExecutor, as_completed

    # Wait for corpus_metadata.json to be created by Pipeline 1
    metadata_path = query_folder / "corpus_metadata.json"

    while not metadata_path.exists():
        if queue_mgr.pipeline1_done.is_set() and not metadata_path.exists():
            queue_mgr.pipeline3_complete()
            return
        time.sleep(1)

    # Load metadata once (updated in-place by worker function)
    all_papers = load_corpus_json(query_folder)

    # Track results
    filtered_out_pdfs = set()
    total_llm_time = 0.0
    papers_submitted = 0

    # Create persistent executor with 4 workers
    executor = ThreadPoolExecutor(max_workers=4)

    # Track pending futures
    future_to_paper = {}  # future -> paper dict

    try:
        while True:
            # Get paper from queue (non-blocking check)
            paper = queue_mgr.pipeline3_get(timeout=0.5)

            if paper is not None:
                # Submit paper to executor immediately
                future = executor.submit(
                    filter_single_paper_for_queue,
                    paper['pdf_num'],
                    paper['text_path'],
                    query_folder,
                    all_papers
                )
                future_to_paper[future] = paper
                papers_submitted += 1

            # Check for completed results (don't block)
            completed_futures = []
            for future in list(future_to_paper.keys()):
                if future.done():
                    completed_futures.append(future)
                    paper = future_to_paper[future]
                    try:
                        result = future.result()

                        # Track LLM time
                        total_llm_time += result["llm_time"]

                        # Track filtered out PDFs
                        if result["filtered_out"]:
                            filtered_out_pdfs.add(result["pdf_num"])

                        # Mark task done
                        queue_mgr.pipeline3_task_done()

                    except Exception:
                        pass

            # Remove completed futures
            for future in completed_futures:
                del future_to_paper[future]

            # Exit condition: Pipeline 2 done AND no more papers in queue AND no pending futures
            if queue_mgr.pipeline2_done.is_set() and paper is None and len(future_to_paper) == 0:
                break

    finally:
        # Clean up executor
        executor.shutdown(wait=True)

    # Save updated metadata
    save_corpus_json(all_papers, query_folder)

    # Save LLM timer
    if total_llm_time > 0:
        from Timers.timing_utils import save_timer
        save_timer(query_folder, "7_llm_filter_total", total_llm_time)

    # Clean up table detections JSON
    if filtered_out_pdfs:
        import json
        table_detections_json = query_folder / "yolo_table_detections.json"

        if table_detections_json.exists():
            # Load table detections
            with open(table_detections_json, 'r', encoding='utf-8') as f:
                all_detections = json.load(f)

            # Filter out detections for filtered-out PDFs
            kept_detections = [
                det for det in all_detections
                if det.get("pdf_num") not in filtered_out_pdfs
            ]

            # Save cleaned detections
            with open(table_detections_json, 'w', encoding='utf-8') as f:
                json.dump(kept_detections, f, indent=2)

    # Print Pipeline 3 summary
    passed_papers = [p for p in all_papers if not p.get("breakpoint", "").strip()]
    print(f"Papers filtered for relevance: {len(filtered_out_pdfs)}")
    print(f"Final Corpus Size: {len(passed_papers)}")

    # Signal Pipeline 3 is done
    queue_mgr.pipeline3_complete()


def main():
    overall_start = time.time()

    # Create queue manager
    queue_mgr = PipelineQueueManager()

    # Get/create output directory
    output_base = Path(__file__).parent / "Output"
    output_base.mkdir(parents=True, exist_ok=True)

    # =====================================================================
    # PIPELINE 1 — GenerateCorpus (threading internally)
    # MUST complete BEFORE Pipeline 2 spawns processes
    # =====================================================================
    run_pipeline1_with_queue(queue_mgr, output_base)

    # =====================================================================
    # PIPELINE 2 — NativeTextExtract (multiprocessing, GPU + CPU)
    # MUST run with NO other threads active (Windows spawn requirement)
    # =====================================================================
    run_pipeline2_with_queue(queue_mgr, output_base)

    # =====================================================================
    # PIPELINE 3 — LLMFilter (threading internally)
    # =====================================================================
    run_pipeline3_with_queue(queue_mgr, output_base)

    # ========================================================================
    # Pipeline 4a & 4b: Run in parallel (load models once)
    # ========================================================================
    # Get query folder from Pipeline 1
    query_folder = queue_mgr.query_folder

    pdf_folder_4b = query_folder / "Corpus PDFs"

    # Load Surya models once for Pipeline 4b
    import torch
    import gc
    from surya.table_rec import TableRecPredictor
    from surya.detection import DetectionPredictor

    # CRITICAL: Force garbage collection and clear GPU memory from Pipeline 2
    # (yolo_model from Pipeline 2 is out of scope but may not be freed yet)
    gc.collect()
    torch.cuda.empty_cache()

    table_rec = TableRecPredictor()
    det_predictor = DetectionPredictor()

    # CRITICAL: Set models to eval mode to prevent autograd memory leak
    table_rec.model.eval()
    det_predictor.model.eval()

    # Run 4a and 4b in parallel
    import concurrent.futures

    # Import pipeline functions
    sys.path.insert(0, str(Path(__file__).parent / "4a_LLMTextExtract"))
    from llm_extract_values import extract_all_values as pipeline4a_extract

    sys.path.insert(0, str(Path(__file__).parent / "4b_TableOCR"))
    from table_extraction import run_pipeline_4b

    # Get extracted text folder for Pipeline 4a
    extracted_text_folder = query_folder / "Extracted Text"

    # Wrapper to suppress stderr for Pipeline 4b (Surya progress bars)
    def run_4b_silent(pdf_folder, table_rec_model, det_model):
        import contextlib
        import io
        with contextlib.redirect_stderr(io.StringIO()):
            return run_pipeline_4b(pdf_folder, table_rec_model, det_model)

    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        # Submit both pipelines with correct paths
        future_4a = executor.submit(pipeline4a_extract, extracted_text_folder)

        if pdf_folder_4b and pdf_folder_4b.exists():
            future_4b = executor.submit(run_4b_silent, pdf_folder_4b, table_rec, det_predictor)
        else:
            future_4b = None

        # Wait for both to complete
        try:
            future_4a.result()
            if future_4b:
                future_4b.result()
        except Exception as e:
            raise

    # Clear GPU memory after 4b completes
    torch.cuda.empty_cache()
    del table_rec
    del det_predictor

    # ========================================================================
    # Pipeline 5: GenerateOutputs
    # ========================================================================
    sys.path.insert(0, str(Path(__file__).parent / "5_GenerateOutputs"))
    from GenerateOutputs import run_combiner as pipeline5_run
    pipeline5_run(query_folder)

if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    main()
