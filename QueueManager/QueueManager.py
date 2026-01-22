"""
Download and Validation Queue System with Rolling Window

ARCHITECTURE:
1. Paper Queue (500 papers) - Workers pull on-demand
2. Download Workers (20) - Each downloads one paper's mirrors sequentially
3. Validation Queue - Files waiting for validation
4. Validation Workers (10) - Check format and assign numbers with lock

ROLLING WINDOW:
- 20 workers actively downloading at all times
- As Worker X finishes Paper N, immediately grabs Paper N+20
- Maintains constant throughput
"""

import threading
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from queue import Queue, Empty
import time
import fitz  # PyMuPDF
import uuid


class DownloadValidationManager:
    """Manages download and validation with rolling window of workers."""

    def __init__(self, staging_folder, corpus_folder, mirror_speeds,
                 max_download_workers=20, max_validation_workers=10, verbose=False):
        """
        Initialize queue manager.

        Args:
            staging_folder: Path to Staging Downloads folder
            corpus_folder: Path to Corpus PDFs folder
            mirror_speeds: Dict mapping domain -> speed (bytes/sec)
            max_download_workers: Number of download workers (default: 20)
            max_validation_workers: Number of validation workers (default: 10)
            verbose: Enable debug output for downloads/validation (default: False)
        """
        self.staging_folder = Path(staging_folder)
        self.corpus_folder = Path(corpus_folder)
        self.mirror_speeds = mirror_speeds
        self.max_download_workers = max_download_workers
        self.max_validation_workers = max_validation_workers
        self.verbose = verbose

        # Paper Queue - all papers to download
        self.paper_queue = Queue()

        # Validation Queue - files waiting for validation
        self.validation_queue = Queue()

        # Thread-safe counter for study numbers
        self.study_counter = 1
        self.counter_lock = threading.Lock()

        # Stats tracking
        self.downloaded_count = 0
        self.failed_download_count = 0
        self.valid_pdf_count = 0
        self.paywall_count = 0
        self.unreadable_count = 0
        self.total_bytes = 0

        # Lock for stats
        self.stats_lock = threading.Lock()

        # Flag to signal workers to stop
        self.stop_flag = threading.Event()

    def download_worker(self, worker_id, download_func):
        """
        Download worker - pulls papers from queue and tries all mirrors.

        Args:
            worker_id: Worker identifier for debugging
            download_func: Function to download from URL
        """
        while not self.stop_flag.is_set():
            try:
                # Get next paper (timeout to allow checking stop_flag)
                paper = self.paper_queue.get(timeout=1)
            except Empty:
                # Queue empty, worker done
                break

            temp_idx = paper.get("temp_index", "unknown")

            # Generate unique filename for staging
            temp_uuid = str(uuid.uuid4())
            staging_filename = f"temp_{temp_uuid}.pdf"
            staging_path = self.staging_folder / staging_filename

            # Try all mirrors sequentially until we find a PDF
            success = False
            best_non_pdf = None

            locations = paper.get("oa_locations", [])
            if not locations:
                # No locations available
                paper["status"] = "failed"
                with self.stats_lock:
                    self.failed_download_count += 1
                self.paper_queue.task_done()
                continue

            # Sort locations by speed (fastest first)
            def get_speed(loc):
                url = loc.get("url_for_pdf") or loc.get("url") or ""
                from urllib.parse import urlparse
                try:
                    parsed = urlparse(url)
                    domain = parsed.netloc.replace('www.', '')
                    return self.mirror_speeds.get(domain, 0)
                except:
                    return 0

            sorted_locations = sorted(locations, key=get_speed, reverse=True)

            if self.verbose:
                print(f"\n[DOWNLOAD] Paper {temp_idx}: Starting download ({len(sorted_locations)} mirrors)")

            # Try each mirror
            for mirror_num, loc in enumerate(sorted_locations, 1):
                url = loc.get("url_for_pdf") or loc.get("url")
                if not url:
                    continue

                # Extract mirror name for debug output
                if self.verbose:
                    from urllib.parse import urlparse
                    parsed = urlparse(url)
                    mirror_name = parsed.netloc.replace('www.', '')
                    print(f"[DOWNLOAD] Paper {temp_idx}: Trying mirror #{mirror_num}/{len(sorted_locations)} ({mirror_name})")

                # Download from this mirror
                dl_success, size, file_format, source_name, timeout = download_func(url, staging_path, temp_idx)

                if self.verbose:
                    print(f"[DOWNLOAD] Paper {temp_idx}: Mirror #{mirror_num} result: {file_format if dl_success else 'TIMEOUT' if timeout else 'FAILED'}")

                if dl_success:
                    # Check format
                    if file_format in ["HTML", "XML", "JSON"]:
                        # Not a PDF - delete and save as fallback
                        if staging_path.exists():
                            staging_path.unlink()
                        if best_non_pdf is None:
                            best_non_pdf = (dl_success, size, file_format, source_name, timeout)
                        # Try next mirror
                        continue
                    elif file_format == "UNKNOWN":
                        # Unknown format - delete and save as fallback
                        if staging_path.exists():
                            staging_path.unlink()
                        if best_non_pdf is None:
                            best_non_pdf = (dl_success, size, file_format, source_name, timeout)
                        # Try next mirror
                        continue
                    else:
                        # PDF found! Add to validation queue
                        paper["download_source"] = source_name
                        with self.stats_lock:
                            self.downloaded_count += 1
                            self.total_bytes += size

                        if self.verbose:
                            print(f"[DOWNLOAD] Paper {temp_idx}: PDF found! Added to validation queue ({source_name})")

                        # Add to validation queue
                        self.validation_queue.put((staging_path, paper, file_format))
                        success = True
                        break
                elif timeout:
                    # Track timeout - paper["timeout_occurred"] will be checked later
                    paper["timeout_occurred"] = True

            # If no PDF found, handle fallback
            if not success:
                if best_non_pdf:
                    # Re-download the best non-PDF (we deleted it earlier)
                    fallback_success = False
                    for loc in sorted_locations:
                        url = loc.get("url_for_pdf") or loc.get("url")
                        if not url:
                            continue
                        dl_success, size, file_format, source_name, timeout = download_func(url, staging_path, temp_idx)
                        if dl_success and file_format == best_non_pdf[2]:
                            # Got the fallback file
                            paper["download_source"] = source_name
                            with self.stats_lock:
                                self.downloaded_count += 1
                                self.total_bytes += size
                            # Add to validation queue (will be rejected as paywall)
                            self.validation_queue.put((staging_path, paper, file_format))
                            fallback_success = True
                            break
                        elif timeout:
                            paper["timeout_occurred"] = True

                    if not fallback_success:
                        # Fallback download failed
                        paper["status"] = "failed"
                        with self.stats_lock:
                            self.failed_download_count += 1
                else:
                    # All downloads failed
                    paper["status"] = "failed"
                    with self.stats_lock:
                        self.failed_download_count += 1

            self.paper_queue.task_done()

    def validation_worker(self, worker_id):
        """
        Validation worker - validates files and assigns numbers.

        Args:
            worker_id: Worker identifier for debugging
        """
        while not self.stop_flag.is_set():
            try:
                # Get next file to validate (timeout to allow checking stop_flag)
                staging_path, paper, file_format = self.validation_queue.get(timeout=1)
            except Empty:
                # Check if downloads are still ongoing
                if self.paper_queue.unfinished_tasks == 0:
                    break
                continue

            try:
                # Check if it's a non-PDF file (HTML, XML, JSON, UNKNOWN)
                if file_format in ["HTML", "XML", "JSON", "UNKNOWN"]:
                    # Not a PDF - delete
                    if staging_path.exists():
                        staging_path.unlink()
                    paper["status"] = "Not PDF"
                    paper["file_format"] = file_format  # Store actual format for display
                    with self.stats_lock:
                        self.paywall_count += 1
                    self.validation_queue.task_done()
                    continue

                # Validate with PyMuPDF
                doc = fitz.open(staging_path)
                page_count = len(doc)

                # Check if file is readable
                has_text = False
                try:
                    if page_count > 0:
                        first_page_text = doc[0].get_text()
                        has_text = len(first_page_text.strip()) > 0
                except:
                    has_text = False
                finally:
                    doc.close()

                # Determine if valid (UNKNOWN already rejected above)
                is_valid = page_count > 0 and has_text

                if is_valid:
                    # Valid PDF - use pre-assigned study number
                    study_num = paper.get("study_number")
                    if not study_num:
                        # Fallback if somehow missing (shouldn't happen)
                        with self.counter_lock:
                            study_num = self.study_counter
                            self.study_counter += 1
                            paper["study_number"] = study_num

                    # Move to Corpus PDFs with assigned number
                    final_path = self.corpus_folder / f"{study_num}.pdf"
                    staging_path.rename(final_path)

                    if self.verbose:
                        temp_idx = paper.get("temp_index", "unknown")
                        source = paper.get("download_source", "unknown")
                        print(f"[VALIDATE] Paper {temp_idx}: Validated and renamed to {study_num}.pdf (source: {source})")

                    # Update paper status
                    if file_format == "UNKNOWN":
                        paper["status"] = "unknown"
                    else:
                        paper["status"] = "downloaded"

                    with self.stats_lock:
                        self.valid_pdf_count += 1
                else:
                    # Invalid PDF - delete
                    if staging_path.exists():
                        staging_path.unlink()
                    paper["status"] = "unreadable"
                    with self.stats_lock:
                        self.unreadable_count += 1

            except Exception as e:
                # Error during validation - delete file
                if staging_path.exists():
                    staging_path.unlink()
                paper["status"] = "unreadable"
                with self.stats_lock:
                    self.unreadable_count += 1

            self.validation_queue.task_done()

    def process_papers(self, papers, download_func, progress_callback=None):
        """
        Process all papers with rolling window of workers.

        Args:
            papers: List of paper dicts to download
            download_func: Function(url, filepath, temp_idx) -> (success, size, file_format, source_name)
            progress_callback: Optional callback(downloaded, failed, validated) for progress updates

        Returns:
            list: Updated papers list with study_numbers assigned
        """
        total_papers = len(papers)

        # Print setup info FIRST (before adding to queue and starting workers)
        print(f"\n[QUEUE 1] Paper Queue: {total_papers} papers")
        print(f"[WORKERS] Starting {self.max_download_workers} download workers...")
        print(f"[WORKERS] Starting {self.max_validation_workers} validation workers...")

        # Add all papers to queue AFTER printing
        for paper in papers:
            self.paper_queue.put(paper)

        # Start download workers
        download_executor = ThreadPoolExecutor(max_workers=self.max_download_workers)
        download_futures = []
        for i in range(self.max_download_workers):
            future = download_executor.submit(self.download_worker, i, download_func)
            download_futures.append(future)

        # Start validation workers
        validation_executor = ThreadPoolExecutor(max_workers=self.max_validation_workers)
        validation_futures = []
        for i in range(self.max_validation_workers):
            future = validation_executor.submit(self.validation_worker, i)
            validation_futures.append(future)

        # Monitor progress
        while self.paper_queue.unfinished_tasks > 0 or self.validation_queue.unfinished_tasks > 0:
            if progress_callback:
                progress_callback(self.downloaded_count, self.failed_download_count, self.valid_pdf_count)
            time.sleep(0.5)

        # Signal workers to stop
        self.stop_flag.set()

        # Wait for all workers to finish
        download_executor.shutdown(wait=True)
        validation_executor.shutdown(wait=True)

        # Final progress update
        if progress_callback:
            progress_callback(self.downloaded_count, self.failed_download_count, self.valid_pdf_count)

        return papers

    def get_stats(self):
        """Get download and validation statistics."""
        return {
            "downloaded": self.downloaded_count,
            "failed_downloads": self.failed_download_count,
            "valid_pdfs": self.valid_pdf_count,
            "paywall_files": self.paywall_count,
            "unreadable_files": self.unreadable_count,
            "total_bytes": self.total_bytes
        }
