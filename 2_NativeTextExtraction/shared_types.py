"""Shared types for GPU/CPU worker communication."""

from dataclasses import dataclass
from typing import List, Tuple
import numpy as np

# GPU completion sentinel
GPU_DONE = "__GPU_DONE__"


@dataclass
class PageBatch:
    """Batch of page images sent to GPU worker."""
    batch_id: int
    pdf_num: int
    worker_id: int  # which worker should receive results
    page_indices: List[int]  # page numbers in this PDF
    images: List[np.ndarray]  # H x W x 3 uint8


@dataclass
class DetectionResult:
    """Detection results from GPU worker."""
    batch_id: int
    pdf_num: int
    worker_id: int  # which worker this result belongs to
    page_indices: List[int]
    # per-page list of list of (x1, y1, x2, y2, cls_id, conf)
    boxes_per_page: List[List[Tuple[float, float, float, float, int, float]]]
