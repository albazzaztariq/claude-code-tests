"""GPU owner process - single process that owns CUDA + YOLO."""

import torch
from ultralytics import YOLO
from multiprocessing import Queue
from shared_types import PageBatch, DetectionResult, GPU_DONE


def gpu_worker(model_path: str,
               in_queue: Queue,
               out_queues: dict,
               stop_token: str = "__STOP__",
               device: str = "cuda:0",
               conf_thres: float = 0.175):
    """
    Single process that owns CUDA + YOLO.
    Receives PageBatch objects, returns DetectionResult objects.

    Args:
        model_path: Path to YOLO model
        in_queue: Queue receiving PageBatch objects
        out_queues: Dict mapping worker_id -> Queue (per-worker output queues)
        stop_token: String to signal shutdown
        device: CUDA device
        conf_thres: Confidence threshold for detections
    """
    # Track which workers have sent batches
    active_workers = set()

    # Disable gradient computation globally
    torch.set_grad_enabled(False)

    # Load model once
    model = YOLO(model_path)
    model.to(device)
    model.model.eval()

    batches_processed = 0

    while True:
        item = in_queue.get()
        if item == stop_token:
            break

        assert isinstance(item, PageBatch), f"Expected PageBatch, got {type(item)}"

        # Track active workers
        active_workers.add(item.worker_id)

        images = item.images  # list of np arrays

        # Run YOLO inference
        with torch.no_grad():
            results = model.predict(
                images,
                imgsz=640,
                conf=conf_thres,
                device=device,
                verbose=False
            )

        # Extract boxes to CPU
        boxes_per_page = []
        for res in results:
            page_boxes = []
            for box in res.boxes:
                x1, y1, x2, y2 = box.xyxy[0].tolist()
                cls_id = int(box.cls[0])
                conf = float(box.conf[0])
                page_boxes.append((x1, y1, x2, y2, cls_id, conf))
            boxes_per_page.append(page_boxes)

        # Release GPU tensors
        del results

        # Send results directly to worker's dedicated queue (no routing needed)
        det = DetectionResult(
            batch_id=item.batch_id,
            pdf_num=item.pdf_num,
            worker_id=item.worker_id,
            page_indices=item.page_indices,
            boxes_per_page=boxes_per_page
        )

        out_queues[item.worker_id].put(det)

        batches_processed += 1

    # CRITICAL: Send GPU_DONE sentinel to ALL workers (not just active ones)
    # Workers that never got PDFs still need to exit
    num_workers = len(out_queues)
    for wid in range(num_workers):
        out_queues[wid].put((GPU_DONE, wid))

    # Cleanup
    del model
    torch.cuda.empty_cache()
