# -*- coding: utf-8 -*-
"""Screenspace analysis engine for clipgen.

Provides image analysis primitives, analysis workflows, a background
task queue worker, and manifest persistence for the Screenspace feature.
"""

from __future__ import annotations

import copy
import difflib
import json
import queue
import re
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import cv2
import imagehash
import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity as ssim

import config
import utils
import video


# ---------------------------------------------------------------------------
# Analysis primitives
# ---------------------------------------------------------------------------


def extract_region(frame: np.ndarray, region: Dict[str, int]) -> np.ndarray:
    """Crop a rectangular region from a frame.

    Args:
        frame: BGR image as numpy array (H, W, 3).
        region: Dict with keys ``x``, ``y``, ``w``, ``h`` in pixels.

    Returns:
        Cropped region as numpy array.
    """
    h_frame, w_frame = frame.shape[:2]
    x = max(0, region["x"])
    y = max(0, region["y"])
    x2 = min(w_frame, x + region["w"])
    y2 = min(h_frame, y + region["h"])
    return frame[y:y2, x:x2]


def average_color_hsv(region_pixels: np.ndarray) -> Dict[str, float]:
    """Compute mean HSV color of a region.

    Args:
        region_pixels: BGR image region as numpy array.

    Returns:
        Dict with keys ``h`` (0-180), ``s`` (0-255), ``v`` (0-255).
    """
    hsv = cv2.cvtColor(region_pixels, cv2.COLOR_BGR2HSV)
    mean = np.mean(hsv, axis=(0, 1))
    return {"h": float(mean[0]), "s": float(mean[1]), "v": float(mean[2])}


def color_matches(
    region_pixels: np.ndarray,
    target_color: Dict[str, float],
    tolerance: Dict[str, float],
) -> bool:
    """Check if region's average HSV color is within tolerance of target.

    Handles hue wraparound (red at 0/180 boundary).
    """
    avg = average_color_hsv(region_pixels)
    hue_diff = abs(avg["h"] - target_color["h"])
    hue_ok = min(hue_diff, 180.0 - hue_diff) <= tolerance["h"]
    s_ok = abs(avg["s"] - target_color["s"]) <= tolerance["s"]
    v_ok = abs(avg["v"] - target_color["v"]) <= tolerance["v"]
    return hue_ok and s_ok and v_ok


def compute_frame_diff(
    region_a: np.ndarray,
    region_b: np.ndarray,
    noise_threshold: int = 0,
) -> float:
    """Compute pixel difference ratio between two same-sized regions.

    Applies Gaussian blur, thresholds noise, and morphological opening.

    Returns:
        Change ratio 0.0-1.0 (fraction of pixels that changed).
    """
    if noise_threshold <= 0:
        noise_threshold = config.SCREENSPACE_NOISE_THRESHOLD
    k = config.SCREENSPACE_BLUR_KERNEL
    a_blur = cv2.GaussianBlur(region_a, (k, k), 0)
    b_blur = cv2.GaussianBlur(region_b, (k, k), 0)
    a_gray = cv2.cvtColor(a_blur, cv2.COLOR_BGR2GRAY)
    b_gray = cv2.cvtColor(b_blur, cv2.COLOR_BGR2GRAY)
    diff = cv2.absdiff(a_gray, b_gray)
    _, mask = cv2.threshold(diff, noise_threshold, 255, cv2.THRESH_BINARY)
    mk = config.SCREENSPACE_MORPH_KERNEL
    kernel = np.ones((mk, mk), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)
    if mask.size == 0:
        return 0.0
    return float(np.count_nonzero(mask)) / float(mask.size)


def regions_are_similar(
    region_a: np.ndarray,
    region_b: np.ndarray,
    threshold: float = 0.0,
) -> Tuple[bool, float]:
    """SSIM-based similarity check with blur preprocessing.

    Returns:
        Tuple of (is_similar, ssim_score).
    """
    if threshold <= 0.0:
        threshold = config.SCREENSPACE_SSIM_THRESHOLD
    k = config.SCREENSPACE_BLUR_KERNEL
    a_blur = cv2.GaussianBlur(region_a, (k, k), 0)
    b_blur = cv2.GaussianBlur(region_b, (k, k), 0)
    a_gray = cv2.cvtColor(a_blur, cv2.COLOR_BGR2GRAY)
    b_gray = cv2.cvtColor(b_blur, cv2.COLOR_BGR2GRAY)
    score = float(ssim(a_gray, b_gray))
    return score >= threshold, score


def compute_phash(region_pixels: np.ndarray) -> imagehash.ImageHash:
    """Compute perceptual hash of a region for fast similarity scanning."""
    rgb = cv2.cvtColor(region_pixels, cv2.COLOR_BGR2RGB)
    pil_img = Image.fromarray(rgb)
    return imagehash.phash(pil_img)


def scan_video_frames(
    video_path: str,
    region: Dict[str, int],
    interval_seconds: float,
    callback: Callable[[float, np.ndarray], Optional[bool]],
    *,
    start_seconds: float = 0.0,
    end_seconds: Optional[float] = None,
) -> None:
    """Iterate through video at interval, extract region, call callback.

    The *callback* receives ``(timestamp_seconds, region_pixels)`` and may
    return ``False`` to stop iteration early.
    """
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return

    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    duration = total_frames / fps if fps > 0 else 0.0
    if end_seconds is None or end_seconds > duration:
        end_seconds = duration

    ts = start_seconds
    while ts <= end_seconds:
        cap.set(cv2.CAP_PROP_POS_MSEC, ts * 1000.0)
        ret, frame = cap.read()
        if not ret:
            break
        cropped = extract_region(frame, region)
        result = callback(ts, cropped)
        if result is False:
            break
        ts += interval_seconds

    cap.release()


def build_timelapse_command(
    video_path: str,
    region: Dict[str, int],
    speedup_factor: float,
    output_path: str,
    output_format: str = "mp4",
) -> List[str]:
    """Construct ffmpeg argv for a cropped timelapse."""
    x, y, w, h = region["x"], region["y"], region["w"], region["h"]
    vf = f"crop={w}:{h}:{x}:{y},setpts=PTS/{speedup_factor}"

    cmd = [
        "ffmpeg",
        "-y",
        "-loglevel",
        config.FFMPEG_LOGLEVEL,
        "-i",
        video_path,
        "-vf",
        vf,
        "-an",
    ]

    if output_format == "gif":
        cmd.extend(["-loop", "0"])
    else:
        cmd.extend(["-c:v", "libx264", "-preset", "fast", "-crf", "23"])

    cmd.append(output_path)
    return cmd


# ---------------------------------------------------------------------------
# Analysis workflows
# ---------------------------------------------------------------------------


def scan_color(
    video_path: str,
    region: Dict[str, int],
    target_color: Dict[str, float],
    tolerance: Dict[str, float],
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: Optional[float] = None,
    on_progress: Optional[Callable[[float], None]] = None,
    cancel_flag: Optional[Callable[[], bool]] = None,
) -> List[Dict[str, Any]]:
    """Scan video for frames where region color matches target.

    Returns list of ``{start, end, duration}`` spans (consecutive matches
    merged).
    """
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return []
    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    duration = total_frames / fps if fps > 0 else 0.0
    cap.release()

    if end_seconds is None or end_seconds > duration:
        end_seconds = duration
    total_range = end_seconds - start_seconds

    matches: List[float] = []

    def _cb(ts: float, pixels: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False
        if color_matches(pixels, target_color, tolerance):
            matches.append(ts)
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
    )

    if on_progress:
        on_progress(1.0)
    return _merge_timestamp_spans(matches, interval_seconds)


def scan_changes(
    video_path: str,
    region: Dict[str, int],
    threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    noise_threshold: int = 0,
    start_seconds: float = 0.0,
    end_seconds: Optional[float] = None,
    on_progress: Optional[Callable[[float], None]] = None,
    cancel_flag: Optional[Callable[[], bool]] = None,
) -> List[Dict[str, Any]]:
    """Scan video for change points in a region.

    Returns list of ``{timestamp, magnitude}`` dicts.
    """
    if threshold <= 0:
        threshold = config.SCREENSPACE_CHANGE_RATIO_THRESHOLD
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL
    if noise_threshold <= 0:
        noise_threshold = config.SCREENSPACE_NOISE_THRESHOLD

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return []
    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    duration = total_frames / fps if fps > 0 else 0.0
    cap.release()

    if end_seconds is None or end_seconds > duration:
        end_seconds = duration
    total_range = end_seconds - start_seconds

    results: List[Dict[str, Any]] = []
    prev_pixels: List[Optional[np.ndarray]] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False
        if prev_pixels[0] is not None:
            mag = compute_frame_diff(prev_pixels[0], pixels, noise_threshold)
            if mag >= threshold:
                results.append({"timestamp": ts, "magnitude": round(mag, 4)})
        prev_pixels[0] = pixels.copy()
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
    )

    if on_progress:
        on_progress(1.0)
    return results


def scan_similarity(
    video_path: str,
    region: Dict[str, int],
    reference_frame: np.ndarray,
    threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: Optional[float] = None,
    on_progress: Optional[Callable[[float], None]] = None,
    cancel_flag: Optional[Callable[[], bool]] = None,
) -> List[Dict[str, Any]]:
    """Find frames where region is similar to a reference.

    Returns list of ``{timestamp, score}`` dicts, sorted by score
    descending.
    """
    if threshold <= 0:
        threshold = config.SCREENSPACE_SSIM_THRESHOLD
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return []
    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    duration = total_frames / fps if fps > 0 else 0.0
    cap.release()

    if end_seconds is None or end_seconds > duration:
        end_seconds = duration
    total_range = end_seconds - start_seconds

    results: List[Dict[str, Any]] = []

    def _cb(ts: float, pixels: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False
        if pixels.shape == reference_frame.shape:
            is_sim, score = regions_are_similar(pixels, reference_frame, threshold)
            if is_sim:
                results.append({"timestamp": ts, "score": round(score, 4)})
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
    )

    if on_progress:
        on_progress(1.0)
    results.sort(key=lambda r: r["score"], reverse=True)
    return results


def scan_text(
    video_path: str,
    region: Dict[str, int],
    search_string: str,
    interval_seconds: float = 2.0,
    *,
    fuzzy_threshold: float = 0.0,
    languages: Optional[List[str]] = None,
    start_seconds: float = 0.0,
    end_seconds: Optional[float] = None,
    on_progress: Optional[Callable[[float], None]] = None,
    cancel_flag: Optional[Callable[[], bool]] = None,
) -> List[Dict[str, Any]]:
    """Scan for text appearances in a region using EasyOCR.

    EasyOCR is lazy-imported. Raises ``ImportError`` with install
    instructions if missing.
    """
    try:
        import easyocr  # type: ignore[import-untyped]
    except ImportError:
        raise ImportError(
            "EasyOCR is required for text scan. Install with: uv add easyocr"
        ) from None

    if fuzzy_threshold <= 0:
        fuzzy_threshold = config.SCREENSPACE_OCR_FUZZY_THRESHOLD
    if languages is None:
        languages = ["en"]

    reader = easyocr.Reader(languages, verbose=False)

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return []
    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    duration = total_frames / fps if fps > 0 else 0.0
    cap.release()

    if end_seconds is None or end_seconds > duration:
        end_seconds = duration
    total_range = end_seconds - start_seconds

    results: List[Dict[str, Any]] = []
    search_lower = search_string.lower()

    def _cb(ts: float, pixels: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False
        ocr_results = reader.readtext(pixels, detail=1)
        for _, text, conf in ocr_results:
            ratio = difflib.SequenceMatcher(
                None, search_lower, text.lower()
            ).ratio()
            if ratio >= fuzzy_threshold:
                results.append(
                    {
                        "timestamp": ts,
                        "text_found": text,
                        "confidence": round(conf, 4),
                    }
                )
                break
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
    )

    if on_progress:
        on_progress(1.0)
    return results


_NUMBERS_RE = re.compile(r"-?\d+(?:\.\d+)?")
_VALID_OPERATORS = ("eq", "gt", "lt", "gte", "lte", "range")


def scan_numbers(
    video_path: str,
    region: Dict[str, int],
    operator: str,
    target_value: float = 0,
    interval_seconds: float = 2.0,
    *,
    range_min: Optional[float] = None,
    range_max: Optional[float] = None,
    languages: Optional[List[str]] = None,
    start_seconds: float = 0.0,
    end_seconds: Optional[float] = None,
    on_progress: Optional[Callable[[float], None]] = None,
    cancel_flag: Optional[Callable[[], bool]] = None,
) -> List[Dict[str, Any]]:
    """Scan for numeric values in a region and apply a comparison.

    Uses EasyOCR to detect text, parses numbers from it, and returns
    timestamps where the detected number satisfies the comparison.
    """
    if operator not in _VALID_OPERATORS:
        raise ValueError(
            f"Unknown operator '{operator}'. Must be one of: {', '.join(_VALID_OPERATORS)}"
        )

    try:
        import easyocr  # type: ignore[import-untyped]
    except ImportError:
        raise ImportError(
            "EasyOCR is required for numbers scan. Install with: uv add easyocr"
        ) from None

    if languages is None:
        languages = ["en"]

    reader = easyocr.Reader(languages, verbose=False)

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return []
    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    duration = total_frames / fps if fps > 0 else 0.0
    cap.release()

    if end_seconds is None or end_seconds > duration:
        end_seconds = duration
    total_range = end_seconds - start_seconds

    results: List[Dict[str, Any]] = []

    def _check(value: float) -> bool:
        if operator == "eq":
            return value == target_value
        elif operator == "gt":
            return value > target_value
        elif operator == "lt":
            return value < target_value
        elif operator == "gte":
            return value >= target_value
        elif operator == "lte":
            return value <= target_value
        elif operator == "range":
            return (range_min is not None and range_max is not None
                    and range_min <= value <= range_max)
        return False

    def _cb(ts: float, pixels: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False
        ocr_results = reader.readtext(pixels, detail=1)
        for _, text, _conf in ocr_results:
            cleaned = text.replace(",", "")
            for match in _NUMBERS_RE.findall(cleaned):
                num = float(match)
                if _check(num):
                    results.append({"timestamp": ts, "number_found": num})
                    return None
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_frames(
        video_path,
        region,
        interval_seconds,
        _cb,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
    )

    if on_progress:
        on_progress(1.0)
    return results


def generate_timelapse(
    video_path: str,
    region: Dict[str, int],
    speedup_factor: float,
    output_path: str,
    output_format: str = "mp4",
) -> Optional[str]:
    """Generate a cropped timelapse via ffmpeg.

    Returns output file path on success, ``None`` on failure.
    """
    cmd = build_timelapse_command(
        video_path, region, speedup_factor, output_path, output_format
    )
    result = video.run_ffmpeg_process(
        cmd,
        input_file=video_path,
        output_file=output_path,
        os_error_message="Failed to generate timelapse",
    )
    if result is None or result.returncode != 0:
        return None
    return output_path


def _merge_timestamp_spans(
    timestamps: List[float], interval: float
) -> List[Dict[str, Any]]:
    """Merge consecutive matched timestamps into spans."""
    if not timestamps:
        return []
    timestamps.sort()
    spans: List[Dict[str, Any]] = []
    start = timestamps[0]
    end = timestamps[0]
    gap = interval * 1.5

    for ts in timestamps[1:]:
        if ts - end <= gap:
            end = ts
        else:
            spans.append(
                {
                    "start": round(start, 2),
                    "end": round(end, 2),
                    "duration": round(end - start, 2),
                }
            )
            start = ts
            end = ts

    spans.append(
        {
            "start": round(start, 2),
            "end": round(end, 2),
            "duration": round(end - start, 2),
        }
    )
    return spans


# ---------------------------------------------------------------------------
# Task queue and worker
# ---------------------------------------------------------------------------

TASK_STATUS_QUEUED = "queued"
TASK_STATUS_RUNNING = "running"
TASK_STATUS_COMPLETED = "completed"
TASK_STATUS_FAILED = "failed"
TASK_STATUS_CANCELLED = "cancelled"

_SENTINEL = object()


def create_task(
    task_type: str,
    participant: str,
    source_video: str,
    video_path: str,
    region_name: str,
    region_coords: Dict[str, int],
    parameters: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Build a new task dict with all fields initialized."""
    return {
        "id": f"ss_{uuid.uuid4().hex[:8]}",
        "type": task_type,
        "participant": participant,
        "source_video": source_video,
        "video_path": video_path,
        "region": region_name,
        "region_coords": region_coords,
        "parameters": parameters or {},
        "status": TASK_STATUS_QUEUED,
        "progress": 0.0,
        "priority": 100,
        "result": None,
        "error": None,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "completed_at": None,
        "_cancelled": False,
    }


class ScreenspaceWorker:
    """Background thread that processes analysis tasks sequentially."""

    def __init__(self) -> None:
        self._queue: queue.PriorityQueue[Any] = queue.PriorityQueue()
        self._tasks: Dict[str, Dict[str, Any]] = {}
        self._lock = threading.Lock()
        self._thread: Optional[threading.Thread] = None
        self._running = False
        self.on_task_complete: Optional[Callable[[], None]] = None

    def start(self) -> None:
        """Start the worker thread."""
        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        """Signal the worker thread to stop."""
        self._running = False
        self._queue.put((0, "", _SENTINEL))
        if self._thread is not None:
            self._thread.join(timeout=5)

    def enqueue(self, task: Dict[str, Any]) -> str:
        """Add a task to the queue. Returns the task ID."""
        task_id = task["id"]
        with self._lock:
            self._tasks[task_id] = task
        self._queue.put((task.get("priority", 100), task["created_at"], task_id))
        return task_id

    def cancel(self, task_id: str) -> bool:
        """Cancel a queued or running task. Returns True if cancelled."""
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return False
            if task["status"] == TASK_STATUS_QUEUED:
                task["status"] = TASK_STATUS_CANCELLED
                return True
            if task["status"] == TASK_STATUS_RUNNING:
                task["_cancelled"] = True
                return True
        return False

    def get_task(self, task_id: str) -> Optional[Dict[str, Any]]:
        """Return a task dict by ID (thread-safe copy)."""
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return None
            return copy.deepcopy(task)

    def get_all_tasks(self) -> List[Dict[str, Any]]:
        """Return all tasks (thread-safe copies)."""
        with self._lock:
            return [copy.deepcopy(t) for t in self._tasks.values()]

    def reorder(self, task_ids: List[str]) -> bool:
        """Reorder queued tasks by the given ID sequence.

        Assigns new priorities so that earlier IDs in the list have
        lower (higher-priority) values.
        """
        with self._lock:
            for i, tid in enumerate(task_ids):
                task = self._tasks.get(tid)
                if task and task["status"] == TASK_STATUS_QUEUED:
                    task["priority"] = i + 1
        return True

    def remove_task(self, task_id: str) -> bool:
        """Cancel (if active) and fully remove a task."""
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return False
            if task["status"] == TASK_STATUS_RUNNING:
                task["_cancelled"] = True
            self._tasks.pop(task_id, None)
            return True

    def _run(self) -> None:
        """Worker loop."""
        while self._running:
            try:
                _, _, task_id = self._queue.get(timeout=1)
            except queue.Empty:
                continue
            if task_id is _SENTINEL:
                break

            with self._lock:
                task = self._tasks.get(task_id)
                if task is None or task["status"] != TASK_STATUS_QUEUED:
                    continue
                task["status"] = TASK_STATUS_RUNNING

            self._execute_task(task)

            if self.on_task_complete:
                self.on_task_complete()

    def _execute_task(self, task: Dict[str, Any]) -> None:
        """Dispatch task to the appropriate workflow function."""
        task_id = task["id"]

        def _on_progress(progress: float) -> None:
            with self._lock:
                t = self._tasks.get(task_id)
                if t:
                    t["progress"] = min(progress, 1.0)

        def _cancel_flag() -> bool:
            with self._lock:
                t = self._tasks.get(task_id)
                return bool(t and t.get("_cancelled"))

        try:
            result = self._dispatch(task, _on_progress, _cancel_flag)
            with self._lock:
                t = self._tasks.get(task_id)
                if t:
                    if t.get("_cancelled"):
                        t["status"] = TASK_STATUS_CANCELLED
                    else:
                        t["status"] = TASK_STATUS_COMPLETED
                        t["result"] = result
                        t["progress"] = 1.0
                    t["completed_at"] = datetime.now(timezone.utc).isoformat()
        except Exception as exc:
            with self._lock:
                t = self._tasks.get(task_id)
                if t:
                    t["status"] = TASK_STATUS_FAILED
                    t["error"] = str(exc)
                    t["completed_at"] = datetime.now(timezone.utc).isoformat()

    def _dispatch(
        self,
        task: Dict[str, Any],
        on_progress: Callable[[float], None],
        cancel_flag: Callable[[], bool],
    ) -> Any:
        """Route task to the correct workflow function."""
        video_path = task["video_path"]
        region = task["region_coords"]
        params = task.get("parameters", {})
        task_type = task["type"]

        if task_type == "color":
            return scan_color(
                video_path,
                region,
                target_color=params.get("target_color", {"h": 0, "s": 0, "v": 0}),
                tolerance=params.get("tolerance", {"h": 10, "s": 50, "v": 50}),
                interval_seconds=params.get("interval", 0),
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
            )
        elif task_type == "change":
            return scan_changes(
                video_path,
                region,
                threshold=params.get("threshold", 0),
                interval_seconds=params.get("interval", 0),
                noise_threshold=params.get("noise_threshold", 0),
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
            )
        elif task_type == "similarity":
            ref_frame = params.get("reference_frame")
            if ref_frame is None:
                raise ValueError("Similarity scan requires a reference_frame parameter")
            return scan_similarity(
                video_path,
                region,
                reference_frame=ref_frame,
                threshold=params.get("threshold", 0),
                interval_seconds=params.get("interval", 0),
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
            )
        elif task_type == "text":
            return scan_text(
                video_path,
                region,
                search_string=params.get("search_string", ""),
                interval_seconds=params.get("interval", 2.0),
                fuzzy_threshold=params.get("fuzzy_threshold", 0),
                languages=params.get("languages"),
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
            )
        elif task_type == "numbers":
            return scan_numbers(
                video_path,
                region,
                operator=params.get("operator", "gt"),
                target_value=params.get("target_value", 0),
                interval_seconds=params.get("interval", 2.0),
                range_min=params.get("range_min"),
                range_max=params.get("range_max"),
                languages=params.get("languages"),
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
            )
        elif task_type == "timelapse":
            output_path = params.get("output_path", "")
            if not output_path:
                ext = "gif" if params.get("output_format") == "gif" else "mp4"
                output_path = str(
                    Path(utils.get_effective_output_dir())
                    / f"timelapse_{task['id']}.{ext}"
                )
            return generate_timelapse(
                video_path,
                region,
                speedup_factor=params.get("speedup_factor", 10.0),
                output_path=output_path,
                output_format=params.get("output_format", "mp4"),
            )
        else:
            raise ValueError(f"Unknown task type: {task_type}")


# ---------------------------------------------------------------------------
# Manifest I/O
# ---------------------------------------------------------------------------


def _empty_screenspace_manifest() -> Dict[str, Any]:
    return {"regions": {}, "tasks": []}


def load_screenspace_manifest() -> Dict[str, Any]:
    """Load the screenspace manifest from the output directory.

    Returns a dict with ``regions`` and ``tasks`` keys.
    """
    manifest_path = (
        Path(utils.get_effective_output_dir()) / config.SCREENSPACE_MANIFEST_FILENAME
    )
    if not manifest_path.is_file():
        return _empty_screenspace_manifest()
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return _empty_screenspace_manifest()
    if not isinstance(data, dict):
        return _empty_screenspace_manifest()
    return {
        "regions": data.get("regions", {}),
        "tasks": data.get("tasks", []),
    }


def save_screenspace_manifest(
    regions: Dict[str, Dict[str, Any]],
    tasks: List[Dict[str, Any]],
) -> Optional[Path]:
    """Write the screenspace manifest to disk.

    Strips internal fields (prefixed with ``_``) from tasks before writing.
    Returns the manifest path on success, or ``None`` on failure.
    """
    manifest_path = (
        Path(utils.get_effective_output_dir()) / config.SCREENSPACE_MANIFEST_FILENAME
    )
    clean_tasks = []
    for task in tasks:
        ct = {k: v for k, v in task.items() if not k.startswith("_")}
        if "parameters" in ct:
            ct["parameters"] = {
                k: v for k, v in ct["parameters"].items() if k != "reference_frame"
            }
        clean_tasks.append(ct)
    try:
        manifest_path.write_text(
            json.dumps(
                {"regions": regions, "tasks": clean_tasks},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        return manifest_path
    except OSError as e:
        utils.warning_print(f"Could not write screenspace manifest: {e}")
        return None
