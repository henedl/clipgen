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
import time
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
# Module-level caches
# ---------------------------------------------------------------------------

_ocr_readers: Dict[tuple, Any] = {}


def _get_ocr_reader(languages: List[str]) -> Any:
    """Return a cached EasyOCR Reader for the given language set."""
    import easyocr

    key = tuple(sorted(languages))
    if key not in _ocr_readers:
        _ocr_readers[key] = easyocr.Reader(list(key), verbose=False)
    return _ocr_readers[key]


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
    h, w = region_pixels.shape[:2]
    if h > 64 or w > 64:
        region_pixels = cv2.resize(
            region_pixels, (min(w, 64), min(h, 64)), interpolation=cv2.INTER_AREA
        )
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
    max_dim = 256
    h, w = region_a.shape[:2]
    if h > max_dim or w > max_dim:
        scale = max_dim / max(h, w)
        new_w, new_h = int(w * scale), int(h * scale)
        region_a = cv2.resize(region_a, (new_w, new_h), interpolation=cv2.INTER_AREA)
        region_b = cv2.resize(region_b, (new_w, new_h), interpolation=cv2.INTER_AREA)
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
    fps: float = 0.0,
    duration: float = 0.0,
) -> None:
    """Iterate through video at interval, extract region, call callback.

    The *callback* receives ``(timestamp_seconds, region_pixels)`` and may
    return ``False`` to stop iteration early.

    When *fps* and *duration* are provided, skips internal metadata reads.
    Uses sequential frame reading (grab/retrieve) for small intervals
    to avoid expensive H.264 seeking.
    """
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return

    if fps <= 0:
        fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    if duration <= 0:
        total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
        duration = total_frames / fps if fps > 0 else 0.0
    if end_seconds is None or end_seconds > duration:
        end_seconds = duration

    use_sequential = interval_seconds <= config.SCREENSPACE_SEQUENTIAL_READ_MAX_INTERVAL

    if use_sequential:
        frame_interval = max(1, round(interval_seconds * fps))
        if start_seconds > 0:
            cap.set(cv2.CAP_PROP_POS_MSEC, start_seconds * 1000.0)
        frame_idx = 0
        end_frame = int(end_seconds * fps)
        start_frame = int(start_seconds * fps)
        while True:
            grabbed = cap.grab()
            if not grabbed:
                break
            abs_frame = start_frame + frame_idx
            if abs_frame > end_frame:
                break
            if frame_idx % frame_interval == 0:
                ret, frame = cap.retrieve()
                if not ret:
                    break
                ts = start_seconds + frame_idx / fps
                cropped = extract_region(frame, region)
                result = callback(ts, cropped)
                if result is False:
                    break
            frame_idx += 1
    else:
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
    on_result: Optional[Callable[[Dict[str, Any]], None]] = None,
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
    vid_fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    vid_duration = total_frames / vid_fps if vid_fps > 0 else 0.0
    cap.release()

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    matches: List[float] = []

    def _cb(ts: float, pixels: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False
        avg = average_color_hsv(pixels)
        hue_diff = abs(avg["h"] - target_color["h"])
        hue_dist = min(hue_diff, 180.0 - hue_diff)
        s_dist = abs(avg["s"] - target_color["s"])
        v_dist = abs(avg["v"] - target_color["v"])
        if (
            hue_dist <= tolerance["h"]
            and s_dist <= tolerance["s"]
            and v_dist <= tolerance["v"]
        ):
            matches.append(ts)
            if on_result:
                conf = max(
                    0.0,
                    1.0
                    - max(
                        hue_dist / max(tolerance["h"], 1e-6),
                        s_dist / max(tolerance["s"], 1e-6),
                        v_dist / max(tolerance["v"], 1e-6),
                    ),
                )
                on_result({"timestamp": ts, "_confidence": conf})
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
        fps=vid_fps,
        duration=vid_duration,
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
    on_result: Optional[Callable[[Dict[str, Any]], None]] = None,
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
    vid_fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    vid_duration = total_frames / vid_fps if vid_fps > 0 else 0.0
    cap.release()

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: List[Dict[str, Any]] = []
    prev_gray: List[Optional[np.ndarray]] = [None]
    k = config.SCREENSPACE_BLUR_KERNEL
    mk = config.SCREENSPACE_MORPH_KERNEL
    morph_kernel = np.ones((mk, mk), np.uint8)

    def _cb(ts: float, pixels: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False
        curr_gray = cv2.cvtColor(
            cv2.GaussianBlur(pixels, (k, k), 0), cv2.COLOR_BGR2GRAY
        )
        if prev_gray[0] is not None:
            diff = cv2.absdiff(prev_gray[0], curr_gray)
            _, mask = cv2.threshold(diff, noise_threshold, 255, cv2.THRESH_BINARY)
            mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, morph_kernel)
            mag = float(np.count_nonzero(mask)) / float(mask.size) if mask.size else 0.0
            if mag >= threshold:
                rd = {"timestamp": ts, "magnitude": round(mag, 4)}
                results.append(rd)
                if on_result:
                    on_result(rd)
        prev_gray[0] = curr_gray
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
        fps=vid_fps,
        duration=vid_duration,
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
    on_result: Optional[Callable[[Dict[str, Any]], None]] = None,
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
    vid_fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    vid_duration = total_frames / vid_fps if vid_fps > 0 else 0.0
    cap.release()

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: List[Dict[str, Any]] = []
    ref_phash = compute_phash(reference_frame)
    phash_threshold = config.SCREENSPACE_PHASH_THRESHOLD

    # Pre-resize and preprocess reference frame once
    max_dim = 256
    rh, rw = reference_frame.shape[:2]
    needs_resize = rh > max_dim or rw > max_dim
    if needs_resize:
        scale = max_dim / max(rh, rw)
        new_w, new_h = int(rw * scale), int(rh * scale)
        ref_resized = cv2.resize(
            reference_frame, (new_w, new_h), interpolation=cv2.INTER_AREA
        )
    else:
        ref_resized = reference_frame
    bk = config.SCREENSPACE_BLUR_KERNEL
    ref_gray = cv2.cvtColor(
        cv2.GaussianBlur(ref_resized, (bk, bk), 0), cv2.COLOR_BGR2GRAY
    )

    prev_skip_gray: List[Optional[np.ndarray]] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False
        if pixels.shape == reference_frame.shape:
            # Static-frame skip
            gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
            if prev_skip_gray[0] is not None:
                if float(np.mean(cv2.absdiff(prev_skip_gray[0], gray))) < 2.0:
                    if on_progress and total_range > 0:
                        on_progress((ts - start_seconds) / total_range)
                    return None
            prev_skip_gray[0] = gray

            frame_phash = compute_phash(pixels)
            if ref_phash - frame_phash <= phash_threshold:
                if needs_resize:
                    cand = cv2.resize(
                        pixels, (new_w, new_h), interpolation=cv2.INTER_AREA
                    )
                else:
                    cand = pixels
                cand_gray = cv2.cvtColor(
                    cv2.GaussianBlur(cand, (bk, bk), 0), cv2.COLOR_BGR2GRAY
                )
                score = float(ssim(ref_gray, cand_gray))
                if score >= threshold:
                    rd = {"timestamp": ts, "score": round(score, 4)}
                    results.append(rd)
                    if on_result:
                        on_result(rd)
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
        fps=vid_fps,
        duration=vid_duration,
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
    on_result: Optional[Callable[[Dict[str, Any]], None]] = None,
) -> List[Dict[str, Any]]:
    """Scan for text appearances in a region using EasyOCR.

    EasyOCR is lazy-imported. Raises ``ImportError`` with install
    instructions if missing.
    """
    try:
        import easyocr  # noqa: F401 — validate availability
    except ImportError:
        raise ImportError(
            "EasyOCR is required for text scan. Install with: uv add easyocr"
        ) from None

    if fuzzy_threshold <= 0:
        fuzzy_threshold = config.SCREENSPACE_OCR_FUZZY_THRESHOLD
    if languages is None:
        languages = ["en"]

    reader = _get_ocr_reader(languages)

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return []
    vid_fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    vid_duration = total_frames / vid_fps if vid_fps > 0 else 0.0
    cap.release()

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: List[Dict[str, Any]] = []
    search_lower = search_string.lower()
    prev_gray: List[Optional[np.ndarray]] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False
        gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
        if prev_gray[0] is not None:
            diff = float(np.mean(cv2.absdiff(prev_gray[0], gray)))
            if diff < 2.0:
                if on_progress and total_range > 0:
                    on_progress((ts - start_seconds) / total_range)
                return None
        prev_gray[0] = gray
        ocr_results = reader.readtext(pixels, detail=1)
        for _, text, conf in ocr_results:
            ratio = difflib.SequenceMatcher(None, search_lower, text.lower()).ratio()
            if ratio >= fuzzy_threshold:
                rd = {
                    "timestamp": ts,
                    "text_found": text,
                    "confidence": round(conf, 4),
                }
                results.append(rd)
                if on_result:
                    on_result(rd)
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
        fps=vid_fps,
        duration=vid_duration,
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
    on_result: Optional[Callable[[Dict[str, Any]], None]] = None,
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
        import easyocr  # noqa: F401 — validate availability
    except ImportError:
        raise ImportError(
            "EasyOCR is required for numbers scan. Install with: uv add easyocr"
        ) from None

    if languages is None:
        languages = ["en"]

    reader = _get_ocr_reader(languages)

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return []
    vid_fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    vid_duration = total_frames / vid_fps if vid_fps > 0 else 0.0
    cap.release()

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: List[Dict[str, Any]] = []
    prev_gray: List[Optional[np.ndarray]] = [None]

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
            return (
                range_min is not None
                and range_max is not None
                and range_min <= value <= range_max
            )
        return False

    def _cb(ts: float, pixels: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False
        gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
        if prev_gray[0] is not None:
            diff = float(np.mean(cv2.absdiff(prev_gray[0], gray)))
            if diff < 2.0:
                if on_progress and total_range > 0:
                    on_progress((ts - start_seconds) / total_range)
                return None
        prev_gray[0] = gray
        ocr_results = reader.readtext(pixels, detail=1)
        for _, text, _conf in ocr_results:
            cleaned = text.replace(",", "")
            for match in _NUMBERS_RE.findall(cleaned):
                num = float(match)
                if _check(num):
                    rd = {"timestamp": ts, "number_found": num}
                    results.append(rd)
                    if on_result:
                        on_result(rd)
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
        fps=vid_fps,
        duration=vid_duration,
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
TASK_STATUS_PAUSED = "paused"

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
        self._paused = threading.Event()
        self.on_task_complete: Optional[Callable[[], None]] = None

    def start(self) -> None:
        """Start the worker thread."""
        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def restore_tasks(self, tasks: List[Dict[str, Any]]) -> None:
        """Load historical tasks into the worker (completed/failed/cancelled).

        Only non-queued tasks are restored for display; queued tasks are not
        re-enqueued since the analysis context (video caps) is gone.
        """
        with self._lock:
            for t in tasks:
                if t.get("id"):
                    self._tasks[t["id"]] = copy.deepcopy(t)

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
        """Cancel a queued, running, or paused task. Returns True if cancelled."""
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return False
            if task["status"] in (TASK_STATUS_QUEUED, TASK_STATUS_PAUSED):
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

    @property
    def is_paused(self) -> bool:
        """Return whether the queue is paused."""
        return self._paused.is_set()

    def pause(self) -> None:
        """Pause the queue. Stops the running task so it yields partial results."""
        self._paused.set()
        with self._lock:
            for task in self._tasks.values():
                if task["status"] == TASK_STATUS_RUNNING:
                    task["_paused_flag"] = True

    def resume(self) -> None:
        """Resume the queue. Re-enqueues paused tasks from where they left off."""
        self._paused.clear()
        to_resume: List[Dict[str, Any]] = []
        with self._lock:
            for task in self._tasks.values():
                if task["status"] == TASK_STATUS_PAUSED:
                    to_resume.append(task)

        for task in to_resume:
            progress = task.get("progress", 0.0)
            params = task.get("parameters", {})
            start = params.get("start_seconds", 0.0)
            end = params.get("end_seconds")
            if end is None:
                cap = cv2.VideoCapture(task["video_path"])
                fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
                total = cap.get(cv2.CAP_PROP_FRAME_COUNT)
                end = total / fps if fps > 0 else 0.0
                cap.release()
            resume_at = start + progress * (end - start)

            with self._lock:
                task["_partial_results"] = task.get("result", [])
                task["_progress_offset"] = progress
                task["_progress_scale"] = max(1.0 - progress, 0.001)
                task.pop("result", None)
                task.pop("_paused_flag", None)
                params["start_seconds"] = resume_at
                task["status"] = TASK_STATUS_QUEUED

            self._queue.put((task.get("priority", 100), task["created_at"], task["id"]))

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

    def drain_new_events(self) -> List[Dict[str, Any]]:
        """Collect and clear ``_generated_events`` from all tasks. Thread-safe."""
        events: List[Dict[str, Any]] = []
        with self._lock:
            for t in self._tasks.values():
                events.extend(t.pop("_generated_events", []))
        return events

    def _generate_events_from_results(
        self, task: Dict[str, Any], raw_results: List[Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        """Convert raw task results into ScreenspaceEvent records."""
        task_type = task.get("type", "")
        if task_type == "timelapse":
            return []
        events: List[Dict[str, Any]] = []
        for r in raw_results:
            ts = r.get("timestamp", r.get("start", 0.0))
            metadata: Dict[str, Any] = {}
            if task_type == "color":
                confidence = r.get("_confidence", 1.0)
            elif task_type == "change":
                confidence = r.get("magnitude", 0.0)
                metadata["magnitude"] = r.get("magnitude", 0.0)
            elif task_type == "similarity":
                confidence = r.get("score", 0.0)
                metadata["score"] = r.get("score", 0.0)
            elif task_type == "text":
                confidence = r.get("confidence", 0.0)
                metadata["text_found"] = r.get("text_found", "")
            elif task_type == "numbers":
                confidence = 1.0
                metadata["value"] = r.get("number_found", 0)
            else:
                confidence = 1.0
            events.append(create_event(task, ts, confidence, metadata))
        return events

    def _run(self) -> None:
        """Worker loop."""
        while self._running:
            try:
                item = self._queue.get(timeout=1)
            except queue.Empty:
                continue
            priority, created_at, task_id = item
            if task_id is _SENTINEL:
                break

            if self._paused.is_set():
                self._queue.put(item)
                time.sleep(0.25)
                continue

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

        # Seed incremental results (includes partial results on resume)
        with self._lock:
            t = self._tasks.get(task_id)
            if t:
                t["result"] = list(t.get("_partial_results", []))
                t["_raw_results"] = list(t.get("_partial_results", []))

        def _on_progress(progress: float) -> None:
            with self._lock:
                t = self._tasks.get(task_id)
                if t and not t.get("_paused_flag"):
                    offset = t.get("_progress_offset", 0.0)
                    scale = t.get("_progress_scale", 1.0)
                    t["progress"] = min(offset + progress * scale, 1.0)

        def _cancel_flag() -> bool:
            with self._lock:
                t = self._tasks.get(task_id)
                return bool(t and (t.get("_cancelled") or t.get("_paused_flag")))

        def _on_result(result_dict: Dict[str, Any]) -> None:
            with self._lock:
                t = self._tasks.get(task_id)
                if t:
                    if isinstance(t.get("result"), list):
                        t["result"].append(result_dict)
                    if isinstance(t.get("_raw_results"), list):
                        t["_raw_results"].append(result_dict)

        try:
            result = self._dispatch(task, _on_progress, _cancel_flag, _on_result)
            with self._lock:
                t = self._tasks.get(task_id)
                if t:
                    if t.get("_paused_flag"):
                        t["status"] = TASK_STATUS_PAUSED
                        t["result"] = result
                    elif t.get("_cancelled"):
                        t["status"] = TASK_STATUS_CANCELLED
                    else:
                        t["status"] = TASK_STATUS_COMPLETED
                        partial = t.pop("_partial_results", None)
                        if partial and isinstance(result, list):
                            result = partial + result
                        t["result"] = result
                        t["progress"] = 1.0
                        t.pop("_progress_offset", None)
                        t.pop("_progress_scale", None)
                        raw = t.pop("_raw_results", [])
                        t["_generated_events"] = self._generate_events_from_results(
                            t, raw
                        )
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
        on_result: Optional[Callable[[Dict[str, Any]], None]] = None,
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
                on_result=on_result,
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
                on_result=on_result,
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
                on_result=on_result,
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
                on_result=on_result,
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
                on_result=on_result,
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
    return {"regions": {}, "tasks": [], "events": []}


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
        "events": data.get("events", []),
    }


def save_screenspace_manifest(
    regions: Dict[str, Dict[str, Any]],
    tasks: List[Dict[str, Any]],
    events: Optional[List[Dict[str, Any]]] = None,
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
                {
                    "regions": regions,
                    "tasks": clean_tasks,
                    "events": events or [],
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        return manifest_path
    except OSError as e:
        utils.warning_print(f"Could not write screenspace manifest: {e}")
        return None


def create_event(
    task: Dict[str, Any],
    timestamp: float,
    confidence: float,
    metadata: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Create a ScreenspaceEvent from a task result entry."""
    event_label = task.get("parameters", {}).get("event_label", "")
    if not event_label:
        event_label = task["type"] + ": " + task.get("region", "")
    return {
        "id": f"ev_{uuid.uuid4().hex[:8]}",
        "source_video": task.get("source_video", ""),
        "participant": task.get("participant", ""),
        "detector": task["type"],
        "event_type": event_label,
        "time_in": round(timestamp, 2),
        "time_out": round(timestamp, 2),
        "confidence": round(max(0.0, min(1.0, confidence)), 4),
        "metadata": metadata or {},
        "excluded": False,
        "task_id": task["id"],
        "region": task.get("region", ""),
    }
