# -*- coding: utf-8 -*-
"""Screenspace analysis engine for clipgen.

Ten analysis tools (passed as 'type' when creating a task):
  multitool  – chain multiple tools; each subsequent step only checks frames that passed previous steps
  color      – frames where a region's average HSV color matches a target within tolerance
  change     – frames where pixel diff ratio exceeds SCREENSPACE_CHANGE_RATIO_THRESHOLD
  similarity – frames matching a reference capture via SSIM (SCREENSPACE_SSIM_THRESHOLD)
  text       – OCR fuzzy search for a query string (SCREENSPACE_OCR_FUZZY_THRESHOLD); requires EasyOCR
  numbers    – OCR numeric comparison with a relational condition (eq/gt/lt/gte/lte/range)
  timelapse  – sped-up video of a region over a time range
  template   – find a reference image/template anywhere in the full frame via cv2.matchTemplate
  flow       – detect motion in a region via dense optical flow (cv2.calcOpticalFlowFarneback)
  scene      – classify frames by similarity to user-captured reference scenes

Workflow: user draws regions on a frame → enqueues tasks → ScreenspaceWorker processes in
a background thread → results are timestamps or artifact files → state persisted to
screenspace_manifest.json. Region coordinates are normalized (0–1); source_width/source_height
are stored for denormalization to target video resolution.
"""

from __future__ import annotations

import copy
import difflib
import json
import math
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


def match_template(
    frame: np.ndarray,
    template: np.ndarray,
    threshold: float = 0.0,
    nms_overlap: float = 0.0,
    mask: Optional[np.ndarray] = None,
) -> List[Dict[str, Any]]:
    """Find all locations where template appears in frame.

    Uses ``cv2.matchTemplate`` with ``TM_CCOEFF_NORMED``.  Non-maximum
    suppression removes overlapping detections.  An optional *mask*
    (same size as *template*, single-channel) restricts matching to
    non-transparent regions — useful for uploaded PNGs with alpha.

    Returns:
        List of ``{x, y, w, h, score}`` dicts for each match above *threshold*.
    """
    if threshold <= 0.0:
        threshold = config.SCREENSPACE_TEMPLATE_MATCH_THRESHOLD
    if nms_overlap <= 0.0:
        nms_overlap = config.SCREENSPACE_TEMPLATE_NMS_OVERLAP

    k = config.SCREENSPACE_BLUR_KERNEL
    frame_gray = cv2.cvtColor(cv2.GaussianBlur(frame, (k, k), 0), cv2.COLOR_BGR2GRAY)
    tmpl_gray = cv2.cvtColor(cv2.GaussianBlur(template, (k, k), 0), cv2.COLOR_BGR2GRAY)

    th, tw = tmpl_gray.shape[:2]
    if th > frame_gray.shape[0] or tw > frame_gray.shape[1]:
        return []

    # A constant (zero-variance) template produces undefined TM_CCOEFF_NORMED
    # results — every position may score ~1.0.  Bail out early.
    if float(np.std(tmpl_gray)) < 1e-6:
        return []

    # Blur the mask to match the blurred template/frame
    gray_mask = None
    if mask is not None:
        gray_mask = cv2.GaussianBlur(mask, (k, k), 0)

    result = cv2.matchTemplate(
        frame_gray, tmpl_gray, cv2.TM_CCOEFF_NORMED, mask=gray_mask
    )
    locs = np.where(result >= threshold)
    if len(locs[0]) == 0:
        return []

    # Collect raw detections sorted by score descending.
    # TM_CCOEFF_NORMED can produce inf/nan when a frame patch has zero
    # variance (constant colour) — skip those positions.
    detections: List[Dict[str, Any]] = []
    for pt_y, pt_x in zip(locs[0], locs[1]):
        score = float(result[pt_y, pt_x])
        if not math.isfinite(score):
            continue
        detections.append(
            {"x": int(pt_x), "y": int(pt_y), "w": tw, "h": th, "score": score}
        )
    detections.sort(key=lambda d: d["score"], reverse=True)

    # Non-maximum suppression
    kept: List[Dict[str, Any]] = []
    for det in detections:
        overlaps = False
        for k_det in kept:
            # Compute IoU
            xa = max(det["x"], k_det["x"])
            ya = max(det["y"], k_det["y"])
            xb = min(det["x"] + det["w"], k_det["x"] + k_det["w"])
            yb = min(det["y"] + det["h"], k_det["y"] + k_det["h"])
            inter = max(0, xb - xa) * max(0, yb - ya)
            area_a = det["w"] * det["h"]
            area_b = k_det["w"] * k_det["h"]
            union = area_a + area_b - inter
            if union > 0 and inter / union > nms_overlap:
                overlaps = True
                break
        if not overlaps:
            kept.append(det)
    return kept


def compute_optical_flow(
    prev_gray: np.ndarray,
    curr_gray: np.ndarray,
    pyr_scale: float = 0.0,
    return_grid: bool = False,
) -> Dict[str, Any]:
    """Compute dense optical flow between two grayscale frames.

    Returns:
        Dict with ``magnitude`` (mean flow vector length),
        ``angle`` (dominant direction in degrees, 0-360), and optionally
        ``flow_grid`` (sparse grid of motion vectors for visualization).
    """
    if pyr_scale <= 0.0:
        pyr_scale = config.SCREENSPACE_FLOW_PYRE_SCALE

    # Resize to max 256px for speed
    max_dim = 256
    h, w = prev_gray.shape[:2]
    if h > max_dim or w > max_dim:
        scale = max_dim / max(h, w)
        new_w, new_h = int(w * scale), int(h * scale)
        prev_gray = cv2.resize(prev_gray, (new_w, new_h), interpolation=cv2.INTER_AREA)
        curr_gray = cv2.resize(curr_gray, (new_w, new_h), interpolation=cv2.INTER_AREA)

    flow_out = np.zeros((*prev_gray.shape[:2], 2), dtype=np.float32)
    flow = cv2.calcOpticalFlowFarneback(
        prev_gray, curr_gray, flow_out, pyr_scale, 3, 15, 3, 5, 1.2, 0
    )
    mag, ang = cv2.cartToPolar(flow[..., 0], flow[..., 1], angleInDegrees=True)
    mean_mag = float(np.mean(mag))

    # Dominant angle: weighted mean by magnitude
    if mean_mag > 0:
        # Use circular mean to avoid wraparound issues
        rad = np.deg2rad(ang)
        sin_sum = float(np.sum(mag * np.sin(rad)))
        cos_sum = float(np.sum(mag * np.cos(rad)))
        dominant_angle = float(np.rad2deg(np.arctan2(sin_sum, cos_sum))) % 360.0
    else:
        dominant_angle = 0.0

    result: Dict[str, Any] = {
        "magnitude": round(mean_mag, 4),
        "angle": round(dominant_angle, 1),
    }

    if return_grid:
        grid_size = config.SCREENSPACE_FLOW_GRID_SIZE
        min_mag = config.SCREENSPACE_FLOW_GRID_MIN_MAG
        gh, gw = mag.shape[:2]
        step_y = max(1, gh // grid_size)
        step_x = max(1, gw // grid_size)
        grid: List[Dict[str, float]] = []
        for gy in range(0, gh, step_y):
            for gx in range(0, gw, step_x):
                cell_mag = float(np.mean(mag[gy : gy + step_y, gx : gx + step_x]))
                if cell_mag < min_mag:
                    continue
                cell_ang = float(np.mean(ang[gy : gy + step_y, gx : gx + step_x]))
                grid.append(
                    {
                        "x": round((gx + step_x / 2) / gw, 3),
                        "y": round((gy + step_y / 2) / gh, 3),
                        "mag": round(cell_mag, 2),
                        "ang": round(cell_ang, 1),
                    }
                )
        result["flow_grid"] = grid

    return result


def compute_scene_fingerprint(region_pixels: np.ndarray) -> Dict[str, Any]:
    """Compute a feature-based fingerprint for scene classification.

    Combines HSV histogram, edge density, and color statistics into a
    fingerprint suitable for comparison via :func:`compare_scene_fingerprints`.
    """
    # Resize to standardize
    max_dim = 128
    h, w = region_pixels.shape[:2]
    if h > max_dim or w > max_dim:
        scale = max_dim / max(h, w)
        region_pixels = cv2.resize(
            region_pixels,
            (int(w * scale), int(h * scale)),
            interpolation=cv2.INTER_AREA,
        )

    bins = config.SCREENSPACE_SCENE_HISTOGRAM_BINS
    hsv = cv2.cvtColor(region_pixels, cv2.COLOR_BGR2HSV)
    # 3D histogram flattened
    hist = cv2.calcHist(
        [hsv],
        [0, 1, 2],
        None,
        [bins, bins, bins],
        [0, 180, 0, 256, 0, 256],
    )
    cv2.normalize(hist, hist)

    # Edge density
    gray = cv2.cvtColor(region_pixels, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 100, 200)
    edge_density = (
        float(np.count_nonzero(edges)) / float(edges.size) if edges.size > 0 else 0.0
    )

    # Color stats per channel
    color_stats: List[float] = []
    for ch in range(3):
        channel = region_pixels[:, :, ch].astype(np.float64)
        color_stats.extend([float(np.mean(channel)), float(np.std(channel))])

    return {
        "histogram": hist,
        "edge_density": edge_density,
        "color_stats": color_stats,
    }


def compare_scene_fingerprints(
    fp_a: Dict[str, Any],
    fp_b: Dict[str, Any],
) -> float:
    """Compare two scene fingerprints.

    Returns similarity score 0.0–1.0.
    """
    # Histogram correlation: range [-1, 1] → [0, 1]
    hist_corr = cv2.compareHist(
        fp_a["histogram"].astype(np.float32),
        fp_b["histogram"].astype(np.float32),
        cv2.HISTCMP_CORREL,
    )
    hist_sim = (hist_corr + 1.0) / 2.0

    # Edge density similarity
    edge_sim = 1.0 - abs(fp_a["edge_density"] - fp_b["edge_density"])

    # Color stats similarity (normalized Euclidean distance)
    stats_a = np.array(fp_a["color_stats"], dtype=np.float64)
    stats_b = np.array(fp_b["color_stats"], dtype=np.float64)
    max_dist = np.sqrt(len(stats_a)) * 255.0  # theoretical max
    dist = float(np.linalg.norm(stats_a - stats_b))
    color_sim = 1.0 - (dist / max_dist) if max_dist > 0 else 1.0

    # Weighted average
    score = 0.6 * hist_sim + 0.2 * edge_sim + 0.2 * color_sim
    return max(0.0, min(1.0, score))


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


def scan_video_full_frames(
    video_path: str,
    interval_seconds: float,
    callback: Callable[[float, np.ndarray], Optional[bool]],
    *,
    start_seconds: float = 0.0,
    end_seconds: Optional[float] = None,
    fps: float = 0.0,
    duration: float = 0.0,
) -> None:
    """Like :func:`scan_video_frames` but passes the full frame (no region crop).

    Used by template detection which searches the entire frame.
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
                result = callback(ts, frame)
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
            result = callback(ts, frame)
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


def scan_template(
    video_path: str,
    region: Dict[str, int],
    template_image: np.ndarray,
    threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    template_mask: Optional[np.ndarray] = None,
    start_seconds: float = 0.0,
    end_seconds: Optional[float] = None,
    on_progress: Optional[Callable[[float], None]] = None,
    cancel_flag: Optional[Callable[[], bool]] = None,
    on_result: Optional[Callable[[Dict[str, Any]], None]] = None,
) -> List[Dict[str, Any]]:
    """Scan video for frames containing the template image.

    The template is searched across the **full frame** (not limited to a
    region).  *region* is unused for cropping but kept in the signature
    for consistency with other workflows.  An optional *template_mask*
    restricts matching to non-transparent regions of an uploaded PNG.

    Returns list of ``{timestamp, matches, best_score, match_count}`` dicts.
    """
    if threshold <= 0:
        threshold = config.SCREENSPACE_TEMPLATE_MATCH_THRESHOLD
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

    def _cb(ts: float, frame: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False
        matches = match_template(
            frame, template_image, threshold=threshold, mask=template_mask
        )
        if matches:
            best = max(m["score"] for m in matches)
            rd = {
                "timestamp": ts,
                "matches": matches,
                "best_score": round(best, 4),
                "match_count": len(matches),
            }
            results.append(rd)
            if on_result:
                on_result(
                    {
                        "timestamp": ts,
                        "best_score": rd["best_score"],
                        "match_count": rd["match_count"],
                    }
                )
        if on_progress and total_range > 0:
            on_progress((ts - start_seconds) / total_range)
        return None

    scan_video_full_frames(
        video_path,
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


def scan_flow(
    video_path: str,
    region: Dict[str, int],
    magnitude_threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: Optional[float] = None,
    on_progress: Optional[Callable[[float], None]] = None,
    cancel_flag: Optional[Callable[[], bool]] = None,
    on_result: Optional[Callable[[Dict[str, Any]], None]] = None,
) -> List[Dict[str, Any]]:
    """Scan video for motion in a region using dense optical flow.

    Returns list of ``{timestamp, magnitude, angle}`` dicts where
    magnitude exceeds *magnitude_threshold*.
    """
    if magnitude_threshold <= 0:
        magnitude_threshold = config.SCREENSPACE_FLOW_MAGNITUDE_THRESHOLD
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
    prev_gray: List[Optional[np.ndarray]] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False
        curr_gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
        if prev_gray[0] is not None:
            flow_result = compute_optical_flow(
                prev_gray[0], curr_gray, return_grid=True
            )
            if flow_result["magnitude"] >= magnitude_threshold:
                rd: Dict[str, Any] = {
                    "timestamp": ts,
                    "magnitude": flow_result["magnitude"],
                    "angle": flow_result["angle"],
                    "flow_grid": flow_result.get("flow_grid", []),
                }
                results.append(rd)
                if on_result:
                    # Keep incremental updates lightweight (no grid)
                    on_result(
                        {
                            "timestamp": ts,
                            "magnitude": flow_result["magnitude"],
                            "angle": flow_result["angle"],
                        }
                    )
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


def scan_scene(
    video_path: str,
    region: Dict[str, int],
    reference_scenes: List[Dict[str, Any]],
    threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: Optional[float] = None,
    on_progress: Optional[Callable[[float], None]] = None,
    cancel_flag: Optional[Callable[[], bool]] = None,
    on_result: Optional[Callable[[Dict[str, Any]], None]] = None,
) -> List[Dict[str, Any]]:
    """Classify frames by similarity to reference scene fingerprints.

    *reference_scenes* is a list of ``{name: str, frame: np.ndarray}``
    dicts.  Each frame's region is fingerprinted and compared against
    all references; the best match above *threshold* is reported.

    Returns list of ``{timestamp, scene_name, score}`` dicts.
    """
    if threshold <= 0:
        threshold = config.SCREENSPACE_SCENE_SIMILARITY_THRESHOLD
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    # Pre-compute fingerprints for reference scenes
    ref_fps: List[Tuple[str, Dict[str, Any]]] = []
    for ref in reference_scenes:
        fp = compute_scene_fingerprint(ref["frame"])
        ref_fps.append((ref["name"], fp))

    if not ref_fps:
        return []

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
    prev_skip_gray: List[Optional[np.ndarray]] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> Optional[bool]:
        if cancel_flag and cancel_flag():
            return False

        # Static-frame skip (same pattern as similarity scan)
        curr_gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY).astype(np.float32)
        if prev_skip_gray[0] is not None:
            if abs(float(np.mean(curr_gray)) - float(np.mean(prev_skip_gray[0]))) < 2.0:
                if on_progress and total_range > 0:
                    on_progress((ts - start_seconds) / total_range)
                return None
        prev_skip_gray[0] = curr_gray

        fp = compute_scene_fingerprint(pixels)
        best_name = ""
        best_score = 0.0
        for ref_name, ref_fp in ref_fps:
            score = compare_scene_fingerprints(fp, ref_fp)
            if score > best_score:
                best_score = score
                best_name = ref_name

        if best_score >= threshold:
            rd = {
                "timestamp": ts,
                "scene_name": best_name,
                "score": round(best_score, 4),
            }
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
    return results


# ---------------------------------------------------------------------------
# Multitool: per-frame evaluation and multi-factor scan
# ---------------------------------------------------------------------------

_NUMBERS_CHECK_RE = re.compile(r"-?\d+(?:\.\d+)?")


def _extract_confidence(tool_type: str, result: Dict[str, Any]) -> float:
    """Extract a normalized [0, 1] confidence from a tool-specific result dict."""
    if tool_type == "color":
        return result.get("_confidence", 1.0)
    elif tool_type == "change":
        return min(result.get("magnitude", 0.0), 1.0)
    elif tool_type == "similarity":
        return result.get("score", 0.0)
    elif tool_type == "text":
        return result.get("confidence", 0.0)
    elif tool_type == "numbers":
        return 1.0
    elif tool_type == "template":
        return result.get("best_score", 0.0)
    elif tool_type == "flow":
        return min(result.get("magnitude", 0.0) / 10.0, 1.0)
    elif tool_type == "scene":
        return result.get("score", 0.0)
    elif tool_type == "multitool":
        return result.get("min_confidence", 0.0)
    return 1.0


def check_frame_for_tool(
    frame: np.ndarray,
    prev_frame: Optional[np.ndarray],
    region: Dict[str, int],
    tool_type: str,
    parameters: Dict[str, Any],
) -> Tuple[bool, Optional[Dict[str, Any]]]:
    """Evaluate whether a single frame passes a tool's criteria.

    Used by :func:`scan_multitool` for steps 1+ in the chain.  Returns
    ``(passed, result_dict)`` where *result_dict* contains tool-specific
    metadata when the check passes, or ``None`` when it does not.

    For **change** and **flow** tools *prev_frame* is required (the frame
    immediately before the candidate timestamp).  If it is ``None`` the
    check is skipped (returns ``(False, None)``).
    """
    if tool_type == "color":
        pixels = extract_region(frame, region)
        target = parameters.get("target_color", {"h": 0, "s": 0, "v": 0})
        tol = parameters.get("tolerance", {"h": 10, "s": 50, "v": 50})
        avg = average_color_hsv(pixels)
        hue_diff = abs(avg["h"] - target["h"])
        hue_dist = min(hue_diff, 180.0 - hue_diff)
        s_dist = abs(avg["s"] - target["s"])
        v_dist = abs(avg["v"] - target["v"])
        if hue_dist <= tol["h"] and s_dist <= tol["s"] and v_dist <= tol["v"]:
            conf = max(
                0.0,
                1.0
                - max(
                    hue_dist / max(tol["h"], 1e-6),
                    s_dist / max(tol["s"], 1e-6),
                    v_dist / max(tol["v"], 1e-6),
                ),
            )
            return True, {"_confidence": conf}
        return False, None

    elif tool_type == "change":
        if prev_frame is None:
            return False, None
        pixels = extract_region(frame, region)
        prev_pixels = extract_region(prev_frame, region)
        threshold = parameters.get(
            "threshold", config.SCREENSPACE_CHANGE_RATIO_THRESHOLD
        )
        noise_threshold = parameters.get(
            "noise_threshold", config.SCREENSPACE_NOISE_THRESHOLD
        )
        k = config.SCREENSPACE_BLUR_KERNEL
        mk = config.SCREENSPACE_MORPH_KERNEL
        morph_kernel = np.ones((mk, mk), np.uint8)
        curr_gray = cv2.cvtColor(
            cv2.GaussianBlur(pixels, (k, k), 0), cv2.COLOR_BGR2GRAY
        )
        prev_gray = cv2.cvtColor(
            cv2.GaussianBlur(prev_pixels, (k, k), 0), cv2.COLOR_BGR2GRAY
        )
        diff = cv2.absdiff(prev_gray, curr_gray)
        _, mask = cv2.threshold(diff, noise_threshold, 255, cv2.THRESH_BINARY)
        mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, morph_kernel)
        mag = float(np.count_nonzero(mask)) / float(mask.size) if mask.size else 0.0
        if mag >= threshold:
            return True, {"magnitude": round(mag, 4)}
        return False, None

    elif tool_type == "similarity":
        ref = parameters.get("reference_frame")
        if ref is None:
            return False, None
        pixels = extract_region(frame, region)
        threshold = parameters.get("threshold", config.SCREENSPACE_SSIM_THRESHOLD)
        is_sim, score = regions_are_similar(pixels, ref, threshold)
        if is_sim:
            return True, {"score": round(score, 4)}
        return False, None

    elif tool_type == "text":
        pixels = extract_region(frame, region)
        search_string = parameters.get("search_string", "")
        if not search_string:
            return False, None
        fuzzy_threshold = parameters.get(
            "fuzzy_threshold", config.SCREENSPACE_OCR_FUZZY_THRESHOLD
        )
        languages = parameters.get("languages") or ["en"]
        reader = _get_ocr_reader(languages)
        ocr_results = reader.readtext(pixels, detail=1)
        search_lower = search_string.lower()
        for _, text, conf in ocr_results:
            ratio = difflib.SequenceMatcher(None, search_lower, text.lower()).ratio()
            if ratio >= fuzzy_threshold:
                return True, {"text_found": text, "confidence": round(conf, 4)}
        return False, None

    elif tool_type == "numbers":
        pixels = extract_region(frame, region)
        operator = parameters.get("operator", "gt")
        target_value = parameters.get("target_value", 0)
        range_min = parameters.get("range_min")
        range_max = parameters.get("range_max")
        languages = parameters.get("languages") or ["en"]
        reader = _get_ocr_reader(languages)
        ocr_results = reader.readtext(pixels, detail=1)
        for _, text, _conf in ocr_results:
            cleaned = text.replace(",", "")
            for match in _NUMBERS_CHECK_RE.findall(cleaned):
                num = float(match)
                passed = False
                if operator == "eq":
                    passed = num == target_value
                elif operator == "gt":
                    passed = num > target_value
                elif operator == "lt":
                    passed = num < target_value
                elif operator == "gte":
                    passed = num >= target_value
                elif operator == "lte":
                    passed = num <= target_value
                elif operator == "range":
                    passed = (
                        range_min is not None
                        and range_max is not None
                        and range_min <= num <= range_max
                    )
                if passed:
                    return True, {"number_found": num}
        return False, None

    elif tool_type == "template":
        template_img = parameters.get("template_image")
        if template_img is None:
            return False, None
        threshold = parameters.get(
            "threshold", config.SCREENSPACE_TEMPLATE_MATCH_THRESHOLD
        )
        template_mask = parameters.get("template_mask")
        matches = match_template(
            frame, template_img, threshold=threshold, mask=template_mask
        )
        if matches:
            best = max(m["score"] for m in matches)
            return True, {
                "best_score": round(best, 4),
                "match_count": len(matches),
            }
        return False, None

    elif tool_type == "flow":
        if prev_frame is None:
            return False, None
        pixels = extract_region(frame, region)
        prev_pixels = extract_region(prev_frame, region)
        curr_gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
        prev_gray_f = cv2.cvtColor(prev_pixels, cv2.COLOR_BGR2GRAY)
        magnitude_threshold = parameters.get(
            "magnitude_threshold", config.SCREENSPACE_FLOW_MAGNITUDE_THRESHOLD
        )
        flow_result = compute_optical_flow(prev_gray_f, curr_gray)
        if flow_result["magnitude"] >= magnitude_threshold:
            return True, {
                "magnitude": flow_result["magnitude"],
                "angle": flow_result["angle"],
            }
        return False, None

    elif tool_type == "scene":
        ref_scenes = parameters.get("reference_scenes")
        if not ref_scenes:
            return False, None
        threshold = parameters.get(
            "threshold", config.SCREENSPACE_SCENE_SIMILARITY_THRESHOLD
        )
        pixels = extract_region(frame, region)
        fp = compute_scene_fingerprint(pixels)
        best_name = ""
        best_score = 0.0
        for ref in ref_scenes:
            ref_fp = compute_scene_fingerprint(ref["frame"])
            score = compare_scene_fingerprints(fp, ref_fp)
            if score > best_score:
                best_score = score
                best_name = ref["name"]
        if best_score >= threshold:
            return True, {"scene_name": best_name, "score": round(best_score, 4)}
        return False, None

    return False, None


def scan_multitool(
    video_path: str,
    region: Dict[str, int],
    steps: List[Dict[str, Any]],
    *,
    start_seconds: float = 0.0,
    end_seconds: Optional[float] = None,
    on_progress: Optional[Callable[[float], None]] = None,
    cancel_flag: Optional[Callable[[], bool]] = None,
    on_result: Optional[Callable[[Dict[str, Any]], None]] = None,
) -> List[Dict[str, Any]]:
    """Run a multi-factor scan chaining several tool types.

    Step 0 runs a full scan on the video.  Each subsequent step only
    checks the timestamps that passed the previous step.  The final
    results are timestamps that survived all steps.

    Each entry in *steps* is a dict with ``"type"`` plus the tool's
    own parameters (e.g. ``target_color``, ``tolerance`` for color).

    Returns a list of ``{timestamp, tool_types, steps, min_confidence}``
    dicts.
    """
    num_steps = len(steps)
    if num_steps < 2:
        raise ValueError("Multitool requires at least 2 steps")

    # ---- Step 0: full scan via the appropriate scan_* function ----
    step0 = steps[0]
    step0_type = step0["type"]
    step0_collected: List[Dict[str, Any]] = []

    def _collect(rd: Dict[str, Any]) -> None:
        step0_collected.append(rd)

    def _scaled_progress_0(p: float) -> None:
        if on_progress:
            on_progress(p / num_steps)

    def _cancel() -> bool:
        return bool(cancel_flag and cancel_flag())

    interval = step0.get("interval", config.SCREENSPACE_DEFAULT_INTERVAL)
    s0_start = step0.get("start_seconds", start_seconds)
    s0_end = step0.get("end_seconds", end_seconds)

    if step0_type == "color":
        scan_color(
            video_path,
            region,
            target_color=step0.get("target_color", {"h": 0, "s": 0, "v": 0}),
            tolerance=step0.get("tolerance", {"h": 10, "s": 50, "v": 50}),
            interval_seconds=interval,
            start_seconds=s0_start,
            end_seconds=s0_end,
            on_progress=_scaled_progress_0,
            cancel_flag=_cancel,
            on_result=_collect,
        )
    elif step0_type == "change":
        scan_changes(
            video_path,
            region,
            threshold=step0.get("threshold", 0),
            interval_seconds=interval,
            noise_threshold=step0.get("noise_threshold", 0),
            start_seconds=s0_start,
            end_seconds=s0_end,
            on_progress=_scaled_progress_0,
            cancel_flag=_cancel,
            on_result=_collect,
        )
    elif step0_type == "similarity":
        ref_frame = step0.get("reference_frame")
        if ref_frame is None:
            raise ValueError("Similarity step requires a reference_frame parameter")
        scan_similarity(
            video_path,
            region,
            reference_frame=ref_frame,
            threshold=step0.get("threshold", 0),
            interval_seconds=interval,
            start_seconds=s0_start,
            end_seconds=s0_end,
            on_progress=_scaled_progress_0,
            cancel_flag=_cancel,
            on_result=_collect,
        )
    elif step0_type == "text":
        scan_text(
            video_path,
            region,
            search_string=step0.get("search_string", ""),
            interval_seconds=interval,
            fuzzy_threshold=step0.get("fuzzy_threshold", 0),
            languages=step0.get("languages"),
            start_seconds=s0_start,
            end_seconds=s0_end,
            on_progress=_scaled_progress_0,
            cancel_flag=_cancel,
            on_result=_collect,
        )
    elif step0_type == "numbers":
        scan_numbers(
            video_path,
            region,
            operator=step0.get("operator", "gt"),
            target_value=step0.get("target_value", 0),
            interval_seconds=interval,
            range_min=step0.get("range_min"),
            range_max=step0.get("range_max"),
            languages=step0.get("languages"),
            start_seconds=s0_start,
            end_seconds=s0_end,
            on_progress=_scaled_progress_0,
            cancel_flag=_cancel,
            on_result=_collect,
        )
    elif step0_type == "template":
        template_img = step0.get("template_image")
        if template_img is None:
            raise ValueError("Template step requires a template_image parameter")
        scan_template(
            video_path,
            region,
            template_image=template_img,
            threshold=step0.get("threshold", 0),
            interval_seconds=interval,
            template_mask=step0.get("template_mask"),
            start_seconds=s0_start,
            end_seconds=s0_end,
            on_progress=_scaled_progress_0,
            cancel_flag=_cancel,
            on_result=_collect,
        )
    elif step0_type == "flow":
        scan_flow(
            video_path,
            region,
            magnitude_threshold=step0.get("magnitude_threshold", 0),
            interval_seconds=interval,
            start_seconds=s0_start,
            end_seconds=s0_end,
            on_progress=_scaled_progress_0,
            cancel_flag=_cancel,
            on_result=_collect,
        )
    elif step0_type == "scene":
        ref_scenes = step0.get("reference_scenes")
        if not ref_scenes:
            raise ValueError("Scene step requires reference_scenes parameter")
        scan_scene(
            video_path,
            region,
            reference_scenes=ref_scenes,
            threshold=step0.get("threshold", 0),
            interval_seconds=interval,
            start_seconds=s0_start,
            end_seconds=s0_end,
            on_progress=_scaled_progress_0,
            cancel_flag=_cancel,
            on_result=_collect,
        )
    else:
        raise ValueError(f"Unsupported step 0 type: {step0_type}")

    if _cancel():
        return []

    # Build working set: {timestamp: [step0_result_dict]}
    working: Dict[float, List[Dict[str, Any]]] = {}
    for rd in step0_collected:
        ts = rd.get("timestamp", rd.get("start", 0.0))
        working[ts] = [rd]

    if not working:
        if on_progress:
            on_progress(1.0)
        return []

    # ---- Steps 1..N-1: check surviving timestamps ----
    for step_idx in range(1, num_steps):
        if _cancel():
            return []
        step = steps[step_idx]
        step_type = step["type"]
        new_working: Dict[float, List[Dict[str, Any]]] = {}

        cap = cv2.VideoCapture(video_path)
        if not cap.isOpened():
            break

        timestamps = sorted(working.keys())
        total_ts = len(timestamps)

        for ti, ts in enumerate(timestamps):
            if _cancel():
                break

            # Read frame at timestamp
            cap.set(cv2.CAP_PROP_POS_MSEC, ts * 1000.0)
            ret, frame = cap.read()
            if not ret:
                continue

            # For change/flow, read previous frame
            prev_frame = None
            if step_type in ("change", "flow"):
                prev_ts = max(0.0, ts - interval)
                cap.set(cv2.CAP_PROP_POS_MSEC, prev_ts * 1000.0)
                ret2, prev_frame = cap.read()
                if not ret2:
                    prev_frame = None

            passed, result_dict = check_frame_for_tool(
                frame, prev_frame, region, step_type, step
            )
            if passed and result_dict is not None:
                new_working[ts] = working[ts] + [result_dict]

            if on_progress and total_ts > 0:
                step_base = step_idx / num_steps
                step_frac = (ti + 1) / total_ts / num_steps
                on_progress(step_base + step_frac)

        cap.release()
        working = new_working

        if not working:
            break

    # ---- Build final results ----
    tool_types = [s["type"] for s in steps]
    results: List[Dict[str, Any]] = []
    for ts in sorted(working.keys()):
        step_results = working[ts]
        confidences = []
        for i, sr in enumerate(step_results):
            confidences.append(_extract_confidence(steps[i]["type"], sr))
        min_conf = min(confidences) if confidences else 0.0
        rd = {
            "timestamp": round(ts, 2),
            "tool_types": tool_types,
            "steps": step_results,
            "min_confidence": round(min_conf, 4),
        }
        results.append(rd)
        if on_result:
            on_result(rd)

    if on_progress:
        on_progress(1.0)
    return results


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


# ---------------------------------------------------------------------------
# Heatmap generation
# ---------------------------------------------------------------------------


def generate_template_heatmap(
    results: List[Dict[str, Any]],
    frame_width: int,
    frame_height: int,
    output_path: str,
) -> Optional[str]:
    """Generate a heatmap PNG from accumulated template match bounding boxes.

    Each pixel's intensity reflects how many times (weighted by score) it fell
    inside a match bounding box across all scanned frames.
    """
    accumulator = np.zeros((frame_height, frame_width), dtype=np.float32)
    for r in results:
        for m in r.get("matches", []):
            x, y, w, h = int(m["x"]), int(m["y"]), int(m["w"]), int(m["h"])
            y2 = min(y + h, frame_height)
            x2 = min(x + w, frame_width)
            accumulator[y:y2, x:x2] += m.get("score", 1.0)

    if accumulator.max() == 0:
        return None

    normalized = (accumulator / accumulator.max() * 255).astype(np.uint8)
    normalized = cv2.GaussianBlur(normalized, (15, 15), 0)
    heatmap = cv2.applyColorMap(normalized, cv2.COLORMAP_JET)
    cv2.imwrite(output_path, heatmap)
    return output_path


def generate_flow_heatmap(
    results: List[Dict[str, Any]],
    region_width: int,
    region_height: int,
    output_path: str,
) -> Optional[str]:
    """Generate a heatmap PNG from accumulated optical flow magnitudes.

    Uses ``flow_grid`` data from each result to paint per-cell motion
    intensity across all frames.
    """
    acc_size = 256
    accumulator = np.zeros((acc_size, acc_size), dtype=np.float32)
    for r in results:
        for cell in r.get("flow_grid", []):
            cx = int(cell["x"] * (acc_size - 1))
            cy = int(cell["y"] * (acc_size - 1))
            radius = max(1, acc_size // 16)
            cv2.circle(accumulator, (cx, cy), radius, float(cell["mag"]), -1)

    if accumulator.max() == 0:
        return None

    normalized = (accumulator / accumulator.max() * 255).astype(np.uint8)
    normalized = cv2.GaussianBlur(normalized, (15, 15), 0)
    heatmap = cv2.applyColorMap(normalized, cv2.COLORMAP_JET)
    heatmap = cv2.resize(
        heatmap, (region_width, region_height), interpolation=cv2.INTER_LINEAR
    )
    cv2.imwrite(output_path, heatmap)
    return output_path


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
    def is_alive(self) -> bool:
        """Return whether the worker thread is alive."""
        return self._thread is not None and self._thread.is_alive()

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
            confidence = _extract_confidence(task_type, r)
            metadata: Dict[str, Any] = {}
            if task_type == "change":
                metadata["magnitude"] = r.get("magnitude", 0.0)
            elif task_type == "similarity":
                metadata["score"] = r.get("score", 0.0)
            elif task_type == "text":
                metadata["text_found"] = r.get("text_found", "")
            elif task_type == "numbers":
                metadata["value"] = r.get("number_found", 0)
            elif task_type == "template":
                metadata["match_count"] = r.get("match_count", 0)
                metadata["best_score"] = r.get("best_score", 0.0)
            elif task_type == "flow":
                metadata["magnitude"] = r.get("magnitude", 0.0)
                metadata["angle"] = r.get("angle", 0.0)
            elif task_type == "scene":
                metadata["scene_name"] = r.get("scene_name", "")
                metadata["score"] = r.get("score", 0.0)
            elif task_type == "multitool":
                metadata["tool_types"] = r.get("tool_types", [])
                metadata["steps"] = r.get("steps", [])
            events.append(create_event(task, ts, confidence, metadata))
        return events

    def _generate_heatmap(
        self, task: Dict[str, Any], results: List[Dict[str, Any]]
    ) -> None:
        """Generate a heatmap PNG for template or flow tasks."""
        task_type = task.get("type", "")
        task_id = task["id"]
        if task_type == "template":
            heatmap_path = str(
                Path(utils.get_effective_output_dir()) / f"heatmap_{task_id}.png"
            )
            props = video.probe_video_properties(task["video_path"])
            fw = props.get("width", 1920) if props else 1920
            fh = props.get("height", 1080) if props else 1080
            hp = generate_template_heatmap(results, fw, fh, heatmap_path)
            if hp:
                task["heatmap"] = Path(hp).name
        elif task_type == "flow":
            heatmap_path = str(
                Path(utils.get_effective_output_dir()) / f"heatmap_{task_id}.png"
            )
            rc = task.get("region_coords", {})
            rw = rc.get("w", 256)
            rh = rc.get("h", 256)
            hp = generate_flow_heatmap(results, rw, rh, heatmap_path)
            if hp:
                task["heatmap"] = Path(hp).name

    def _run(self) -> None:
        """Worker loop."""
        while self._running:
            try:
                item = self._queue.get(timeout=1)
            except queue.Empty:
                continue

            try:
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
                    try:
                        self.on_task_complete()
                    except Exception as exc:
                        utils.warning_print(f"on_task_complete callback failed: {exc}")
            except Exception as exc:
                utils.warning_print(f"Worker loop error: {exc}")

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
                    if t.get("parameters", {}).get("detect_first"):
                        t["_cancelled"] = True

        try:
            result = self._dispatch(task, _on_progress, _cancel_flag, _on_result)
            with self._lock:
                t = self._tasks.get(task_id)
                if t:
                    if t.get("_paused_flag"):
                        t["status"] = TASK_STATUS_PAUSED
                        t["result"] = result
                    elif t.get("_cancelled"):
                        if t.get("parameters", {}).get("detect_first"):
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
                        else:
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
                        # Generate heatmap artifacts for template/flow tasks
                        if isinstance(result, list) and result:
                            self._generate_heatmap(t, result)
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
        elif task_type == "template":
            template_img = params.get("template_image")
            if template_img is None:
                raise ValueError("Template scan requires a template_image parameter")
            return scan_template(
                video_path,
                region,
                template_image=template_img,
                threshold=params.get("threshold", 0),
                interval_seconds=params.get("interval", 0),
                template_mask=params.get("template_mask"),
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
                on_result=on_result,
            )
        elif task_type == "flow":
            return scan_flow(
                video_path,
                region,
                magnitude_threshold=params.get("magnitude_threshold", 0),
                interval_seconds=params.get("interval", 0),
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
                on_result=on_result,
            )
        elif task_type == "scene":
            ref_scenes = params.get("reference_scenes")
            if not ref_scenes:
                raise ValueError("Scene scan requires reference_scenes parameter")
            return scan_scene(
                video_path,
                region,
                reference_scenes=ref_scenes,
                threshold=params.get("threshold", 0),
                interval_seconds=params.get("interval", 0),
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
                on_result=on_result,
            )
        elif task_type == "multitool":
            steps = params.get("steps", [])
            if not steps or len(steps) < 2:
                raise ValueError("Multitool requires at least 2 steps")
            return scan_multitool(
                video_path,
                region,
                steps=steps,
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
                on_result=on_result,
            )
        else:
            raise ValueError(f"Unknown task type: {task_type}")


# ---------------------------------------------------------------------------
# Manifest I/O
# ---------------------------------------------------------------------------


def _empty_screenspace_manifest() -> Dict[str, Any]:
    return {"regions": {}, "tasks": [], "events": [], "stashes": []}


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
        "stashes": data.get("stashes", []),
    }


def save_screenspace_manifest(
    regions: Dict[str, Dict[str, Any]],
    tasks: List[Dict[str, Any]],
    events: Optional[List[Dict[str, Any]]] = None,
    stashes: Optional[List[Dict[str, Any]]] = None,
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
            _binary_keys = (
                "reference_frame",
                "template_image",
                "template_mask",
                "reference_scenes",
            )
            ct["parameters"] = {
                k: v for k, v in ct["parameters"].items() if k not in _binary_keys
            }
            # Strip binary data from multitool step parameters
            if "steps" in ct["parameters"]:
                ct["parameters"]["steps"] = [
                    {k: v for k, v in s.items() if k not in _binary_keys}
                    for s in ct["parameters"]["steps"]
                ]
        # Strip flow_grid from results (large per-frame data, not needed on disk)
        if isinstance(ct.get("result"), list):
            ct["result"] = [
                {k: v for k, v in r.items() if k != "flow_grid"} for r in ct["result"]
            ]
        clean_tasks.append(ct)
    try:
        manifest_path.write_text(
            json.dumps(
                {
                    "regions": regions,
                    "tasks": clean_tasks,
                    "events": events or [],
                    "stashes": stashes or [],
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
