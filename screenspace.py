# -*- coding: utf-8 -*-
"""Screenspace analysis engine for clipgen.

Eleven analysis tools (passed as 'type' when creating a task):
  multitool   – chain multiple tools; each subsequent step only checks frames that passed previous steps
  color       – frames where a region's average HSV color matches a target within tolerance
  change      – frames where pixel diff ratio exceeds SCREENSPACE_CHANGE_RATIO_THRESHOLD
  similarity  – frames matching a reference capture via SSIM (SCREENSPACE_SSIM_THRESHOLD)
  text        – OCR fuzzy search for a query string (SCREENSPACE_OCR_FUZZY_THRESHOLD); requires EasyOCR
  numbers     – OCR numeric comparison with a relational condition (eq/gt/lt/gte/lte/range)
  timelapse   – sped-up video of a region over a time range
  template    – find a reference image/template anywhere in the full frame via cv2.matchTemplate
  flow        – detect motion in a region via dense optical flow (cv2.calcOpticalFlowFarneback)
  scene       – classify frames by similarity to user-captured reference scenes
  inactivity  – detect spans of near-duplicate frames via perceptual hashing (loading screens, frozen states)

Workflow: user draws regions on a frame → enqueues tasks → ScreenspaceWorker processes in
a background thread → results are timestamps or artifact files → state persisted to
screenspace_manifest.json. Region coordinates are normalized (0–1); source_width/source_height
are stored for denormalization to target video resolution.
"""

import copy
import difflib

import math
import queue
import re
import shutil
import subprocess
import threading
import time
import uuid
from concurrent.futures import Future, ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any, Callable

import cv2
import numpy as np

if TYPE_CHECKING:
    import imagehash

import config
import utils
import video


# ---------------------------------------------------------------------------
# Module-level caches
# ---------------------------------------------------------------------------

_ocr_readers: dict[tuple, Any] = {}
_ocr_lock = threading.Lock()


def _get_ocr_reader(languages: list[str]) -> Any:
    """Return a cached EasyOCR Reader for the given language set."""
    import easyocr

    key = tuple(sorted(languages))
    with _ocr_lock:
        if key not in _ocr_readers:
            _ocr_readers[key] = easyocr.Reader(list(key), verbose=False)
        return _ocr_readers[key]


# ---------------------------------------------------------------------------
# Analysis primitives
# ---------------------------------------------------------------------------


def extract_region(frame: np.ndarray, region: dict[str, int]) -> np.ndarray:
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


def average_color_hsv(region_pixels: np.ndarray) -> dict[str, float]:
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
    target_color: dict[str, float],
    tolerance: dict[str, float],
) -> tuple[bool, float]:
    """Check if region's average HSV color is within tolerance of target.

    Handles hue wraparound (red at 0/180 boundary).

    Returns:
        Tuple of (matches, confidence) where confidence is 0.0–1.0.
    """
    avg = average_color_hsv(region_pixels)
    hue_diff = abs(avg["h"] - target_color["h"])
    hue_dist = min(hue_diff, 180.0 - hue_diff)
    s_dist = abs(avg["s"] - target_color["s"])
    v_dist = abs(avg["v"] - target_color["v"])
    matched = (
        hue_dist <= tolerance["h"]
        and s_dist <= tolerance["s"]
        and v_dist <= tolerance["v"]
    )
    conf = max(
        0.0,
        1.0
        - max(
            hue_dist / max(tolerance["h"], 1e-6),
            s_dist / max(tolerance["s"], 1e-6),
            v_dist / max(tolerance["v"], 1e-6),
        ),
    )
    return matched, conf


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
) -> tuple[bool, float]:
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
    from skimage.metrics import structural_similarity as ssim

    score = float(ssim(a_gray, b_gray))
    return score >= threshold, score


def compute_phash(region_pixels: np.ndarray) -> "imagehash.ImageHash":
    """Compute perceptual hash of a region for fast similarity scanning."""
    import imagehash
    from PIL import Image

    rgb = cv2.cvtColor(region_pixels, cv2.COLOR_BGR2RGB)
    pil_img = Image.fromarray(rgb)
    return imagehash.phash(pil_img)


def match_template(
    frame: np.ndarray,
    template: np.ndarray,
    threshold: float = 0.0,
    nms_overlap: float = 0.0,
    mask: np.ndarray | None = None,
) -> list[dict[str, Any]]:
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
    detections: list[dict[str, Any]] = []
    for pt_y, pt_x in zip(locs[0], locs[1]):
        score = float(result[pt_y, pt_x])
        if not math.isfinite(score):
            continue
        detections.append(
            {"x": int(pt_x), "y": int(pt_y), "w": tw, "h": th, "score": score}
        )
    detections.sort(key=lambda d: d["score"], reverse=True)

    # Non-maximum suppression
    kept: list[dict[str, Any]] = []
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
) -> dict[str, Any]:
    """Compute dense optical flow between two grayscale frames.

    Returns:
        Dict with ``magnitude`` (mean flow vector length),
        ``angle`` (dominant direction in degrees, 0-360), and optionally
        ``flow_grid`` (sparse grid of motion vectors for visualization).
    """
    if pyr_scale <= 0.0:
        pyr_scale = config.SCREENSPACE_FLOW_PYR_SCALE

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

    result: dict[str, Any] = {
        "magnitude": round(mean_mag, 4),
        "angle": round(dominant_angle, 1),
    }

    if return_grid:
        grid_size = config.SCREENSPACE_FLOW_GRID_SIZE
        min_mag = config.SCREENSPACE_FLOW_GRID_MIN_MAG
        gh, gw = mag.shape[:2]
        step_y = max(1, gh // grid_size)
        step_x = max(1, gw // grid_size)
        grid: list[dict[str, float]] = []
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


def compute_scene_fingerprint(region_pixels: np.ndarray) -> dict[str, Any]:
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
    color_stats: list[float] = []
    for ch in range(3):
        channel = region_pixels[:, :, ch].astype(np.float64)
        color_stats.extend([float(np.mean(channel)), float(np.std(channel))])

    return {
        "histogram": hist,
        "edge_density": edge_density,
        "color_stats": color_stats,
    }


def compare_scene_fingerprints(
    fp_a: dict[str, Any],
    fp_b: dict[str, Any],
) -> float:
    """Compare two scene fingerprints.

    Returns similarity score 0.0–1.0.
    """
    # Histogram correlation: range [-1, 1] → [0, 1]
    # Flatten 3D histograms to 1D — cv2.compareHist returns incorrect
    # results for multidimensional arrays.
    hist_corr = cv2.compareHist(
        fp_a["histogram"].flatten().astype(np.float32),
        fp_b["histogram"].flatten().astype(np.float32),
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
    region: dict[str, int] | None,
    interval_seconds: float,
    callback: Callable[[float, np.ndarray], bool | None],
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    fps: float = 0.0,
    duration: float = 0.0,
    fast_opts: dict[str, Any] | None = None,
) -> None:
    """Iterate through video at interval, extract region, call callback.

    The *callback* receives ``(timestamp_seconds, region_pixels)`` and may
    return ``False`` to stop iteration early.  When *region* is ``None``
    the full frame is passed (used by template detection).

    When *fps* and *duration* are provided, skips internal metadata reads.
    Uses sequential frame reading (grab/retrieve) for small intervals
    to avoid expensive H.264 seeking.

    *fast_opts* enables fast-scan optimizations when provided:
    - ``phash_skip``: skip frames whose perceptual hash is unchanged
    - ``max_region_dim``: downscale extracted region to this max dimension
    """
    full_frame = region is None

    # Extract frames via ffmpeg pipe (faster H.264 decoding, no cv2.VideoCapture)
    if not _scan_via_ffmpeg_pipe(
        video_path,
        None if full_frame else region,
        interval_seconds,
        callback,
        start_seconds=start_seconds,
        end_seconds=end_seconds if end_seconds is not None else 0.0,
        fps=fps,
        duration=duration,
        fast_opts=fast_opts,
        full_frame=full_frame,
    ):
        utils.warning_print(
            f"ffmpeg pipe extraction failed for {Path(video_path).name}"
        )


def scan_video_full_frames(
    video_path: str,
    interval_seconds: float,
    callback: Callable[[float, np.ndarray], bool | None],
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    fps: float = 0.0,
    duration: float = 0.0,
    fast_opts: dict[str, Any] | None = None,
) -> None:
    """Like :func:`scan_video_frames` but passes the full frame (no region crop).

    Thin wrapper that calls ``scan_video_frames`` with ``region=None``.
    """
    scan_video_frames(
        video_path,
        None,
        interval_seconds,
        callback,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        fps=fps,
        duration=duration,
        fast_opts=fast_opts,
    )


# ---------------------------------------------------------------------------
# Batch frame extraction via ffmpeg pipe (experiment 2E)
# ---------------------------------------------------------------------------


def _ffmpeg_pipe_frames(
    video_path: str,
    interval_seconds: float,
    *,
    start_seconds: float = 0.0,
    end_seconds: float = 0.0,
    region: dict[str, int] | None = None,
    frame_width: int = 0,
    frame_height: int = 0,
    max_dim: int = 0,
) -> Iterator[tuple[float, np.ndarray]]:
    """Yield ``(timestamp, frame)`` tuples extracted via an ffmpeg pipe.

    Uses a single ffmpeg process with ``-f rawvideo`` piped to stdout,
    which is typically faster than per-frame ``cv2.VideoCapture`` seeking
    for H.264 content.

    *region* applies an ffmpeg ``crop`` filter so only the ROI pixels are
    decoded and transferred.  *max_dim* adds a ``scale`` filter to cap the
    largest output dimension (useful for fast-scan downscaling).

    The caller can stop iteration at any time (e.g. on cancel); the
    ``finally`` block ensures the subprocess is cleaned up.
    """
    if not shutil.which("ffmpeg"):
        return

    # Determine output dimensions
    filters = [f"fps=1/{interval_seconds}"]

    if region:
        filters.append(f"crop={region['w']}:{region['h']}:{region['x']}:{region['y']}")
        out_w, out_h = region["w"], region["h"]
    else:
        out_w, out_h = frame_width, frame_height

    if out_w <= 0 or out_h <= 0:
        return

    if max_dim > 0 and (out_w > max_dim or out_h > max_dim):
        scale = max_dim / max(out_w, out_h)
        out_w = int(out_w * scale)
        out_h = int(out_h * scale)
        # Ensure even dimensions for rawvideo
        out_w += out_w % 2
        out_h += out_h % 2
        filters.append(f"scale={out_w}:{out_h}")

    cmd: list[str] = ["ffmpeg"]
    if start_seconds > 0:
        cmd += ["-ss", str(start_seconds)]
    cmd += ["-i", video_path]
    if end_seconds > start_seconds:
        cmd += ["-t", str(end_seconds - start_seconds)]
    cmd += [
        "-vf",
        ",".join(filters),
        "-pix_fmt",
        "bgr24",
        "-f",
        "rawvideo",
        "-loglevel",
        "error",
        "pipe:1",
    ]

    frame_size = out_w * out_h * 3
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    assert proc.stdout is not None  # guaranteed by stdout=PIPE
    frame_idx = 0
    try:
        while True:
            raw = proc.stdout.read(frame_size)
            if len(raw) < frame_size:
                break
            frame = np.frombuffer(raw, dtype=np.uint8).reshape((out_h, out_w, 3)).copy()
            ts = start_seconds + frame_idx * interval_seconds
            if end_seconds > 0 and ts > end_seconds:
                break
            yield (ts, frame)
            frame_idx += 1
    finally:
        if proc.stdout:
            proc.stdout.close()
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait()


def _scan_via_ffmpeg_pipe(
    video_path: str,
    region: dict[str, int] | None,
    interval_seconds: float,
    callback: Callable[[float, np.ndarray], bool | None],
    *,
    start_seconds: float = 0.0,
    end_seconds: float = 0.0,
    fps: float = 0.0,
    duration: float = 0.0,
    fast_opts: dict[str, Any] | None = None,
    full_frame: bool = False,
) -> bool:
    """Try to scan frames via ffmpeg pipe, calling *callback* for each.

    Returns ``True`` if the ffmpeg path succeeded (caller should skip the
    cv2 fallback), ``False`` if it failed and the caller should fall back
    to cv2-based extraction.
    """
    if not shutil.which("ffmpeg"):
        return False

    props = video.probe_video_properties(video_path)
    if not props:
        return False
    frame_width = props.get("width", 0)
    frame_height = props.get("height", 0)
    if frame_width <= 0 or frame_height <= 0:
        return False

    if end_seconds <= 0:
        end_seconds = duration

    # Fast-scan state
    _phash_skip = bool(fast_opts and fast_opts.get("phash_skip"))
    _max_dim = (fast_opts or {}).get("max_region_dim", 0)
    _phash_thresh = (fast_opts or {}).get(
        "phash_threshold", config.SCREENSPACE_FAST_SCAN_PHASH_THRESHOLD
    )
    _prev_phash: list["imagehash.ImageHash | None"] = [None]

    pipe_region = None if full_frame else region
    # For the pipe, push max_dim downscaling into ffmpeg when phash_skip is off.
    # When phash_skip is on, we need the un-downscaled frame for hashing, so
    # we downscale in Python after the hash check.
    pipe_max_dim = _max_dim if (not _phash_skip and _max_dim > 0) else 0

    try:
        for ts, frame in _ffmpeg_pipe_frames(
            video_path,
            interval_seconds,
            start_seconds=start_seconds,
            end_seconds=end_seconds,
            region=pipe_region,
            frame_width=frame_width,
            frame_height=frame_height,
            max_dim=pipe_max_dim,
        ):
            if _phash_skip:
                fh = compute_phash(frame)
                if _prev_phash[0] is not None and fh - _prev_phash[0] <= _phash_thresh:
                    continue
                _prev_phash[0] = fh
                # Downscale after phash if needed
                if _max_dim > 0:
                    rh, rw = frame.shape[:2]
                    if rh > _max_dim or rw > _max_dim:
                        sc = _max_dim / max(rh, rw)
                        frame = cv2.resize(
                            frame,
                            (int(rw * sc), int(rh * sc)),
                            interpolation=cv2.INTER_AREA,
                        )

            result = callback(ts, frame)
            if result is False:
                break

        return True  # ffmpeg pipe succeeded (even if video had 0 frames)
    except Exception as exc:
        utils.warning_print(f"ffmpeg pipe scan failed, falling back to cv2: {exc}")
        return False


def build_timelapse_command(
    video_path: str,
    region: dict[str, int],
    speedup_factor: float,
    output_path: str,
    output_format: str = "mp4",
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    sample_interval: float = 0.0,
) -> list[str]:
    """Construct ffmpeg argv for a cropped timelapse.

    *sample_interval* (seconds) controls frame sampling: when > 0, only one
    frame per interval is kept before cropping and speed-up.  0 means every
    frame is used (default).
    """
    x, y, w, h = region["x"], region["y"], region["w"], region["h"]
    filters: list[str] = []
    if sample_interval > 0:
        filters.append(f"fps=1/{sample_interval}")
    filters.append(f"crop={w}:{h}:{x}:{y}")
    filters.append(f"setpts=PTS/{speedup_factor}")
    vf = ",".join(filters)

    cmd = [
        "ffmpeg",
        "-y",
        "-loglevel",
        config.FFMPEG_LOGLEVEL,
    ]

    if start_seconds > 0:
        cmd += ["-ss", str(start_seconds)]

    cmd += ["-i", video_path]

    if end_seconds is not None and end_seconds > start_seconds:
        cmd += ["-t", str(end_seconds - start_seconds)]

    cmd += ["-vf", vf, "-an"]

    if output_format == "gif":
        cmd.extend(["-loop", "0"])
    else:
        cmd.extend(["-c:v", "libx264", "-preset", "fast", "-crf", "23"])

    cmd.append(output_path)
    return cmd


def _probe_video_meta(video_path: str) -> tuple[float, float]:
    """Return ``(fps, duration)`` via ffprobe."""
    props = video.probe_video_properties(video_path)
    if props and props.get("fps", 0) > 0 and props.get("duration", 0) > 0:
        return (props["fps"], props["duration"])
    return (0.0, 0.0)


# ---------------------------------------------------------------------------
# Analysis workflows
# ---------------------------------------------------------------------------


def scan_color(
    video_path: str,
    region: dict[str, int],
    target_color: dict[str, float],
    tolerance: dict[str, float],
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan video for frames where region color matches target.

    Returns list of ``{start, end, duration}`` spans (consecutive matches
    merged).
    """
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    matches: list[float] = []

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        matched, conf = color_matches(pixels, target_color, tolerance)
        if matched:
            matches.append(ts)
            if on_result:
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
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return _merge_timestamp_spans(matches, interval_seconds)


def scan_changes(
    video_path: str,
    region: dict[str, int],
    threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    noise_threshold: int = 0,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan video for change points in a region.

    Returns list of ``{timestamp, magnitude}`` dicts.
    """
    if threshold <= 0:
        threshold = config.SCREENSPACE_CHANGE_RATIO_THRESHOLD
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL
    if noise_threshold <= 0:
        noise_threshold = config.SCREENSPACE_NOISE_THRESHOLD

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    prev_gray: list[np.ndarray | None] = [None]
    k = config.SCREENSPACE_BLUR_KERNEL
    mk = config.SCREENSPACE_MORPH_KERNEL
    morph_kernel = np.ones((mk, mk), np.uint8)

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
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
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def scan_similarity(
    video_path: str,
    region: dict[str, int],
    reference_frame: np.ndarray,
    threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Find frames where region is similar to a reference.

    Returns list of ``{timestamp, score}`` dicts, sorted by score
    descending.
    """
    from skimage.metrics import structural_similarity as ssim

    if threshold <= 0:
        threshold = config.SCREENSPACE_SSIM_THRESHOLD
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    ref_phash = compute_phash(reference_frame)
    phash_threshold = config.SCREENSPACE_PHASH_THRESHOLD

    # Pre-resize and preprocess reference frame once
    max_dim = 256
    rh, rw = reference_frame.shape[:2]
    if rh > max_dim or rw > max_dim:
        scale = max_dim / max(rh, rw)
        cmp_w, cmp_h = int(rw * scale), int(rh * scale)
        ref_resized = cv2.resize(
            reference_frame, (cmp_w, cmp_h), interpolation=cv2.INTER_AREA
        )
    else:
        cmp_w, cmp_h = rw, rh
        ref_resized = reference_frame
    bk = config.SCREENSPACE_BLUR_KERNEL
    ref_gray = cv2.cvtColor(
        cv2.GaussianBlur(ref_resized, (bk, bk), 0), cv2.COLOR_BGR2GRAY
    )

    prev_skip_gray: list[np.ndarray | None] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
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
            # Always resize candidate to match reference dimensions for SSIM
            ph, pw = pixels.shape[:2]
            if pw != cmp_w or ph != cmp_h:
                cand = cv2.resize(pixels, (cmp_w, cmp_h), interpolation=cv2.INTER_AREA)
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
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    results.sort(key=lambda r: r["score"], reverse=True)
    return results


def scan_text(
    video_path: str,
    region: dict[str, int],
    search_string: str,
    interval_seconds: float = 2.0,
    *,
    fuzzy_threshold: float = 0.0,
    languages: list[str] | None = None,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
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

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    search_lower = search_string.lower()
    prev_gray: list[np.ndarray | None] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
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
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


_NUMBERS_RE = re.compile(r"-?\d+(?:\.\d+)?")
_VALID_OPERATORS = ("eq", "gt", "lt", "gte", "lte", "range")


def _number_matches(
    value: float,
    operator: str,
    target_value: float = 0,
    range_min: float | None = None,
    range_max: float | None = None,
) -> bool:
    """Check if *value* satisfies the given numeric comparison."""
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


def scan_numbers(
    video_path: str,
    region: dict[str, int],
    operator: str,
    target_value: float = 0,
    interval_seconds: float = 2.0,
    *,
    range_min: float | None = None,
    range_max: float | None = None,
    languages: list[str] | None = None,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
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

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    prev_gray: list[np.ndarray | None] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
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
                if _number_matches(num, operator, target_value, range_min, range_max):
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
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def generate_timelapse(
    video_path: str,
    region: dict[str, int],
    speedup_factor: float,
    output_path: str,
    output_format: str = "mp4",
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    sample_interval: float = 0.0,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
) -> str | None:
    """Generate a cropped timelapse via ffmpeg.

    *sample_interval* controls frame sampling (seconds between captured
    frames).  0 means every frame is used.

    Reports encoding progress through *on_progress* by parsing ffmpeg's
    ``-progress`` output.  Checks *cancel_flag* periodically to allow
    early termination.

    Returns output file path on success, ``None`` on failure.
    """
    cmd = build_timelapse_command(
        video_path,
        region,
        speedup_factor,
        output_path,
        output_format,
        start_seconds=start_seconds,
        end_seconds=end_seconds,
        sample_interval=sample_interval,
    )

    # Estimate output duration for progress calculation.
    input_duration = 0.0
    if end_seconds is not None and end_seconds > start_seconds:
        input_duration = end_seconds - start_seconds
    else:
        _, vid_dur = _probe_video_meta(video_path)
        if vid_dur > 0:
            input_duration = max(vid_dur - start_seconds, 0.0)

    expected_out_us = (
        (input_duration / speedup_factor * 1_000_000) if input_duration > 0 else 0
    )

    # Use Popen with -progress to get real-time encoding updates
    cmd += ["-progress", "pipe:1"]
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        assert proc.stdout is not None  # guaranteed by stdout=PIPE
    except (FileNotFoundError, OSError):
        return None

    if on_progress:
        on_progress(0.0)

    try:
        for line in proc.stdout:
            text = line.decode("utf-8", errors="replace").strip()
            if text.startswith("out_time_us=") and expected_out_us > 0 and on_progress:
                try:
                    us = int(text.split("=", 1)[1])
                    on_progress(min(us / expected_out_us, 0.99))
                except (ValueError, ZeroDivisionError):
                    pass
            if cancel_flag and cancel_flag():
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    proc.wait()
                return None
    finally:
        if proc.stdout:
            proc.stdout.close()
        proc.wait()

    if on_progress:
        on_progress(1.0)
    if proc.returncode != 0:
        return None
    return output_path


def scan_template(
    video_path: str,
    region: dict[str, int],
    template_image: np.ndarray,
    threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    template_mask: np.ndarray | None = None,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
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

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []

    _tmpl_downscale = bool(fast_opts and fast_opts.get("template_downscale"))

    def _cb(ts: float, frame: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        work_frame = frame
        scale_back = 1
        if _tmpl_downscale:
            fh, fw = frame.shape[:2]
            nw, nh = fw // 2, fh // 2
            if nw > 0 and nh > 0:
                work_frame = cv2.resize(frame, (nw, nh), interpolation=cv2.INTER_AREA)
                scale_back = 2
        matches = match_template(
            work_frame, template_image, threshold=threshold, mask=template_mask
        )
        if matches:
            if scale_back > 1:
                for m in matches:
                    m["x"] *= scale_back
                    m["y"] *= scale_back
                    m["w"] *= scale_back
                    m["h"] *= scale_back
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
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def scan_flow(
    video_path: str,
    region: dict[str, int],
    magnitude_threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan video for motion in a region using dense optical flow.

    Returns list of ``{timestamp, magnitude, angle}`` dicts where
    magnitude exceeds *magnitude_threshold*.
    """
    if magnitude_threshold <= 0:
        magnitude_threshold = config.SCREENSPACE_FLOW_MAGNITUDE_THRESHOLD
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    prev_gray: list[np.ndarray | None] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False
        curr_gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
        if prev_gray[0] is not None:
            flow_result = compute_optical_flow(
                prev_gray[0], curr_gray, return_grid=True
            )
            if flow_result["magnitude"] >= magnitude_threshold:
                rd: dict[str, Any] = {
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
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def scan_scene(
    video_path: str,
    region: dict[str, int],
    reference_scenes: list[dict[str, Any]],
    threshold: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Classify frames by similarity to reference scene fingerprints.

    *reference_scenes* is a list of ``{name: str, frame: np.ndarray}``
    dicts, optionally with a per-scene ``threshold`` float.  Each
    frame's region is fingerprinted and compared against all references;
    the best match above its threshold is reported.

    Returns list of ``{timestamp, scene_name, score}`` dicts.
    """
    default_threshold = (
        threshold if threshold > 0 else config.SCREENSPACE_SCENE_SIMILARITY_THRESHOLD
    )
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    # Pre-compute fingerprints for reference scenes (with per-scene thresholds)
    ref_fps: list[tuple[str, dict[str, Any], float]] = []
    for ref in reference_scenes:
        fp = compute_scene_fingerprint(ref["frame"])
        ref_thresh = float(ref.get("threshold", default_threshold))
        ref_fps.append((ref["name"], fp, ref_thresh))

    if not ref_fps:
        return []

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    prev_skip_gray: list[np.ndarray | None] = [None]

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
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
        best_thresh = default_threshold
        for ref_name, ref_fp, ref_thresh in ref_fps:
            score = compare_scene_fingerprints(fp, ref_fp)
            if score > best_score:
                best_score = score
                best_name = ref_name
                best_thresh = ref_thresh

        if best_score >= best_thresh:
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
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def scan_inactivity(
    video_path: str,
    region: dict[str, int],
    threshold: int = 0,
    min_duration: float = 0.0,
    interval_seconds: float = 0.0,
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Scan video for inactivity spans (near-duplicate consecutive frames).

    Uses perceptual hashing to compare consecutive frames.  When frames
    remain similar (Hamming distance <= *threshold*) for at least
    *min_duration* seconds, a span result is emitted.

    Returns list of ``{start, end, duration, avg_distance}`` dicts.
    """
    if threshold <= 0:
        threshold = config.SCREENSPACE_INACTIVITY_PHASH_THRESHOLD
    if min_duration <= 0:
        min_duration = config.SCREENSPACE_INACTIVITY_MIN_DURATION
    if interval_seconds <= 0:
        interval_seconds = config.SCREENSPACE_DEFAULT_INTERVAL

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if end_seconds is None or end_seconds > vid_duration:
        end_seconds = vid_duration
    total_range = end_seconds - start_seconds

    results: list[dict[str, Any]] = []
    prev_hash: list["imagehash.ImageHash | None"] = [None]
    span_start: list[float | None] = [None]
    span_distances: list[list[int]] = [[]]
    last_ts: list[float] = [start_seconds]

    def _flush_span(span_end: float) -> None:
        if span_start[0] is not None:
            dur = span_end - span_start[0]
            if dur >= min_duration:
                dists = span_distances[0]
                avg_dist = sum(dists) / len(dists) if dists else 0.0
                rd = {
                    "start": round(span_start[0], 2),
                    "end": round(span_end, 2),
                    "duration": round(dur, 2),
                    "avg_distance": round(avg_dist, 2),
                }
                results.append(rd)
                if on_result:
                    on_result(rd)
        span_start[0] = None
        span_distances[0] = []

    def _cb(ts: float, pixels: np.ndarray) -> bool | None:
        if cancel_flag and cancel_flag():
            return False

        curr_hash = compute_phash(pixels)
        last_ts[0] = ts

        if prev_hash[0] is not None:
            dist = int(curr_hash - prev_hash[0])
            if dist <= threshold:
                # Frame is similar — extend or start span
                if span_start[0] is None:
                    span_start[0] = ts - interval_seconds
                span_distances[0].append(dist)
            else:
                # Frame changed — flush any active span
                _flush_span(ts - interval_seconds)
        prev_hash[0] = curr_hash

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
        fast_opts=fast_opts,
    )

    # Flush final span if video ended during an inactive period
    _flush_span(last_ts[0])

    if on_progress:
        on_progress(1.0)
    return results


# ---------------------------------------------------------------------------
# Multitool: per-frame evaluation and multi-factor scan
# ---------------------------------------------------------------------------


def _extract_confidence(tool_type: str, result: dict[str, Any]) -> float:
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
    elif tool_type == "inactivity":
        return min(result.get("duration", 0.0) / 30.0, 1.0)
    return 1.0


def check_frame_for_tool(
    frame: np.ndarray,
    prev_frame: np.ndarray | None,
    region: dict[str, int],
    tool_type: str,
    parameters: dict[str, Any],
) -> tuple[bool, dict[str, Any] | None]:
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
        matched, conf = color_matches(pixels, target, tol)
        if matched:
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
            for match in _NUMBERS_RE.findall(cleaned):
                num = float(match)
                if _number_matches(num, operator, target_value, range_min, range_max):
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
            # Cache fingerprint on the reference dict to avoid recomputing per frame
            ref_fp = ref.get("_cached_fingerprint")
            if ref_fp is None:
                ref_fp = compute_scene_fingerprint(ref["frame"])
                ref["_cached_fingerprint"] = ref_fp
            score = compare_scene_fingerprints(fp, ref_fp)
            if score > best_score:
                best_score = score
                best_name = ref["name"]
        if best_score >= threshold:
            return True, {"scene_name": best_name, "score": round(best_score, 4)}
        return False, None

    elif tool_type == "inactivity":
        if prev_frame is None:
            return False, None
        pixels = extract_region(frame, region)
        prev_pixels = extract_region(prev_frame, region)
        thresh = parameters.get(
            "threshold", config.SCREENSPACE_INACTIVITY_PHASH_THRESHOLD
        )
        curr_h = compute_phash(pixels)
        prev_h = compute_phash(prev_pixels)
        dist = int(curr_h - prev_h)
        if dist <= thresh:
            return True, {"distance": dist}
        return False, None

    return False, None


def scan_multitool(
    video_path: str,
    region: dict[str, int],
    steps: list[dict[str, Any]],
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Run a multi-factor scan chaining several tool types.

    Iterates the video once, checking all steps per frame.  A frame
    must pass every step (in order) to be included in the results.
    Results are emitted incrementally via *on_result* as each passing
    frame is found.

    Each entry in *steps* is a dict with ``"type"`` plus the tool's
    own parameters (e.g. ``target_color``, ``tolerance`` for color).

    Returns a list of ``{timestamp, tool_types, steps, min_confidence}``
    dicts.
    """
    if len(steps) < 2:
        raise ValueError("Multitool requires at least 2 steps")

    interval = steps[0].get("interval", config.SCREENSPACE_DEFAULT_INTERVAL)

    # Compute iteration range (intersection of all per-step ranges)
    scan_start = start_seconds
    scan_end = end_seconds
    for step in steps:
        s_start = step.get("start_seconds")
        if s_start is not None:
            scan_start = max(scan_start, s_start)
        s_end = step.get("end_seconds")
        if s_end is not None:
            scan_end = min(scan_end, s_end) if scan_end is not None else s_end

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if scan_end is None or scan_end > vid_duration:
        scan_end = vid_duration
    total_range = scan_end - scan_start

    tool_types = [s["type"] for s in steps]
    step_regions = [s.get("region_coords", region) for s in steps]
    prev_frame: list[np.ndarray | None] = [None]
    results: list[dict[str, Any]] = []

    def _cancel() -> bool:
        return bool(cancel_flag and cancel_flag())

    def _cb(ts: float, frame: np.ndarray) -> bool | None:
        if _cancel():
            return False

        step_results: list[dict[str, Any]] = []
        for i, step in enumerate(steps):
            passed, rd = check_frame_for_tool(
                frame, prev_frame[0], step_regions[i], step["type"], step
            )
            if not passed or rd is None:
                break
            step_results.append(rd)

        prev_frame[0] = frame

        if len(step_results) == len(steps):
            confidences = [
                _extract_confidence(steps[i]["type"], sr)
                for i, sr in enumerate(step_results)
            ]
            rd = {
                "timestamp": round(ts, 2),
                "tool_types": tool_types,
                "steps": step_results,
                "min_confidence": round(min(confidences), 4),
            }
            results.append(rd)
            if on_result:
                on_result(rd)

        if on_progress and total_range > 0:
            on_progress((ts - scan_start) / total_range)
        return None

    scan_video_full_frames(
        video_path,
        interval,
        _cb,
        start_seconds=scan_start,
        end_seconds=scan_end,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def _merge_timestamp_spans(
    timestamps: list[float], interval: float
) -> list[dict[str, Any]]:
    """Merge consecutive matched timestamps into spans."""
    if not timestamps:
        return []
    timestamps.sort()
    spans: list[dict[str, Any]] = []
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
    region_coords: dict[str, int],
    parameters: dict[str, Any] | None = None,
) -> dict[str, Any]:
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
    results: list[dict[str, Any]],
    frame_width: int,
    frame_height: int,
    output_path: str,
) -> str | None:
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
    results: list[dict[str, Any]],
    region_width: int,
    region_height: int,
    output_path: str,
) -> str | None:
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


def _accumulate_heatmap_result(
    accumulator: np.ndarray,
    result: dict[str, Any],
    heatmap_type: str,
) -> None:
    """Add a single result's contribution to a heatmap accumulator."""
    acc_h, acc_w = accumulator.shape[:2]
    if heatmap_type == "template":
        for m in result.get("matches", []):
            x, y, w, h = int(m["x"]), int(m["y"]), int(m["w"]), int(m["h"])
            y2 = min(y + h, acc_h)
            x2 = min(x + w, acc_w)
            accumulator[y:y2, x:x2] += m.get("score", 1.0)
    elif heatmap_type == "flow":
        for cell in result.get("flow_grid", []):
            cx = int(cell["x"] * (acc_w - 1))
            cy = int(cell["y"] * (acc_h - 1))
            radius = max(1, acc_w // 16)
            cv2.circle(accumulator, (cx, cy), radius, float(cell["mag"]), -1)


def generate_heatmap_gif(
    results: list[dict[str, Any]],
    width: int,
    height: int,
    output_path: str,
    heatmap_type: str = "template",
    num_frames: int = 24,
    frame_duration_ms: int = 120,
) -> str | None:
    """Generate an animated GIF showing heatmap accumulation over time.

    Divides *results* into *num_frames* temporal buckets, progressively
    accumulates heatmap data, and writes frames as an animated GIF.
    """
    from PIL import Image

    if not results:
        return None

    num_frames = min(num_frames, len(results))
    if num_frames < 2:
        return None

    acc_h, acc_w = height, width
    if heatmap_type == "flow":
        acc_h = acc_w = 256

    # First pass: compute global max for consistent normalization
    global_acc = np.zeros((acc_h, acc_w), dtype=np.float32)
    for r in results:
        _accumulate_heatmap_result(global_acc, r, heatmap_type)
    global_max = float(global_acc.max())
    if global_max == 0:
        return None

    # Second pass: build progressive frames
    accumulator = np.zeros((acc_h, acc_w), dtype=np.float32)
    frames: list[Image.Image] = []
    bucket_size = max(1, len(results) // num_frames)

    for frame_idx in range(num_frames):
        start_idx = frame_idx * bucket_size
        end_idx = (
            len(results)
            if frame_idx == num_frames - 1
            else min((frame_idx + 1) * bucket_size, len(results))
        )
        for r_idx in range(start_idx, end_idx):
            _accumulate_heatmap_result(accumulator, results[r_idx], heatmap_type)

        normalized = (accumulator / global_max * 255).astype(np.uint8)
        normalized = cv2.GaussianBlur(normalized, (15, 15), 0)
        colored = cv2.applyColorMap(normalized, cv2.COLORMAP_JET)

        if heatmap_type == "flow":
            colored = cv2.resize(
                colored, (width, height), interpolation=cv2.INTER_LINEAR
            )

        rgb = cv2.cvtColor(colored, cv2.COLOR_BGR2RGB)
        frames.append(Image.fromarray(rgb))

    if not frames:
        return None

    frames[0].save(
        output_path,
        save_all=True,
        append_images=frames[1:],
        duration=frame_duration_ms,
        loop=0,
    )
    return output_path


class ScreenspaceWorker:
    """Background thread that processes analysis tasks sequentially."""

    def __init__(self) -> None:
        self._queue: queue.PriorityQueue[Any] = queue.PriorityQueue()
        self._tasks: dict[str, dict[str, Any]] = {}
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._running = False
        self._paused = threading.Event()
        self.on_task_complete: Callable[[], None] | None = None
        self.on_progress_update: Callable[[], None] | None = None

    def start(self) -> None:
        """Start the worker thread."""
        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def restore_tasks(self, tasks: list[dict[str, Any]]) -> None:
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
            self._thread.join(timeout=15)

    def enqueue(self, task: dict[str, Any]) -> str:
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

    def get_task(self, task_id: str) -> dict[str, Any] | None:
        """Return a task dict by ID (thread-safe copy)."""
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return None
            return copy.deepcopy(task)

    def get_all_tasks(self) -> list[dict[str, Any]]:
        """Return all tasks (thread-safe copies)."""
        with self._lock:
            return [copy.deepcopy(t) for t in self._tasks.values()]

    def reorder(self, task_ids: list[str]) -> bool:
        """Reorder queued tasks by the given ID sequence.

        Assigns new priorities so that earlier IDs in the list have
        lower (higher-priority) values.  Drains and re-inserts queue
        items so the PriorityQueue heap reflects the new ordering.
        """
        with self._lock:
            for i, tid in enumerate(task_ids):
                task = self._tasks.get(tid)
                if task and task["status"] == TASK_STATUS_QUEUED:
                    task["priority"] = i + 1

            # Drain the queue and re-insert with updated priorities
            pending_items: list[Any] = []
            while not self._queue.empty():
                try:
                    pending_items.append(self._queue.get_nowait())
                except queue.Empty:
                    break
            for priority, created_at, tid in pending_items:
                task = self._tasks.get(tid) if isinstance(tid, str) else None
                new_priority = task["priority"] if task else priority
                self._queue.put((new_priority, created_at, tid))
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
        to_resume: list[dict[str, Any]] = []
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
                _, end = _probe_video_meta(task["video_path"])
            resume_at = start + progress * (end - start)

            with self._lock:
                task["_partial_results"] = list(task.get("result") or [])
                task["_progress_offset"] = progress
                task["_progress_scale"] = max(1.0 - progress, 0.001)
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

    def drain_new_events(self) -> list[dict[str, Any]]:
        """Collect and clear ``_generated_events`` from all tasks. Thread-safe."""
        events: list[dict[str, Any]] = []
        with self._lock:
            for t in self._tasks.values():
                events.extend(t.pop("_generated_events", []))
        return events

    def _generate_events_from_results(
        self, task: dict[str, Any], raw_results: list[dict[str, Any]]
    ) -> list[dict[str, Any]]:
        """Convert raw task results into ScreenspaceEvent records."""
        task_type = task.get("type", "")
        if task_type == "timelapse":
            return []
        events: list[dict[str, Any]] = []
        for r in raw_results:
            ts = r.get("timestamp", r.get("start", 0.0))
            confidence = _extract_confidence(task_type, r)
            metadata: dict[str, Any] = {}
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
            elif task_type == "inactivity":
                metadata["duration"] = r.get("duration", 0.0)
                metadata["avg_distance"] = r.get("avg_distance", 0.0)
            ev = create_event(task, ts, confidence, metadata)
            if task_type == "inactivity" and "end" in r:
                ev["time_out"] = round(r["end"], 2)
            events.append(ev)
        return events

    def _generate_heatmap(
        self, task: dict[str, Any], results: list[dict[str, Any]]
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
            gif_path = str(
                Path(utils.get_effective_output_dir()) / f"heatmap_{task_id}.gif"
            )
            gp = generate_heatmap_gif(
                results, fw, fh, gif_path, heatmap_type="template"
            )
            if gp:
                task["heatmap_gif"] = Path(gp).name
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
            gif_path = str(
                Path(utils.get_effective_output_dir()) / f"heatmap_{task_id}.gif"
            )
            gp = generate_heatmap_gif(results, rw, rh, gif_path, heatmap_type="flow")
            if gp:
                task["heatmap_gif"] = Path(gp).name

    def _run(self) -> None:
        """Worker loop with concurrent task execution via ThreadPoolExecutor."""
        max_workers = config.SCREENSPACE_PARALLEL_WORKERS
        with ThreadPoolExecutor(max_workers=max_workers) as pool:
            active: dict[str, Future[None]] = {}

            while self._running:
                try:
                    # 1. Collect completed futures
                    done_ids = [tid for tid, f in active.items() if f.done()]
                    for tid in done_ids:
                        future = active.pop(tid)
                        try:
                            future.result()
                        except Exception as exc:
                            utils.warning_print(f"Worker task {tid} raised: {exc}")
                        if self.on_task_complete:
                            try:
                                self.on_task_complete()
                            except Exception as exc:
                                utils.warning_print(
                                    f"on_task_complete callback failed: {exc}"
                                )
                        if self.on_progress_update:
                            try:
                                self.on_progress_update()
                            except Exception:
                                pass

                    # 2. If paused, wait
                    if self._paused.is_set():
                        time.sleep(0.25)
                        continue

                    # 3. Submit new tasks if capacity available
                    if len(active) >= max_workers:
                        time.sleep(0.1)
                        continue

                    try:
                        item = self._queue.get(timeout=0.25)
                    except queue.Empty:
                        continue

                    priority, created_at, task_id = item
                    if task_id is _SENTINEL:
                        # Wait for active tasks to finish
                        for f in active.values():
                            try:
                                f.result(timeout=10)
                            except Exception:
                                pass
                        if self.on_task_complete:
                            try:
                                self.on_task_complete()
                            except Exception:
                                pass
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

                    future = pool.submit(self._execute_task, task)
                    active[task_id] = future

                except Exception as exc:
                    utils.warning_print(f"Worker loop error: {exc}")

    def _execute_task(self, task: dict[str, Any]) -> None:
        """Dispatch task to the appropriate workflow function."""
        task_id = task["id"]

        # Seed incremental results (includes partial results on resume)
        with self._lock:
            t = self._tasks.get(task_id)
            if t:
                t["result"] = list(t.get("_partial_results", []))
                t["_raw_results"] = list(t.get("_partial_results", []))

        # Per-task throttle for SSE notifications (avoids cross-task race)
        last_notify: list[float] = [0.0]

        def _throttled_notify() -> None:
            now = time.monotonic()
            if now - last_notify[0] >= 0.5:
                last_notify[0] = now
                if self.on_progress_update:
                    self.on_progress_update()

        def _on_progress(progress: float) -> None:
            with self._lock:
                t = self._tasks.get(task_id)
                if t and not t.get("_paused_flag"):
                    offset = t.get("_progress_offset", 0.0)
                    scale = t.get("_progress_scale", 1.0)
                    t["progress"] = min(offset + progress * scale, 1.0)
            _throttled_notify()

        def _cancel_flag() -> bool:
            with self._lock:
                t = self._tasks.get(task_id)
                return bool(t and (t.get("_cancelled") or t.get("_paused_flag")))

        def _on_result(result_dict: dict[str, Any]) -> None:
            with self._lock:
                t = self._tasks.get(task_id)
                if t:
                    if isinstance(t.get("result"), list):
                        t["result"].append(result_dict)
                    if isinstance(t.get("_raw_results"), list):
                        t["_raw_results"].append(result_dict)
                    if t.get("parameters", {}).get("detect_first"):
                        t["_cancelled"] = True
            _throttled_notify()

        try:
            result = self._dispatch(task, _on_progress, _cancel_flag, _on_result)
            with self._lock:
                t = self._tasks.get(task_id)
                if t:
                    if t.get("_paused_flag"):
                        t["status"] = TASK_STATUS_PAUSED
                        t["result"] = result
                    elif t.get("_cancelled") and not t.get("parameters", {}).get(
                        "detect_first"
                    ):
                        t["status"] = TASK_STATUS_CANCELLED
                    else:
                        # Normal completion or detect_first early stop
                        partial = t.pop("_partial_results", None)
                        if partial and isinstance(result, list):
                            result = partial + result
                        t["status"] = TASK_STATUS_COMPLETED
                        t["result"] = result
                        t["progress"] = 1.0
                        t.pop("_progress_offset", None)
                        t.pop("_progress_scale", None)
                        raw = t.pop("_raw_results", [])
                        t["_generated_events"] = self._generate_events_from_results(
                            t, raw
                        )
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
        task: dict[str, Any],
        on_progress: Callable[[float], None],
        cancel_flag: Callable[[], bool],
        on_result: Callable[[dict[str, Any]], None] | None = None,
    ) -> Any:
        """Route task to the correct workflow function."""
        video_path = task["video_path"]
        region = task["region_coords"]
        params = task.get("parameters", {})
        task_type = task["type"]

        # Fast scan: apply interval multiplier and build optimization dict
        # (timelapse has its own sample_interval and does not use fast scan)
        scan_mode = params.get("scan_mode", "normal")
        fast_opts: dict[str, Any] | None = None
        if scan_mode == "fast" and task_type != "timelapse":
            multiplier = config.SCREENSPACE_FAST_SCAN_INTERVAL_MULTIPLIER
            if params.get("interval", 0) > 0:
                params["interval"] = params["interval"] * multiplier
            _fast_dims = {
                "color": 32,
                "change": 128,
                "similarity": 128,
                "flow": 128,
                "scene": 64,
                "inactivity": 64,
            }
            fast_opts = {
                "phash_skip": True,
                "max_region_dim": _fast_dims.get(task_type, 0),
            }
            if task_type == "template":
                fast_opts["template_downscale"] = True

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
                fast_opts=fast_opts,
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
                fast_opts=fast_opts,
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
                fast_opts=fast_opts,
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
                fast_opts=fast_opts,
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
                fast_opts=fast_opts,
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
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                sample_interval=params.get("sample_interval", 0.0),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
            )
        elif task_type == "template":
            template_img = params.get("template_image")
            if template_img is None:
                raise ValueError("Template scan requires a template_image parameter")
            tmpl_mask = params.get("template_mask")
            # Fast scan: downscale template + mask by 2x
            if scan_mode == "fast" and template_img is not None:
                th, tw = template_img.shape[:2]
                ntw, nth = tw // 2, th // 2
                if ntw > 0 and nth > 0:
                    template_img = cv2.resize(
                        template_img, (ntw, nth), interpolation=cv2.INTER_AREA
                    )
                    if tmpl_mask is not None:
                        tmpl_mask = cv2.resize(
                            tmpl_mask, (ntw, nth), interpolation=cv2.INTER_AREA
                        )
            return scan_template(
                video_path,
                region,
                template_image=template_img,
                threshold=params.get("threshold", 0),
                interval_seconds=params.get("interval", 0),
                template_mask=tmpl_mask,
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
                on_result=on_result,
                fast_opts=fast_opts,
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
                fast_opts=fast_opts,
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
                fast_opts=fast_opts,
            )
        elif task_type == "multitool":
            steps = params.get("steps", [])
            if not steps or len(steps) < 2:
                raise ValueError("Multitool requires at least 2 steps")
            # Fast scan: multiply the interval used by scan_multitool
            # (reads from steps[0]["interval"] with fallback to default)
            if scan_mode == "fast" and steps:
                mt_interval = steps[0].get(
                    "interval", config.SCREENSPACE_DEFAULT_INTERVAL
                )
                steps[0]["interval"] = (
                    mt_interval * config.SCREENSPACE_FAST_SCAN_INTERVAL_MULTIPLIER
                )
            return scan_multitool(
                video_path,
                region,
                steps=steps,
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
                on_result=on_result,
                fast_opts=fast_opts,
            )
        elif task_type == "inactivity":
            return scan_inactivity(
                video_path,
                region,
                threshold=params.get("threshold", 0),
                min_duration=params.get("min_duration", 0.0),
                interval_seconds=params.get("interval", 0),
                start_seconds=params.get("start_seconds", 0.0),
                end_seconds=params.get("end_seconds"),
                on_progress=on_progress,
                cancel_flag=cancel_flag,
                on_result=on_result,
                fast_opts=fast_opts,
            )
        else:
            raise ValueError(f"Unknown task type: {task_type}")


# ---------------------------------------------------------------------------
# Manifest I/O
# ---------------------------------------------------------------------------


def _empty_screenspace_manifest() -> dict[str, Any]:
    return {"regions": {}, "tasks": [], "events": [], "stashes": []}


def load_screenspace_manifest() -> dict[str, Any]:
    """Load the screenspace manifest from the output directory.

    Returns a dict with ``regions`` and ``tasks`` keys.
    """
    data = utils.load_json_manifest(config.SCREENSPACE_MANIFEST_FILENAME)
    if not isinstance(data, dict):
        return _empty_screenspace_manifest()
    return {
        "regions": data.get("regions", {}),
        "tasks": data.get("tasks", []),
        "events": data.get("events", []),
        "stashes": data.get("stashes", []),
    }


def save_screenspace_manifest(
    regions: dict[str, dict[str, Any]],
    tasks: list[dict[str, Any]],
    events: list[dict[str, Any]] | None = None,
    stashes: list[dict[str, Any]] | None = None,
) -> Path | None:
    """Write the screenspace manifest to disk.

    Strips internal fields (prefixed with ``_``) from tasks before writing.
    Returns the manifest path on success, or ``None`` on failure.
    """
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
            # Strip binary data and internal coords from multitool step parameters
            if "steps" in ct["parameters"]:
                _step_strip_keys = _binary_keys + ("region_coords",)
                ct["parameters"]["steps"] = [
                    {k: v for k, v in s.items() if k not in _step_strip_keys}
                    for s in ct["parameters"]["steps"]
                ]
        # Strip flow_grid from results (large per-frame data, not needed on disk)
        if isinstance(ct.get("result"), list):
            ct["result"] = [
                {k: v for k, v in r.items() if k != "flow_grid"} for r in ct["result"]
            ]
        clean_tasks.append(ct)
    return utils.save_json_manifest(
        config.SCREENSPACE_MANIFEST_FILENAME,
        {
            "regions": regions,
            "tasks": clean_tasks,
            "events": events or [],
            "stashes": stashes or [],
        },
        warn_label="screenspace manifest",
    )


def create_event(
    task: dict[str, Any],
    timestamp: float,
    confidence: float,
    metadata: dict[str, Any] | None = None,
) -> dict[str, Any]:
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
