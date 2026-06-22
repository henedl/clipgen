# -*- coding: utf-8 -*-
"""Screenspace heatmap generation (pure cv2/PIL leaf).

Template-, flow-, and change-match heatmap PNGs plus animated-GIF views: a
cumulative accumulation and a rolling-window (recent-only, fading) variant.
No sibling-module dependencies.
"""

from typing import TYPE_CHECKING, Any

import cv2
import numpy as np

if TYPE_CHECKING:
    from PIL import Image


# ---------------------------------------------------------------------------
# Heatmap generation
# ---------------------------------------------------------------------------


def _colorize_accumulator(accumulator: np.ndarray, max_val: float) -> np.ndarray:
    """Normalize by *max_val*, blur, and apply the JET colormap → BGR uint8."""
    normalized = (accumulator / max_val * 255).astype(np.uint8)
    normalized = cv2.GaussianBlur(normalized, (15, 15), 0)
    return cv2.applyColorMap(normalized, cv2.COLORMAP_JET)


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

    heatmap = _colorize_accumulator(accumulator, accumulator.max())
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

    heatmap = _colorize_accumulator(accumulator, accumulator.max())
    heatmap = cv2.resize(
        heatmap, (region_width, region_height), interpolation=cv2.INTER_LINEAR
    )
    cv2.imwrite(output_path, heatmap)
    return output_path


def generate_change_heatmap(
    results: list[dict[str, Any]],
    region_width: int,
    region_height: int,
    output_path: str,
) -> str | None:
    """Generate a heatmap PNG from accumulated per-frame change-mask grids.

    Uses ``change_grid`` data (downsampled change masks) from each result to
    paint where pixels changed most often across all detected change frames.
    """
    acc_size = 256
    accumulator = np.zeros((acc_size, acc_size), dtype=np.float32)
    for r in results:
        _accumulate_heatmap_result(accumulator, r, "change")

    if accumulator.max() == 0:
        return None

    heatmap = _colorize_accumulator(accumulator, accumulator.max())
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
    elif heatmap_type in ("flow", "change"):
        grid_key = "flow_grid" if heatmap_type == "flow" else "change_grid"
        for cell in result.get(grid_key, []):
            cx = int(cell["x"] * (acc_w - 1))
            cy = int(cell["y"] * (acc_h - 1))
            radius = max(1, acc_w // 16)
            cv2.circle(accumulator, (cx, cy), radius, float(cell["mag"]), -1)


def _heatmap_frame_image(
    accumulator: np.ndarray,
    global_max: float,
    heatmap_type: str,
    width: int,
    height: int,
) -> "Image.Image":
    """Colorize a heatmap accumulator into a PIL frame for animated GIFs.

    Region-scoped heatmaps (flow, change) accumulate at a fixed resolution and
    are resized to the requested frame size; template accumulates frame-native.
    """
    from PIL import Image

    colored = _colorize_accumulator(accumulator, global_max)
    if heatmap_type in ("flow", "change"):
        colored = cv2.resize(colored, (width, height), interpolation=cv2.INTER_LINEAR)
    rgb = cv2.cvtColor(colored, cv2.COLOR_BGR2RGB)
    return Image.fromarray(rgb)


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
    if not results:
        return None

    num_frames = min(num_frames, len(results))
    if num_frames < 2:
        return None

    acc_h, acc_w = height, width
    if heatmap_type in ("flow", "change"):
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

        frames.append(
            _heatmap_frame_image(accumulator, global_max, heatmap_type, width, height)
        )

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


def generate_rolling_heatmap_gif(
    results: list[dict[str, Any]],
    width: int,
    height: int,
    output_path: str,
    heatmap_type: str = "template",
    num_frames: int = 24,
    window_frames: int = 6,
    frame_duration_ms: int = 120,
) -> str | None:
    """Generate an animated GIF showing a sliding-window heatmap over time.

    Like :func:`generate_heatmap_gif`, but each frame accumulates only the
    *window_frames* most recent buckets instead of everything seen so far, so
    older heat fades out as the window advances ("rolling window"). Brightness
    is normalized against the densest window for stable frame-to-frame contrast.
    """
    if not results:
        return None

    num_frames = min(num_frames, len(results))
    if num_frames < 2:
        return None

    window_frames = max(1, window_frames)
    acc_h, acc_w = height, width
    if heatmap_type in ("flow", "change"):
        acc_h = acc_w = 256
    bucket_size = max(1, len(results) // num_frames)

    def _accumulate_window(frame_idx: int) -> np.ndarray:
        acc = np.zeros((acc_h, acc_w), dtype=np.float32)
        win_start = max(0, frame_idx - window_frames + 1)
        for bucket in range(win_start, frame_idx + 1):
            start_idx = bucket * bucket_size
            end_idx = (
                len(results)
                if bucket == num_frames - 1
                else min((bucket + 1) * bucket_size, len(results))
            )
            for r_idx in range(start_idx, end_idx):
                _accumulate_heatmap_result(acc, results[r_idx], heatmap_type)
        return acc

    # First pass: densest window sets the shared normalization ceiling.
    global_max = 0.0
    for frame_idx in range(num_frames):
        global_max = max(global_max, float(_accumulate_window(frame_idx).max()))
    if global_max == 0:
        return None

    # Second pass: render each window as a frame.
    frames: list[Image.Image] = []
    for frame_idx in range(num_frames):
        frames.append(
            _heatmap_frame_image(
                _accumulate_window(frame_idx), global_max, heatmap_type, width, height
            )
        )

    frames[0].save(
        output_path,
        save_all=True,
        append_images=frames[1:],
        duration=frame_duration_ms,
        loop=0,
    )
    return output_path


# ---------------------------------------------------------------------------
# Analysis tools (strategy registry)
# ---------------------------------------------------------------------------
#
# Each tool is a small class wrapping the corresponding module-level ``scan_*``
# function (preserved for tests that monkeypatch them). The two registry-level
# dispatch points are:
#   - :func:`check_frame_for_tool` (single-frame eval used by multitool)
#   - :meth:`ScreenspaceWorker._dispatch` (full-video scan, called by the worker)
# Both look up the tool by name in ``TOOLS`` and delegate to its methods.
