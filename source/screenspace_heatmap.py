"""Screenspace heatmap generation (pure cv2/PIL leaf).

Template-, flow-, change-, and attention-heatmap PNGs plus animated-GIF views:
a cumulative accumulation and a rolling-window (recent-only, fading) variant.
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


# Heatmap types that accumulate sparse normalized {x, y, mag} grid cells at a
# fixed 256×256 resolution (template accumulates match boxes frame-native).
_GRID_KEYS: dict[str, str] = {
    "flow": "flow_grid",
    "change": "change_grid",
    "attention": "saliency_grid",
}


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


def generate_attention_heatmap(
    results: list[dict[str, Any]],
    frame_width: int,
    frame_height: int,
    output_path: str,
) -> str | None:
    """Generate a heatmap PNG from accumulated per-frame saliency grids.

    Uses ``saliency_grid`` data (downsampled saliency maps, one per sampled
    frame) so heat reflects predicted attention dwell across the whole scan —
    the eye-tracking-style deliverable. Full-frame: sized to the video frame.
    """
    acc_size = 256
    accumulator = np.zeros((acc_size, acc_size), dtype=np.float32)
    for r in results:
        _accumulate_heatmap_result(accumulator, r, "attention")

    if accumulator.max() == 0:
        return None

    heatmap = _colorize_accumulator(accumulator, accumulator.max())
    heatmap = cv2.resize(
        heatmap, (frame_width, frame_height), interpolation=cv2.INTER_LINEAR
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
    elif heatmap_type in _GRID_KEYS:
        for cell in result.get(_GRID_KEYS[heatmap_type], []):
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

    Grid-based heatmaps (flow, change, attention) accumulate at a fixed
    resolution and are resized to the requested frame size; template
    accumulates frame-native.
    """
    from PIL import Image

    colored = _colorize_accumulator(accumulator, global_max)
    if heatmap_type in _GRID_KEYS:
        colored = cv2.resize(colored, (width, height), interpolation=cv2.INTER_LINEAR)
    rgb = cv2.cvtColor(colored, cv2.COLOR_BGR2RGB)
    return Image.fromarray(rgb)


def _frame_bucket_bounds(
    frame_idx: int, total: int, num_frames: int
) -> tuple[int, int]:
    """Return the ``[start, end)`` result indices for one GIF frame.

    Proportional split (``floor(i*total/n) .. floor((i+1)*total/n)``) so results
    spread evenly across frames — adjacent bucket sizes differ by at most one.
    Floor division of ``total // num_frames`` instead piles every remainder onto
    the final frame (e.g. 47 results / 24 frames → 23 frames of 1, last of 24).
    The last frame's end is exactly ``total``.
    """
    return (frame_idx * total) // num_frames, ((frame_idx + 1) * total) // num_frames


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
    if heatmap_type in _GRID_KEYS:
        acc_h = acc_w = 256

    # Pass 1: accumulate everything to find the shared ceiling. Accumulation is
    # monotonic, so the final cumulative state already carries the global max —
    # no per-frame snapshots needed to compute it.
    accumulator = np.zeros((acc_h, acc_w), dtype=np.float32)
    for r in results:
        _accumulate_heatmap_result(accumulator, r, heatmap_type)
    global_max = float(accumulator.max())
    if global_max == 0:
        return None

    # Pass 2: replay the accumulation bucket-by-bucket, colorizing each frame
    # inline against that shared max. Only one float32 accumulator is ever
    # resident, so peak memory is ~one frame instead of all `num_frames`
    # snapshots at once (~190MB → ~8MB at 1080p). `_heatmap_frame_image` reads
    # the accumulator without mutating it, so accumulation safely continues.
    accumulator = np.zeros((acc_h, acc_w), dtype=np.float32)
    frames: list[Image.Image] = []
    for frame_idx in range(num_frames):
        start_idx, end_idx = _frame_bucket_bounds(frame_idx, len(results), num_frames)
        for r_idx in range(start_idx, end_idx):
            _accumulate_heatmap_result(accumulator, results[r_idx], heatmap_type)
        frames.append(
            _heatmap_frame_image(accumulator, global_max, heatmap_type, width, height)
        )

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
    if heatmap_type in _GRID_KEYS:
        acc_h = acc_w = 256

    def _accumulate_window(frame_idx: int) -> np.ndarray:
        acc = np.zeros((acc_h, acc_w), dtype=np.float32)
        win_start = max(0, frame_idx - window_frames + 1)
        for bucket in range(win_start, frame_idx + 1):
            start_idx, end_idx = _frame_bucket_bounds(bucket, len(results), num_frames)
            for r_idx in range(start_idx, end_idx):
                _accumulate_heatmap_result(acc, results[r_idx], heatmap_type)
        return acc

    # Pass 1: build each window once to find the shared ceiling, discarding each
    # array immediately so only one window is ever resident.
    global_max = 0.0
    for i in range(num_frames):
        global_max = max(global_max, float(_accumulate_window(i).max()))
    if global_max == 0:
        return None

    # Pass 2: rebuild each window and colorize inline against that shared max —
    # peak memory is ~one window instead of all `num_frames` at once. Windows are
    # independent, so rebuilding them is cheap relative to holding them all.
    frames: list[Image.Image] = []
    for idx in range(num_frames):
        frames.append(
            _heatmap_frame_image(
                _accumulate_window(idx), global_max, heatmap_type, width, height
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
