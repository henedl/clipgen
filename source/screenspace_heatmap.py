"""Screenspace heatmap generation (pure cv2/PIL leaf).

Template/flow/change/attention PNGs plus two animated-GIF views: cumulative, and
a rolling window (recent-only, fading). A GIF gives the browser no way to seek,
so build_gif_sprite_bytes() re-tiles one into a sprite sheet for hover-scrub.
"""

import io
import math
from pathlib import Path
from typing import TYPE_CHECKING, Any

import cv2
import numpy as np

import config
import utils

if TYPE_CHECKING:
    from PIL import Image


def _normalize_blur(accumulator: np.ndarray, max_val: float) -> np.ndarray:
    """Normalize by *max_val* and blur → uint8 intensity (JET palette indexes)."""
    normalized = (accumulator / max_val * 255).astype(np.uint8)
    return cv2.GaussianBlur(normalized, (15, 15), 0)


def _colorize_accumulator(accumulator: np.ndarray, max_val: float) -> np.ndarray:
    """Normalize by *max_val*, blur, and apply the JET colormap → BGR uint8."""
    return cv2.applyColorMap(_normalize_blur(accumulator, max_val), cv2.COLORMAP_JET)


# 256-entry JET palette as RGB bytes for PIL "P"-mode GIF frames, built once
# from the same cv2 colormap the PNG path applies (index i is exactly
# applyColorMap's color for gray value i).
_JET_PALETTE: bytes | None = None


def _jet_palette() -> bytes:
    global _JET_PALETTE
    if _JET_PALETTE is None:
        ramp = np.arange(256, dtype=np.uint8).reshape(1, 256)
        _JET_PALETTE = cv2.applyColorMap(ramp, cv2.COLORMAP_JET)[0, :, ::-1].tobytes()
    return _JET_PALETTE


def _write_png(output_path: str, image: np.ndarray) -> bool:
    """Encode *image* to PNG and write it, returning whether it landed.

    Deliberately not ``cv2.imwrite``: that returns ``False`` (no exception) on a
    missing parent directory, a permission error, a full disk, and on paths its
    own file layer cannot open — and every caller here used to discard that
    boolean and report success, so a heatmap that was never written still got
    its filename published to the manifest and the browser. Encoding in memory
    and writing through ``pathlib`` keeps path handling in Python, where a
    failure is a real ``OSError``.
    """
    try:
        ok, buf = cv2.imencode(".png", image)
        if not ok:
            utils.warning_print(f"Heatmap PNG encoding failed: {output_path}")
            return False
        Path(output_path).write_bytes(buf.tobytes())
    except (OSError, cv2.error) as exc:
        utils.warning_print(f"Could not write heatmap PNG {output_path}: {exc}")
        return False
    return True


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
    if not _write_png(output_path, heatmap):
        return None
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
    if not _write_png(output_path, heatmap):
        return None
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
    if not _write_png(output_path, heatmap):
        return None
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
    if not _write_png(output_path, heatmap):
        return None
    return output_path


def _accumulate_heatmap_result(
    accumulator: np.ndarray,
    result: dict[str, Any],
    heatmap_type: str,
    mask_out: np.ndarray | None = None,
) -> None:
    """Add a single result's contribution to a heatmap accumulator.

    *mask_out* (uint8, accumulator-shaped) additionally records which pixels
    the grid branch drew — the geometry, not the values, so a ``mag`` of 0
    still marks its pixels. The rolling-GIF bucket layers replay overwrites
    through it (see :func:`generate_rolling_heatmap_gif`).
    """
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
            if mask_out is not None:
                cv2.circle(mask_out, (cx, cy), radius, 1, -1)


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

    Frames are built in palette ("P") mode: the JET colormap maps the 256
    normalized intensity values onto exactly 256 colors, so the blurred
    intensity image *is* the palette index image. Handing PIL RGB frames
    instead made the GIF encoder re-derive a 256-color palette per frame
    (quantizing ~1M pixels each) — the dominant cost of heatmap GIF
    generation: a 24-frame 1280×720 attention GIF drops 1.59 s → 0.66 s
    (rolling 1.87 s → 0.98 s), file size roughly unchanged. Grid types now
    interpolate in intensity space rather than between mapped colors;
    decoded output differs from the old quantized frames by ≤ ~5% per
    channel, comparable to the quantizer's own error.
    """
    from PIL import Image

    idx = _normalize_blur(accumulator, global_max)
    if heatmap_type in _GRID_KEYS:
        idx = cv2.resize(idx, (width, height), interpolation=cv2.INTER_LINEAR)
    frame = Image.fromarray(idx, mode="P")
    frame.putpalette(_jet_palette())
    return frame


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


def _save_animation(
    frames: list["Image.Image"],
    output_path: str,
    frame_duration_ms: int,
) -> dict[str, Any]:
    """Write *frames* as an animated GIF and describe it for the frontend.

    Returns ``{"path", "frames", "w", "h"}``. Everything here is intrinsic to the
    animation, never derived from config, so a stored descriptor can never drift
    away from the sprite sheet :func:`build_gif_sprite_bytes` renders from the
    same GIF later.

    ``frames`` is re-read from the written file rather than taken from ``len()``:
    PIL collapses a run of identical frames into one, which happens whenever
    consecutive buckets add no new heat (``cv2.circle`` sets rather than
    accumulates, so repeat detections at one spot render identically). Trusting
    the input count would tell the frontend to scrub to cells the sheet doesn't
    have.
    """
    from PIL import Image

    frames[0].save(
        output_path,
        save_all=True,
        append_images=frames[1:],
        duration=frame_duration_ms,
        loop=0,
    )
    written = len(frames)
    try:
        with Image.open(output_path) as anim:
            written = int(getattr(anim, "n_frames", written))
    except (OSError, ValueError):
        pass
    return {
        "path": output_path,
        "frames": written,
        "w": frames[0].size[0],
        "h": frames[0].size[1],
    }


def sprite_grid(frame_count: int) -> tuple[int, int]:
    """Return the ``(cols, rows)`` sprite grid for *frame_count* frames."""
    cols = max(1, min(frame_count, int(config.SCREENSPACE_HEATMAP_SPRITE_COLS)))
    return cols, math.ceil(frame_count / cols)


def build_gif_sprite_bytes(gif_path: str, cols: int) -> bytes | None:
    """Tile an animated GIF's frames into a single sprite-sheet PNG.

    A GIF cannot be seeked in a browser, so the heatmap thumbs rest on a sprite
    cell and map hover position to a frame. Rendered on demand from the GIF that
    is already on disk (and cached in memory by the route) rather than written
    alongside it — sprites are a derived view everywhere else in clipgen too, and
    the output directory is the user's, not a scratch space.

    Frames are downscaled to ``SCREENSPACE_HEATMAP_SPRITE_FRAME_WIDTH``: a
    frame-native template heatmap is 1080p, and 24 of those in one sheet is a
    ~50-megapixel PNG for a thumbnail that renders ~200px wide. Aspect ratio is
    preserved, so the caller's ``w``/``h`` still describe the cell shape.
    """
    from PIL import Image, ImageSequence

    try:
        with Image.open(gif_path) as anim:
            frames = [f.convert("RGB") for f in ImageSequence.Iterator(anim)]
    except Exception as exc:
        # Deliberately broad: decoding a truncated or half-written GIF (a scan
        # killed mid-save) leaks whatever PIL's frame walk hits — IndexError and
        # struct.error as readily as OSError — and all of them mean the same
        # thing here: no sprite, fall back to plain playback rather than 500.
        utils.warning_print(f"Could not read heatmap GIF {gif_path}: {exc}")
        return None
    if not frames:
        return None

    src_w, src_h = frames[0].size
    target_w = max(1, min(int(config.SCREENSPACE_HEATMAP_SPRITE_FRAME_WIDTH), src_w))
    target_h = max(1, round(src_h * target_w / src_w))
    cols = max(1, min(len(frames), cols))
    rows = math.ceil(len(frames) / cols)

    sheet = Image.new("RGB", (cols * target_w, rows * target_h))
    for idx, frame in enumerate(frames):
        scaled = frame.resize((target_w, target_h), Image.Resampling.BILINEAR)
        sheet.paste(scaled, ((idx % cols) * target_w, (idx // cols) * target_h))

    buf = io.BytesIO()
    sheet.save(buf, format="PNG")
    return buf.getvalue()


def generate_heatmap_gif(
    results: list[dict[str, Any]],
    width: int,
    height: int,
    output_path: str,
    heatmap_type: str = "template",
    num_frames: int = 24,
    frame_duration_ms: int = 120,
) -> dict[str, Any] | None:
    """Generate an animated GIF showing heatmap accumulation over time.

    Divides *results* into *num_frames* temporal buckets, progressively
    accumulates heatmap data, and writes frames as an animated GIF. Returns the
    :func:`_save_animation` descriptor, or ``None`` when there is nothing to draw.
    """
    if not results:
        return None

    num_frames = min(num_frames, len(results))
    if num_frames < 2:
        return None

    acc_h, acc_w = height, width
    if heatmap_type in _GRID_KEYS:
        acc_h = acc_w = 256

    # Pass 1: find the shared ceiling. Accumulation is monotonic, so the final
    # state already carries the global max — no per-frame snapshots needed.
    accumulator = np.zeros((acc_h, acc_w), dtype=np.float32)
    for r in results:
        _accumulate_heatmap_result(accumulator, r, heatmap_type)
    global_max = float(accumulator.max())
    if global_max == 0:
        return None

    # Pass 2: replay bucket-by-bucket, colorizing inline against that max. Only
    # one float32 accumulator stays resident, so peak memory is ~one frame rather
    # than all `num_frames` snapshots (~190MB → ~8MB at 1080p).
    # `_heatmap_frame_image` reads without mutating, so accumulation continues.
    accumulator = np.zeros((acc_h, acc_w), dtype=np.float32)
    frames: list[Image.Image] = []
    for frame_idx in range(num_frames):
        start_idx, end_idx = _frame_bucket_bounds(frame_idx, len(results), num_frames)
        for r_idx in range(start_idx, end_idx):
            _accumulate_heatmap_result(accumulator, results[r_idx], heatmap_type)
        frames.append(
            _heatmap_frame_image(accumulator, global_max, heatmap_type, width, height)
        )

    return _save_animation(frames, output_path, frame_duration_ms)


def generate_rolling_heatmap_gif(
    results: list[dict[str, Any]],
    width: int,
    height: int,
    output_path: str,
    heatmap_type: str = "template",
    num_frames: int = 24,
    window_frames: int = 6,
    frame_duration_ms: int = 120,
) -> dict[str, Any] | None:
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

    if heatmap_type in _GRID_KEYS:
        # Grid draws *set* pixels (cv2.circle, last draw wins), so a window is
        # reproducible from per-bucket layers: overwrite each bucket's drawn
        # pixels in bucket order — bit-identical to replaying the results (the
        # mask records drawn geometry, values carry each bucket's own overlap
        # resolution). Building every bucket once instead of rebuilding each
        # window from raw results twice (max pass + colorize pass) cuts the
        # Python-level circle draws ~12x: 596 → 145 ms of accumulate on a
        # 600-result attention scan. Layers are 256×256, so all 24 together
        # are ~3 MB — nothing like the full-res frames the two-pass shape
        # exists to avoid holding.
        layers: list[tuple[np.ndarray, np.ndarray]] = []
        for bucket in range(num_frames):
            vals = np.zeros((acc_h, acc_w), dtype=np.float32)
            mask = np.zeros((acc_h, acc_w), dtype=np.uint8)
            start_idx, end_idx = _frame_bucket_bounds(bucket, len(results), num_frames)
            for r_idx in range(start_idx, end_idx):
                _accumulate_heatmap_result(
                    vals, results[r_idx], heatmap_type, mask_out=mask
                )
            layers.append((vals, mask.astype(bool)))

        def _accumulate_window(frame_idx: int) -> np.ndarray:
            acc = np.zeros((acc_h, acc_w), dtype=np.float32)
            for bucket in range(max(0, frame_idx - window_frames + 1), frame_idx + 1):
                vals, mask = layers[bucket]
                acc[mask] = vals[mask]
            return acc

    else:
        # Template accumulates additively (`+=`), where bucket-layer sums would
        # reorder float additions and drift off the replayed result — and its
        # accumulator is frame-native, so 24 resident layers would be the
        # memory problem the rebuild shape avoids. Keep rebuilding from results.
        def _accumulate_window(frame_idx: int) -> np.ndarray:
            acc = np.zeros((acc_h, acc_w), dtype=np.float32)
            win_start = max(0, frame_idx - window_frames + 1)
            for bucket in range(win_start, frame_idx + 1):
                start_idx, end_idx = _frame_bucket_bounds(
                    bucket, len(results), num_frames
                )
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

    return _save_animation(frames, output_path, frame_duration_ms)
