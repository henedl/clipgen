"""Screenspace image-analysis primitives (pure cv2/numpy).

Region cropping/denormalization/resolution, HSV color math, frame-diff, SSIM,
perceptual hashing, template matching, optical flow, and scene fingerprinting,
plus the small scan-support helpers (morphology kernel cache, consecutive-match
buffer, static-frame skip). No file or ffmpeg I/O lives here.
"""

import functools
import math
import statistics
from collections.abc import Callable
from typing import TYPE_CHECKING, Any

import cv2

try:
    # cv2.data.haarcascades is a submodule attribute, so ty needs the explicit
    # import. Absent on cv2 builds without the bundled cascade data — the face
    # channel feature-detects it and degrades to a zeros map.
    import cv2.data
except ImportError:  # pragma: no cover - depends on the installed cv2 build
    pass
import numpy as np

if TYPE_CHECKING:
    import imagehash

import config


@functools.cache
def _morph_kernel(size: int) -> np.ndarray:
    """Return a shared square uint8 kernel for cv2 morphology ops.

    cv2.morphologyEx treats the kernel as read-only, so callers share one cached
    array instead of reallocating np.ones() per call/frame.
    """
    return np.ones((size, size), np.uint8)


ScanCallback = Callable[[float, np.ndarray], bool | None]
"""Per-frame callback signature for scan_video_frames and friends.

Receives ``(timestamp_seconds, region_pixels)`` and may return ``False`` to stop
iteration early. Used by all eight scan tools (color, change, similarity, text,
numbers, flow, scene, inactivity)."""


class _ConsecutiveBuffer:
    """Emit one event only after N consecutive matching sampled frames.

    ``push()`` records a matching frame; once ``size`` matches accumulate it
    returns a single event -- re-stamped with the median timestamp of the run and
    carrying the payload of the frame nearest that median -- and resets.
    ``reset()`` (called on any non-match) discards a partial run. ``size == 1``
    emits on every push, so the default reproduces pre-temporal-coherence
    behavior exactly.
    """

    def __init__(self, size: int) -> None:
        self.size = max(1, int(size))
        self._events: list[dict[str, Any]] = []
        self._timestamps: list[float] = []

    def push(self, ts: float, event: dict[str, Any]) -> dict[str, Any] | None:
        self._events.append(event)
        self._timestamps.append(ts)
        if len(self._events) >= self.size:
            median_ts = statistics.median(self._timestamps)
            # Pair the median timestamp with the nearest frame's payload: the exact
            # middle for odd-length runs, and for even lengths (median interpolated
            # between two frames) the nearer real frame rather than an arbitrary
            # upper-middle one.
            nearest = min(
                range(len(self._timestamps)),
                key=lambda i: abs(self._timestamps[i] - median_ts),
            )
            emitted = dict(self._events[nearest])
            emitted["timestamp"] = median_ts
            self.reset()
            return emitted
        return None

    def carry(self, ts: float) -> dict[str, Any] | None:
        """Extend an active run when a frame is skipped as static.

        A static frame is near-identical to the last processed one, so a match
        that was on screen still is — the run continues rather than breaking.
        Re-pushes the most recent matched event under *ts*, emitting if that
        completes the run. No-op when no run is active, which leaves ``size == 1``
        (where ``push`` emits and resets every frame) on its original path.
        """
        if not self._events:
            return None
        return self.push(ts, self._events[-1])

    def reset(self) -> None:
        self._events = []
        self._timestamps = []


def _is_static_skip(
    ts: float,
    pixels: np.ndarray,
    prev_gray: list[np.ndarray | None],
    buf: _ConsecutiveBuffer,
    results: list[dict[str, Any]],
    on_result: Callable[[dict[str, Any]], None] | None,
    on_progress: Callable[[float], None] | None,
    start_seconds: float,
    total_range: float,
) -> bool:
    """Decide whether *pixels* is a near-duplicate of the previous frame.

    True when the mean grayscale diff from the last processed frame is below
    ``SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD``; the caller then returns from its
    per-frame callback. Content and any active match are unchanged, so this
    carries the consecutive-match run via ``buf.carry`` (emitting through
    *results*/*on_result*) and reports progress instead of breaking it. Otherwise
    records *pixels* as the new baseline and returns False. Shared by
    scan_text/scan_numbers.
    """
    gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
    if prev_gray[0] is not None:
        diff = mean_gray_diff(prev_gray[0], gray)
        if diff < config.SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD:
            emitted = buf.carry(ts)
            if emitted is not None:
                results.append(emitted)
                if on_result:
                    on_result(emitted)
            if on_progress and total_range > 0:
                on_progress((ts - start_seconds) / total_range)
            return True
    prev_gray[0] = gray
    return False


def _frame_is_static(prev_gray: np.ndarray | None, curr_gray: np.ndarray) -> bool:
    """True when *curr_gray* is a near-duplicate of the last processed gray frame.

    "Static" means a mean grayscale diff below
    ``SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD``; False when *prev_gray* is
    ``None`` (no baseline yet). Unlike :func:`_is_static_skip` this is a bare
    predicate with no buffer/progress side effects, so each caller applies the
    skip semantics its scan needs (reset a motion run, extend a span, carry the
    last result). Shared by scan_similarity, scan_flow, scan_inactivity,
    scan_boundaries, scan_scene, and scan_template.
    """
    if prev_gray is None:
        return False
    return (
        mean_gray_diff(prev_gray, curr_gray)
        < config.SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD
    )


def mean_gray_diff(a: np.ndarray, b: np.ndarray) -> float:
    """Mean absolute difference between two same-shaped uint8 arrays.

    Equals ``float(np.mean(cv2.absdiff(a, b)))`` exactly (both accumulate the
    integer L1 sum in double), but ``cv2.norm`` fuses absdiff+sum in one SIMD
    pass with no full-size temporaries — and this line runs on every tool's
    per-frame path via the static-frame skip above.
    """
    return cv2.norm(a, b, cv2.NORM_L1) / a.size


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


def region_mask_for(region: dict[str, Any], h: int, w: int) -> np.ndarray | None:
    """Rasterize a shaped region's contours to a uint8 0/255 mask of shape (h, w).

    ``region["mask_points"]`` holds a list of polygon contours whose vertices
    are bbox-relative (0-1 fractions of the region's own bounding rect), so the
    mask can be rasterized directly at whatever size the cropped pixels actually
    arrive at — after ffmpeg's ``cv_scale``/``max_dim`` rescales or a tool's
    internal downscale — with no resize chain. Contours are filled one at a time
    (union semantics — a multi-part region masks every part; overlaps never
    XOR out). Returns ``None`` for rectangular regions (no ``mask_points``),
    letting callers keep their unmasked fast path.
    """
    contours = region.get("mask_points")
    if not contours or h <= 0 or w <= 0:
        return None
    mask = np.zeros((h, w), dtype=np.uint8)
    for contour in contours:
        if len(contour) < 3:
            continue
        poly = np.array([[u * w, v * h] for u, v in contour], dtype=np.float64)
        cv2.fillPoly(mask, [np.round(poly).astype(np.int32)], 255)
    return mask


def mask_points_key(contours: Any) -> tuple[Any, ...]:
    """Hashable key for a region's ``mask_points`` contour list (or None/[]).

    Cache keys (per-frame memos, pin-OCR readings) must distinguish same-bbox/
    different-shape regions; nested lists aren't hashable, so every keyed cache
    routes through this helper instead of hand-rolled ``tuple(map(tuple, …))``.
    """
    return tuple(tuple(tuple(p) for p in contour) for contour in contours or ())


def filter_matches_by_region_mask(
    matches: list[dict[str, Any]], region: dict[str, Any]
) -> list[dict[str, Any]]:
    """Drop template matches whose center falls outside a shaped region.

    *matches* carry frame-pixel ``x/y/w/h`` boxes; *region* is the pixel
    ``region_coords`` dict. Rect regions (no ``mask_points``) and degenerate
    bboxes pass through unchanged — template search is full-frame for rects,
    so only the polygon adds a restriction.
    """
    points = region.get("mask_points")
    rw, rh = region.get("w", 0), region.get("h", 0)
    if not points or rw <= 0 or rh <= 0:
        return matches
    rx, ry = region.get("x", 0), region.get("y", 0)
    return [
        m
        for m in matches
        if point_in_mask_points(
            (m["x"] + m["w"] / 2.0 - rx) / rw,
            (m["y"] + m["h"] / 2.0 - ry) / rh,
            points,
        )
    ]


def region_masker(
    region: dict[str, Any],
) -> Callable[[np.ndarray], np.ndarray | None]:
    """Return ``fn(pixels) -> mask|None`` caching the rasterized mask per shape.

    Scans receive crops at a constant (but not statically known) size — after
    ffmpeg's ``cv_scale``/``max_dim`` rescales — so the mask is rasterized on
    first use and reused for every subsequent frame. Rect regions cost one dict
    lookup per frame (the cached value is ``None``).
    """
    cache: dict[tuple[int, int], np.ndarray | None] = {}

    def _mask_for(pixels: np.ndarray) -> np.ndarray | None:
        key = pixels.shape[:2]
        if key not in cache:
            cache[key] = region_mask_for(region, key[0], key[1])
        return cache[key]

    return _mask_for


def point_in_mask_points(u: float, v: float, contours: list[Any] | None) -> bool:
    """Test a point against a shaped region's contour list in bbox-relative space.

    Scale-free companion to :func:`region_mask_for` for consumers that test
    individual centers (OCR readings, template match boxes) rather than
    rasterizing a full mask. True when the point falls inside ANY contour
    (union semantics, matching the per-contour fill in ``region_mask_for``).
    """
    return any(_point_in_polygon(u, v, contour) for contour in contours or ())


def _point_in_polygon(u: float, v: float, points: list[Any]) -> bool:
    """Ray-cast point-in-polygon test for one implicitly-closed contour."""
    n = len(points)
    if n < 3:
        return False
    inside = False
    j = n - 1
    for i in range(n):
        ui, vi = float(points[i][0]), float(points[i][1])
        uj, vj = float(points[j][0]), float(points[j][1])
        if (vi > v) != (vj > v):
            cross_u = (uj - ui) * (v - vi) / (vj - vi) + ui
            if u < cross_u:
                inside = not inside
        j = i
    return inside


def denormalize_region(
    region: dict[str, Any], target_w: int, target_h: int
) -> dict[str, Any]:
    """Convert a normalized region (0–1 floats) to integer pixel coordinates.

    *region* carries normalized ``x``/``y``/``w``/``h`` plus, for shaped regions,
    bbox-relative ``points`` + ``shape``; those pass through to the result as
    ``mask_points``/``shape``.
    """
    out: dict[str, Any] = {
        "x": round(region["x"] * target_w),
        "y": round(region["y"] * target_h),
        "w": round(region["w"] * target_w),
        "h": round(region["h"] * target_h),
    }
    # Contour vertices are bbox-relative (0-1 of the region's own rect), so they
    # survive denormalization verbatim. Copying here threads the mask into every
    # region_coords consumer (task snapshots, multitool steps, workflows,
    # calibration) with no further plumbing.
    points = region.get("points")
    if points:
        out["mask_points"] = points
        if region.get("shape"):
            out["shape"] = region["shape"]
    return out


FULL_FRAME_REGION_NAME = "full_frame"
FULL_FRAME_REGION: dict[str, Any] = {
    "x": 0.0,
    "y": 0.0,
    "w": 1.0,
    "h": 1.0,
    "source_width": 0,
    "source_height": 0,
}


def resolve_region_request(
    region_name: str,
    region_ref: Any,
    manifest: dict[str, Any],
) -> tuple[str, dict[str, Any]]:
    """Resolve a (region_name, region_ref) request to (resolved_name, normalized_region).

    Mirrors the active/stash/full_frame precedence the Screenspace server and CLI both
    rely on: a ``region_ref`` (``{source: "active"|"stash"|"full_frame", name?, stash_id?}``)
    pins the lookup to a specific source, while a bare ``region_name`` falls back to
    active-regions-first then stashes. Pure — reads only the passed ``manifest`` and never
    flattens stashes, so duplicate region names across stashes stay distinguishable.

    Returns the resolved name and the normalized (0–1) region dict. Raises ``ValueError``
    with a human-readable message on any failure (caller maps it to a 400 or a CLI error).
    """
    active_regions = manifest.get("regions", {})
    if not isinstance(active_regions, dict):
        active_regions = {}

    if region_ref is None:
        if region_name == FULL_FRAME_REGION_NAME:
            return FULL_FRAME_REGION_NAME, dict(FULL_FRAME_REGION)
        if region_name in active_regions:
            return region_name, active_regions[region_name]
        for stash in manifest.get("stashes", []):
            stash_regions = stash.get("regions", {})
            if isinstance(stash_regions, dict) and region_name in stash_regions:
                return region_name, stash_regions[region_name]
        raise ValueError(f"Region '{region_name}' not found")

    if not isinstance(region_ref, dict):
        # ValueError, not TypeError: this resolver's contract is "ValueError means
        # bad request", and every caller catches it to render a hint or a 400.
        raise ValueError("region_ref must be an object")  # noqa: TRY004

    source = str(region_ref.get("source", "")).strip()

    if source == "full_frame":
        return FULL_FRAME_REGION_NAME, dict(FULL_FRAME_REGION)

    name = str(region_ref.get("name", "")).strip()
    if not name:
        raise ValueError("region_ref.name is required")

    if source == "active":
        if name not in active_regions:
            raise ValueError(f"Region '{name}' not found")
        return name, active_regions[name]

    if source == "stash":
        stash_id = str(region_ref.get("stash_id", "")).strip()
        if not stash_id:
            raise ValueError("region_ref.stash_id is required")
        for stash in manifest.get("stashes", []):
            if stash.get("id") != stash_id:
                continue
            stash_regions = stash.get("regions", {})
            if not isinstance(stash_regions, dict) or name not in stash_regions:
                raise ValueError(f"Region '{name}' not found in stash '{stash_id}'")
            return name, stash_regions[name]
        raise ValueError(f"Stash '{stash_id}' not found")

    raise ValueError("region_ref.source must be 'active', 'stash', or 'full_frame'")


def average_color_hsv(
    region_pixels: np.ndarray, mask: np.ndarray | None = None
) -> dict[str, float]:
    """Compute mean HSV color of a region.

    Args:
        region_pixels: BGR image region as numpy array.
        mask: Optional uint8 shaped-region mask (same shape as the crop);
            when given, the mean is computed over mask pixels only.

    Returns:
        Dict with keys ``h`` (0-180), ``s`` (0-255), ``v`` (0-255).
    """
    if region_pixels.size == 0:
        # A region cropped fully off-frame: np.mean on the empty crop would
        # yield NaN (same guard as color_present).
        return {"h": 0.0, "s": 0.0, "v": 0.0}
    h, w = region_pixels.shape[:2]
    if h > 64 or w > 64:
        new_w, new_h = min(w, 64), min(h, 64)
        region_pixels = cv2.resize(
            region_pixels, (new_w, new_h), interpolation=cv2.INTER_AREA
        )
        if mask is not None:
            mask = cv2.resize(mask, (new_w, new_h), interpolation=cv2.INTER_NEAREST)
    hsv = cv2.cvtColor(region_pixels, cv2.COLOR_BGR2HSV)
    if mask is not None and np.any(mask):
        mean = hsv[mask > 0].mean(axis=0)
    else:
        mean = np.mean(hsv, axis=(0, 1))
    return {"h": float(mean[0]), "s": float(mean[1]), "v": float(mean[2])}


def color_matches(
    region_pixels: np.ndarray,
    target_color: dict[str, float],
    tolerance: dict[str, float],
    mask: np.ndarray | None = None,
) -> tuple[bool, float]:
    """Check if region's average HSV color is within tolerance of target.

    Handles hue wraparound (red at 0/180 boundary).

    Returns:
        Tuple of (matches, confidence) where confidence is 0.0–1.0.
    """
    avg = average_color_hsv(region_pixels, mask=mask)
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


def _channel_runs(flags: np.ndarray) -> tuple[tuple[int, int], ...]:
    """Contiguous ``(lo, hi)`` index runs of a 256-entry boolean table."""
    idx = np.flatnonzero(flags)
    if idx.size == 0:
        return ()
    breaks = np.flatnonzero(np.diff(idx) > 1)
    starts = np.concatenate(([idx[0]], idx[breaks + 1]))
    ends = np.concatenate((idx[breaks], [idx[-1]]))
    return tuple(zip(starts.tolist(), ends.tolist()))


@functools.cache
def _hsv_match_bands(
    h: float, s: float, v: float, tol_h: float, tol_s: float, tol_v: float
) -> tuple[tuple[tuple[int, int, int], tuple[int, int, int]], ...]:
    """``cv2.inRange`` (lower, upper) HSV bands equivalent to the per-pixel test.

    Each of the three per-channel predicates depends only on that channel's
    ``uint8`` value, so evaluating the *same float32 expressions* over all 256
    possible values yields exact membership tables — the bands below are those
    tables, not an approximation of them. Hue wraparound splits into two bands;
    everything else is one. Empty result means no pixel value can match.

    Cached per (target, tolerance): constant for a whole scan, and bounded by
    the number of distinct color configurations a user has set up.
    """
    vals = np.arange(256, dtype=np.float32)
    hue_diff = np.abs(vals - h)
    hue_ok = np.minimum(hue_diff, 180.0 - hue_diff) <= tol_h
    sat_ok = np.abs(vals - s) <= tol_s
    val_ok = np.abs(vals - v) <= tol_v
    return tuple(
        ((h_lo, s_lo, v_lo), (h_hi, s_hi, v_hi))
        for h_lo, h_hi in _channel_runs(hue_ok)
        for s_lo, s_hi in _channel_runs(sat_ok)
        for v_lo, v_hi in _channel_runs(val_ok)
    )


def color_present(
    region_pixels: np.ndarray,
    target_color: dict[str, float],
    tolerance: dict[str, float],
    min_coverage: float = 0.0,
    mask: np.ndarray | None = None,
) -> tuple[bool, float]:
    """Check whether the target color appears *anywhere* in the region.

    Where :func:`color_matches` averages the whole region, this builds a per-pixel
    HSV match mask and fires once the matching fraction reaches ``min_coverage``
    — so a small patch of target color in an otherwise neutral region is caught
    here but averaged away there. Hue wraparound is handled identically.

    Scanned at full resolution (no ``INTER_AREA`` downscale) so small patches
    survive; fast-scan callers must skip ``max_region_dim`` on this path.

    Returns:
        ``(matches, coverage)``, coverage being the 0.0–1.0 matching fraction.
        ``min_coverage <= 0`` fires on a single pixel.
    """
    if region_pixels.size == 0:
        return False, 0.0
    hsv = cv2.cvtColor(region_pixels, cv2.COLOR_BGR2HSV)
    # Integer band tests on the uint8 HSV rather than three full-size float32
    # casts and comparisons: identical mask (see _hsv_match_bands), ~12x faster,
    # and no ~25 MB of per-frame temporaries at 1080p — this path deliberately
    # runs at full resolution, so it is the hottest per-pixel routine here.
    bands = _hsv_match_bands(
        float(target_color["h"]),
        float(target_color["s"]),
        float(target_color["v"]),
        float(tolerance["h"]),
        float(tolerance["s"]),
        float(tolerance["v"]),
    )
    hits: np.ndarray | None = None
    for lower, upper in bands:
        band = cv2.inRange(hsv, np.array(lower, np.uint8), np.array(upper, np.uint8))
        hits = band if hits is None else cv2.bitwise_or(hits, band)
    match = (
        hits.astype(bool) if hits is not None else np.zeros(hsv.shape[:2], dtype=bool)
    )
    # Shaped regions: only pixels inside the polygon count, in both the match
    # numerator and the coverage denominator. A mask emptied by extreme downscale
    # falls back to the full rect, as average_color_hsv does.
    if mask is not None and not np.any(mask):
        mask = None
    denom = match.size
    if mask is not None:
        match &= mask > 0
        denom = int(np.count_nonzero(mask))
    count = int(np.count_nonzero(match))
    coverage = count / denom if denom else 0.0
    matched = count > 0 if min_coverage <= 0 else coverage >= min_coverage
    return matched, float(coverage)


def blur_gray(region: np.ndarray) -> np.ndarray:
    """The frame-diff front-end: Gaussian blur then grayscale one BGR region.

    Split out of the diff so per-frame callers (``scan_changes``,
    ``ChangeTool``) can compute it once per frame and carry it forward —
    frame N is otherwise blurred twice, as ``region_b`` at step N and again
    as ``region_a`` at step N+1, and this 3-channel blur is the most
    expensive op on the Change hot path.
    """
    k = config.SCREENSPACE_BLUR_KERNEL
    return cv2.cvtColor(cv2.GaussianBlur(region, (k, k), 0), cv2.COLOR_BGR2GRAY)


def _frame_diff_mask(
    region_a: np.ndarray,
    region_b: np.ndarray,
    noise_threshold: int = 0,
) -> np.ndarray:
    """Blur, grayscale, absdiff, threshold and morph-open two same-sized BGR regions.

    Returns the binary change mask; callers derive a change ratio from it or
    downsample it to a grid (the Change heatmap). Single source of truth for the
    frame-diff computation shared by ``compute_frame_diff``, ``ChangeTool`` and
    ``scan_changes`` — the per-frame callers go through the ``_gray`` variants
    with a carried-forward ``blur_gray`` result instead of this pairwise form.
    """
    return _frame_diff_mask_gray(
        blur_gray(region_a), blur_gray(region_b), noise_threshold
    )


def _frame_diff_mask_gray(
    a_gray: np.ndarray,
    b_gray: np.ndarray,
    noise_threshold: int = 0,
) -> np.ndarray:
    """Diff back-end: absdiff, threshold and morph-open two ``blur_gray`` outputs."""
    if noise_threshold <= 0:
        noise_threshold = config.SCREENSPACE_NOISE_THRESHOLD
    diff = cv2.absdiff(a_gray, b_gray)
    _, mask = cv2.threshold(diff, noise_threshold, 255, cv2.THRESH_BINARY)
    kernel = _morph_kernel(config.SCREENSPACE_MORPH_KERNEL)
    return cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)


def compute_frame_diff(
    region_a: np.ndarray,
    region_b: np.ndarray,
    noise_threshold: int = 0,
    mask: np.ndarray | None = None,
) -> float:
    """Compute pixel difference ratio between two same-sized regions.

    Applies Gaussian blur, thresholds noise, and morphological opening.
    For shaped regions, *mask* (uint8, crop-sized) restricts both the counted
    changes and the denominator to polygon pixels.

    Returns:
        Change ratio 0.0-1.0 (fraction of pixels that changed).
    """
    return compute_frame_diff_gray(
        blur_gray(region_a), blur_gray(region_b), noise_threshold, mask
    )


def compute_frame_diff_gray(
    a_gray: np.ndarray,
    b_gray: np.ndarray,
    noise_threshold: int = 0,
    mask: np.ndarray | None = None,
) -> float:
    """``compute_frame_diff`` on precomputed ``blur_gray`` outputs (hot paths)."""
    diff = _frame_diff_mask_gray(a_gray, b_gray, noise_threshold)
    if diff.size == 0:
        return 0.0
    if mask is not None and np.any(mask):
        diff = cv2.bitwise_and(diff, mask)
        return float(np.count_nonzero(diff)) / float(np.count_nonzero(mask))
    return float(np.count_nonzero(diff)) / float(diff.size)


def _ssim_preprocess(
    region_a: np.ndarray, region_b: np.ndarray
) -> tuple[np.ndarray, np.ndarray]:
    """Shared SSIM front-end: resize the pair to <=256 px, blur, grayscale.

    Single source of truth for the preprocessing used by both
    ``regions_are_similar`` (scalar SSIM on the scan hot path) and
    ``ssim_diff_map`` (per-pixel map for the Model view preview), so the preview
    mirrors exactly what the scan scores.
    """
    max_dim = 256
    h, w = region_a.shape[:2]
    if h > max_dim or w > max_dim:
        scale = max_dim / max(h, w)
        new_w, new_h = int(w * scale), int(h * scale)
        region_a = cv2.resize(region_a, (new_w, new_h), interpolation=cv2.INTER_AREA)
        region_b = cv2.resize(region_b, (new_w, new_h), interpolation=cv2.INTER_AREA)
    k = config.SCREENSPACE_BLUR_KERNEL
    a_gray = cv2.cvtColor(cv2.GaussianBlur(region_a, (k, k), 0), cv2.COLOR_BGR2GRAY)
    b_gray = cv2.cvtColor(cv2.GaussianBlur(region_b, (k, k), 0), cv2.COLOR_BGR2GRAY)
    return a_gray, b_gray


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
    a_gray, b_gray = _ssim_preprocess(region_a, region_b)
    from skimage.metrics import structural_similarity as ssim

    score = float(ssim(a_gray, b_gray))
    return score >= threshold, score


def ssim_diff_map(
    region_a: np.ndarray, region_b: np.ndarray
) -> tuple[float, np.ndarray]:
    """SSIM score plus the per-pixel structural-similarity map.

    Runs ``structural_similarity(..., full=True)`` over the same <=256/blur/gray
    preprocessing as ``regions_are_similar``. Preview-only (off the scan hot
    path), so the extra full-map cost never touches per-frame scanning. Returns
    ``(score, ssim_map)`` where ``ssim_map`` is float in roughly [-1, 1] at the
    preprocessed (<=256 px) resolution (higher = more similar).
    """
    a_gray, b_gray = _ssim_preprocess(region_a, region_b)
    from skimage.metrics import structural_similarity as ssim

    score, smap = ssim(a_gray, b_gray, full=True)
    return float(score), np.asarray(smap, dtype=np.float32)


def compute_phash(
    region_pixels: np.ndarray, gray: np.ndarray | None = None
) -> "imagehash.ImageHash":
    """Compute perceptual hash of a region for fast similarity scanning.

    Mirrors ``imagehash.phash`` (grayscale → 32×32 → 2D DCT → top-left 8×8 →
    median threshold) natively in cv2, skipping the per-frame BGR→RGB +
    ``PIL.Image`` round-trip that dominates this hot scan-callback path.
    Callers that already hold the unblurred grayscale (every scan whose
    static-skip check converts the frame) pass it as *gray* to skip the
    second conversion.
    """
    import imagehash

    if gray is None:
        gray = cv2.cvtColor(region_pixels, cv2.COLOR_BGR2GRAY)
    small = cv2.resize(gray, (32, 32), interpolation=cv2.INTER_AREA).astype(np.float32)
    dct = cv2.dct(small)
    dctlowfreq = dct[:8, :8]
    diff = dctlowfreq > np.median(dctlowfreq)
    return imagehash.ImageHash(diff)


# Prepared template payload shared across frames in a scan: grayscale blurred
# template, grayscale blurred mask (or None), and a "degenerate" flag set when
# the template has near-zero variance (TM_CCOEFF_NORMED is undefined there).
_PreparedTemplate = tuple[np.ndarray, "np.ndarray | None", bool]


def _scale_template(
    template: np.ndarray,
    mask: np.ndarray | None,
    template_scale: float,
) -> tuple[np.ndarray, np.ndarray | None]:
    """Resize a template (and its mask) for matching.

    Combines the user-supplied *template_scale* with the global CV resolution
    scale so the template matches at the same relative size on the (possibly
    upscaled) extracted frames. The opt-in template_scale slider lets users
    compensate when their uploaded PNG is captured at a different pixel scale
    than its in-video rendering (e.g. a 50x50 icon appearing as 24x24 on
    screen). Returns the pair unchanged when the effective scale is ~1.0.
    """
    cv_scale = (
        config.SCREENSPACE_CV_RESOLUTION_SCALE
        if config.SCREENSPACE_CV_RESOLUTION_SCALE > 0
        else 1.0
    )
    effective = template_scale * cv_scale
    if not (effective > 0 and abs(effective - 1.0) > 1e-6):
        return template, mask
    th, tw = template.shape[:2]
    nw = max(8, round(tw * effective))
    nh = max(8, round(th * effective))
    interp = cv2.INTER_AREA if effective < 1.0 else cv2.INTER_CUBIC
    scaled_template = cv2.resize(template, (nw, nh), interpolation=interp)
    scaled_mask = (
        cv2.resize(mask, (nw, nh), interpolation=cv2.INTER_NEAREST)
        if mask is not None
        else None
    )
    return scaled_template, scaled_mask


def _prepare_template(
    template: np.ndarray, mask: np.ndarray | None
) -> _PreparedTemplate:
    """Compute the per-scan-constant grayscale template and mask once.

    Hoisted out of :func:`match_template` so callers that run the same
    template against many frames (scan_template, evaluate_region) can pay
    the blur+cvtColor cost a single time instead of per-frame.
    """
    k = config.SCREENSPACE_BLUR_KERNEL
    tmpl_gray = cv2.cvtColor(cv2.GaussianBlur(template, (k, k), 0), cv2.COLOR_BGR2GRAY)
    # Binarize the alpha mask (>= 128 -> 255) rather than blurring it: soft masks
    # let semi-transparent edge pixels contribute partially to cv2.matchTemplate,
    # inflating TM_CCOEFF_NORMED for mostly-transparent PNG icons.
    if mask is not None:
        _, gray_mask = cv2.threshold(mask, 127, 255, cv2.THRESH_BINARY)
    else:
        gray_mask = None
    # Degenerate if the pixels that actually contribute to matching have no
    # variance: TM_CCOEFF_NORMED normalizes by the masked template std, so at ~0
    # (a mostly-transparent PNG with a flat opaque patch, especially scaled down)
    # the denominator underflows and every position scores ~1.0.
    if gray_mask is not None:
        masked = tmpl_gray[gray_mask > 0]
        contributing_std = float(masked.std()) if masked.size else 0.0
    else:
        contributing_std = float(np.std(tmpl_gray))
    degenerate = contributing_std < 1.0
    return (tmpl_gray, gray_mask, degenerate)


def _template_correlation_map(
    frame: np.ndarray, prepared: _PreparedTemplate
) -> np.ndarray | None:
    """Compute the finite TM_CCOEFF_NORMED correlation map for a prepared template.

    Returns ``None`` when matching is undefined — a degenerate (constant)
    template, or one larger than the frame. inf/nan cells (near-zero variance
    patches) are neutralized to ``-1.0`` so callers can safely threshold the map
    or read its peak (the threshold-independent scalar used by calibration).
    """
    tmpl_gray, gray_mask, degenerate = prepared
    if degenerate:
        # Zero-variance template: see _prepare_template's degeneracy note.
        return None
    k = config.SCREENSPACE_BLUR_KERNEL
    frame_gray = cv2.cvtColor(cv2.GaussianBlur(frame, (k, k), 0), cv2.COLOR_BGR2GRAY)
    th, tw = tmpl_gray.shape[:2]
    if th > frame_gray.shape[0] or tw > frame_gray.shape[1]:
        return None
    result = cv2.matchTemplate(
        frame_gray, tmpl_gray, cv2.TM_CCOEFF_NORMED, mask=gray_mask
    )
    if not np.all(np.isfinite(result)):
        result = np.where(np.isfinite(result), result, -1.0)
    return result


def _match_template_prepared(
    frame: np.ndarray,
    prepared: _PreparedTemplate,
    threshold: float,
    nms_overlap: float,
    corr: np.ndarray | None = None,
) -> list[dict[str, Any]]:
    """Match a frame against an already-prepared template payload.

    Callers that already computed this frame's correlation map (to read its
    threshold-independent peak) pass it as *corr* — the map is the single most
    expensive op in the tool and recomputing it here doubled the cost of every
    passing frame.
    """
    result = _template_correlation_map(frame, prepared) if corr is None else corr
    if result is None:
        return []
    tmpl_gray, _gray_mask, _degenerate = prepared
    th, tw = tmpl_gray.shape[:2]
    locs = np.where(result >= threshold)
    if len(locs[0]) == 0:
        return []

    # Pathological matchTemplate output (low-variance masked templates at certain
    # scales) can yield tens of thousands of above-threshold candidates, and the
    # O(n^2) NMS below would freeze the worker. Cap by raw score.
    _MAX_CANDIDATES = 5000
    scores = result[locs]
    if len(locs[0]) > _MAX_CANDIDATES:
        top_idx = np.argpartition(scores, -_MAX_CANDIDATES)[-_MAX_CANDIDATES:]
        ys, xs = locs[0][top_idx], locs[1][top_idx]
        scores = scores[top_idx]
    else:
        ys, xs = locs[0], locs[1]

    detections: list[dict[str, Any]] = []
    for pt_y, pt_x, raw in zip(ys, xs, scores):
        score = float(raw)
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


def match_template(
    frame: np.ndarray,
    template: np.ndarray,
    threshold: float = 0.0,
    nms_overlap: float = 0.0,
    mask: np.ndarray | None = None,
    *,
    prepared: _PreparedTemplate | None = None,
    corr: np.ndarray | None = None,
) -> list[dict[str, Any]]:
    """Find all locations where template appears in frame.

    ``cv2.matchTemplate`` with ``TM_CCOEFF_NORMED``, then non-maximum suppression
    to drop overlapping detections. An optional *mask* (template-sized,
    single-channel) restricts matching to non-transparent regions — for uploaded
    PNGs with alpha.

    Across many frames with one template, build *prepared* once via
    :func:`_prepare_template` to skip the per-call blur and grayscale conversion.
    A caller that already holds this frame's correlation map (from
    :func:`_template_correlation_map`) passes it as *corr* to skip recomputing it.

    Returns:
        ``{x, y, w, h, score}`` dicts for each match above *threshold*.
    """
    if threshold <= 0.0:
        threshold = config.SCREENSPACE_TEMPLATE_MATCH_THRESHOLD
    if nms_overlap <= 0.0:
        nms_overlap = config.SCREENSPACE_TEMPLATE_NMS_OVERLAP

    if prepared is None:
        prepared = _prepare_template(template, mask)
    return _match_template_prepared(frame, prepared, threshold, nms_overlap, corr)


def flow_downscale(
    gray: np.ndarray, mask: np.ndarray | None = None
) -> tuple[np.ndarray, np.ndarray | None]:
    """The <=256px downscale ``compute_optical_flow`` applies to its inputs.

    Exposed so per-frame callers can downscale each frame once and carry the
    result forward as the next pair's "previous" side — otherwise every frame
    is INTER_AREA-resized twice (as curr at step N, as prev at step N+1).
    No-op (returns the inputs) when *gray* already fits.
    """
    max_dim = 256
    h, w = gray.shape[:2]
    if h <= max_dim and w <= max_dim:
        return gray, mask
    scale = max_dim / max(h, w)
    new_w, new_h = int(w * scale), int(h * scale)
    small = cv2.resize(gray, (new_w, new_h), interpolation=cv2.INTER_AREA)
    if mask is not None:
        mask = cv2.resize(mask, (new_w, new_h), interpolation=cv2.INTER_NEAREST)
    return small, mask


def compute_optical_flow(
    prev_gray: np.ndarray,
    curr_gray: np.ndarray,
    pyr_scale: float = 0.0,
    return_grid: bool = False,
    mask: np.ndarray | None = None,
) -> dict[str, Any]:
    """Compute dense optical flow between two grayscale frames.

    Farneback takes no mask, so shaped regions compute flow over the full rect and
    *mask* (uint8, crop-sized) restricts the statistics instead. Vectors within
    ~one window of the polygon edge still see outside pixels — acceptable
    contamination for motion detection.

    Returns:
        ``magnitude`` (mean vector length), ``angle`` (dominant direction, 0-360),
        and optionally ``flow_grid`` (sparse vectors for visualization).
    """
    if pyr_scale <= 0.0:
        pyr_scale = config.SCREENSPACE_FLOW_PYR_SCALE

    # Resize to max 256px for speed. Per-frame callers (scan_flow) pre-downscale
    # via flow_downscale and carry the previous frame's result forward, in which
    # case the inputs arrive <=256px and these are no-ops.
    prev_gray, _ = flow_downscale(prev_gray)
    curr_gray, mask = flow_downscale(curr_gray, mask)
    if mask is not None and not np.any(mask):
        mask = None

    flow_out = np.zeros((*prev_gray.shape[:2], 2), dtype=np.float32)
    flow = cv2.calcOpticalFlowFarneback(
        prev_gray, curr_gray, flow_out, pyr_scale, 3, 15, 3, 5, 1.2, 0
    )
    mag, ang = cv2.cartToPolar(flow[..., 0], flow[..., 1], angleInDegrees=True)
    inside = mask > 0 if mask is not None else None
    mean_mag = (
        float(np.mean(mag[inside])) if inside is not None else float(np.mean(mag))
    )

    # Dominant angle: weighted mean by magnitude
    if mean_mag > 0:
        # Use circular mean to avoid wraparound issues
        rad = np.deg2rad(ang)
        weights = mag if inside is None else mag * inside
        sin_sum = float(np.sum(weights * np.sin(rad)))
        cos_sum = float(np.sum(weights * np.cos(rad)))
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
                if (
                    inside is not None
                    and not inside[
                        min(gy + step_y // 2, gh - 1), min(gx + step_x // 2, gw - 1)
                    ]
                ):
                    continue
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


def compute_scene_fingerprint(
    region_pixels: np.ndarray, mask: np.ndarray | None = None
) -> dict[str, Any]:
    """Compute a feature-based fingerprint for scene classification.

    Combines HSV histogram, edge density, and color statistics into a
    fingerprint suitable for comparison via :func:`compare_scene_fingerprints`.

    For shaped regions, *mask* (uint8, crop-sized) restricts every component to
    polygon pixels. Fingerprints are only comparable when computed with the
    same mask, so callers must mask their reference fingerprints too —
    rasterized at each crop's own size (reference crops are source-resolution,
    scan crops may be rescaled).
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
        if mask is not None:
            mask = cv2.resize(
                mask, region_pixels.shape[1::-1], interpolation=cv2.INTER_NEAREST
            )
    if mask is not None and not np.any(mask):
        mask = None

    bins = config.SCREENSPACE_SCENE_HISTOGRAM_BINS
    hsv = cv2.cvtColor(region_pixels, cv2.COLOR_BGR2HSV)
    # 3D histogram flattened
    hist = cv2.calcHist(
        [hsv],
        [0, 1, 2],
        mask,
        [bins, bins, bins],
        [0, 180, 0, 256, 0, 256],
    )
    cv2.normalize(hist, hist)

    # Edge density
    gray = cv2.cvtColor(region_pixels, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 100, 200)
    if mask is not None:
        masked_edges = cv2.bitwise_and(edges, mask)
        denom = float(np.count_nonzero(mask))
        edge_density = float(np.count_nonzero(masked_edges)) / denom if denom else 0.0
    else:
        edge_density = (
            float(np.count_nonzero(edges)) / float(edges.size)
            if edges.size > 0
            else 0.0
        )

    # Color stats per channel
    inside = mask > 0 if mask is not None else None
    color_stats: list[float] = []
    for ch in range(3):
        channel = region_pixels[:, :, ch].astype(np.float64)
        if inside is not None:
            channel = channel[inside]
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

    # Weighted average. cv2.compareHist can return NaN on degenerate
    # (e.g. all-zero) histograms; clamp would not catch it because
    # NaN comparisons return False.
    score = 0.6 * hist_sim + 0.2 * edge_sim + 0.2 * color_sim
    if not math.isfinite(score):
        return 0.0
    return max(0.0, min(1.0, score))


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


# ── Attention (computational saliency) ───────────────────────────────
# Classic bottom-up composite (no learned model, no contrib modules):
# spectral residual + Lab center-surround contrast + frame-diff motion
# [+ optional Haar faces], multiplied by a center-weighted prior. The
# per-frame map feeds the same grid/heatmap pipeline as flow/change.


@functools.cache
def _dft_friendly(shape: tuple[int, int]) -> bool:
    """Whether ``cv2.dft`` is the faster transform for this frame shape.

    cv2's DFT is fastest on sizes that factor into 2/3/5 and is markedly
    *slower* than numpy on awkward ones — measured on this project's benchmark
    frames: 0.163 ms vs 0.400 ms at 144x256, but 4.021 ms vs 1.254 ms at
    151x257. The attention working dim pins only the longest axis to 256; the
    other follows the source aspect ratio and can land on a prime, so the fast
    path is gated rather than unconditional.
    """
    return all(cv2.getOptimalDFTSize(n) == n for n in shape)


def compute_spectral_residual(gray: np.ndarray) -> np.ndarray:
    """Spectral-residual saliency (Hou & Zhang 2007), float32 in [0, 1].

    The log-amplitude spectrum minus its local average isolates the
    "unexpected" frequency content; recombining it with the original phase
    highlights the spatial locations responsible for it.

    Two transforms, same math: on DFT-friendly shapes cv2 runs it in float32
    (the numpy branch promotes to complex128 and allocates four more full-size
    complex temporaries), recovering the phase term as ``z/|z|`` from the
    magnitude already in hand rather than through ``angle`` + a complex ``exp``.
    Agreement between the branches is ~1e-6 and is asserted in the tests.
    """
    f32 = gray.astype(np.float32)
    if _dft_friendly(f32.shape[:2]):
        spectrum = cv2.dft(f32, flags=cv2.DFT_COMPLEX_OUTPUT)
        real, imag = spectrum[:, :, 0], spectrum[:, :, 1]
        mag = cv2.magnitude(real, imag)
        log_amp = np.log1p(mag)
        residual = log_amp - cv2.blur(log_amp, (3, 3))
        scale = np.exp(residual)
        unit = np.divide(scale, mag, out=np.zeros_like(scale), where=mag > 0)
        out_real = real * unit
        out_imag = imag * unit
        # A zero coefficient (uniform frame) has no direction; the numpy branch
        # reads angle(0) as 0, i.e. a unit vector of 1+0j. Match that.
        flat = mag <= 0
        np.copyto(out_real, scale, where=flat)
        np.copyto(out_imag, np.float32(0.0), where=flat)
        inverse = cv2.idft(cv2.merge([out_real, out_imag]), flags=cv2.DFT_SCALE)
        sal = inverse[:, :, 0] ** 2 + inverse[:, :, 1] ** 2
    else:
        fft = np.fft.fft2(f32)
        log_amp = np.log1p(np.abs(fft)).astype(np.float32)
        phase = np.angle(fft)
        residual = log_amp - cv2.blur(log_amp, (3, 3))
        sal = np.abs(np.fft.ifft2(np.exp(residual) * np.exp(1j * phase))) ** 2
    sal = cv2.GaussianBlur(sal.astype(np.float32), (9, 9), 2.5)
    peak = float(sal.max())
    return sal / peak if peak > 0 else sal


def compute_color_contrast(bgr: np.ndarray) -> np.ndarray:
    """Lab center-surround contrast, float32 in [0, 1].

    Sum over L/a/b of |channel − wide Gaussian blur of channel|: bright/colored
    elements that differ from their surround score high regardless of hue.

    The surround is deliberately computed at **half scale**. Its sigma is
    ``max_dim / 8``, which on a float32 input asks OpenCV for a ~8σ+1 tap
    kernel — wider than the image itself at the attention working dim, and 66%
    of the whole attention callback. Halving the resolution costs 5x less
    (3.22 ms → 0.61 ms per frame) and moves the normalized map by at most
    0.012; a blur that broad is smooth enough that the downscale loses nothing
    it was measuring. Blurring the interleaved 3-channel Lab in one call is the
    obvious alternative and was measured *slower* (3.66 ms) — the kernel width
    is the cost, not the call count.
    """
    lab = cv2.cvtColor(bgr, cv2.COLOR_BGR2LAB).astype(np.float32)
    height, width = lab.shape[:2]
    sigma = max(3.0, max(height, width) / 8.0)
    small = cv2.resize(
        lab,
        (max(1, width // 2), max(1, height // 2)),
        interpolation=cv2.INTER_AREA,
    )
    surround = cv2.resize(
        cv2.GaussianBlur(small, (0, 0), sigma / 2.0),
        (width, height),
        interpolation=cv2.INTER_LINEAR,
    )
    diff = cv2.absdiff(lab, surround)
    contrast = diff[:, :, 0] + diff[:, :, 1] + diff[:, :, 2]
    peak = float(contrast.max())
    return contrast / peak if peak > 0 else contrast


def compute_motion_saliency(
    prev_gray: np.ndarray | None, curr_gray: np.ndarray
) -> np.ndarray:
    """Frame-diff motion map, float32 in [0, 1].

    Deliberately absolute (scaled by 255, not per-frame max): a static frame
    contributes ~zero motion instead of amplified sensor noise. Frame-diff
    rather than Farneback flow — attention needs *where* changed, not
    direction, at a fraction of the cost. Zeros when there is no comparable
    previous frame (first frame, or a resolution change between videos).
    """
    if prev_gray is None or prev_gray.shape != curr_gray.shape:
        return np.zeros(curr_gray.shape[:2], dtype=np.float32)
    diff = cv2.absdiff(prev_gray, curr_gray).astype(np.float32) / 255.0
    return cv2.GaussianBlur(diff, (9, 9), 2.5)


# Lazy Haar-cascade singleton. Typed Any because the legacy CascadeClassifier
# API only exists on OpenCV 4.x wheels — opencv-python-headless 5.x removed it,
# so it must never be named in annotations or called unguarded.
_face_cascade: Any | None = None


def _get_face_cascade() -> Any | None:
    """Lazy singleton for the bundled frontal-face Haar cascade.

    Returns ``None`` when this cv2 build ships neither the legacy
    ``CascadeClassifier`` API nor the bundled cascade data (removed in
    opencv-python-headless 5.x); the face channel then degrades to a zeros
    map and :func:`compute_saliency_map` leaves its weight out of the mix.
    """
    global _face_cascade
    if _face_cascade is None:
        cascade_cls = getattr(cv2, "CascadeClassifier", None)
        haar_dir = getattr(getattr(cv2, "data", None), "haarcascades", None)
        if cascade_cls is None or not haar_dir:
            return None
        _face_cascade = cascade_cls(haar_dir + "haarcascade_frontalface_default.xml")
    return _face_cascade


def face_detection_available() -> bool:
    """Whether this cv2 build supports the Haar face channel (OpenCV 4.x)."""
    return _get_face_cascade() is not None


def compute_face_saliency(gray: np.ndarray) -> np.ndarray:
    """Gaussian blobs at Haar face detections, float32 in [0, 1].

    Zeros map when no faces are found, and when the cv2 build has no
    CascadeClassifier support at all. Opt-in via
    ``SCREENSPACE_ATTENTION_FACE_CHANNEL`` — Haar false-positives on UI
    avatars/icons distort the map on footage without a webcam PiP.
    """
    blobs = np.zeros(gray.shape[:2], dtype=np.float32)
    cascade = _get_face_cascade()
    if cascade is None:
        return blobs
    faces = cascade.detectMultiScale(
        gray, scaleFactor=1.1, minNeighbors=4, minSize=(12, 12)
    )
    if len(faces) == 0:
        return blobs
    for x, y, w, h in faces:
        cv2.ellipse(
            blobs, (x + w // 2, y + h // 2), (w // 2, h // 2), 0, 0, 360, 1.0, -1
        )
    sigma = max(2.0, max(gray.shape[:2]) / 32.0)
    blobs = cv2.GaussianBlur(blobs, (0, 0), sigma)
    peak = float(blobs.max())
    return blobs / peak if peak > 0 else blobs


@functools.cache
def _center_prior(shape: tuple[int, int], bias: float) -> np.ndarray:
    """Center-weighted multiplicative prior: (1 − bias) + bias·gaussian.

    Cached per (shape, bias); treat the returned array as read-only.
    """
    h, w = shape
    ys = (np.arange(h, dtype=np.float32) - (h - 1) / 2.0) / max(1.0, 0.35 * h)
    xs = (np.arange(w, dtype=np.float32) - (w - 1) / 2.0) / max(1.0, 0.35 * w)
    gauss = np.exp(-0.5 * (ys[:, None] ** 2 + xs[None, :] ** 2))
    return ((1.0 - bias) + bias * gauss).astype(np.float32)


def compute_saliency_map(
    bgr: np.ndarray,
    prev_gray: np.ndarray | None,
    *,
    weights: dict[str, float] | None = None,
    center_bias: float | None = None,
    include_face: bool | None = None,
) -> tuple[np.ndarray, np.ndarray]:
    """Combined saliency map for one frame.

    Returns ``(map, curr_gray)`` — the caller rolls ``curr_gray`` forward as
    the next frame's ``prev_gray``. The map is the weighted mean of the
    enabled channels times the center prior, clipped to [0, 1]; it is NOT
    re-normalized to max=1, so peak strength stays comparable across frames
    (a static low-contrast screen genuinely pulls less attention).
    """
    if weights is None:
        weights = {
            "spectral": config.SCREENSPACE_ATTENTION_WEIGHT_SPECTRAL,
            "contrast": config.SCREENSPACE_ATTENTION_WEIGHT_CONTRAST,
            "motion": config.SCREENSPACE_ATTENTION_WEIGHT_MOTION,
            "face": config.SCREENSPACE_ATTENTION_WEIGHT_FACE,
        }
    if center_bias is None:
        center_bias = config.SCREENSPACE_ATTENTION_CENTER_BIAS
    if include_face is None:
        include_face = config.SCREENSPACE_ATTENTION_FACE_CHANNEL
    if include_face and not face_detection_available():
        # No CascadeClassifier in this cv2 build (opencv 5.x wheels): keep the
        # face weight out of the denominator so the map isn't dimmed by a
        # channel that can only ever contribute zeros.
        include_face = False

    curr_gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    combined = weights.get("spectral", 0.0) * compute_spectral_residual(curr_gray)
    combined += weights.get("contrast", 0.0) * compute_color_contrast(bgr)
    combined += weights.get("motion", 0.0) * compute_motion_saliency(
        prev_gray, curr_gray
    )
    total = (
        weights.get("spectral", 0.0)
        + weights.get("contrast", 0.0)
        + weights.get("motion", 0.0)
    )
    if include_face:
        combined += weights.get("face", 0.0) * compute_face_saliency(curr_gray)
        total += weights.get("face", 0.0)
    if total > 0:
        combined /= total
    combined *= _center_prior(combined.shape[:2], float(center_bias))
    return np.clip(combined, 0.0, 1.0).astype(np.float32), curr_gray


_SALIENCY_WEIGHT_PARAMS = {
    "spectral": "weight_spectral",
    "contrast": "weight_contrast",
    "motion": "weight_motion",
    "face": "weight_face",
}


def saliency_kwargs_from_params(params: dict[str, Any]) -> dict[str, Any]:
    """Map task/preview parameter dicts onto :func:`compute_saliency_map` kwargs.

    Shared by the tool layer and the preview builders so a per-task weight
    override tunes the scan and the Model view identically. Absent keys fall
    back to the ``SCREENSPACE_ATTENTION_*`` config defaults; a ``weight_face``
    of 0 disables the face channel entirely (no checkbox needed).
    """
    kwargs: dict[str, Any] = {}
    if any(pk in params for pk in _SALIENCY_WEIGHT_PARAMS.values()):
        kwargs["weights"] = {
            channel: float(
                params.get(
                    param_key,
                    getattr(config, "SCREENSPACE_ATTENTION_WEIGHT_" + channel.upper()),
                )
            )
            for channel, param_key in _SALIENCY_WEIGHT_PARAMS.items()
        }
    if "weight_face" in params:
        kwargs["include_face"] = float(params["weight_face"]) > 0
    if "center_bias" in params:
        kwargs["center_bias"] = float(params["center_bias"])
    return kwargs


def saliency_grid_from_map(
    sal: np.ndarray, grid_n: int, min_mag: float
) -> list[dict[str, float]]:
    """Downsample a saliency map to sparse normalized grid cells.

    Same ``{"x", "y", "mag"}`` cell shape as ``flow_grid``/``change_grid`` so
    the heatmap pipeline consumes it unchanged. Per-frame normalized by the
    grid max (dwell weighting: every sampled frame contributes one unit of
    attention wherever it looked, however weak the frame's absolute saliency).
    """
    grid_n = max(1, int(grid_n))
    cells = cv2.resize(sal, (grid_n, grid_n), interpolation=cv2.INTER_AREA)
    peak = float(cells.max())
    if peak <= 0:
        return []
    grid: list[dict[str, float]] = []
    for gy in range(grid_n):
        for gx in range(grid_n):
            mag = float(cells[gy, gx]) / peak
            if mag < min_mag:
                continue
            grid.append(
                {
                    "x": round((gx + 0.5) / grid_n, 3),
                    "y": round((gy + 0.5) / grid_n, 3),
                    "mag": round(mag, 3),
                }
            )
    return grid


def saliency_peak(sal: np.ndarray) -> tuple[float, float, float]:
    """Locate the attention peak: (peak_x, peak_y, peak_value), coords 0-1.

    The argmax runs on a blurred copy so a single hot pixel can't make the
    peak jitter between frames.
    """
    h, w = sal.shape[:2]
    blurred = cv2.GaussianBlur(sal, (0, 0), max(1.0, max(h, w) / 32.0))
    _, peak_val, _, peak_loc = cv2.minMaxLoc(blurred)
    return (
        round((peak_loc[0] + 0.5) / w, 4),
        round((peak_loc[1] + 0.5) / h, 4),
        round(float(peak_val), 4),
    )
