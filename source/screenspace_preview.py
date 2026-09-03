"""Preprocessing preview generation for the Screenspace "Model view" pane.

Renders what a tool's CV pipeline actually operates on — grayscale crops, diff
masks, edge maps, flow vectors, pHash bit grids — from a frame (plus optional
prev_frame / reference_frame / template_image), a region, and the tool's params.
Output is always a (H, W, 3) BGR uint8 ndarray, ready for ``cv2.imencode``.

Entry points:

- :func:`build_preview` — labeled multi-panel composite for the side panel.
- :func:`build_overlay_layer` — one layer at native region/frame resolution, for
  painting over the live frame canvas. Per-tool catalog: :data:`OVERLAY_LAYERS`.
"""

from typing import Any

import cv2
import numpy as np

import config
import screenspace_ocr
import screenspace_primitives


# Composite preview width cap; the UI pane is small.
_MAX_WIDTH = 512
_PANEL_GAP = 6
_LABEL_HEIGHT = 16

# Diff/magnitude colormap; JET matches screenspace_heatmap's _colorize_accumulator.
_DIFF_COLORMAP = cv2.COLORMAP_JET


def build_preview(
    frame: "np.ndarray",
    prev_frame: "np.ndarray | None",
    region: dict[str, Any] | None,
    tool: str,
    params: dict[str, Any],
) -> "np.ndarray":
    """Build a BGR preview image for the given tool and frame."""
    if tool == "color":
        return _preview_color(frame, region, params)
    if tool == "change":
        return _preview_change(frame, prev_frame, region, params)
    if tool == "similarity":
        return _preview_similarity(frame, region, params)
    if tool in ("text", "numbers"):
        return _preview_text_numbers(frame, region, params)
    if tool == "timelapse":
        return _preview_timelapse(frame, region)
    if tool == "template":
        return _preview_template(frame, region, params)
    if tool == "shape":
        return _preview_shape(frame, region, params)
    if tool == "flow":
        return _preview_flow(frame, prev_frame, region, params)
    if tool == "scene":
        return _preview_scene(frame, region, params)
    if tool in ("inactivity", "boundary"):
        # Boundary compares consecutive pHashes; show the same bit grid.
        return _preview_inactivity(frame, region)
    if tool == "attention":
        return _preview_attention(frame, prev_frame, params)
    if tool == "multitool":
        steps = params.get("steps") or []
        if steps:
            step = steps[0]
            return build_preview(
                frame,
                prev_frame,
                region,
                step.get("type", ""),
                step.get("parameters") or {},
            )
        return _placeholder("Add a step to see its preview")
    return _placeholder(f"Unknown tool: {tool}")


# ---------------------------------------------------------------------------
# Overlay layers: blink-comparator images over the live frame.
# Non-pixel-aligned tools (timelapse, inactivity) omitted.
# ---------------------------------------------------------------------------


# (layer_id, label, scope) per tool; scope "region" or "frame" sets the size.
OVERLAY_LAYERS: dict[str, list[tuple[str, str, str]]] = {
    "color": [("region", "Region (≤64 px)", "region")],
    "change": [
        ("changes", "Changes on frame", "region"),
        ("gray_blur", "Gray blur", "region"),
        ("abs_diff", "Abs diff", "region"),
        ("mask", "Threshold mask", "region"),
    ],
    "similarity": [
        ("gray", "Current gray", "region"),
        ("ssim_diff", "SSIM diff", "region"),
    ],
    "text": [("gray", "OCR input (gray)", "region")],
    "numbers": [("gray", "OCR input (gray)", "region")],
    "template": [("match_heatmap", "Match heatmap", "frame")],
    "shape": [
        ("edges", "Edge ridges", "frame"),
        ("match_heatmap", "Match heatmap", "frame"),
    ],
    "flow": [("flow_vectors", "Flow vectors", "region")],
    "scene": [("edges", "Canny edges", "region")],
    "attention": [("saliency_map", "Saliency map", "frame")],
}


def overlay_layer_scope(tool: str, layer: str) -> str | None:
    """Return the scope ("region" or "frame") for a given tool/layer, or None."""
    if tool == "multitool":
        return None
    for layer_id, _label, scope in OVERLAY_LAYERS.get(tool, []):
        if layer_id == layer:
            return scope
    return None


def build_overlay_layer(
    frame: "np.ndarray",
    prev_frame: "np.ndarray | None",
    region: dict[str, Any] | None,
    tool: str,
    layer: str,
    params: dict[str, Any],
) -> "np.ndarray | None":
    """Build a single overlay layer for the given tool, sized to the region or frame.

    Returns BGR uint8. Returns None if the layer can't be produced (e.g. missing
    prev frame for change/flow). Caller is responsible for validating that
    ``(tool, layer)`` is in :data:`OVERLAY_LAYERS`.
    """
    if tool == "multitool":
        steps = params.get("steps") or []
        if not steps:
            return None
        step = steps[0]
        return build_overlay_layer(
            frame,
            prev_frame,
            region,
            step.get("type", ""),
            layer,
            step.get("parameters") or {},
        )

    scope = overlay_layer_scope(tool, layer)
    if scope is None:
        return None

    if scope == "region":
        pixels = _clip_region_pixels(frame, region)
        if pixels is None:
            return None
    # scope == "frame" doesn't need region pixels

    if tool == "color" and layer == "region":
        return pixels.copy()

    if tool == "change":
        return _overlay_change(pixels, prev_frame, region, layer, params)

    if tool == "similarity" and layer == "gray":
        k = config.SCREENSPACE_BLUR_KERNEL
        gray = cv2.cvtColor(cv2.GaussianBlur(pixels, (k, k), 0), cv2.COLOR_BGR2GRAY)
        return _gray_to_bgr(gray)

    if tool == "similarity" and layer == "ssim_diff":
        return _overlay_ssim_diff(pixels, params)

    if tool in ("text", "numbers") and layer == "gray":
        if params.get("ocr_preprocess"):
            pixels = screenspace_ocr._preprocess_for_ocr(pixels)
        return _gray_to_bgr(cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY))

    if tool == "scene" and layer == "edges":
        gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
        edges = screenspace_primitives.canny_edges(gray)
        # Dilate so 1-px Canny lines survive browser downscaling; capped to
        # avoid chunky lines.
        thickness = max(1, min(2, min(gray.shape[:2]) // 300))
        if thickness > 1:
            kernel = np.ones((thickness, thickness), np.uint8)
            edges = cv2.dilate(edges, kernel, iterations=1)
        return _gray_to_bgr(edges)

    if tool == "flow" and layer == "flow_vectors":
        return _overlay_flow(pixels, prev_frame, region, params)

    if tool == "template" and layer == "match_heatmap":
        return _overlay_template_heatmap(frame, region, params)

    if tool == "shape" and layer == "edges":
        edges = screenspace_primitives._frame_edge_map(frame)
        return _gray_to_bgr(np.clip(edges, 0, 255).astype(np.uint8))

    if tool == "shape" and layer == "match_heatmap":
        return _overlay_shape_heatmap(frame, region, params)

    if tool == "attention" and layer == "saliency_map":
        return _overlay_attention_saliency(frame, prev_frame, params)

    return None


def _overlay_change(
    pixels: "np.ndarray",
    prev_frame: "np.ndarray | None",
    region: dict[str, Any] | None,
    layer: str,
    params: dict[str, Any],
) -> "np.ndarray | None":
    k = config.SCREENSPACE_BLUR_KERNEL
    curr_gray = cv2.cvtColor(cv2.GaussianBlur(pixels, (k, k), 0), cv2.COLOR_BGR2GRAY)
    if layer == "gray_blur":
        return _gray_to_bgr(curr_gray)
    if prev_frame is None:
        return None
    prev_pixels = _clip_region_pixels(prev_frame, region)
    if prev_pixels is None or prev_pixels.shape[:2] != pixels.shape[:2]:
        return None
    prev_gray = cv2.cvtColor(
        cv2.GaussianBlur(prev_pixels, (k, k), 0), cv2.COLOR_BGR2GRAY
    )
    diff = cv2.absdiff(prev_gray, curr_gray)
    if layer == "abs_diff":
        # Colorize so faint change reads; zero pixels stay black to blend.
        return _colorize_diff(diff, keep_zero_black=True)
    if layer in ("mask", "changes"):
        noise = int(params.get("noise_threshold", config.SCREENSPACE_NOISE_THRESHOLD))
        _, mask = cv2.threshold(diff, noise, 255, cv2.THRESH_BINARY)
        mk = config.SCREENSPACE_MORPH_KERNEL
        mask_clean = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((mk, mk), np.uint8))
        # Shaped region: zero changes outside the polygon, as the scan does.
        if region is not None:
            region_mask = screenspace_primitives.region_mask_for(
                region, *mask_clean.shape[:2]
            )
            if region_mask is not None:
                mask_clean = cv2.bitwise_and(mask_clean, region_mask)
        if layer == "mask":
            # Cyan-on-black reads against any frame content when alpha-blended.
            out = np.zeros(
                (mask_clean.shape[0], mask_clean.shape[1], 3), dtype=np.uint8
            )
            out[mask_clean > 0] = (220, 220, 0)  # BGR cyan-ish
            return out
        # "changes": tint changed pixels; unchanged pixels keep the live frame.
        return _tint_changes(pixels, diff, mask_clean)
    return None


def _overlay_ssim_diff(pixels: "np.ndarray", params: dict[str, Any]) -> "np.ndarray":
    """Colorized SSIM dissimilarity map (region-native). Warm = dissimilar.

    Similar (dissimilarity 0) stays black so it blends cleanly over the live
    frame. Mirrors the Similarity scan's preprocessing via
    ``screenspace_primitives.ssim_diff_map``. With no captured reference there is
    nothing to diff, so returns the live region (a no-op overlay) rather than
    ``None`` — which would 500 the preview route and leave the overlay stale.
    """
    ref = params.get("reference_frame")
    if not (isinstance(ref, np.ndarray) and ref.size > 0):
        return pixels.copy()
    if ref.shape[:2] != pixels.shape[:2]:
        ref = cv2.resize(
            ref, (pixels.shape[1], pixels.shape[0]), interpolation=cv2.INTER_AREA
        )
    _score, smap = screenspace_primitives.ssim_diff_map(pixels, ref)
    dis = np.clip((1.0 - smap) * 0.5, 0.0, 1.0)
    colored = _colorize_diff((dis * 255).astype(np.uint8), keep_zero_black=True)
    # ssim_diff_map runs at <=256; upscale to native so the overlay aligns.
    if colored.shape[:2] != pixels.shape[:2]:
        colored = cv2.resize(
            colored,
            (pixels.shape[1], pixels.shape[0]),
            interpolation=cv2.INTER_LINEAR,
        )
    return colored


def _draw_flow_arrows(
    vis: "np.ndarray",
    flow: "np.ndarray",
    grid_size: int,
    min_mag: float,
    coord_scale: float = 1.0,
) -> None:
    """Draw an 8x8-ish grid of flow arrows on ``vis`` in-place.

    Arrow length and line thickness scale with the grid cell size on ``vis``
    so arrows stay visible whether ``vis`` is a 256-px preview or a 1080-px
    overlay that the browser will subsequently scale down to display.

    ``coord_scale`` multiplies the flow grid coordinates so arrows can be
    drawn on a higher-resolution ``vis`` than the resolution at which
    ``flow`` was computed (used by the overlay to keep the gray background
    crisp while matching the actual CV pipeline's flow detection).
    """
    fh, fw = flow.shape[:2]
    step_y = max(1, fh // grid_size)
    step_x = max(1, fw // grid_size)
    vis_step_x = step_x * coord_scale
    vis_step_y = step_y * coord_scale
    vis_cell = min(vis_step_x, vis_step_y)
    max_arrow_len = vis_cell * 0.7
    thickness = max(1, min(2, round(vis_cell / 32)))

    samples: list[tuple[float, float, float, float, float]] = []
    max_mag = 0.0
    for gy in range(0, fh, step_y):
        for gx in range(0, fw, step_x):
            cy = gy + step_y // 2
            cx = gx + step_x // 2
            if cy >= fh or cx >= fw:
                continue
            fx, fy = float(flow[cy, cx, 0]), float(flow[cy, cx, 1])
            m = float(np.sqrt(fx * fx + fy * fy))
            if m < min_mag:
                continue
            samples.append((cx * coord_scale, cy * coord_scale, fx, fy, m))
            max_mag = max(max_mag, m)

    if not samples or max_mag <= 0:
        return
    arrow_scale = max_arrow_len / max_mag
    for vx, vy, fx, fy, m in samples:
        sx, sy = round(vx), round(vy)
        ex = round(vx + fx * arrow_scale)
        ey = round(vy + fy * arrow_scale)
        # Color-code by magnitude: slow = blue, fast = red (JET).
        color = _magnitude_color(m / max_mag)
        cv2.arrowedLine(vis, (sx, sy), (ex, ey), color, thickness, tipLength=0.3)


def _overlay_flow(
    pixels: "np.ndarray",
    prev_frame: "np.ndarray | None",
    region: dict[str, Any] | None,
    params: dict[str, Any],  # magnitude param affects threshold display only
) -> "np.ndarray | None":
    if prev_frame is None:
        return None
    prev_pixels = _clip_region_pixels(prev_frame, region)
    if prev_pixels is None or prev_pixels.shape[:2] != pixels.shape[:2]:
        return None
    curr_gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
    prev_gray = cv2.cvtColor(prev_pixels, cv2.COLOR_BGR2GRAY)

    # Match compute_optical_flow's downscale so arrows show the scored vectors;
    # coords rescale via coord_scale.
    h, w = prev_gray.shape[:2]
    max_dim = 256
    if h > max_dim or w > max_dim:
        scale = max_dim / max(h, w)
        small_w, small_h = int(w * scale), int(h * scale)
        prev_small = cv2.resize(
            prev_gray, (small_w, small_h), interpolation=cv2.INTER_AREA
        )
        curr_small = cv2.resize(
            curr_gray, (small_w, small_h), interpolation=cv2.INTER_AREA
        )
        coord_scale = w / float(small_w)
    else:
        prev_small, curr_small = prev_gray, curr_gray
        coord_scale = 1.0

    flow_out = np.zeros((*prev_small.shape[:2], 2), dtype=np.float32)
    flow = cv2.calcOpticalFlowFarneback(
        prev_small,
        curr_small,
        flow_out,
        config.SCREENSPACE_FLOW_PYR_SCALE,
        3,
        15,
        3,
        5,
        1.2,
        0,
    )
    vis = cv2.cvtColor(curr_gray, cv2.COLOR_GRAY2BGR)
    _draw_flow_arrows(
        vis,
        flow,
        config.SCREENSPACE_FLOW_GRID_SIZE,
        config.SCREENSPACE_FLOW_GRID_MIN_MAG,
        coord_scale=coord_scale,
    )
    return vis


def _overlay_template_heatmap(
    frame: "np.ndarray", region: dict[str, Any] | None, params: dict[str, Any]
) -> "np.ndarray | None":
    template = params.get("template_image")
    if not (isinstance(template, np.ndarray) and template.size > 0):
        return None
    mask = params.get("template_mask")
    if not (isinstance(mask, np.ndarray) and mask.size > 0):
        mask = None
    # Reuse the scan's template prep and correlation so the preview matches
    # a real scan.
    prepared = screenspace_primitives._prepare_template(template, mask)
    result = screenspace_primitives._template_correlation_map(frame, prepared)
    if result is None:
        return None
    tmpl_gray = prepared[0]
    window = screenspace_primitives.region_search_window(region or {})
    if window is not None:
        th_t, tw_t = tmpl_gray.shape[:2]
        result = screenspace_primitives._mask_corr_outside_window(
            result, tw_t, th_t, window
        )
        if result is None:
            return None
    norm = np.empty_like(result)
    cv2.normalize(result, norm, 0, 255, cv2.NORM_MINMAX)
    heat = cv2.applyColorMap(norm.astype(np.uint8), cv2.COLORMAP_JET)
    # matchTemplate anchors top-left: offset by half the template; replicate
    # edges to fill the frame.
    fh, fw = frame.shape[:2]
    hh, hw = heat.shape[:2]
    if (hh, hw) == (fh, fw):
        return heat
    th, tw = tmpl_gray.shape[:2]
    top = th // 2
    left = tw // 2
    bottom = fh - hh - top
    right = fw - hw - left
    return cv2.copyMakeBorder(
        heat, top, bottom, left, right, borderType=cv2.BORDER_REPLICATE
    )


def _shape_prepared(params: dict[str, Any]) -> "list[dict[str, Any]] | None":
    """Prepared shape reference from preview params, or None without one."""
    shape_img = params.get("shape_image")
    if not (isinstance(shape_img, np.ndarray) and shape_img.size > 0):
        return None
    mask = params.get("shape_mask")
    if not (isinstance(mask, np.ndarray) and mask.size > 0):
        mask = None
    return screenspace_primitives._prepare_shape_reference(
        shape_img,
        mask,
        float(params.get("scale_min", 0) or 0),
        float(params.get("scale_max", 0) or 0),
        int(params.get("scale_steps", 0) or 0),
        float(params.get("scale_y_min", 0) or 0),
        float(params.get("scale_y_max", 0) or 0),
        int(params.get("scale_y_steps", 0) or 0),
    )


def _shape_best_corr(
    frame_edges: "np.ndarray",
    prepared: "list[dict[str, Any]]",
    window: "tuple[float, float, float, float] | None" = None,
) -> "tuple[float, np.ndarray, dict[str, Any]] | None":
    """Best cross-scale correlation map — the same matchTemplate call and
    region-window masking the scan runs (see match_shape), so the preview
    reflects exactly what a scan sees."""
    best: tuple[float, np.ndarray, dict[str, Any]] | None = None
    fh, fw = frame_edges.shape[:2]
    for entry in prepared:
        if entry["h"] > fh or entry["w"] > fw:
            continue
        result = cv2.matchTemplate(frame_edges, entry["edges"], cv2.TM_CCOEFF_NORMED)
        if not np.all(np.isfinite(result)):
            result = np.where(np.isfinite(result), result, -1.0)
        if window is not None:
            result = screenspace_primitives._mask_corr_outside_window(
                result, entry["w"], entry["h"], window
            )
            if result is None:
                continue
        peak = float(result.max()) if result.size else -1.0
        if best is None or peak > best[0]:
            best = (peak, result, entry)
    return best


def _overlay_shape_heatmap(
    frame: "np.ndarray", region: dict[str, Any] | None, params: dict[str, Any]
) -> "np.ndarray | None":
    prepared = _shape_prepared(params)
    if not prepared:
        return None
    window = screenspace_primitives.region_search_window(region or {})
    best = _shape_best_corr(
        screenspace_primitives._frame_edge_map(frame), prepared, window
    )
    if best is None:
        return None
    _peak, result, entry = best
    norm = np.empty_like(result)
    cv2.normalize(result, norm, 0, 255, cv2.NORM_MINMAX)
    heat = cv2.applyColorMap(norm.astype(np.uint8), cv2.COLORMAP_JET)
    # Center like the template overlay: half-size offset, replicated edges.
    fh, fw = frame.shape[:2]
    hh, hw = heat.shape[:2]
    if (hh, hw) == (fh, fw):
        return heat
    top = entry["h"] // 2
    left = entry["w"] // 2
    bottom = fh - hh - top
    right = fw - hw - left
    return cv2.copyMakeBorder(
        heat, top, bottom, left, right, borderType=cv2.BORDER_REPLICATE
    )


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------


def _placeholder(message: str, w: int = 320, h: int = 120) -> "np.ndarray":
    """Return a dark panel with a centered message — used for empty states."""
    img = np.full((h, w, 3), 32, dtype=np.uint8)
    cv2.putText(
        img,
        message,
        (12, h // 2),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.5,
        (200, 200, 200),
        1,
        cv2.LINE_AA,
    )
    return img


def _gray_to_bgr(gray: "np.ndarray") -> "np.ndarray":
    if gray.ndim == 3:
        return gray
    return cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)


def _colorize_diff(
    gray: "np.ndarray", *, keep_zero_black: bool = False
) -> "np.ndarray":
    """Colorize a single-channel 0-255 magnitude map with JET → BGR uint8.

    Uses a fixed 0-255 scale (no per-frame normalization) so the color reflects
    the *absolute* change magnitude and faint diffs that read as near-black in
    grayscale become visible color. With ``keep_zero_black``, exactly-zero
    magnitudes are forced back to black (JET maps 0 to blue) so the layer
    alpha-blends cleanly when painted over the live frame.
    """
    if gray.ndim == 3:
        gray = cv2.cvtColor(gray, cv2.COLOR_BGR2GRAY)
    colored = cv2.applyColorMap(gray, _DIFF_COLORMAP)
    if keep_zero_black:
        colored[gray == 0] = 0
    return colored


def _magnitude_color(t: float) -> tuple[int, int, int]:
    """Map a normalized magnitude in [0, 1] to a BGR color via the diff colormap."""
    v = round(max(0.0, min(1.0, t)) * 255)
    bgr = cv2.applyColorMap(np.array([[v]], dtype=np.uint8), _DIFF_COLORMAP)[0, 0]
    return int(bgr[0]), int(bgr[1]), int(bgr[2])


def _tint_changes(
    pixels: "np.ndarray", diff: "np.ndarray", mask_clean: "np.ndarray"
) -> "np.ndarray":
    """Region pixels with the changed pixels tinted warm (colorized magnitude).

    Unchanged pixels keep the real frame content, so this reads as change in
    context — and as an on-frame overlay it does not darken the live frame, which
    ``renderOverlay`` alpha-blends whole at 0.7 (a black background would drop
    unchanged areas to ~30% brightness).
    """
    out = pixels.copy()
    sel = mask_clean > 0
    if np.any(sel):
        colored = _colorize_diff(diff, keep_zero_black=True)
        out[sel] = (0.35 * pixels[sel] + 0.65 * colored[sel]).astype(np.uint8)
    return out


def _fit_width(img: "np.ndarray", target_w: int) -> "np.ndarray":
    """Upscale/downscale preserving aspect so width == target_w."""
    h, w = img.shape[:2]
    if w == target_w:
        return img
    scale = target_w / float(w)
    new_h = max(1, round(h * scale))
    interp = cv2.INTER_AREA if scale < 1.0 else cv2.INTER_NEAREST
    return cv2.resize(img, (target_w, new_h), interpolation=interp)


def _label_panel(img: "np.ndarray", label: str) -> "np.ndarray":
    """Stack a small label strip above an image."""
    img = _gray_to_bgr(img)
    _h, w = img.shape[:2]
    strip = np.full((_LABEL_HEIGHT, w, 3), 24, dtype=np.uint8)
    cv2.putText(
        strip,
        label,
        (4, 12),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.38,
        (210, 210, 210),
        1,
        cv2.LINE_AA,
    )
    return np.vstack([strip, img])


def _hstack_panels(panels: list["np.ndarray"]) -> "np.ndarray":
    """Horizontally stack equal-height panels with a small gap between them."""
    if not panels:
        return _placeholder("No preview")
    target_h = max(p.shape[0] for p in panels)
    resized = []
    for p in panels:
        h, w = p.shape[:2]
        if h != target_h:
            scale = target_h / float(h)
            new_w = max(1, round(w * scale))
            p = cv2.resize(
                _gray_to_bgr(p), (new_w, target_h), interpolation=cv2.INTER_AREA
            )
        else:
            p = _gray_to_bgr(p)
        resized.append(p)
    gap = np.full((target_h, _PANEL_GAP, 3), 16, dtype=np.uint8)
    out = resized[0]
    for p in resized[1:]:
        out = np.hstack([out, gap, p])
    return out


def _clip_region_pixels(
    frame: "np.ndarray", region: dict[str, Any] | None
) -> "np.ndarray | None":
    if region is None:
        return None
    pixels = screenspace_primitives.extract_region(frame, region)
    if pixels.size == 0:
        return None
    # Shaped region: dim outside the polygon to show what the mask weighs.
    mask = screenspace_primitives.region_mask_for(region, *pixels.shape[:2])
    if mask is not None:
        pixels = pixels.copy()
        pixels[mask == 0] //= 4
    return pixels


# ---------------------------------------------------------------------------
# Per-tool previews
# ---------------------------------------------------------------------------


def _preview_color(
    frame: "np.ndarray",
    region: dict[str, Any] | None,
    params: dict[str, Any],
) -> "np.ndarray":
    pixels = _clip_region_pixels(frame, region)
    if pixels is None:
        return _placeholder("Select a region to preview")

    # Downscale for display only; average_color_hsv uses the full-resolution crop.
    h, w = pixels.shape[:2]
    if h > 64 or w > 64:
        down = cv2.resize(
            pixels, (min(w, 64), min(h, 64)), interpolation=cv2.INTER_AREA
        )
    else:
        down = pixels.copy()
    # Mask the mean to match the scan; inside pixels are undimmed.
    region_mask = (
        screenspace_primitives.region_mask_for(region, *pixels.shape[:2])
        if region is not None
        else None
    )
    mean_hsv = screenspace_primitives.average_color_hsv(pixels, mask=region_mask)

    # Target swatch from params (falls back to computed mean)
    tgt_h = float(params.get("h", mean_hsv["h"]))
    tgt_s = float(params.get("s", mean_hsv["s"]))
    tgt_v = float(params.get("v", mean_hsv["v"]))
    tgt_bgr = cv2.cvtColor(
        np.array([[[int(tgt_h), int(tgt_s), int(tgt_v)]]], dtype=np.uint8),
        cv2.COLOR_HSV2BGR,
    )[0, 0]
    mean_bgr = cv2.cvtColor(
        np.array(
            [[[int(mean_hsv["h"]), int(mean_hsv["s"]), int(mean_hsv["v"])]]],
            dtype=np.uint8,
        ),
        cv2.COLOR_HSV2BGR,
    )[0, 0]

    swatch_h = 80
    mean_swatch = np.full((swatch_h, 80, 3), mean_bgr, dtype=np.uint8)
    target_swatch = np.full((swatch_h, 80, 3), tgt_bgr, dtype=np.uint8)

    crop_panel = _label_panel(_fit_width(down, 120), "region (≤64 px)")
    mean_panel = _label_panel(
        mean_swatch,
        f"mean HSV {int(mean_hsv['h'])},{int(mean_hsv['s'])},{int(mean_hsv['v'])}",
    )
    target_panel = _label_panel(
        target_swatch, f"target {int(tgt_h)},{int(tgt_s)},{int(tgt_v)}"
    )
    return _hstack_panels([crop_panel, mean_panel, target_panel])


def _preview_change(
    frame: "np.ndarray",
    prev_frame: "np.ndarray | None",
    region: dict[str, Any] | None,
    params: dict[str, Any],
) -> "np.ndarray":
    pixels = _clip_region_pixels(frame, region)
    if pixels is None:
        return _placeholder("Select a region to preview")
    k = config.SCREENSPACE_BLUR_KERNEL
    curr_blur = cv2.GaussianBlur(pixels, (k, k), 0)
    curr_gray = cv2.cvtColor(curr_blur, cv2.COLOR_BGR2GRAY)

    if prev_frame is None:
        panel = _label_panel(_fit_width(curr_gray, 200), "gray-blur (no prev)")
        return panel

    prev_pixels = _clip_region_pixels(prev_frame, region)
    if prev_pixels is None or prev_pixels.shape[:2] != pixels.shape[:2]:
        panel = _label_panel(_fit_width(curr_gray, 200), "gray-blur (no prev)")
        return panel

    prev_blur = cv2.GaussianBlur(prev_pixels, (k, k), 0)
    prev_gray = cv2.cvtColor(prev_blur, cv2.COLOR_BGR2GRAY)

    noise = int(params.get("noise_threshold", config.SCREENSPACE_NOISE_THRESHOLD))
    diff = cv2.absdiff(prev_gray, curr_gray)
    _, mask = cv2.threshold(diff, noise, 255, cv2.THRESH_BINARY)
    mk = config.SCREENSPACE_MORPH_KERNEL
    mask_clean = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((mk, mk), np.uint8))

    # Shaped region: count and normalize inside the polygon only, like scan_changes.
    region_mask = (
        screenspace_primitives.region_mask_for(region, *mask_clean.shape[:2])
        if region is not None
        else None
    )
    if region_mask is not None:
        mask_clean = cv2.bitwise_and(mask_clean, region_mask)
        denom = float(np.count_nonzero(region_mask))
    else:
        denom = float(mask_clean.size)
    ratio = float(np.count_nonzero(mask_clean)) / denom if denom else 0.0

    # Tint the region where the cleaned mask fired, colorized by magnitude.
    changes = _tint_changes(pixels, diff, mask_clean)

    return _hstack_panels(
        [
            _label_panel(_fit_width(curr_gray, 160), "gray-blur"),
            _label_panel(_fit_width(_colorize_diff(diff), 160), "abs-diff"),
            _label_panel(_fit_width(changes, 160), f"changes (ratio {ratio:.3f})"),
        ]
    )


def _preview_similarity(
    frame: "np.ndarray",
    region: dict[str, Any] | None,
    params: dict[str, Any],
) -> "np.ndarray":
    pixels = _clip_region_pixels(frame, region)
    if pixels is None:
        return _placeholder("Select a region to preview")

    max_dim = 256
    h, w = pixels.shape[:2]
    if h > max_dim or w > max_dim:
        scale = max_dim / max(h, w)
        pixels_small = cv2.resize(
            pixels,
            (int(w * scale), int(h * scale)),
            interpolation=cv2.INTER_AREA,
        )
    else:
        pixels_small = pixels
    k = config.SCREENSPACE_BLUR_KERNEL
    curr_gray = cv2.cvtColor(
        cv2.GaussianBlur(pixels_small, (k, k), 0), cv2.COLOR_BGR2GRAY
    )

    ref_frame = params.get("reference_frame")
    panels = [_label_panel(_fit_width(curr_gray, 200), "current gray (≤256)")]
    if isinstance(ref_frame, np.ndarray) and ref_frame.size > 0:
        rh, rw = ref_frame.shape[:2]
        if rh > max_dim or rw > max_dim:
            rs = max_dim / max(rh, rw)
            ref_small = cv2.resize(
                ref_frame,
                (int(rw * rs), int(rh * rs)),
                interpolation=cv2.INTER_AREA,
            )
        else:
            ref_small = ref_frame
        ref_gray = cv2.cvtColor(
            cv2.GaussianBlur(ref_small, (k, k), 0), cv2.COLOR_BGR2GRAY
        )
        panels.append(_label_panel(_fit_width(ref_gray, 200), "reference gray"))

        # SSIM difference map, warm = dissimilar; reuses the scan's preprocessing.
        ref_for_ssim = ref_frame
        if ref_for_ssim.shape[:2] != pixels.shape[:2]:
            ref_for_ssim = cv2.resize(
                ref_for_ssim,
                (pixels.shape[1], pixels.shape[0]),
                interpolation=cv2.INTER_AREA,
            )
        score, smap = screenspace_primitives.ssim_diff_map(pixels, ref_for_ssim)
        dis = np.clip((1.0 - smap) * 0.5, 0.0, 1.0)
        heat = _colorize_diff((dis * 255).astype(np.uint8))
        panels.append(_label_panel(_fit_width(heat, 200), f"SSIM diff ({score:.3f})"))
    return _hstack_panels(panels)


def _preview_text_numbers(
    frame: "np.ndarray",
    region: dict[str, Any] | None,
    params: dict[str, Any],
) -> "np.ndarray":
    pixels = _clip_region_pixels(frame, region)
    if pixels is None:
        return _placeholder("Select a region to preview")
    label = "OCR input (gray)"
    if params.get("ocr_preprocess"):
        pixels = screenspace_ocr._preprocess_for_ocr(pixels)
        label = "OCR input (enhanced)"
    gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
    return _label_panel(_fit_width(gray, 300), label)


def _preview_timelapse(
    frame: "np.ndarray",
    region: dict[str, Any] | None,
) -> "np.ndarray":
    pixels = _clip_region_pixels(frame, region)
    if pixels is None:
        return _placeholder("Select a region to preview")
    return _label_panel(_fit_width(pixels, 300), "region crop (no preprocessing)")


def _preview_template(
    frame: "np.ndarray",
    region: dict[str, Any] | None,
    params: dict[str, Any],
) -> "np.ndarray":
    k = config.SCREENSPACE_BLUR_KERNEL
    frame_gray = cv2.cvtColor(cv2.GaussianBlur(frame, (k, k), 0), cv2.COLOR_BGR2GRAY)
    panels = [_label_panel(_fit_width(frame_gray, 240), "frame gray-blur")]

    template = params.get("template_image")
    if isinstance(template, np.ndarray) and template.size > 0:
        tmpl_gray = cv2.cvtColor(
            cv2.GaussianBlur(template, (k, k), 0), cv2.COLOR_BGR2GRAY
        )
        panels.append(_label_panel(_fit_width(tmpl_gray, 120), "template"))

        # Reuse the scan's prepared-template pipeline; a blurred mask would
        # inflate TM_CCOEFF_NORMED.
        mask = params.get("template_mask")
        if not (isinstance(mask, np.ndarray) and mask.size > 0):
            mask = None
        prepared = screenspace_primitives._prepare_template(template, mask)
        result = screenspace_primitives._template_correlation_map(frame, prepared)
        window = screenspace_primitives.region_search_window(region or {})
        if result is not None and window is not None:
            th_t, tw_t = prepared[0].shape[:2]
            result = screenspace_primitives._mask_corr_outside_window(
                result, tw_t, th_t, window
            )
        if result is not None:
            norm = np.empty_like(result)
            cv2.normalize(result, norm, 0, 255, cv2.NORM_MINMAX)
            heat = cv2.applyColorMap(norm.astype(np.uint8), cv2.COLORMAP_JET)
            panels.append(_label_panel(_fit_width(heat, 240), "match heatmap"))
    else:
        panels.append(_label_panel(_placeholder("no template", 120, 80), "template"))
    return _hstack_panels(panels)


def _preview_shape(
    frame: "np.ndarray",
    region: dict[str, Any] | None,
    params: dict[str, Any],
) -> "np.ndarray":
    # Built only from the scan's own helpers so panels match the real model.
    frame_edges = screenspace_primitives._frame_edge_map(frame)
    edges_u8 = np.clip(frame_edges, 0, 255).astype(np.uint8)
    panels = [_label_panel(_fit_width(edges_u8, 240), "frame edges")]

    prepared = _shape_prepared(params)
    if prepared is None:
        panels.append(_label_panel(_placeholder("no reference", 120, 80), "reference"))
        return _hstack_panels(panels)
    if not prepared:
        panels.append(
            _label_panel(_placeholder("no usable edges", 140, 80), "reference")
        )
        return _hstack_panels(panels)
    window = screenspace_primitives.region_search_window(region or {})
    best = _shape_best_corr(frame_edges, prepared, window)
    if best is None:
        panels.append(
            _label_panel(
                _placeholder("reference larger than frame", 180, 80), "reference"
            )
        )
        return _hstack_panels(panels)
    _peak, result, entry = best
    ref_u8 = np.clip(entry["edges"], 0, 255).astype(np.uint8)
    # ASCII only: cv2.putText renders non-ASCII glyphs (a multiply sign) as ??.
    scale_label = f"reference x{entry['scale']:.2f}"
    if entry.get("scale_y") not in (None, entry["scale"]):
        scale_label = f"reference x{entry['scale']:.2f}/{entry['scale_y']:.2f}"
    panels.append(_label_panel(_fit_width(ref_u8, 120), scale_label))
    norm = np.empty_like(result)
    cv2.normalize(result, norm, 0, 255, cv2.NORM_MINMAX)
    heat = cv2.applyColorMap(norm.astype(np.uint8), cv2.COLORMAP_JET)
    panels.append(_label_panel(_fit_width(heat, 240), "match heatmap"))
    return _hstack_panels(panels)


def _preview_flow(
    frame: "np.ndarray",
    prev_frame: "np.ndarray | None",
    region: dict[str, Any] | None,
    params: dict[str, Any],
) -> "np.ndarray":
    pixels = _clip_region_pixels(frame, region)
    if pixels is None:
        return _placeholder("Select a region to preview")
    curr_gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)

    if prev_frame is None:
        return _label_panel(_fit_width(curr_gray, 200), "current gray (no prev)")
    prev_pixels = _clip_region_pixels(prev_frame, region)
    if prev_pixels is None or prev_pixels.shape[:2] != pixels.shape[:2]:
        return _label_panel(_fit_width(curr_gray, 200), "current gray (no prev)")
    prev_gray = cv2.cvtColor(prev_pixels, cv2.COLOR_BGR2GRAY)

    max_dim = 256
    h, w = prev_gray.shape[:2]
    if h > max_dim or w > max_dim:
        scale = max_dim / max(h, w)
        new_w, new_h = int(w * scale), int(h * scale)
        prev_r = cv2.resize(prev_gray, (new_w, new_h), interpolation=cv2.INTER_AREA)
        curr_r = cv2.resize(curr_gray, (new_w, new_h), interpolation=cv2.INTER_AREA)
    else:
        prev_r, curr_r = prev_gray, curr_gray

    flow_out = np.zeros((*prev_r.shape[:2], 2), dtype=np.float32)
    flow = cv2.calcOpticalFlowFarneback(
        prev_r, curr_r, flow_out, config.SCREENSPACE_FLOW_PYR_SCALE, 3, 15, 3, 5, 1.2, 0
    )
    mag, _ = cv2.cartToPolar(flow[..., 0], flow[..., 1], angleInDegrees=True)

    vis = cv2.cvtColor(curr_r, cv2.COLOR_GRAY2BGR)
    _draw_flow_arrows(
        vis,
        flow,
        config.SCREENSPACE_FLOW_GRID_SIZE,
        config.SCREENSPACE_FLOW_GRID_MIN_MAG,
    )

    # Shaped region: flow runs on the full rect, stats on polygon pixels only.
    region_mask = (
        screenspace_primitives.region_mask_for(region, *mag.shape[:2])
        if region is not None
        else None
    )
    if region_mask is not None and np.any(region_mask):
        mean_mag = float(np.mean(mag[region_mask > 0]))
    else:
        mean_mag = float(np.mean(mag))
    thresh = float(
        params.get("magnitude_threshold", config.SCREENSPACE_FLOW_MAGNITUDE_THRESHOLD)
    )
    return _hstack_panels(
        [
            _label_panel(_fit_width(prev_gray, 160), "prev gray"),
            _label_panel(
                _fit_width(vis, 160),
                f"flow vectors (mean {mean_mag:.2f}, thr {thresh:.2f})",
            ),
        ]
    )


def _preview_scene(
    frame: "np.ndarray",
    region: dict[str, Any] | None,
    params: dict[str, Any],  # reserved for future scene-ref overlays
) -> "np.ndarray":
    pixels = _clip_region_pixels(frame, region)
    if pixels is None:
        return _placeholder("Select a region to preview")

    max_dim = 128
    h, w = pixels.shape[:2]
    if h > max_dim or w > max_dim:
        scale = max_dim / max(h, w)
        pixels_small = cv2.resize(
            pixels,
            (int(w * scale), int(h * scale)),
            interpolation=cv2.INTER_AREA,
        )
    else:
        pixels_small = pixels

    # Canny at native resolution (as compute_scene_fingerprint), then downscale
    # the edge map for display.
    gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
    edges = screenspace_primitives.canny_edges(gray)
    # Shaped region: only polygon pixels feed edge density and histogram.
    region_mask = (
        screenspace_primitives.region_mask_for(region, *pixels.shape[:2])
        if region is not None
        else None
    )
    if region_mask is not None and np.any(region_mask):
        edge_density = float(np.count_nonzero(cv2.bitwise_and(edges, region_mask))) / (
            float(np.count_nonzero(region_mask))
        )
    else:
        edge_density = (
            float(np.count_nonzero(edges)) / float(edges.size) if edges.size else 0.0
        )
    if edges.shape[:2] != pixels_small.shape[:2]:
        edges = cv2.resize(
            edges,
            (pixels_small.shape[1], pixels_small.shape[0]),
            interpolation=cv2.INTER_AREA,
        )

    # 8-bin hue histogram strip
    hsv = cv2.cvtColor(pixels_small, cv2.COLOR_BGR2HSV)
    small_mask = (
        screenspace_primitives.region_mask_for(region, *pixels_small.shape[:2])
        if region is not None
        else None
    )
    hist = cv2.calcHist([hsv], [0], small_mask, [8], [0, 180]).flatten()
    hist = hist / (hist.max() + 1e-6)
    bar_w, bar_h = 24, 80
    strip = np.full((bar_h, bar_w * len(hist), 3), 20, dtype=np.uint8)
    for i, v in enumerate(hist):
        hue_deg = int((i + 0.5) * (180.0 / len(hist)))
        bgr = cv2.cvtColor(
            np.array([[[hue_deg, 200, 220]]], dtype=np.uint8), cv2.COLOR_HSV2BGR
        )[0, 0]
        h_px = max(1, int(v * (bar_h - 4)))
        cv2.rectangle(
            strip,
            (i * bar_w + 2, bar_h - h_px),
            ((i + 1) * bar_w - 2, bar_h - 1),
            (int(bgr[0]), int(bgr[1]), int(bgr[2])),
            -1,
        )

    return _hstack_panels(
        [
            _label_panel(_fit_width(pixels_small, 160), "region (≤128)"),
            _label_panel(_fit_width(edges, 160), f"Canny edges ({edge_density:.2%})"),
            _label_panel(_fit_width(strip, 160), "hue histogram (8 bins)"),
        ]
    )


def _preview_inactivity(
    frame: "np.ndarray",
    region: dict[str, Any] | None,
) -> "np.ndarray":
    pixels = _clip_region_pixels(frame, region)
    if pixels is None:
        return _placeholder("Select a region to preview")

    ph = screenspace_primitives.compute_phash(pixels)
    # PHash exposes .hash as a 2D bool ndarray (typically 8×8 for phash)
    bits = np.asarray(ph.hash, dtype=np.uint8) * 255
    # Upscale to a visible grid
    cell = 16
    grid = np.kron(bits, np.ones((cell, cell), dtype=np.uint8))
    grid_bgr = cv2.cvtColor(grid, cv2.COLOR_GRAY2BGR)
    # Draw gridlines for readability
    n = bits.shape[0]
    for i in range(1, n):
        cv2.line(grid_bgr, (i * cell, 0), (i * cell, n * cell), (60, 60, 60), 1)
        cv2.line(grid_bgr, (0, i * cell), (n * cell, i * cell), (60, 60, 60), 1)

    return _hstack_panels(
        [
            _label_panel(_fit_width(pixels, 160), "region"),
            _label_panel(_fit_width(grid_bgr, 160), f"pHash bits ({n}×{n})"),
        ]
    )


def _attention_working_frames(
    frame: "np.ndarray", prev_frame: "np.ndarray | None"
) -> "tuple[np.ndarray, np.ndarray | None]":
    """Downscale to the scan's working size; prev gray only when comparable."""
    max_dim = config.SCREENSPACE_ATTENTION_WORKING_DIM
    h, w = frame.shape[:2]
    if h > max_dim or w > max_dim:
        scale = max_dim / max(h, w)
        size = (int(w * scale), int(h * scale))
        small = cv2.resize(frame, size, interpolation=cv2.INTER_AREA)
    else:
        small = frame
    prev_gray = None
    if prev_frame is not None and prev_frame.shape[:2] == frame.shape[:2]:
        prev_small = cv2.resize(
            prev_frame, (small.shape[1], small.shape[0]), interpolation=cv2.INTER_AREA
        )
        prev_gray = cv2.cvtColor(prev_small, cv2.COLOR_BGR2GRAY)
    return small, prev_gray


def _saliency_to_bgr(sal: "np.ndarray", *, colorize: bool = False) -> "np.ndarray":
    """Map a [0, 1] float saliency map to a displayable BGR uint8 image."""
    img = np.clip(sal * 255.0, 0, 255).astype(np.uint8)
    if colorize:
        return cv2.applyColorMap(img, _DIFF_COLORMAP)
    return _gray_to_bgr(img)


def _preview_attention(
    frame: "np.ndarray",
    prev_frame: "np.ndarray | None",
    params: dict[str, Any] | None = None,
) -> "np.ndarray":
    small, prev_gray = _attention_working_frames(frame, prev_frame)
    curr_gray = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY)

    spectral = screenspace_primitives.compute_spectral_residual(curr_gray)
    contrast = screenspace_primitives.compute_color_contrast(small)
    combined, _ = screenspace_primitives.compute_saliency_map(
        small,
        prev_gray,
        **screenspace_primitives.saliency_kwargs_from_params(params or {}),
    )
    _px, _py, peak_value = screenspace_primitives.saliency_peak(combined)

    if prev_gray is not None:
        motion = screenspace_primitives.compute_motion_saliency(prev_gray, curr_gray)
        motion_panel = _label_panel(_fit_width(_saliency_to_bgr(motion), 160), "motion")
    else:
        motion_panel = _label_panel(
            _placeholder("no prev frame", w=160, h=90), "motion"
        )

    return _hstack_panels(
        [
            _label_panel(_fit_width(small, 160), "input (≤256 px)"),
            _label_panel(_fit_width(_saliency_to_bgr(spectral), 160), "spectral"),
            _label_panel(_fit_width(_saliency_to_bgr(contrast), 160), "contrast"),
            motion_panel,
            _label_panel(
                _fit_width(_saliency_to_bgr(combined, colorize=True), 160),
                f"saliency (peak {peak_value:.2f})",
            ),
        ]
    )


def _overlay_attention_saliency(
    frame: "np.ndarray",
    prev_frame: "np.ndarray | None",
    params: dict[str, Any] | None = None,
) -> "np.ndarray":
    """Combined saliency map colorized and resized to native frame resolution."""
    small, prev_gray = _attention_working_frames(frame, prev_frame)
    combined, _ = screenspace_primitives.compute_saliency_map(
        small,
        prev_gray,
        **screenspace_primitives.saliency_kwargs_from_params(params or {}),
    )
    heat = _saliency_to_bgr(combined, colorize=True)
    fh, fw = frame.shape[:2]
    if heat.shape[:2] != (fh, fw):
        heat = cv2.resize(heat, (fw, fh), interpolation=cv2.INTER_LINEAR)
    return heat


def encode_png(img: "np.ndarray", *, cap_width: bool = True) -> bytes:
    """Encode a BGR image as PNG bytes — convenience for the Flask route.

    By default caps width at :data:`_MAX_WIDTH` for the side-panel composite.
    Pass ``cap_width=False`` for overlay layers, which need native resolution
    so they paint pixel-true on top of the frame canvas.
    """
    if cap_width and img.shape[1] > _MAX_WIDTH:
        img = _fit_width(img, _MAX_WIDTH)
    ok, buf = cv2.imencode(".png", img)
    if not ok:
        return b""
    return buf.tobytes()
