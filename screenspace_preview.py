# -*- coding: utf-8 -*-
"""Preprocessing preview generation for the Screenspace "Model view" pane.

Given a frame (and optionally a prev_frame / reference_frame / template_image),
a region, a tool type, and the tool's parameters, returns a composite BGR image
showing what that tool's CV pipeline actually operates on — grayscale crops,
diff masks, edge maps, flow vectors, pHash bit grids, etc.

The produced image is always a single ``numpy.ndarray`` of shape (H, W, 3) in
BGR uint8, suitable for ``cv2.imencode('.png', img)``.

Entry point: :func:`build_preview`.
"""

from typing import Any

import cv2
import numpy as np

import config
import screenspace


# Max width (px) of the composite preview image.  Kept modest: the UI pane is
# small and generating larger images just wastes bandwidth.
_MAX_WIDTH = 512
_PANEL_GAP = 6
_LABEL_HEIGHT = 16


def build_preview(
    frame: "np.ndarray",
    prev_frame: "np.ndarray | None",
    region: dict[str, int] | None,
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
        return _preview_template(frame, params)
    if tool == "flow":
        return _preview_flow(frame, prev_frame, region, params)
    if tool == "scene":
        return _preview_scene(frame, region, params)
    if tool == "inactivity":
        return _preview_inactivity(frame, region)
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


def _fit_width(img: "np.ndarray", target_w: int) -> "np.ndarray":
    """Upscale/downscale preserving aspect so width == target_w."""
    h, w = img.shape[:2]
    if w == target_w:
        return img
    scale = target_w / float(w)
    new_h = max(1, int(round(h * scale)))
    interp = cv2.INTER_AREA if scale < 1.0 else cv2.INTER_NEAREST
    return cv2.resize(img, (target_w, new_h), interpolation=interp)


def _label_panel(img: "np.ndarray", label: str) -> "np.ndarray":
    """Stack a small label strip above an image."""
    img = _gray_to_bgr(img)
    h, w = img.shape[:2]
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
            new_w = max(1, int(round(w * scale)))
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
    frame: "np.ndarray", region: dict[str, int] | None
) -> "np.ndarray | None":
    if region is None:
        return None
    pixels = screenspace.extract_region(frame, region)
    if pixels.size == 0:
        return None
    return pixels


# ---------------------------------------------------------------------------
# Per-tool previews
# ---------------------------------------------------------------------------


def _preview_color(
    frame: "np.ndarray",
    region: dict[str, int] | None,
    params: dict[str, Any],
) -> "np.ndarray":
    pixels = _clip_region_pixels(frame, region)
    if pixels is None:
        return _placeholder("Select a region to preview")

    # Downscaled crop (mirrors average_color_hsv's ≤64 resize)
    h, w = pixels.shape[:2]
    if h > 64 or w > 64:
        down = cv2.resize(
            pixels, (min(w, 64), min(h, 64)), interpolation=cv2.INTER_AREA
        )
    else:
        down = pixels.copy()
    mean_hsv = screenspace.average_color_hsv(pixels)

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
    region: dict[str, int] | None,
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

    ratio = (
        float(np.count_nonzero(mask_clean)) / float(mask_clean.size)
        if mask_clean.size
        else 0.0
    )

    return _hstack_panels(
        [
            _label_panel(_fit_width(curr_gray, 160), "gray-blur"),
            _label_panel(_fit_width(diff, 160), "abs-diff"),
            _label_panel(_fit_width(mask_clean, 160), f"mask (ratio {ratio:.3f})"),
        ]
    )


def _preview_similarity(
    frame: "np.ndarray",
    region: dict[str, int] | None,
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
    return _hstack_panels(panels)


def _preview_text_numbers(
    frame: "np.ndarray",
    region: dict[str, int] | None,
    params: dict[str, Any],
) -> "np.ndarray":
    pixels = _clip_region_pixels(frame, region)
    if pixels is None:
        return _placeholder("Select a region to preview")
    gray = cv2.cvtColor(pixels, cv2.COLOR_BGR2GRAY)
    return _label_panel(_fit_width(gray, 300), "OCR input (gray)")


def _preview_timelapse(
    frame: "np.ndarray",
    region: dict[str, int] | None,
) -> "np.ndarray":
    pixels = _clip_region_pixels(frame, region)
    if pixels is None:
        return _placeholder("Select a region to preview")
    return _label_panel(_fit_width(pixels, 300), "region crop (no preprocessing)")


def _preview_template(
    frame: "np.ndarray",
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

        # Match heatmap (only when template fits inside the frame)
        if (
            tmpl_gray.shape[0] <= frame_gray.shape[0]
            and tmpl_gray.shape[1] <= frame_gray.shape[1]
            and float(np.std(tmpl_gray)) >= 1e-6
        ):
            mask = params.get("template_mask")
            gray_mask = None
            if isinstance(mask, np.ndarray) and mask.size > 0:
                gray_mask = cv2.GaussianBlur(mask, (k, k), 0)
            result = cv2.matchTemplate(
                frame_gray, tmpl_gray, cv2.TM_CCOEFF_NORMED, mask=gray_mask
            )
            result = np.nan_to_num(result, nan=0.0, posinf=1.0, neginf=-1.0)
            norm = np.empty_like(result)
            cv2.normalize(result, norm, 0, 255, cv2.NORM_MINMAX)
            heat = cv2.applyColorMap(norm.astype(np.uint8), cv2.COLORMAP_JET)
            panels.append(_label_panel(_fit_width(heat, 240), "match heatmap"))
    else:
        panels.append(_label_panel(_placeholder("no template", 120, 80), "template"))
    return _hstack_panels(panels)


def _preview_flow(
    frame: "np.ndarray",
    prev_frame: "np.ndarray | None",
    region: dict[str, int] | None,
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

    # Draw flow arrows on the current-frame panel
    vis = cv2.cvtColor(curr_r, cv2.COLOR_GRAY2BGR)
    grid_size = config.SCREENSPACE_FLOW_GRID_SIZE
    gh, gw = mag.shape[:2]
    step_y = max(1, gh // grid_size)
    step_x = max(1, gw // grid_size)
    min_mag = config.SCREENSPACE_FLOW_GRID_MIN_MAG
    for gy in range(0, gh, step_y):
        for gx in range(0, gw, step_x):
            cy = gy + step_y // 2
            cx = gx + step_x // 2
            if cy >= gh or cx >= gw:
                continue
            fx, fy = float(flow[cy, cx, 0]), float(flow[cy, cx, 1])
            m = float(np.sqrt(fx * fx + fy * fy))
            if m < min_mag:
                continue
            scale = 4.0
            end = (int(cx + fx * scale), int(cy + fy * scale))
            cv2.arrowedLine(vis, (cx, cy), end, (40, 220, 40), 1, tipLength=0.3)

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
    region: dict[str, int] | None,
    params: dict[str, Any],  # noqa: ARG001 — reserved for future scene-ref overlays
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

    gray = cv2.cvtColor(pixels_small, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 100, 200)
    edge_density = (
        float(np.count_nonzero(edges)) / float(edges.size) if edges.size else 0.0
    )

    # 8-bin hue histogram strip
    hsv = cv2.cvtColor(pixels_small, cv2.COLOR_BGR2HSV)
    hist = cv2.calcHist([hsv], [0], None, [8], [0, 180]).flatten()
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
    region: dict[str, int] | None,
) -> "np.ndarray":
    pixels = _clip_region_pixels(frame, region)
    if pixels is None:
        return _placeholder("Select a region to preview")

    ph = screenspace.compute_phash(pixels)
    # imagehash.ImageHash exposes .hash as a 2D bool ndarray (typically 8×8 for phash)
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


def encode_png(img: "np.ndarray") -> bytes:
    """Encode a BGR image as PNG bytes — convenience for the Flask route."""
    # Cap overall width so the UI pane never has to scale huge images down.
    if img.shape[1] > _MAX_WIDTH:
        img = _fit_width(img, _MAX_WIDTH)
    ok, buf = cv2.imencode(".png", img)
    if not ok:
        return b""
    return buf.tobytes()
