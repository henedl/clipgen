"""Composer annotation renderer: PIL RGBA overlays, no Flask and no ffmpeg drawtext."""

from __future__ import annotations

import math
from pathlib import Path
from typing import Any

import config

# Probed in order; load_default() never fails. Text metrics differ from the Inter
# preview (see composer-annotate.js).
_ANNOTATION_FONT_PATHS = (
    "/System/Library/Fonts/Helvetica.ttc",
    "/Library/Fonts/Arial.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "C:/Windows/Fonts/arial.ttf",
)
_annotation_font_cache: dict[int, Any] = {}


def _annotation_font(size: int) -> Any:
    if size in _annotation_font_cache:
        return _annotation_font_cache[size]
    from PIL import ImageFont

    font: Any = None
    for path in _ANNOTATION_FONT_PATHS:
        if Path(path).is_file():
            try:
                font = ImageFont.truetype(path, size)
                break
            except OSError:
                continue
    if font is None:
        font = ImageFont.load_default()
    _annotation_font_cache[size] = font
    return font


def _parse_hex_color(value: str) -> tuple[int, int, int, int]:
    raw = str(value or "").lstrip("#")
    try:
        if len(raw) == 3:
            raw = "".join(ch * 2 for ch in raw)
        r, g, b = (int(raw[i : i + 2], 16) for i in (0, 2, 4))
        return (r, g, b, 255)
    except (ValueError, IndexError):
        return (240, 90, 60, 255)  # config default's RGB


def _sample_polyline(
    pts: list[tuple[float, float]], spacing: float
) -> list[tuple[float, float]]:
    """Points spaced ~*spacing* px along the polyline (endpoints included)."""
    result = [pts[0]]
    next_at = spacing
    dist = 0.0
    for i in range(len(pts) - 1):
        (x0, y0), (x1, y1) = pts[i], pts[i + 1]
        seg = math.hypot(x1 - x0, y1 - y0)
        if seg <= 0:
            continue
        ux, uy = (x1 - x0) / seg, (y1 - y0) / seg
        while next_at <= dist + seg:
            t = next_at - dist
            result.append((x0 + ux * t, y0 + uy * t))
            next_at += spacing
        dist += seg
    result.append(pts[-1])
    return result


def _dash_segments(
    pts: list[tuple[float, float]], on: float, off: float
) -> list[tuple[tuple[float, float], tuple[float, float]]]:
    """(a, b) endpoint pairs for the 'on' runs of a dashed polyline.

    ``pos`` carries the arc length across polyline vertices so the dash pattern
    stays continuous around corners (the same look ``ctx.setLineDash`` gives).
    """
    period = on + off
    segs: list[tuple[tuple[float, float], tuple[float, float]]] = []
    pos = 0.0
    for i in range(len(pts) - 1):
        (x0, y0), (x1, y1) = pts[i], pts[i + 1]
        seg = math.hypot(x1 - x0, y1 - y0)
        if seg <= 0:
            continue
        ux, uy = (x1 - x0) / seg, (y1 - y0) / seg
        d = 0.0
        while d < seg:
            phase = pos % period
            if phase < on:
                run = min(on - phase, seg - d)
                segs.append(
                    (
                        (x0 + ux * d, y0 + uy * d),
                        (x0 + ux * (d + run), y0 + uy * (d + run)),
                    )
                )
            else:
                run = min(period - phase, seg - d)
            d += run
            pos += run
    return segs


def _dash_polyline(
    draw: Any,
    points: list[tuple[float, float]],
    color: Any,
    width: int,
    style: str,
) -> None:
    """Stroke a polyline dashed/dotted — PIL has no native dash support."""
    pts = [(float(x), float(y)) for (x, y) in points]
    if len(pts) < 2:
        return
    if style == "dotted":
        r = max(1.0, width / 2.0)
        for cx, cy in _sample_polyline(pts, max(2.0, width * 2.0)):
            draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=color)
        return
    for a, b in _dash_segments(pts, max(1.0, width * 2.5), max(1.0, width * 2.0)):
        draw.line([a, b], fill=color, width=width, joint="curve")


def _rotate(
    cx: float, cy: float, dx: float, dy: float, cos_r: float, sin_r: float
) -> tuple[float, float]:
    """Rotate offset (dx, dy) about (cx, cy)."""
    return cx + dx * cos_r - dy * sin_r, cy + dx * sin_r + dy * cos_r


def render_annotation_overlay(
    annotations: list[dict[str, Any]], width: int, height: int
) -> Any:
    """Draw 0..1-normalized annotations on a transparent RGBA frame, matching the browser preview."""
    from PIL import Image, ImageDraw

    img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    for ann in annotations:
        style = ann.get("style") or {}
        color = _parse_hex_color(style.get("color", ""))
        geometry = ann.get("geometry") or {}
        stroke_width = style.get("strokeWidth", config.COMPOSER_ANNOTATION_STROKE_WIDTH)
        stroke = max(1, round(float(stroke_width) * width))
        stroke_style = str(style.get("strokeStyle") or "solid")
        if ann.get("type") == "freehand":
            points = [
                (float(p[0]) * width, float(p[1]) * height)
                for p in geometry.get("points", [])
            ]
            if len(points) == 1:
                x, y = points[0]
                r = max(stroke, 2)
                draw.ellipse([x - r, y - r, x + r, y + r], fill=color)
            elif points:
                if stroke_style == "solid":
                    draw.line(points, fill=color, width=stroke, joint="curve")
                else:
                    _dash_polyline(draw, points, color, stroke, stroke_style)
        elif ann.get("type") == "shape":
            cx = float(geometry.get("x", 0)) * width
            cy = float(geometry.get("y", 0)) * height
            sw = float(geometry.get("w", 0)) * width
            sh = float(geometry.get("h", 0)) * height
            rotation = float(geometry.get("rotation", 0) or 0.0)
            if sw < 1 or sh < 1:
                continue
            # Match the browser's ctx.rotate (y-down, positive = clockwise).
            rad = math.radians(rotation)
            cos_r, sin_r = math.cos(rad), math.sin(rad)
            if geometry.get("shape") == "rect":
                # Repeating a corner fills the start joint.
                corners = [
                    _rotate(cx, cy, dx, dy, cos_r, sin_r)
                    for dx, dy in (
                        (-sw / 2, -sh / 2),
                        (sw / 2, -sh / 2),
                        (sw / 2, sh / 2),
                        (-sw / 2, sh / 2),
                    )
                ]
                if stroke_style == "solid":
                    draw.line(
                        corners + [corners[0], corners[1]],
                        fill=color,
                        width=stroke,
                        joint="curve",
                    )
                else:
                    _dash_polyline(
                        draw, corners + [corners[0]], color, stroke, stroke_style
                    )
            elif stroke_style != "solid":
                # PIL cannot dash an ellipse: sample the rotated perimeter as a polygon
                # and dash-walk it.
                a, b = sw / 2, sh / 2
                perim = math.pi * (
                    3 * (a + b) - math.sqrt(max(0.0, (3 * a + b) * (a + 3 * b)))
                )
                n = max(48, int(perim / max(1.0, stroke * 2)))
                poly = []
                for k in range(n + 1):
                    th = 2 * math.pi * k / n
                    ex, ey = a * math.cos(th), b * math.sin(th)
                    poly.append(_rotate(cx, cy, ex, ey, cos_r, sin_r))
                _dash_polyline(draw, poly, color, stroke, stroke_style)
            else:  # solid ellipse
                box = [cx - sw / 2, cy - sh / 2, cx + sw / 2, cy + sh / 2]
                if abs(rotation) < 0.01:
                    draw.ellipse(box, outline=color, width=stroke)
                else:
                    # PIL cannot stroke a rotated ellipse; rotate a temp layer instead.
                    # PIL's angle is counter-clockwise.
                    layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
                    ImageDraw.Draw(layer).ellipse(box, outline=color, width=stroke)
                    layer = layer.rotate(
                        -rotation, resample=Image.Resampling.BICUBIC, center=(cx, cy)
                    )
                    img = Image.alpha_composite(img, layer)
                    draw = ImageDraw.Draw(img)
        elif ann.get("type") == "text":
            text = str(geometry.get("text") or "")
            if not text:
                continue
            size = max(
                8,
                round(
                    float(style.get("fontSize", config.COMPOSER_ANNOTATION_FONT_SIZE))
                    * height
                ),
            )
            font = _annotation_font(size)
            x = float(geometry.get("x", 0)) * width
            y = float(geometry.get("y", 0)) * height
            # Soft dark backing box keeps text legible over any footage.
            bbox = draw.textbbox((x, y), text, font=font)
            pad = max(2, round(size * 0.25))
            draw.rectangle(
                [bbox[0] - pad, bbox[1] - pad, bbox[2] + pad, bbox[3] + pad],
                fill=(0, 0, 0, 110),
            )
            draw.text((x, y), text, fill=color, font=font)
    return img
