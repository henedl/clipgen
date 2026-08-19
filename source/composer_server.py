"""Composer Flask blueprint — source-video cutting timeline + local manifest.

Registered at ``/composer`` by ``server.build_combined_app`` (mutually exclusive
launch with the other web modes, but all blueprints are always mounted). Phase 1
ships the static page routes, participant/video discovery, and the composer
manifest CRUD: in/out **cut pairs** plus persisted UI state (marker-lane
toggles + fold states). Phase 2 adds **trims** — non-destructive per-marker
span overrides keyed by the frontend's marker key (``"sheet:12:P01:0"``,
``"screenspace:<event_id>"``, ``"transcript-mark:<mark_id>"``), stored in
video-global seconds *after* the Convergence per-lane offset (the trim is
against the video, so a later offset change doesn't move it). The source
manifests (sheet / Screenspace / Transcripts) are never mutated. Phase 3 adds
**annotations** — text labels, freehand strokes, and rotatable rect/ellipse
shapes with a visibility span, geometry normalized 0..1 to the video frame
(x/points/strokeWidth to width, y/fontSize to height; shape rotation in
degrees clockwise, applied in pixel space) — plus export endpoints that
render annotations with
PIL (transparent RGBA overlay; no ffmpeg ``drawtext``, sidestepping the
Homebrew libfreetype gotcha) and burn them into a screenshot, GIF, or video
span via the ffmpeg ``overlay`` filter with span-relative
``enable='between(t,...)'`` windows. A GIF/video span that straddles a
recording-part boundary is first stitched into a temp clip (t=0 == span start)
via ``pipeline.cut_global_range`` — the same cut/stitch chain Studio's intake
uses — so the overlay pass always sees one continuous input. Exported artifacts
land in the regular ``clipgen_manifest.json`` via ``viewer.save_manifest``.

All Composer state lives in ``composer_manifest.json`` in the output dir
(load-on-startup, save-after-mutations). Composer never writes to the
spreadsheet — cut pairs feed clip generation through Studio's
``/api/generate-intake`` endpoint, which takes raw start/end seconds.

Times are source-video **global seconds**: a multi-part participant (a session
recorded across several files) is addressed on the stitched timeline, and the
frontend maps global time onto parts using the ``parts[].offset`` values from
``GET /api/participants``.

Module-level state (``_sheet_context``, ``_manifest``) is initialized by
:func:`_init_composer_state`, mirroring the other blueprints. The input
directory is not among them: it is read live per request, since it can move
mid-session via ``POST /api/dirs``.
Mutations hold ``_manifest_lock`` and persist via :func:`_persist_locked`.
"""

from __future__ import annotations

import copy
import math
import os
import threading
import uuid
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from flask import Blueprint, request, send_file

import config
import files
import pipeline
import remux_server
import utils
import video
from server_utils import (
    MediaCache,
    clip_media_response,
    err,
    err_no_video,
    find_by_id,
    ok,
    remove_by_id,
)
import itertools

# ---- Module state (initialized by _init_composer_state) ----

_sheet_context: Any = None
_manifest: dict[str, Any] = {}
_manifest_lock = threading.Lock()

# Shortest allowed cut pair. Anything under this is a misclick, not a clip.
MIN_CUT_SECONDS = 0.2

# Composer's read-only marker lanes are the *other* streams, never itself.
# "composer" is a Convergence swim-lane source (Overview), so strip it here to keep
# the Sheet/Screenspace/Transcript lanes unchanged.
_MARKER_SOURCES: tuple[str, ...] = tuple(
    src for src in config.CONVERGENCE_SOURCES if src != "composer"
)

# ---- Blueprint ----

composer_bp = Blueprint("composer", __name__)

utils.register_static_routes(
    composer_bp,
    "composer.html",
    # Per request, not a snapshot — POST /api/dirs moves config.INPUT_DIR
    # mid-session and never re-inits this blueprint. See transcripts_bp.
    media_dir_getter=lambda: str(utils.get_effective_input_dir()),
    media_error="Input directory not configured",
    icons=True,
)

remux_server.register_remux_routes(composer_bp, lambda: _sheet_context)


# ---- Manifest I/O ----


def _empty_manifest() -> dict[str, Any]:
    return {
        "cuts": [],
        "ui": {
            "markerSources": {src: True for src in _MARKER_SOURCES},
            "markerThumbnails": False,
            "markerAudioScrub": False,
            "followPlayhead": True,
        },
    }


def _load_manifest() -> dict[str, Any]:
    data = utils.load_json_manifest(
        config.COMPOSER_MANIFEST_FILENAME, warn_label="the composer manifest"
    )
    return data if isinstance(data, dict) else _empty_manifest()


def _persist_locked() -> None:
    """Write the composer manifest to disk atomically (tmp + ``os.replace`` via
    :func:`utils.save_json_manifest`) so an interrupted write can't truncate the
    file and silently drop the session's cuts/trims/annotations. Caller must hold
    ``_manifest_lock``."""
    utils.save_json_manifest(
        config.COMPOSER_MANIFEST_FILENAME, _manifest, warn_label="composer manifest"
    )


# ---- Participant / video discovery ----


def _participant_parts(video_paths: list[str]) -> list[dict[str, Any]] | None:
    """Probe ordered part durations → ``[{name, duration, offset}]`` or None.

    ``offset`` is the part's cumulative start on the stitched timeline (the
    same math as ``video.build_source_timeline``, reusing the per-file duration
    cache). Returns None when any part's duration cannot be probed — the
    participant is then listed without timeline data and the frontend falls
    back to the <video> element's own metadata (single-part only).
    """
    parts: list[dict[str, Any]] = []
    offset = 0
    for vp in video_paths:
        duration = video.get_file_duration(str(vp))
        if duration is None:
            return None
        parts.append({"name": Path(vp).name, "duration": duration, "offset": offset})
        offset += duration
    return parts


def _participant_duration(participant: str) -> float | None:
    """Total stitched duration for a participant, or None when unknown."""
    for p in files.resolve_participant_videos(_sheet_context):
        if p["id"] == participant and p.get("has_video"):
            parts = _participant_parts(p["video_paths"])
            if parts is None:
                return None
            return float(sum(part["duration"] for part in parts))
    return None


@composer_bp.route("/api/participants")
def api_participants() -> Any:
    """Participants with source videos, plus part timelines for stitched seeks."""
    participants: list[dict[str, Any]] = []
    resolved = [
        p
        for p in files.resolve_participant_videos(_sheet_context)
        if p.get("has_video")
    ]
    # Every entry below probes its parts for durations and its first part for
    # resolution/fps. Serialized that is one ffprobe subprocess after another and
    # the page cannot render until the last one returns — measured 1.06 s for a
    # 24-participant study on local SSD, all of it ffprobe. Probe the whole set up
    # front so the loop reads the caches instead.
    video.prewarm_probes(vp for p in resolved for vp in p["video_paths"])
    for p in resolved:
        parts = _participant_parts(p["video_paths"])
        # Resolution + fps for the subheader readout, probed from the first part —
        # stitched parts share one source setup. probe_video_properties returns None
        # on unprobeable files and only ever reports width/height > 0.
        props = (
            video.probe_video_properties(str(p["video_paths"][0]))
            if p["video_paths"]
            else None
        )
        participants.append(
            {
                "id": p["id"],
                "has_video": True,
                "in_sheet": p.get("in_sheet", False),
                "browser_seekable": p.get("browser_seekable"),
                "parts": parts or [{"name": Path(vp).name} for vp in p["video_paths"]],
                "total_duration": (
                    sum(part["duration"] for part in parts) if parts else None
                ),
                "width": props.get("width") if props else None,
                "height": props.get("height") if props else None,
                "fps": props.get("fps") if props else None,
                "audio_tracks": props.get("audio_tracks") if props else [],
                "audio_track_count": props.get("audio_track_count") if props else 0,
            }
        )
    # ``has_sheet`` gates the off-sheet badge: with no sheet every entry is
    # ``in_sheet: False``, and marking them all would be noise.
    return ok(
        participants=participants,
        has_sheet=_sheet_context is not None,
        config=utils.get_frontend_config(),
    )


# ---- Manifest + cuts CRUD ----


@composer_bp.route("/api/manifest")
def api_manifest() -> Any:
    with _manifest_lock:
        return ok(manifest=copy.deepcopy(_manifest))


def _clamp_span(participant: str, start: float, end: float) -> tuple[float, float]:
    """Clamp a cut span to ``0 <= start < end <= duration`` at MIN_CUT_SECONDS."""
    duration = _participant_duration(participant)
    start = max(0.0, start)
    if duration is not None:
        start = min(start, max(0.0, duration - MIN_CUT_SECONDS))
        end = min(end, duration)
    end = max(end, start + MIN_CUT_SECONDS)
    return round(start, 3), round(end, 3)


@composer_bp.route("/api/cuts", methods=["POST"])
def api_cut_create() -> Any:
    data = request.get_json(silent=True) or {}
    participant = str(data.get("participant", "")).strip()
    if not participant:
        return err("participant is required")
    try:
        start = float(data.get("start", 0))
        end = float(data.get("end", 0))
    except (TypeError, ValueError):
        return err("start/end must be numbers")
    if end <= start:
        return err("end must be after start")
    start, end = _clamp_span(participant, start, end)
    cut: dict[str, Any] = {
        "id": "cut_" + uuid.uuid4().hex[:8],
        "participant": participant,
        "start": start,
        "end": end,
        "label": str(data.get("label", "")),
        "createdAt": datetime.now(UTC).isoformat(),
    }
    with _manifest_lock:
        _manifest.setdefault("cuts", []).append(cut)
        _persist_locked()
    return ok(cut=cut)


@composer_bp.route("/api/cuts/<cut_id>", methods=["PATCH"])
def api_cut_update(cut_id: str) -> Any:
    data = request.get_json(silent=True) or {}
    with _manifest_lock:
        cut = find_by_id(_manifest.get("cuts", []), cut_id)
        if cut is None:
            return err(f"No cut {cut_id}", 404)
        start = cut["start"]
        end = cut["end"]
        try:
            if data.get("start") is not None:
                start = float(data["start"])
            if data.get("end") is not None:
                end = float(data["end"])
        except (TypeError, ValueError):
            return err("start/end must be numbers")
        if end <= start:
            return err("end must be after start")
        cut["start"], cut["end"] = _clamp_span(cut["participant"], start, end)
        if data.get("label") is not None:
            cut["label"] = str(data["label"])
        _persist_locked()
        return ok(cut=copy.deepcopy(cut))


@composer_bp.route("/api/cuts/<cut_id>", methods=["DELETE"])
def api_cut_delete(cut_id: str) -> Any:
    with _manifest_lock:
        if remove_by_id(_manifest.get("cuts", []), cut_id) is None:
            return err(f"No cut {cut_id}", 404)
        _persist_locked()
    return ok()


@composer_bp.route("/api/trims/<path:key>", methods=["PUT"])
def api_trim_put(key: str) -> Any:
    """Set a non-destructive span override for one source marker.

    Optional ``participant``/``label``/``source`` describe the trimmed marker
    for consumers that can't parse the key (Studio's Composer Intake cards).
    Omitted fields keep any previously stored value — undo/redo re-PUTs carry
    times only.
    """
    data = request.get_json(silent=True) or {}
    try:
        start = float(data["start"])
        end = float(data["end"])
    except (KeyError, TypeError, ValueError):
        return err("start and end are required numbers")
    if end < start + MIN_CUT_SECONDS:
        return err("end must be after start")
    with _manifest_lock:
        trims = _manifest.setdefault("trims", {})
        existing = trims.get(key, {})
        trim = {
            "start": round(max(0.0, start), 3),
            "end": round(end, 3),
            "participant": str(
                data.get("participant") or existing.get("participant", "")
            ),
            "label": str(data.get("label") or existing.get("label", "")),
            "source": str(data.get("source") or existing.get("source", "")),
        }
        trims[key] = trim
        _persist_locked()
    return ok(key=key, trim=trim)


@composer_bp.route("/api/trims/<path:key>", methods=["DELETE"])
def api_trim_delete(key: str) -> Any:
    """Reset a marker to its source span by dropping its override."""
    with _manifest_lock:
        trims = _manifest.get("trims", {})
        if key not in trims:
            return err(f"No trim for {key}", 404)
        del trims[key]
        _persist_locked()
    return ok()


@composer_bp.route("/api/ui", methods=["PUT"])
def api_ui_update() -> Any:
    """Persist UI state (lane toggles, fold states, media toggles) so it
    survives a restart.

    Body may carry ``markerSources`` (lane visibility), ``laneFolds`` (True =
    the lane renders as one compact row), and/or the boolean toggles
    ``markerThumbnails`` / ``markerAudioScrub`` / ``followPlayhead``; at least
    one is required.
    """
    data = request.get_json(silent=True) or {}
    sources = data.get("markerSources")
    folds = data.get("laneFolds")
    thumbs = data.get("markerThumbnails")
    scrub = data.get("markerAudioScrub")
    follow = data.get("followPlayhead")
    if (
        not isinstance(sources, dict)
        and not isinstance(folds, dict)
        and not isinstance(thumbs, bool)
        and not isinstance(scrub, bool)
        and not isinstance(follow, bool)
    ):
        return err(
            "markerSources, laneFolds, markerThumbnails, markerAudioScrub, "
            "or followPlayhead is required"
        )
    response: dict[str, Any] = {}
    with _manifest_lock:
        ui = _manifest.setdefault("ui", {})
        if isinstance(sources, dict):
            ui["markerSources"] = {
                src: bool(sources.get(src, True)) for src in _MARKER_SOURCES
            }
            response["markerSources"] = ui["markerSources"]
        if isinstance(folds, dict):
            fold_lanes = (*_MARKER_SOURCES, "annotations")
            ui["laneFolds"] = {lane: bool(folds.get(lane, True)) for lane in fold_lanes}
            response["laneFolds"] = ui["laneFolds"]
        if isinstance(thumbs, bool):
            ui["markerThumbnails"] = thumbs
            response["markerThumbnails"] = thumbs
        if isinstance(scrub, bool):
            ui["markerAudioScrub"] = scrub
            response["markerAudioScrub"] = scrub
        if isinstance(follow, bool):
            ui["followPlayhead"] = follow
            response["followPlayhead"] = follow
        _persist_locked()
    return ok(**response)


# ---- Annotations CRUD ----

ANNOTATION_TYPES = ("text", "freehand", "shape")
ANNOTATION_SHAPES = ("rect", "ellipse")
ANNOTATION_STROKE_STYLES = ("solid", "dashed", "dotted")


def _clamp01(value: Any) -> float:
    return min(1.0, max(0.0, float(value)))


def _sanitize_annotation_geometry(
    ann_type: str, geometry: Any
) -> dict[str, Any] | None:
    """Validate + normalize geometry for one annotation; None = invalid."""
    if not isinstance(geometry, dict):
        return None
    try:
        if ann_type == "text":
            text = str(geometry.get("text") or "").strip()
            if not text:
                return None
            return {
                "x": round(_clamp01(geometry.get("x", 0)), 4),
                "y": round(_clamp01(geometry.get("y", 0)), 4),
                "text": text,
            }
        if ann_type == "shape":
            # x/y = normalized center, w/h = normalized size, rotation in
            # degrees clockwise around the center (matches ctx.rotate).
            kind = str(geometry.get("shape") or "")
            if kind not in ANNOTATION_SHAPES:
                return None
            w = float(geometry.get("w", 0))
            h = float(geometry.get("h", 0))
            if w <= 0 or h <= 0:
                return None
            return {
                "shape": kind,
                "x": round(_clamp01(geometry.get("x", 0)), 4),
                "y": round(_clamp01(geometry.get("y", 0)), 4),
                "w": max(round(min(w, 1.0), 4), 0.001),
                "h": max(round(min(h, 1.0), 4), 0.001),
                "rotation": round(float(geometry.get("rotation", 0) or 0.0) % 360.0, 2),
            }
        points = geometry.get("points")
        if not isinstance(points, list) or not points:
            return None
        cleaned = [
            [round(_clamp01(p[0]), 4), round(_clamp01(p[1]), 4)]
            for p in points
            if isinstance(p, (list, tuple)) and len(p) >= 2
        ]
        return {"points": cleaned} if cleaned else None
    except (TypeError, ValueError):
        return None


def _sanitize_annotation_style(style: Any) -> dict[str, Any]:
    """Fill missing style fields from the config defaults."""
    style = style if isinstance(style, dict) else {}

    def _num(key: str, default: float) -> float:
        try:
            value = float(style.get(key, default))
        except (TypeError, ValueError):
            return default
        return value if value > 0 else default

    stroke_style = str(style.get("strokeStyle") or "").lower()
    if stroke_style not in ANNOTATION_STROKE_STYLES:
        stroke_style = config.COMPOSER_ANNOTATION_STROKE_STYLE

    return {
        "color": str(style.get("color") or config.COMPOSER_ANNOTATION_COLOR),
        "strokeWidth": _num("strokeWidth", config.COMPOSER_ANNOTATION_STROKE_WIDTH),
        "strokeStyle": stroke_style,
        "fontSize": _num("fontSize", config.COMPOSER_ANNOTATION_FONT_SIZE),
    }


def _parse_annotation_span(span: Any) -> tuple[float, float] | None:
    if not isinstance(span, dict):
        return None
    try:
        start = max(0.0, float(span["start"]))
        end = float(span["end"])
    except (KeyError, TypeError, ValueError):
        return None
    if end < start + MIN_CUT_SECONDS:
        return None
    return round(start, 3), round(end, 3)


@composer_bp.route("/api/annotations", methods=["POST"])
def api_annotation_create() -> Any:
    data = request.get_json(silent=True) or {}
    participant = str(data.get("participant", "")).strip()
    if not participant:
        return err("participant is required")
    ann_type = str(data.get("type", ""))
    if ann_type not in ANNOTATION_TYPES:
        return err(f"type must be one of {', '.join(ANNOTATION_TYPES)}")
    span = _parse_annotation_span(data.get("span"))
    if span is None:
        return err("span with start < end is required")
    geometry = _sanitize_annotation_geometry(ann_type, data.get("geometry"))
    if geometry is None:
        return err("invalid geometry for type " + ann_type)
    annotation: dict[str, Any] = {
        "id": "ann_" + uuid.uuid4().hex[:8],
        "participant": participant,
        "type": ann_type,
        "span": {"start": span[0], "end": span[1]},
        "geometry": geometry,
        "style": _sanitize_annotation_style(data.get("style")),
        "createdAt": datetime.now(UTC).isoformat(),
    }
    with _manifest_lock:
        _manifest.setdefault("annotations", []).append(annotation)
        _persist_locked()
    return ok(annotation=annotation)


@composer_bp.route("/api/annotations/<ann_id>", methods=["PATCH"])
def api_annotation_update(ann_id: str) -> Any:
    data = request.get_json(silent=True) or {}
    with _manifest_lock:
        ann = find_by_id(_manifest.get("annotations", []), ann_id)
        if ann is None:
            return err(f"No annotation {ann_id}", 404)
        if data.get("span") is not None:
            span = _parse_annotation_span(data["span"])
            if span is None:
                return err("span with start < end is required")
            ann["span"] = {"start": span[0], "end": span[1]}
        if data.get("geometry") is not None:
            geometry = _sanitize_annotation_geometry(ann["type"], data["geometry"])
            if geometry is None:
                return err("invalid geometry for type " + str(ann["type"]))
            ann["geometry"] = geometry
        if data.get("style") is not None:
            ann["style"] = _sanitize_annotation_style(data["style"])
        _persist_locked()
        return ok(annotation=copy.deepcopy(ann))


@composer_bp.route("/api/annotations/<ann_id>", methods=["DELETE"])
def api_annotation_delete(ann_id: str) -> Any:
    with _manifest_lock:
        if remove_by_id(_manifest.get("annotations", []), ann_id) is None:
            return err(f"No annotation {ann_id}", 404)
        _persist_locked()
    return ok()


# ---- Annotation rendering (PIL; no ffmpeg drawtext) ----

# Common system font locations, probed in order. load_default() is the
# fixed-size last resort (visibly cruder, but never fails).
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


def _render_annotation_overlay(
    annotations: list[dict[str, Any]], width: int, height: int
) -> Any:
    """Render annotations onto a transparent RGBA PIL image of the frame size.

    Geometry is normalized 0..1 (x/points/strokeWidth to width, y/fontSize to
    height) — the same convention the browser preview canvas uses, so burn-in
    matches the live view at any resolution.
    """
    from PIL import Image, ImageDraw

    img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    for ann in annotations:
        style = ann.get("style") or {}
        color = _parse_hex_color(style.get("color", ""))
        geometry = ann.get("geometry") or {}
        if ann.get("type") == "freehand":
            points = [
                (float(p[0]) * width, float(p[1]) * height)
                for p in geometry.get("points", [])
            ]
            stroke = max(
                1,
                round(
                    float(
                        style.get(
                            "strokeWidth", config.COMPOSER_ANNOTATION_STROKE_WIDTH
                        )
                    )
                    * width
                ),
            )
            stroke_style = str(style.get("strokeStyle") or "solid")
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
            stroke = max(
                1,
                round(
                    float(
                        style.get(
                            "strokeWidth", config.COMPOSER_ANNOTATION_STROKE_WIDTH
                        )
                    )
                    * width
                ),
            )
            stroke_style = str(style.get("strokeStyle") or "solid")
            if geometry.get("shape") == "rect":
                # Rotate the corners with the same transform the browser's
                # ctx.rotate applies (y-down: positive = clockwise), so the
                # burn-in matches the preview exactly. Closing through the
                # second corner again keeps the start joint filled.
                rad = math.radians(rotation)
                cos_r, sin_r = math.cos(rad), math.sin(rad)
                corners = [
                    (cx + dx * cos_r - dy * sin_r, cy + dx * sin_r + dy * cos_r)
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
                # Dashed/dotted ellipse: sample the (rotated) perimeter as a
                # polygon and dash-walk it — PIL can't dash an ellipse outline,
                # and this also handles rotation without the composite trick.
                a, b = sw / 2, sh / 2
                perim = math.pi * (
                    3 * (a + b) - math.sqrt(max(0.0, (3 * a + b) * (a + 3 * b)))
                )
                n = max(48, int(perim / max(1.0, stroke * 2)))
                rad = math.radians(rotation)
                cos_r, sin_r = math.cos(rad), math.sin(rad)
                poly = []
                for k in range(n + 1):
                    th = 2 * math.pi * k / n
                    ex, ey = a * math.cos(th), b * math.sin(th)
                    poly.append(
                        (cx + ex * cos_r - ey * sin_r, cy + ex * sin_r + ey * cos_r)
                    )
                _dash_polyline(draw, poly, color, stroke, stroke_style)
            else:  # solid ellipse
                box = [cx - sw / 2, cy - sh / 2, cx + sw / 2, cy + sh / 2]
                if abs(rotation) < 0.01:
                    draw.ellipse(box, outline=color, width=stroke)
                else:
                    # PIL can't stroke a rotated ellipse directly: draw it
                    # axis-aligned on a temp layer and rotate that around the
                    # center (Image.rotate's positive angle is counter-
                    # clockwise, hence the negation for clockwise rotation).
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


# ---- Exports (annotated screenshot / burned video / GIF) ----

# One export at a time is plenty for a single-user tool; the cancel event
# terminates the in-flight ffmpeg via run_ffmpeg_process's cancel_flag.
# _export_busy enforces the single-export assumption the shared event relies
# on: without it, a second export's clear() would un-cancel the first, and one
# cancel POST would abort both in-flight encodes.
_export_cancel = threading.Event()
_export_busy = threading.Lock()

# Bound on the ffmpeg filter graph: one overlay input per visibility window.
MAX_OVERLAY_WINDOWS = 20


def _find_participant_parts(participant: str) -> list[dict[str, Any]] | None:
    """Ordered ``{name, path, duration, offset}`` parts, or None."""
    p = files.find_participant_record(_sheet_context, participant)
    if p is None or not p.get("has_video"):
        return None
    parts = _participant_parts(p["video_paths"])
    if parts is None:
        return None
    for part, vp in zip(parts, p["video_paths"]):
        part["path"] = str(vp)
    return parts


def _part_for_time(parts: list[dict[str, Any]], t: float) -> dict[str, Any]:
    for part in parts:
        if part["offset"] <= t < part["offset"] + part["duration"]:
            return part
    return parts[-1]


# ---- Scrubber media (sprite sheets / audio snippets) ----

# Backs the marker/cut thumbnail strips + hover audio scrub on the timeline.
# Same LRU + single-flight pattern as Studio's /api/sprite and /api/clip-audio,
# but resolved against Composer's part model and with no spreadsheet dependency.
_sprite_cache = MediaCache(128)
_audio_cache = MediaCache(32)


def _resolve_scrub_source(
    participant: str, start_sec: float, duration: float
) -> tuple[str, float, float] | None:
    """Resolve ``(video_path, local_start, clamped_duration)`` for a global span.

    A span straddling a part boundary samples only the owning part's tail
    (hover preview only — the same tradeoff as Studio's sprite route). Returns
    ``None`` when the participant has no probeable videos.
    """
    parts = _find_participant_parts(participant)
    if not parts:
        return None
    part = _part_for_time(parts, start_sec)
    local_start = max(0.0, start_sec - part["offset"])
    clamped = min(duration, max(0.1, part["duration"] - local_start))
    return part["path"], local_start, clamped


@composer_bp.route("/api/sprite/<participant>")
def api_sprite(participant: str) -> Any:
    """Tiled JPEG sprite sheet for one thumbnail-strip tile.

    The client requests one sprite per zoom-density tile (a power-of-two
    slot-seconds ladder), not per marker span, so zooming in refines the strip
    with additional frames instead of stretching the same ones.
    """
    cols = config.STUDIO_SCRUBBER_SPRITE_COLS
    rows = config.STUDIO_SCRUBBER_SPRITE_ROWS
    return clip_media_response(
        cache=_sprite_cache,
        resolve=lambda start, dur: _resolve_scrub_source(participant, start, dur),
        # seek_frames: per-frame fast seeks keep the cost O(frames), not
        # O(span) — timeline tiles can cover minutes of source video.
        produce=lambda path, local_start, dur: video.extract_sprite_sheet_bytes(
            path, local_start, dur, cols, rows, seek_frames=True
        ),
        mimetype="image/jpeg",
        kind_label="Sprite",
        key_extras=(cols, rows),
        invalid_message="Invalid time range",
    )


@composer_bp.route("/api/audio/<participant>")
def api_audio(participant: str) -> Any:
    """Mono PCM WAV of a marker/cut span for the hover audio scrub.

    Capped at ``config.COMPOSER_SCRUB_MAX_AUDIO_SECONDS`` — markers can span
    minutes, and the WAV is decoded whole in the browser's WebAudio context.
    """
    return clip_media_response(
        cache=_audio_cache,
        resolve=lambda start, dur: _resolve_scrub_source(
            participant, start, min(dur, config.COMPOSER_SCRUB_MAX_AUDIO_SECONDS)
        ),
        produce=video.extract_audio_segment_bytes,
        mimetype="audio/wav",
        kind_label="Audio",
        invalid_message="Invalid time range",
    )


@composer_bp.route("/api/audio-track/<participant>/<int:idx>")
def api_audio_track(participant: str, idx: int) -> Any:
    """Stream one demuxed audio track for the browser's per-track volume mixer."""
    parts = _find_participant_parts(participant)
    if not parts:
        return err_no_video(participant)
    out = video.extract_audio_track(parts[0]["path"], idx)
    if out is None:
        return err("Could not extract audio track", 500)
    response = send_file(str(out), mimetype="audio/mp4", conditional=True)
    response.headers["Cache-Control"] = "no-cache"
    return response


def _unlink_quiet(path: str | None) -> None:
    """Best-effort delete of a temp file; no-op on None or OSError."""
    if not path:
        return
    try:
        os.unlink(path)
    except OSError:
        pass


def _annotations_in_span(
    participant: str, start: float, end: float
) -> list[dict[str, Any]]:
    with _manifest_lock:
        annotations = copy.deepcopy(_manifest.get("annotations", []))
    return [
        a
        for a in annotations
        if a.get("participant") == participant
        and a["span"]["start"] < end
        and a["span"]["end"] > start
    ]


def _annotation_windows(
    annotations: list[dict[str, Any]], start: float, end: float
) -> list[dict[str, Any]]:
    """Split [start, end] into visibility windows with a constant annotation set.

    Window boundaries are the annotation span edges clipped to the export
    range; windows with no visible annotation are dropped. Each window is
    flattened into ONE overlay PNG downstream, so the ffmpeg filter graph is
    bounded by the number of distinct windows, not the annotation count.
    """
    bounds = {start, end}
    for a in annotations:
        bounds.add(min(end, max(start, a["span"]["start"])))
        bounds.add(min(end, max(start, a["span"]["end"])))
    ordered = sorted(bounds)
    windows: list[dict[str, Any]] = []
    for w_start, w_end in itertools.pairwise(ordered):
        if w_end - w_start < 0.01:
            continue
        visible = [
            a
            for a in annotations
            if a["span"]["start"] < w_end and a["span"]["end"] > w_start
        ]
        if visible:
            windows.append({"start": w_start, "end": w_end, "annotations": visible})
    return windows


def _build_overlay_command(
    input_path: str,
    local_start: float,
    duration: float,
    overlay_specs: list[tuple[str, float, float]],
    out_path: str,
    *,
    gif: bool,
    encoder: str = "libx264",
) -> list[str]:
    """ffmpeg argv burning overlay PNGs into a span of *input_path*.

    *overlay_specs* is ``[(png_path, rel_start, rel_end), ...]`` with times
    relative to the span start — input seeking (``-ss`` before ``-i``) resets
    the filter clock to 0, so ``enable='between(t,...)'`` uses span-relative
    times. Only the requested span is decoded/encoded, never the whole file.

    *encoder* selects the H.264 encoder for video output (see
    ``video.resolve_video_encoder``); gif output has none to pick.
    """
    cmd = [
        "ffmpeg",
        "-y",
        "-loglevel",
        config.FFMPEG_LOGLEVEL,
        "-ss",
        f"{max(0.0, local_start):.3f}",
        "-i",
        input_path,
    ]
    for png_path, _, _ in overlay_specs:
        cmd += ["-loop", "1", "-i", png_path]
    chain = []
    prev = "[0:v]"
    for i, (_, rel_start, rel_end) in enumerate(overlay_specs):
        label = f"[v{i + 1}]"
        chain.append(
            f"{prev}[{i + 1}:v]overlay=0:0:"
            f"enable='between(t,{rel_start:.3f},{rel_end:.3f})'{label}"
        )
        prev = label
    if gif:
        chain.append(
            f"{prev}fps={config.GIF_FPS},"
            f"scale={config.GIF_SCALE_WIDTH}:-1:flags=lanczos[vout]"
        )
        prev = "[vout]"
    cmd += ["-filter_complex", ";".join(chain), "-map", prev]
    if gif:
        cmd += ["-t", f"{duration:.3f}", out_path]
    else:
        cmd += ["-map", "0:a?"]
        # veryfast: spans are short and re-encoded once; quality over speed
        # tuning is not worth a config knob here.
        cmd += video.video_encoder_args(encoder, crf=20, preset="veryfast")
        cmd += [
            "-pix_fmt",
            "yuv420p",
            "-c:a",
            "copy",
            "-t",
            f"{duration:.3f}",
            out_path,
        ]
    return cmd


def _composer_artifact(
    participant: str,
    artifact_type: str,
    out_path: str,
    start: float,
    end: float,
    description: str,
    source_video: str,
) -> dict[str, Any]:
    """Artifact record for the clipgen manifest (intake-record shape)."""
    import hashlib

    id_hash = hashlib.md5(
        f"composer|{participant}|{start}|{end}|{Path(out_path).name}".encode()
    ).hexdigest()[:8]
    return {
        "id": f"composer_{id_hash}_s0",
        "type": artifact_type,
        "file": Path(out_path).name,
        "start": round(start, 3),
        "end": round(end, 3),
        "thumbnail": "",
        "study": "",
        "participant": participant,
        "category": "",
        "severity": "",
        "description": description,
        "cellRow": None,
        "cellCol": None,
        "cellA1": "",
        "annotations": [],
        "source": "composer",
        "sourceVideo": source_video,
        "localStart": round(start, 3),
        "localEnd": round(end, 3),
    }


def _save_export_artifact(artifact: dict[str, Any], participant: str) -> None:
    import viewer

    viewer.save_manifest([artifact], participant=participant)


@composer_bp.route("/api/export/cancel", methods=["POST"])
def api_export_cancel() -> Any:
    _export_cancel.set()
    return ok()


@composer_bp.route("/api/export/screenshot", methods=["POST"])
def api_export_screenshot() -> Any:
    """Annotated screenshot at one timestamp (PIL composite; no ffmpeg filter)."""
    data = request.get_json(silent=True) or {}
    participant = str(data.get("participant", "")).strip()
    try:
        at_time = float(data.get("time", 0))
    except (TypeError, ValueError):
        return err("time must be a number")
    parts = _find_participant_parts(participant)
    if not parts:
        return err(f"No video for {participant}", 404)
    part = _part_for_time(parts, at_time)
    frame = video.extract_frame_at_timestamp(part["path"], at_time - part["offset"])
    if frame is None:
        return err("Could not extract a frame at that time", 500)

    annotations = _annotations_in_span(participant, at_time, at_time + 0.001)
    from PIL import Image

    base = Image.fromarray(frame[:, :, ::-1]).convert("RGBA")  # BGR → RGB
    overlay = _render_annotation_overlay(annotations, base.width, base.height)
    composed = Image.alpha_composite(base, overlay).convert("RGB")

    time_tag = utils.seconds_to_timestamp(int(at_time)).replace(":", ".")
    out_path = files.get_unique_filename(
        f"{participant} annotated {time_tag}{config.SCREENSHOT_FORMAT}",
        config.SCREENSHOT_FORMAT,
    )
    try:
        composed.save(out_path)
    except OSError as exc:
        files.release_reservation(out_path)
        return err(f"Could not write screenshot: {exc}", 500)

    artifact = _composer_artifact(
        participant,
        "screen",
        out_path,
        at_time,
        at_time,
        f"Annotated screenshot ({len(annotations)} annotation(s))",
        Path(part["path"]).name,
    )
    _save_export_artifact(artifact, participant)
    return ok(artifact=artifact)


def _run_overlay_export(data: dict[str, Any], *, gif: bool) -> Any:
    """Shared burn/GIF export: validate span, render windows, run ffmpeg."""
    participant = str(data.get("participant", "")).strip()
    try:
        start = float(data.get("start", 0))
        end = float(data.get("end", 0))
    except (TypeError, ValueError):
        return err("start/end must be numbers")
    if end <= start:
        return err("end must be after start")
    parts = _find_participant_parts(participant)
    if not parts:
        return err(f"No video for {participant}", 404)
    annotations = _annotations_in_span(participant, start, end)
    if not annotations:
        return err("No annotations in this span — use Generate for plain clips.")
    windows = _annotation_windows(annotations, start, end)
    # An annotation can overlap the span by less than the 0.01 s a window needs
    # to survive, which clears the "no annotations" check above and still leaves
    # nothing to burn. Without this the filter graph would come out empty and
    # ffmpeg would fail on a `[0:v]` label the graph never defines.
    if not windows:
        return err(
            "No annotation is visible for long enough in this span — widen the "
            "span or the annotation's visibility."
        )
    if len(windows) > MAX_OVERLAY_WINDOWS:
        return err(
            f"Too many distinct annotation windows ({len(windows)}); "
            f"the limit is {MAX_OVERLAY_WINDOWS}."
        )

    import tempfile

    out_dir = str(utils.get_effective_output_dir())
    _export_cancel.clear()

    # Resolve the source the overlay pass decodes. A within-part span decodes the
    # owning part directly with a local seek; a span that straddles a part boundary
    # is first stitched into a temp clip (t=0 == span start) via the same cut chain
    # Studio's intake uses, so the overlay filter sees one continuous input.
    video_paths = [p["path"] for p in parts]
    timeline = video.timeline_or_none(video_paths)
    pieces = (
        utils.map_global_range_to_segments(timeline, start, end)
        if timeline is not None
        else None
    )
    if pieces is not None and not pieces:
        return err("The span is outside the recording", 400)

    stitch_tmp: str | None = None
    if pieces is not None and len(pieces) > 1:
        assert timeline is not None  # a non-None pieces implies a real timeline
        fd, stitch_tmp = tempfile.mkstemp(
            prefix=config.TEMP_ARTIFACT_PREFIX, suffix=config.FILEFORMAT, dir=out_dir
        )
        os.close(fd)
        stitched = pipeline.cut_global_range(
            timeline,
            video_paths[0],
            start,
            end,
            stitch_tmp,
            reencode=config.REENCODING,
            cancel_flag=_export_cancel.is_set,
        )
        if stitched is None:
            _unlink_quiet(stitch_tmp)
            return err(
                "Export cancelled" if _export_cancel.is_set() else "ffmpeg failed"
            )
        overlay_input = stitch_tmp
        overlay_local_start = 0.0
        source_video_name = stitched["sourceVideo"]
    elif pieces:
        assert timeline is not None  # a non-None pieces implies a real timeline
        seg_index, local_start, _ = pieces[0]
        overlay_input = timeline[seg_index][0]
        overlay_local_start = local_start
        source_video_name = Path(overlay_input).name
    else:
        start_part = _part_for_time(parts, start)
        overlay_input = start_part["path"]
        overlay_local_start = start - start_part["offset"]
        source_video_name = Path(overlay_input).name

    # Parts can differ in resolution, so render overlays at the actual input's dims.
    props = video.probe_video_properties(overlay_input)
    if not props or not props.get("width") or not props.get("height"):
        _unlink_quiet(stitch_tmp)
        return err("Could not probe the video resolution", 500)

    fmt = config.GIF_FORMAT if gif else config.FILEFORMAT
    kind = "gif" if gif else "clip"
    time_tag = utils.seconds_to_timestamp(int(start)).replace(":", ".")
    out_path = files.get_unique_filename(
        f"{participant} annotated {kind} {time_tag}{fmt}", fmt
    )
    overlay_specs: list[tuple[str, float, float]] = []
    try:
        for window in windows:
            overlay = _render_annotation_overlay(
                window["annotations"], props["width"], props["height"]
            )
            fd, png_path = tempfile.mkstemp(
                prefix=config.TEMP_ARTIFACT_PREFIX, suffix=".png", dir=out_dir
            )
            os.close(fd)
            # Track before save so a failed overlay.save still gets cleaned up.
            overlay_specs.append(
                (png_path, window["start"] - start, window["end"] - start)
            )
            overlay.save(png_path)

        def build_command(encoder: str) -> list[str]:
            return _build_overlay_command(
                overlay_input,
                overlay_local_start,
                end - start,
                overlay_specs,
                out_path,
                gif=gif,
                encoder=encoder,
            )

        # gif output has no H.264 encoder to choose, so it never probes for one.
        result = video.run_ffmpeg_encode(
            build_command,
            encoder="libx264" if gif else video.resolve_video_encoder(),
            input_file=overlay_input,
            output_file=out_path,
            os_error_message="Annotated export failed.",
            cancel_flag=_export_cancel.is_set,
        )
    except Exception:
        # A failed mkstemp / overlay.save would otherwise leave the placeholder
        # get_unique_filename reserved above sitting on disk; the two explicit
        # ffmpeg failure paths below release it themselves.
        files.release_reservation(out_path)
        raise
    finally:
        for png_path, _, _ in overlay_specs:
            _unlink_quiet(png_path)
        _unlink_quiet(stitch_tmp)
    if result is None:
        files.release_reservation(out_path)
        return err("Export cancelled" if _export_cancel.is_set() else "ffmpeg failed")
    if result.returncode != 0:
        files.release_reservation(out_path)
        return err("ffmpeg failed: " + (result.stderr or "")[-400:], 500)

    artifact = _composer_artifact(
        participant,
        kind,
        out_path,
        start,
        end,
        f"Annotated {kind} ({len(annotations)} annotation(s))",
        source_video_name,
    )
    _save_export_artifact(artifact, participant)
    return ok(artifact=artifact)


@composer_bp.route("/api/export/burn", methods=["POST"])
def api_export_burn() -> Any:
    """Burn annotations into a video span (seek-first; span-only encode)."""
    if not _export_busy.acquire(blocking=False):
        return err("An export is already running", 409)
    try:
        return _run_overlay_export(request.get_json(silent=True) or {}, gif=False)
    finally:
        _export_busy.release()


@composer_bp.route("/api/export/gif", methods=["POST"])
def api_export_gif() -> Any:
    """Burn annotations into an animated GIF of the span."""
    if not _export_busy.acquire(blocking=False):
        return err("An export is already running", 409)
    try:
        return _run_overlay_export(request.get_json(silent=True) or {}, gif=True)
    finally:
        _export_busy.release()


# ---- State init ----


def _init_composer_state(sheet_context: Any = None) -> None:
    """Initialize module-level state for Composer routes.

    Participants are resolved from ``_sheet_context`` + the input dir on demand
    (:func:`files.resolve_participant_videos`), so nothing is snapshotted here.

    Called once, from ``build_combined_app``. A worksheet swap goes through
    :func:`repin_sheet_state` instead — re-running this would reload the
    manifest and could drop a write still sitting in the persist debounce.
    """
    global _sheet_context, _manifest

    _sheet_context = sheet_context
    _manifest = _load_manifest()


def repin_sheet_state(sheet_context: Any = None) -> None:
    """Point the blueprint at a newly opened (or closed) worksheet.

    The sheet-only half of :func:`_init_composer_state`, called by
    ``server._swap_worksheet``. Without it Composer kept the sheet the
    *process* started with — normally none — so opening a spreadsheet from the
    Start overlay left its participant list on bare disk discovery, missing
    every sheet-only column and every filename override.
    """
    global _sheet_context

    _sheet_context = sheet_context
