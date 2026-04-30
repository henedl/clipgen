# -*- coding: utf-8 -*-
"""Screenspace Flask blueprint for clipgen.

Registered at /screenspace/ by start_combined_server(). Works with or without a
spreadsheet; auto-discovers participant videos from the input directory.
Module-level state: _manifest, _worker, _output_dir, _participants (initialized by
_init_screenspace_state()).

API endpoints (all under /screenspace/):
  GET  /media/<filename>                    – serve artifact media files
  GET  /api/participants                    – list discovered participant videos
  GET  /api/video/frame/<participant>/<ts>  – extract a JPEG frame at timestamp
  GET|POST /api/preview/<participant>/<ts>   – PNG of the active tool's preprocessed view (?layer=<id> for single-layer overlay)
  GET  /api/preview/layers                   – JSON catalog of overlay-eligible layers per tool
  GET  /api/video/info/<participant>        – video metadata (duration, resolution, fps)
  GET  /api/video/stream/<participant>     – stream source video (mp4, range-aware)
  GET  /api/regions                         – list regions
  POST /api/regions                         – create or update a region
  DELETE /api/regions/<name>               – delete a region
  GET/POST /api/stashes                    – stash CRUD (save/restore named region sets)
  PUT  /api/stashes/<id>                   – update stash
  DELETE /api/stashes/<id>                 – delete stash
  POST /api/stashes/<id>/restore           – restore a stash
  GET  /api/tasks                          – list task queue
  GET  /api/tasks/<task_id>               – get single task
  POST /api/tasks                          – create and enqueue a new task
  DELETE /api/tasks/<task_id>             – cancel/remove a task
  PUT  /api/tasks/reorder                  – reorder task queue by priority
  POST /api/tasks/pause                    – pause the worker
  POST /api/tasks/resume                   – resume the worker
  GET  /api/tasks/<task_id>/results        – analysis results (timestamps, artifact paths)
  GET  /api/events                         – list result events across all tasks
  GET  /api/export/events                  – export events as analysis-ready JSON (default) or CSV (?format=csv)
  PUT  /api/events/<event_id>/exclude      – mark an event excluded
  PUT  /api/events/<event_id>/include      – mark an event included
  PUT  /api/events/bulk-exclude            – bulk exclude events by task/time range
  PUT  /api/events/bulk-include            – bulk include events by task/time range
"""

import binascii
import copy
import json
import math
import queue as queue_mod
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, TypeGuard, cast

from flask import Blueprint, Response, jsonify, request, send_file

import config
import files
import screenspace
import utils
import video

FlaskResponse = Response | tuple[Response, int]

_VALID_TASK_TYPES = (
    "multitool",
    "color",
    "change",
    "similarity",
    "text",
    "numbers",
    "timelapse",
    "template",
    "flow",
    "scene",
    "inactivity",
)
_VALID_STEP_TYPES = (
    "color",
    "change",
    "similarity",
    "text",
    "numbers",
    "template",
    "flow",
    "scene",
)
_TASK_BINARY_KEYS = (
    "reference_frame",
    "template_image",
    "template_mask",
    "reference_scenes",
)


def _template_bgr_and_mask_from_b64(upload_b64: str) -> tuple[Any, Any]:
    """Decode a base64-encoded image file into a BGR template and optional uint8 mask.

    RGBA inputs yield ``(bgr, alpha_mask)``; RGB/gray yield ``(bgr, None)``.

    Raises:
        ValueError: invalid base64 or undecodable image bytes.
    """
    import base64

    import cv2
    import numpy as np

    try:
        img_bytes = base64.b64decode(upload_b64)
    except (ValueError, binascii.Error) as exc:
        raise ValueError("Could not decode uploaded image") from exc
    img_arr = np.frombuffer(img_bytes, dtype=np.uint8)
    img = cv2.imdecode(img_arr, cv2.IMREAD_UNCHANGED)
    if img is None:
        raise ValueError("Invalid image data")
    if len(img.shape) == 2:
        return cv2.cvtColor(img, cv2.COLOR_GRAY2BGR), None
    if img.shape[2] == 4:
        return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR), img[:, :, 3]
    return img, None


# ---- Module-level state (set once by _init_screenspace_state) ----

_manifest: dict[str, Any] = {}
_worker: "screenspace.ScreenspaceWorker | None" = None
_output_dir: str = ""
_participants: list[dict[str, Any]] = []
_video_metadata_cache: dict[str, dict[str, Any]] = {}
_frame_cache: dict[tuple[str, float, int], bytes] = {}

# ---- SSE (Server-Sent Events) client registry ----

_sse_clients: list[queue_mod.Queue[str]] = []
_sse_clients_lock = threading.Lock()
_manifest_lock = threading.Lock()


def _notify_sse_clients(event_type: str = "update") -> None:
    """Push a notification to all connected SSE clients."""
    with _sse_clients_lock:
        for q in _sse_clients:
            try:
                q.put_nowait(event_type)
            except queue_mod.Full:
                pass


def _sse_task_payload() -> str:
    """Build an SSE data line with current task state."""
    tasks = _worker.get_all_tasks() if _worker else []
    clean = [_clean_task(t) for t in tasks]
    paused = _worker.is_paused if _worker else False
    alive = _worker.is_alive if _worker else False
    data = json.dumps(
        {"ok": True, "tasks": clean, "paused": paused, "worker_alive": alive}
    )
    return f"data: {data}\n\n"


# ---- Blueprint ----

screenspace_bp = Blueprint("screenspace", __name__)

utils.register_static_routes(
    screenspace_bp,
    "screenspace.html",
    media_dir_getter=lambda: _output_dir,
    media_error="Output directory not configured",
    icons=True,
)


# ---- Participants ----


@screenspace_bp.route("/api/participants")
def api_participants() -> FlaskResponse:
    """List participants with source video availability."""
    return jsonify({"ok": True, "participants": _participants})


# ---- Video frame extraction ----


def _find_participant_video(participant_id: str) -> str | None:
    """Resolve the video path for a participant."""
    for p in _participants:
        if p["id"] == participant_id:
            if p["has_video"]:
                return p["video_path"]
            return None
    return None


@screenspace_bp.route("/api/video/frame/<participant>/<timestamp>")
def api_video_frame(participant: str, timestamp: str) -> FlaskResponse:
    """Extract and return a single JPEG frame at the given timestamp.

    Optional query parameter ``w`` requests a scaled-down thumbnail
    (e.g. ``?w=200`` for a 200 px-wide JPEG).  Without ``w`` the frame
    is returned at full resolution.  Results are cached in-memory.
    """
    try:
        ts = float(timestamp)
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "Invalid timestamp"}), 400

    video_path = _find_participant_video(participant)
    if video_path is None:
        return jsonify(
            {"ok": False, "error": f"No video for participant {participant}"}
        ), 404

    width = request.args.get("w", 0, type=int)
    cache_key = (video_path, round(ts, 2), width)
    cached = _frame_cache.get(cache_key)
    if cached is not None:
        return Response(
            cached,
            mimetype="image/jpeg",
            headers={"Cache-Control": "public, max-age=86400"},
        )

    if width > 0:
        jpeg_bytes = video.extract_thumbnail_bytes(video_path, int(ts), width=width)
    else:
        frame = video.extract_frame_at_timestamp(video_path, ts)
        if frame is None:
            return jsonify(
                {"ok": False, "error": "Could not read frame at timestamp"}
            ), 400
        import cv2

        _, jpeg = cv2.imencode(".jpg", frame, [cv2.IMWRITE_JPEG_QUALITY, 85])
        jpeg_bytes = jpeg.tobytes()

    if jpeg_bytes is None:
        return jsonify({"ok": False, "error": "Could not extract frame"}), 400

    _frame_cache[cache_key] = jpeg_bytes
    return Response(
        jpeg_bytes,
        mimetype="image/jpeg",
        headers={"Cache-Control": "public, max-age=86400"},
    )


@screenspace_bp.route("/api/preview/<participant>/<timestamp>", methods=["GET", "POST"])
def api_preview(participant: str, timestamp: str) -> FlaskResponse:
    """Render what the selected tool's CV pipeline sees at ``timestamp``.

    Returns a PNG composite image (grayscale crop, diff mask, edge map, flow
    vectors, pHash bit grid, etc.) tailored to the active tool.  Optional query
    params:

      tool=<name>          one of the 11 screenspace tool types
      region=x,y,w,h       normalized 0–1 region coordinates (required for most tools)
      prev=<seconds>       prior timestamp for change/flow (defaults to ts-1s)
      ref=<seconds>        reference timestamp for similarity (region crop) or template
                           capture preview (same as ``reference_timestamp`` in tasks)
      noise=<int>          change tool's noise_threshold override
      h,s,v=<int>          color tool's target HSV override
      magnitude=<float>    flow tool's magnitude threshold override
      layer=<id>           if set, return that single overlay layer at native
                           region/frame resolution instead of the labeled
                           composite. See ``screenspace_preview.OVERLAY_LAYERS``
                           for valid (tool, layer) pairs.

    For **template** with an **uploaded** PNG, send ``POST`` with JSON body
    ``{"template_image_data": "<base64>"}`` (same field as task enqueue); query
    string still supplies ``tool``, ``region`` (optional), and ``_`` cache-bust.
    """
    import screenspace_preview

    try:
        ts = float(timestamp)
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "Invalid timestamp"}), 400

    video_path = _find_participant_video(participant)
    if video_path is None:
        return jsonify(
            {"ok": False, "error": f"No video for participant {participant}"}
        ), 404

    tool = (request.args.get("tool") or "").strip() or "color"
    if tool not in _VALID_TASK_TYPES:
        return jsonify({"ok": False, "error": f"Unknown tool: {tool}"}), 400

    frame = video.extract_frame_at_timestamp(video_path, ts)
    if frame is None:
        return jsonify({"ok": False, "error": "Could not read frame"}), 400
    frame_h, frame_w = frame.shape[:2]

    region_coords: dict[str, int] | None = None
    region_str = request.args.get("region", "").strip()
    if region_str:
        parts = region_str.split(",")
        if len(parts) == 4:
            try:
                rx, ry, rw, rh = (float(p) for p in parts)
            except ValueError:
                return jsonify({"ok": False, "error": "Invalid region"}), 400
            region_coords = {
                "x": int(round(rx * frame_w)),
                "y": int(round(ry * frame_h)),
                "w": int(round(rw * frame_w)),
                "h": int(round(rh * frame_h)),
            }

    # Prev frame for tools that consume a temporal pair
    prev_frame = None
    if tool in ("change", "flow"):
        prev_ts_raw = request.args.get("prev")
        if prev_ts_raw is not None:
            try:
                prev_ts = float(prev_ts_raw)
            except ValueError:
                prev_ts = max(0.0, ts - 1.0)
        else:
            prev_ts = max(0.0, ts - 1.0)
        if prev_ts < ts:
            prev_frame = video.extract_frame_at_timestamp(video_path, prev_ts)

    # Build params dict for the preview (subset of task parameters)
    params: dict[str, Any] = {}
    if tool == "color":
        for key in ("h", "s", "v"):
            raw = request.args.get(key)
            if raw is not None:
                try:
                    params[key] = float(raw)
                except ValueError:
                    pass
    elif tool == "change":
        raw = request.args.get("noise")
        if raw is not None:
            try:
                params["noise_threshold"] = int(float(raw))
            except ValueError:
                pass
    elif tool == "flow":
        raw = request.args.get("magnitude")
        if raw is not None:
            try:
                params["magnitude_threshold"] = float(raw)
            except ValueError:
                pass
    elif tool == "similarity":
        ref_ts_raw = request.args.get("ref")
        if ref_ts_raw is not None and region_coords is not None:
            try:
                ref_ts = float(ref_ts_raw)
            except ValueError:
                ref_ts = None  # type: ignore[assignment]
            if ref_ts is not None:
                ref_frame = video.extract_frame_at_timestamp(video_path, ref_ts)
                if ref_frame is not None:
                    import screenspace as _ss

                    params["reference_frame"] = _ss.extract_region(
                        ref_frame, region_coords
                    )

    elif tool == "template":
        import screenspace as _ss_tpl

        upload_b64: str | None = None
        if request.method == "POST":
            body = request.get_json(silent=True)
            if isinstance(body, dict):
                raw = body.get("template_image_data")
                if isinstance(raw, str) and raw.strip():
                    upload_b64 = raw.strip()
        if upload_b64:
            try:
                bgr, mask = _template_bgr_and_mask_from_b64(upload_b64)
            except ValueError:
                return jsonify(
                    {"ok": False, "error": "Could not decode uploaded image"}
                ), 400
            params["template_image"] = bgr
            if mask is not None:
                params["template_mask"] = mask
        else:
            ref_ts_raw = request.args.get("ref")
            if ref_ts_raw is not None and region_coords is not None:
                try:
                    ref_ts_tpl = float(ref_ts_raw)
                except ValueError:
                    ref_ts_tpl = None  # type: ignore[assignment]
                if ref_ts_tpl is not None:
                    ref_frame_tpl = video.extract_frame_at_timestamp(
                        video_path, ref_ts_tpl
                    )
                    if ref_frame_tpl is not None:
                        params["template_image"] = _ss_tpl.extract_region(
                            ref_frame_tpl, region_coords
                        )

    layer = (request.args.get("layer") or "").strip()
    if layer:
        # For multitool, the catalog comes from the first step's tool.
        catalog_tool = tool
        if tool == "multitool":
            steps = params.get("steps") or []
            catalog_tool = (
                steps[0].get("type", "") if steps and isinstance(steps[0], dict) else ""
            )
        valid = {
            lid
            for lid, _label, _scope in screenspace_preview.OVERLAY_LAYERS.get(
                catalog_tool, []
            )
        }
        if layer not in valid:
            return jsonify(
                {
                    "ok": False,
                    "error": f"Layer '{layer}' not available for tool '{tool}'",
                }
            ), 400
        layer_img = screenspace_preview.build_overlay_layer(
            frame, prev_frame, region_coords, tool, layer, params
        )
        if layer_img is None or getattr(layer_img, "size", 0) == 0:
            return jsonify({"ok": False, "error": "Could not build overlay layer"}), 500
        png_bytes = screenspace_preview.encode_png(layer_img, cap_width=False)
        if not png_bytes:
            return jsonify({"ok": False, "error": "Could not encode overlay"}), 500
        return Response(
            png_bytes,
            mimetype="image/png",
            headers={"Cache-Control": "no-cache"},
        )

    img = screenspace_preview.build_preview(
        frame, prev_frame, region_coords, tool, params
    )
    if img is None or getattr(img, "size", 0) == 0:
        return jsonify({"ok": False, "error": "Could not build preview"}), 500

    png_bytes = screenspace_preview.encode_png(img)
    if not png_bytes:
        return jsonify({"ok": False, "error": "Could not encode preview"}), 500
    return Response(
        png_bytes,
        mimetype="image/png",
        headers={"Cache-Control": "no-cache"},
    )


@screenspace_bp.route("/api/preview/layers")
def api_preview_layers() -> FlaskResponse:
    """Return the per-tool overlay-layer catalog as JSON.

    Shape: ``{tool: [{id, label, scope}, ...], ...}``. Tools whose previews
    aren't pixel-aligned (timelapse, inactivity) are intentionally absent.
    """
    import screenspace_preview

    out = {
        tool: [
            {"id": lid, "label": label, "scope": scope} for lid, label, scope in layers
        ]
        for tool, layers in screenspace_preview.OVERLAY_LAYERS.items()
    }
    return jsonify({"ok": True, "layers": out})


@screenspace_bp.route("/api/video/info/<participant>")
def api_video_info(participant: str) -> FlaskResponse:
    """Return video metadata (duration, resolution, fps)."""
    if participant in _video_metadata_cache:
        return jsonify({"ok": True, "info": _video_metadata_cache[participant]})

    video_path = _find_participant_video(participant)
    if video_path is None:
        return jsonify(
            {"ok": False, "error": f"No video for participant {participant}"}
        ), 404

    props = video.probe_video_properties(video_path)
    if props is None:
        return jsonify({"ok": False, "error": "Could not probe video file"}), 500

    vid_fps = props.get("fps", 0.0) or 30.0
    width = props.get("width", 0)
    height = props.get("height", 0)
    duration_seconds = props.get("duration", 0.0)
    duration = round(duration_seconds) if duration_seconds > 0 else 0

    info: dict[str, Any] = {
        "participant": participant,
        "duration": duration,
        "fps": vid_fps,
        "width": width if width > 0 else None,
        "height": height if height > 0 else None,
    }
    _video_metadata_cache[participant] = info

    return jsonify({"ok": True, "info": info})


@screenspace_bp.route("/api/video/stream/<participant>")
def api_video_stream(participant: str) -> FlaskResponse:
    """Stream the source video file for a participant (range-request aware)."""
    video_path = _find_participant_video(participant)
    if video_path is None:
        return (
            jsonify({"ok": False, "error": f"No video for participant {participant}"}),
            404,
        )
    return send_file(video_path, mimetype="video/mp4", conditional=True)


# ---- Region coordinate normalization ----


def _normalize_region(
    x: float,
    y: float,
    w: float,
    h: float,
    frame_w: int,
    frame_h: int,
) -> dict[str, Any]:
    """Convert pixel coordinates to normalized 0-1 fractions."""
    return {
        "x": x / frame_w,
        "y": y / frame_h,
        "w": w / frame_w,
        "h": h / frame_h,
        "source_width": frame_w,
        "source_height": frame_h,
    }


def _combined_region_lookup() -> dict[str, Any]:
    """Return regions addressable by legacy name lookups, with active regions winning."""
    regions: dict[str, Any] = {}
    for stash in _manifest.get("stashes", []):
        stash_regions = stash.get("regions", {})
        if isinstance(stash_regions, dict):
            regions.update(stash_regions)
    active_regions = _manifest.get("regions", {})
    if isinstance(active_regions, dict):
        regions.update(active_regions)
    return regions


def _resolve_region_request(
    region_name: str,
    region_ref: Any,
) -> tuple[str, dict[str, Any]] | FlaskResponse:
    """Resolve a task region request without flattening active/stashed duplicates."""
    active_regions = _manifest.get("regions", {})
    if not isinstance(active_regions, dict):
        active_regions = {}

    if region_ref is None:
        if region_name in active_regions:
            return region_name, cast(dict[str, Any], active_regions[region_name])
        for stash in _manifest.get("stashes", []):
            stash_regions = stash.get("regions", {})
            if isinstance(stash_regions, dict) and region_name in stash_regions:
                return region_name, cast(dict[str, Any], stash_regions[region_name])
        return jsonify({"ok": False, "error": f"Region '{region_name}' not found"}), 400

    if not isinstance(region_ref, dict):
        return jsonify({"ok": False, "error": "region_ref must be an object"}), 400

    source = str(region_ref.get("source", "")).strip()
    name = str(region_ref.get("name", "")).strip()
    if not name:
        return jsonify({"ok": False, "error": "region_ref.name is required"}), 400

    if source == "active":
        if name not in active_regions:
            return jsonify({"ok": False, "error": f"Region '{name}' not found"}), 400
        return name, cast(dict[str, Any], active_regions[name])

    if source == "stash":
        stash_id = str(region_ref.get("stash_id", "")).strip()
        if not stash_id:
            return jsonify(
                {"ok": False, "error": "region_ref.stash_id is required"}
            ), 400
        for stash in _manifest.get("stashes", []):
            if stash.get("id") != stash_id:
                continue
            stash_regions = stash.get("regions", {})
            if not isinstance(stash_regions, dict) or name not in stash_regions:
                return (
                    jsonify(
                        {
                            "ok": False,
                            "error": f"Region '{name}' not found in stash '{stash_id}'",
                        }
                    ),
                    400,
                )
            return name, cast(dict[str, Any], stash_regions[name])
        return jsonify({"ok": False, "error": f"Stash '{stash_id}' not found"}), 400

    return (
        jsonify(
            {"ok": False, "error": "region_ref.source must be 'active' or 'stash'"}
        ),
        400,
    )


def _is_flask_error_response(value: Any) -> TypeGuard[FlaskResponse]:
    return isinstance(value, Response) or (
        isinstance(value, tuple) and len(value) == 2 and isinstance(value[1], int)
    )


# ---- Regions CRUD ----


@screenspace_bp.route("/api/regions")
def api_regions_list() -> FlaskResponse:
    """List all saved region definitions."""
    return jsonify({"ok": True, "regions": _manifest.get("regions", {})})


@screenspace_bp.route("/api/regions", methods=["POST"])
def api_regions_create() -> FlaskResponse:
    """Create or update a named region."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "JSON body required"}), 400

    name = data.get("name", "").strip()
    if not name:
        return jsonify({"ok": False, "error": "Region name is required"}), 400

    for field in ("x", "y", "w", "h"):
        val = data.get(field)
        if val is None or not isinstance(val, (int, float)):
            return jsonify({"ok": False, "error": f"'{field}' must be a number"}), 400

    canvas_w = data.get("canvas_width")
    canvas_h = data.get("canvas_height")
    if (
        not isinstance(canvas_w, (int, float))
        or not isinstance(canvas_h, (int, float))
        or canvas_w <= 0
        or canvas_h <= 0
    ):
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "'canvas_width' and 'canvas_height' must be positive numbers",
                }
            ),
            400,
        )

    region = _normalize_region(
        data["x"], data["y"], data["w"], data["h"], int(canvas_w), int(canvas_h)
    )
    if "description" in data:
        region["description"] = str(data["description"])

    with _manifest_lock:
        _manifest.setdefault("regions", {})[name] = region
        _do_persist(drain_events=False)

    return jsonify({"ok": True, "region": region})


@screenspace_bp.route("/api/regions/<name>", methods=["DELETE"])
def api_regions_delete(name: str) -> FlaskResponse:
    """Delete a region definition."""
    regions = _manifest.get("regions", {})
    if name not in regions:
        return jsonify({"ok": False, "error": f"Region '{name}' not found"}), 404

    with _manifest_lock:
        del regions[name]
        _do_persist(drain_events=False)

    return jsonify({"ok": True})


# ---- Stashes CRUD ----


@screenspace_bp.route("/api/stashes")
def api_stashes_list() -> FlaskResponse:
    """List all region stashes."""
    return jsonify({"ok": True, "stashes": _manifest.get("stashes", [])})


@screenspace_bp.route("/api/stashes", methods=["POST"])
def api_stashes_create() -> FlaskResponse:
    """Stash all current regions and clear the active set."""
    regions = _manifest.get("regions", {})
    if not regions:
        return jsonify({"ok": False, "error": "No regions to stash"}), 400

    stash = {
        "id": "stash_" + uuid.uuid4().hex[:8],
        "name": "Stashed Regions",
        "createdAt": datetime.now(timezone.utc).isoformat(),
        "regions": copy.deepcopy(regions),
    }
    with _manifest_lock:
        _manifest.setdefault("stashes", []).append(stash)
        _manifest["regions"] = {}
        _do_persist(drain_events=False)
    return jsonify({"ok": True, "stash": stash})


@screenspace_bp.route("/api/stashes/<stash_id>", methods=["PUT"])
def api_stashes_update(stash_id: str) -> FlaskResponse:
    """Update a stash (rename)."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "JSON body required"}), 400

    stashes = _manifest.get("stashes", [])
    stash = next((s for s in stashes if s["id"] == stash_id), None)
    if stash is None:
        return jsonify({"ok": False, "error": "Stash not found"}), 404

    name = data.get("name", "").strip()
    with _manifest_lock:
        if name:
            stash["name"] = name
        _do_persist(drain_events=False)
    return jsonify({"ok": True, "stash": stash})


@screenspace_bp.route("/api/stashes/<stash_id>", methods=["DELETE"])
def api_stashes_delete(stash_id: str) -> FlaskResponse:
    """Dismiss a region stash."""
    stashes = _manifest.get("stashes", [])
    idx = next((i for i, s in enumerate(stashes) if s["id"] == stash_id), None)
    if idx is None:
        return jsonify({"ok": False, "error": "Stash not found"}), 404

    with _manifest_lock:
        stashes.pop(idx)
        _do_persist(drain_events=False)
    return jsonify({"ok": True})


@screenspace_bp.route("/api/stashes/<stash_id>/restore", methods=["POST"])
def api_stashes_restore(stash_id: str) -> FlaskResponse:
    """Restore a stash: replace active regions with stashed ones (stash is kept)."""
    stashes = _manifest.get("stashes", [])
    idx = next((i for i, s in enumerate(stashes) if s["id"] == stash_id), None)
    if idx is None:
        return jsonify({"ok": False, "error": "Stash not found"}), 404

    stash = stashes[idx]
    with _manifest_lock:
        _manifest["regions"] = copy.deepcopy(stash["regions"])
        _do_persist(drain_events=False)
    return jsonify({"ok": True, "regions": _manifest["regions"]})


# ---- Tasks CRUD ----


@screenspace_bp.route("/api/tasks/stream")
def api_tasks_stream() -> FlaskResponse:
    """SSE endpoint for live task updates (replaces polling)."""
    client_q: queue_mod.Queue[str] = queue_mod.Queue(maxsize=64)
    with _sse_clients_lock:
        _sse_clients.append(client_q)

    def generate():  # type: ignore[no-untyped-def]
        try:
            # Send current state immediately on connect
            yield _sse_task_payload()
            while True:
                try:
                    client_q.get(timeout=15)
                    # Drain any queued notifications (coalesce rapid updates)
                    while not client_q.empty():
                        try:
                            client_q.get_nowait()
                        except queue_mod.Empty:
                            break
                    yield _sse_task_payload()
                except queue_mod.Empty:
                    # Keepalive comment to prevent connection timeout
                    yield ": keepalive\n\n"
        except GeneratorExit:
            pass
        finally:
            with _sse_clients_lock:
                if client_q in _sse_clients:
                    _sse_clients.remove(client_q)

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@screenspace_bp.route("/api/tasks")
def api_tasks_list() -> FlaskResponse:
    """List all tasks with status and progress."""
    tasks = _worker.get_all_tasks() if _worker else []
    clean = [_clean_task(t) for t in tasks]
    paused = _worker.is_paused if _worker else False
    alive = _worker.is_alive if _worker else False
    return jsonify(
        {"ok": True, "tasks": clean, "paused": paused, "worker_alive": alive}
    )


@screenspace_bp.route("/api/tasks/<task_id>")
def api_tasks_get(task_id: str) -> FlaskResponse:
    """Get task detail including results."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500
    task = _worker.get_task(task_id)
    if task is None:
        return jsonify({"ok": False, "error": "Task not found"}), 404
    return jsonify({"ok": True, "task": _clean_task(task)})


# ---- Task creation helpers ----
#
# Private helpers for api_tasks_create below: validation, parameter coercion,
# media extraction, and multitool step preparation. Placed between the read-side
# task routes above and the create/mutate routes below so the create endpoint's
# dependencies are immediately visible when reading top-down.


def _validate_task_request(
    data: dict[str, Any],
) -> (
    tuple[str, str, str, dict[str, Any], dict[str, Any], dict[str, Any] | None]
    | FlaskResponse
):
    """Validate the task creation request body.

    Returns (task_type, participant, region_name, parameters, all_known_regions,
    requested_region) on success, or a Flask error response on failure.
    """
    task_type = data.get("type", "").strip()
    if task_type not in _VALID_TASK_TYPES:
        return jsonify(
            {
                "ok": False,
                "error": f"type must be one of: {', '.join(_VALID_TASK_TYPES)}",
            }
        ), 400

    participant = data.get("participant", "").strip()
    if not participant:
        return jsonify({"ok": False, "error": "participant is required"}), 400

    region_name = data.get("region", "").strip()
    region_ref = data.get("region_ref")
    raw_parameters = data.get("parameters")
    if raw_parameters is None:
        parameters: dict[str, Any] = {}
    elif isinstance(raw_parameters, dict):
        parameters = raw_parameters
    else:
        return jsonify({"ok": False, "error": "parameters must be an object"}), 400

    # Template tasks with an uploaded image scan the full frame; no region needed
    has_uploaded_template = task_type == "template" and parameters.get(
        "template_image_data"
    )

    # Multitool uses per-step regions; others need a global region (unless template upload)
    has_region_request = bool(region_name) or region_ref is not None
    if (
        not has_region_request
        and not has_uploaded_template
        and task_type != "multitool"
    ):
        return jsonify({"ok": False, "error": "region is required"}), 400

    # Early validation for multitool steps
    if task_type == "multitool":
        mt_steps = parameters.get("steps")
        if not mt_steps or not isinstance(mt_steps, list) or len(mt_steps) < 2:
            return jsonify(
                {"ok": False, "error": "Multitool requires at least 2 steps"}
            ), 400
        for i, step_raw in enumerate(mt_steps):
            if not isinstance(step_raw, dict):
                return jsonify(
                    {"ok": False, "error": f"Step {i}: must be an object"}
                ), 400
            step_v = cast(dict[str, Any], step_raw)
            stype = step_v.get("type", "")
            if stype not in _VALID_STEP_TYPES:
                return jsonify(
                    {"ok": False, "error": f"Step {i}: invalid type '{stype}'"}
                ), 400
            logic = step_v.get("logic")
            if logic is not None and logic not in ("AND", "NOT"):
                return jsonify(
                    {"ok": False, "error": f"Step {i}: logic must be 'AND' or 'NOT'"}
                ), 400

    all_known_regions = _combined_region_lookup()
    requested_region: dict[str, Any] | None = None

    # Validate regions
    if task_type == "multitool":
        mt_steps_early: list[dict[str, Any]] = parameters.get("steps", [])
        for i, step in enumerate(mt_steps_early):
            step_region = (step.get("region") or "").strip()
            step_region_ref = step.get("region_ref")
            if not step_region and step_region_ref is None:
                return jsonify(
                    {"ok": False, "error": f"Step {i}: region is required"}
                ), 400
            if step_region_ref is not None:
                resolved = _resolve_region_request(step_region, step_region_ref)
                if _is_flask_error_response(resolved):
                    return resolved
            elif step_region not in all_known_regions:
                return jsonify(
                    {
                        "ok": False,
                        "error": f"Step {i}: region '{step_region}' not found",
                    }
                ), 400
    elif has_region_request:
        resolved = _resolve_region_request(region_name, region_ref)
        if _is_flask_error_response(resolved):
            return resolved
        region_name, requested_region = cast(tuple[str, dict[str, Any]], resolved)

    return (
        task_type,
        participant,
        region_name,
        parameters,
        all_known_regions,
        requested_region,
    )


def _coerce_task_params(
    task_type: str, parameters: dict[str, Any]
) -> dict[str, Any] | FlaskResponse:
    """Coerce and validate type-specific parameter values.

    Returns the updated parameters on success, or a Flask error response on failure.
    """
    try:
        if task_type == "similarity":
            parameters["reference_timestamp"] = _coerce_float(
                parameters.get("reference_timestamp"),
                "reference_timestamp",
                required=True,
            )
        elif task_type == "template" and not parameters.get("template_image_data"):
            parameters["reference_timestamp"] = _coerce_float(
                parameters.get("reference_timestamp"),
                "reference_timestamp",
                required=True,
            )
        elif task_type == "scene":
            parameters["scene_references"] = _validate_scene_references(
                parameters.get("scene_references")
            )
        elif task_type == "multitool":
            for i, step in enumerate(parameters.get("steps", [])):
                step_context = f"Step {i}: "
                if step["type"] == "similarity":
                    step["reference_timestamp"] = _coerce_float(
                        step.get("reference_timestamp"),
                        "reference_timestamp",
                        required=True,
                        context=step_context,
                    )
                elif step["type"] == "template" and not step.get("template_image_data"):
                    step["reference_timestamp"] = _coerce_float(
                        step.get("reference_timestamp"),
                        "reference_timestamp",
                        required=True,
                        context=step_context,
                    )
                elif step["type"] == "scene":
                    step["scene_references"] = _validate_scene_references(
                        step.get("scene_references"),
                        context=step_context,
                    )
                if step.get("type") == "template":
                    _coerce_template_controls(step)
        if task_type == "template":
            _coerce_template_controls(parameters)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    return parameters


def _extract_task_media(
    task_type: str,
    parameters: dict[str, Any],
    video_path: str,
    region_coords: dict[str, Any],
) -> dict[str, Any] | FlaskResponse:
    """Extract reference frames / template images for non-multitool tasks.

    Returns the updated parameters on success, or a Flask error response on failure.
    """
    if task_type == "similarity":
        ref_ts = cast(float, parameters["reference_timestamp"])
        frame = video.extract_frame_at_timestamp(video_path, float(ref_ts))
        if frame is None:
            return jsonify(
                {"ok": False, "error": "Could not read reference frame"}
            ), 400
        ref_region = screenspace.extract_region(frame, region_coords)
        parameters["reference_frame"] = ref_region

    if task_type == "template":
        upload_b64 = parameters.pop("template_image_data", None)
        if upload_b64:
            try:
                bgr, mask = _template_bgr_and_mask_from_b64(upload_b64)
            except ValueError:
                return jsonify(
                    {"ok": False, "error": "Could not decode uploaded image"}
                ), 400
            parameters["template_image"] = bgr
            if mask is not None:
                parameters["template_mask"] = mask
        else:
            ref_ts = cast(float, parameters["reference_timestamp"])
            frame = video.extract_frame_at_timestamp(video_path, float(ref_ts))
            if frame is None:
                return jsonify(
                    {"ok": False, "error": "Could not read template frame"}
                ), 400
            parameters["template_image"] = screenspace.extract_region(
                frame, region_coords
            )

    if task_type == "scene":
        scene_refs = cast(list[dict[str, Any]], parameters["scene_references"])
        reference_scenes = []
        for ref in scene_refs:
            frame = video.extract_frame_at_timestamp(
                video_path, float(ref["timestamp"])
            )
            if frame is None:
                return jsonify(
                    {
                        "ok": False,
                        "error": f"Could not read frame for scene '{ref['name']}'",
                    }
                ), 400
            ref_region = screenspace.extract_region(frame, region_coords)
            scene_entry: dict = {"name": ref["name"], "frame": ref_region}
            if "threshold" in ref:
                scene_entry["threshold"] = float(ref["threshold"])
            reference_scenes.append(scene_entry)
        parameters["reference_scenes"] = reference_scenes

    return parameters


def _prepare_multitool_steps(
    parameters: dict[str, Any],
    all_known_regions: dict[str, Any],
    video_path: str,
    region_coords: dict[str, Any],
    resolve_region_fn: Any,
) -> dict[str, Any] | FlaskResponse:
    """Resolve per-step regions and extract media for multitool tasks.

    Returns the updated parameters on success, or a Flask error response on failure.
    """
    steps = parameters.get("steps", [])
    for i, step in enumerate(steps):
        stype = step.get("type", "")

        # Resolve this step's region to pixel coords
        step_region_name = (step.get("region") or "").strip()
        step_region_ref = step.get("region_ref")
        if step_region_name or step_region_ref is not None:
            resolved = _resolve_region_request(step_region_name, step_region_ref)
            if _is_flask_error_response(resolved):
                return resolved
            resolved_name, resolved_region = cast(tuple[str, dict[str, Any]], resolved)
            step["region"] = resolved_name
            step["region_coords"] = resolve_region_fn(resolved_name, resolved_region)
        else:
            step["region_coords"] = region_coords  # fallback to top-level

        step_rc = step["region_coords"]

        if stype == "similarity":
            ref_ts = cast(float, step["reference_timestamp"])
            frame = video.extract_frame_at_timestamp(video_path, float(ref_ts))
            if frame is None:
                return jsonify(
                    {
                        "ok": False,
                        "error": f"Step {i}: could not read reference frame",
                    }
                ), 400
            step["reference_frame"] = screenspace.extract_region(frame, step_rc)

        elif stype == "template":
            upload_b64 = step.pop("template_image_data", None)
            if upload_b64:
                try:
                    bgr, mask = _template_bgr_and_mask_from_b64(upload_b64)
                except ValueError:
                    return jsonify(
                        {
                            "ok": False,
                            "error": f"Step {i}: could not decode uploaded image",
                        }
                    ), 400
                step["template_image"] = bgr
                if mask is not None:
                    step["template_mask"] = mask
            else:
                ref_ts = cast(float, step["reference_timestamp"])
                frame = video.extract_frame_at_timestamp(video_path, float(ref_ts))
                if frame is None:
                    return jsonify(
                        {
                            "ok": False,
                            "error": f"Step {i}: could not read template frame",
                        }
                    ), 400
                step["template_image"] = screenspace.extract_region(frame, step_rc)

        elif stype == "scene":
            scene_refs = cast(list[dict[str, Any]], step["scene_references"])
            ref_scenes_list = []
            for ref in scene_refs:
                frame = video.extract_frame_at_timestamp(
                    video_path, float(ref["timestamp"])
                )
                if frame is None:
                    return jsonify(
                        {
                            "ok": False,
                            "error": f"Step {i}: could not read frame for scene '{ref['name']}'",
                        }
                    ), 400
                ref_region = screenspace.extract_region(frame, step_rc)
                scene_entry: dict = {"name": ref["name"], "frame": ref_region}
                if "threshold" in ref:
                    scene_entry["threshold"] = float(ref["threshold"])
                ref_scenes_list.append(scene_entry)
            step["reference_scenes"] = ref_scenes_list

    return parameters


@screenspace_bp.route("/api/tasks", methods=["POST"])
def api_tasks_create() -> FlaskResponse:
    """Enqueue a new analysis task."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500

    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "JSON body required"}), 400

    validated = _validate_task_request(data)
    if isinstance(validated, Response) or (
        isinstance(validated, tuple) and len(validated) == 2
    ):
        return cast(FlaskResponse, validated)
    assert isinstance(validated, tuple) and len(validated) == 6  # success tuple
    (
        task_type,
        participant,
        region_name,
        parameters,
        all_known_regions,
        requested_region,
    ) = cast(
        tuple[str, str, str, dict[str, Any], dict[str, Any], dict[str, Any] | None],
        validated,
    )

    video_path = _find_participant_video(participant)
    if video_path is None:
        return jsonify(
            {"ok": False, "error": f"No video for participant {participant}"}
        ), 400

    source_video = ""
    for p in _participants:
        if p["id"] == participant:
            source_video = Path(p["video_path"]).name
            break

    props = video.probe_video_properties(video_path)

    def _resolve_region_coords(
        name: str, region_data: dict[str, Any] | None = None
    ) -> dict[str, Any]:
        """Convert a named region to pixel coordinates."""
        rd = region_data if region_data is not None else all_known_regions[name]
        if props and props.get("width") and props.get("height"):
            return screenspace.denormalize_region(rd, props["width"], props["height"])
        return {k: rd[k] for k in ("x", "y", "w", "h") if k in rd}

    if requested_region is not None:
        region_coords = _resolve_region_coords(region_name, requested_region)
    elif task_type == "multitool":
        first_step_region = parameters.get("steps", [{}])[0].get("region", "")
        if first_step_region and first_step_region in all_known_regions:
            region_name = first_step_region
            region_coords = _resolve_region_coords(first_step_region)
        else:
            region_name = "per_step"
            region_coords = {"x": 0, "y": 0, "w": 0, "h": 0}
    else:
        region_name = "full_frame"
        region_coords = {"x": 0, "y": 0, "w": 0, "h": 0}

    coerced = _coerce_task_params(task_type, parameters)
    if isinstance(coerced, tuple):
        return coerced  # Flask error response
    parameters = cast(dict[str, Any], coerced)

    extracted = _extract_task_media(task_type, parameters, video_path, region_coords)
    if isinstance(extracted, tuple):
        return extracted  # Flask error response
    parameters = cast(dict[str, Any], extracted)

    if task_type == "multitool":
        prepared = _prepare_multitool_steps(
            parameters,
            all_known_regions,
            video_path,
            region_coords,
            _resolve_region_coords,
        )
        if isinstance(prepared, tuple):
            return prepared  # Flask error response
        parameters = cast(dict[str, Any], prepared)

    # Snapshot the global CV resolution scale into the task so the manifest
    # records what scale produced each result (useful when re-running with
    # a different scale to compare outputs).
    parameters.setdefault("cv_resolution_scale", config.SCREENSPACE_CV_RESOLUTION_SCALE)

    task = screenspace.create_task(
        task_type=task_type,
        participant=participant,
        source_video=source_video,
        video_path=video_path,
        region_name=region_name,
        region_coords=region_coords,
        parameters=parameters,
    )
    if requested_region is not None:
        request_region_ref = data.get("region_ref")
        task["region_ref"] = (
            request_region_ref
            if isinstance(request_region_ref, dict)
            else {"source": "active", "name": region_name}
        )

    _worker.enqueue(task)
    _persist_manifest(drain_events=False)
    _notify_sse_clients("task_created")

    return jsonify({"ok": True, "task": _clean_task(task)})


@screenspace_bp.route("/api/tasks/<task_id>", methods=["DELETE"])
def api_tasks_cancel(task_id: str) -> FlaskResponse:
    """Cancel or dismiss a task.  ?dismiss=true fully removes the task."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500
    if request.args.get("dismiss") == "true":
        if not _worker.remove_task(task_id):
            return jsonify({"ok": False, "error": "Task not found"}), 404
        _persist_manifest(drain_events=False)
        _notify_sse_clients("task_dismissed")
        return jsonify({"ok": True})
    if _worker.cancel(task_id):
        _persist_manifest(drain_events=False)
        _notify_sse_clients("task_cancelled")
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Task not found or already finished"}), 400


@screenspace_bp.route("/api/tasks/reorder", methods=["PUT"])
def api_tasks_reorder() -> FlaskResponse:
    """Reorder queued tasks by priority."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500
    data = request.get_json(silent=True)
    if not data or "task_ids" not in data:
        return jsonify({"ok": False, "error": "task_ids list required"}), 400
    _worker.reorder(data["task_ids"])
    _persist_manifest(drain_events=False)
    _notify_sse_clients("reorder")
    return jsonify({"ok": True})


@screenspace_bp.route("/api/tasks/pause", methods=["POST"])
def api_tasks_pause() -> FlaskResponse:
    """Pause the task queue."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500
    _worker.pause()
    _persist_manifest(drain_events=False)
    _notify_sse_clients("pause")
    return jsonify({"ok": True, "paused": True})


@screenspace_bp.route("/api/tasks/resume", methods=["POST"])
def api_tasks_resume() -> FlaskResponse:
    """Resume the task queue."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500
    _worker.resume()
    _persist_manifest(drain_events=False)
    _notify_sse_clients("resume")
    return jsonify({"ok": True, "paused": False})


@screenspace_bp.route("/api/tasks/<task_id>/results")
def api_tasks_results(task_id: str) -> FlaskResponse:
    """Get task results (timestamps, artifacts)."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500
    task = _worker.get_task(task_id)
    if task is None:
        return jsonify({"ok": False, "error": "Task not found"}), 404
    return jsonify({"ok": True, "results": _sanitize_floats(task.get("result"))})


# ---- Events CRUD ----


@screenspace_bp.route("/api/events")
def api_events_list() -> FlaskResponse:
    """List events with optional filtering."""
    events = _manifest.get("events", [])
    excluded_filter = request.args.get("excluded")
    if excluded_filter == "false":
        events = [e for e in events if not e.get("excluded")]
    elif excluded_filter == "true":
        events = [e for e in events if e.get("excluded")]
    participant = request.args.get("participant")
    if participant:
        events = [e for e in events if e.get("participant") == participant]
    task_id = request.args.get("task_id")
    if task_id:
        events = [e for e in events if e.get("task_id") == task_id]
    return jsonify({"ok": True, "events": _sanitize_floats(events)})


@screenspace_bp.route("/api/export/events")
def api_export_events() -> FlaskResponse:
    """Export Screenspace events as analysis-ready JSON or CSV.

    Query params:
      format:      "json" (default) or "csv"
      excluded:    "true" to keep only excluded, "false" to drop excluded; default keeps both
      participant: keep only events for this participant id
      detector:    keep only events from this detector type
    """
    import data_export

    fmt = (request.args.get("format") or "json").lower()
    excluded_filter = request.args.get("excluded")
    participant = request.args.get("participant")
    detector = request.args.get("detector")

    if excluded_filter == "false":
        include_excluded = False
    elif excluded_filter == "true":
        include_excluded = True
    else:
        include_excluded = True

    records = data_export.build_screenspace_events(
        _manifest,
        include_excluded=include_excluded,
        participants=[participant] if participant else None,
        detectors=[detector] if detector else None,
    )
    if excluded_filter == "true":
        records = [r for r in records if r.get("excluded")]

    if fmt == "csv":
        body = data_export.to_csv(
            records,
            preferred_column_order=data_export.SCREENSPACE_EVENT_COLUMNS,
        )
        response = Response(body, mimetype="text/csv")
        response.headers["Content-Disposition"] = (
            'attachment; filename="screenspace_events.csv"'
        )
        return response
    if fmt == "json":
        return jsonify({"ok": True, "events": _sanitize_floats(records)})
    return jsonify({"ok": False, "error": f"Unsupported format: {fmt}"}), 400


@screenspace_bp.route("/api/events/<event_id>/exclude", methods=["PUT"])
def api_event_exclude(event_id: str) -> FlaskResponse:
    """Set an event as excluded."""
    for e in _manifest.get("events", []):
        if e["id"] == event_id:
            with _manifest_lock:
                e["excluded"] = True
                _do_persist(drain_events=False)
            return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Event not found"}), 404


@screenspace_bp.route("/api/events/<event_id>/include", methods=["PUT"])
def api_event_include(event_id: str) -> FlaskResponse:
    """Set an event as included."""
    for e in _manifest.get("events", []):
        if e["id"] == event_id:
            with _manifest_lock:
                e["excluded"] = False
                _do_persist(drain_events=False)
            return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Event not found"}), 404


@screenspace_bp.route("/api/events/bulk-exclude", methods=["PUT"])
def api_events_bulk_exclude() -> FlaskResponse:
    """Bulk-exclude events by ID list."""
    data = request.get_json(silent=True) or {}
    ids = set(data.get("ids", []))
    if not ids:
        return jsonify({"ok": False, "error": "ids list required"}), 400
    with _manifest_lock:
        count = 0
        for e in _manifest.get("events", []):
            if e["id"] in ids:
                e["excluded"] = True
                count += 1
        _do_persist(drain_events=False)
    return jsonify({"ok": True, "updated": count})


@screenspace_bp.route("/api/events/bulk-include", methods=["PUT"])
def api_events_bulk_include() -> FlaskResponse:
    """Bulk-include events by ID list."""
    data = request.get_json(silent=True) or {}
    ids = set(data.get("ids", []))
    if not ids:
        return jsonify({"ok": False, "error": "ids list required"}), 400
    with _manifest_lock:
        count = 0
        for e in _manifest.get("events", []):
            if e["id"] in ids:
                e["excluded"] = False
                count += 1
        _do_persist(drain_events=False)
    return jsonify({"ok": True, "updated": count})


# ---- Helpers ----


def _coerce_float(
    value: Any,
    field_name: str,
    *,
    required: bool = False,
    context: str = "",
) -> float | None:
    """Validate and coerce a request field to a finite float.

    Returns None when value is None and required=False.
    """
    if value is None:
        if required:
            raise ValueError(f"{context}{field_name} is required")
        return None
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{context}{field_name} must be a number") from exc
    if not math.isfinite(number):
        raise ValueError(f"{context}{field_name} must be a finite number")
    return number


def _coerce_template_controls(params: dict[str, Any], *, context: str = "") -> None:
    """Validate template-tool controls: template_scale."""
    import config

    if "template_scale" in params and params["template_scale"] is not None:
        scale = _coerce_float(
            params["template_scale"], "template_scale", context=context
        )
        if scale is None or scale <= 0:
            raise ValueError(f"{context}template_scale must be a positive number")
        lo = config.SCREENSPACE_TEMPLATE_SCALE_MIN
        hi = config.SCREENSPACE_TEMPLATE_SCALE_MAX
        params["template_scale"] = max(lo, min(hi, scale))


def _validate_scene_references(
    scene_refs: Any, *, context: str = ""
) -> list[dict[str, Any]]:
    """Validate scene reference payloads and normalize numeric fields."""
    if not scene_refs or not isinstance(scene_refs, list):
        raise ValueError(f"{context}scene_references must be a non-empty list")
    validated_refs = []
    for i, ref_raw in enumerate(scene_refs):
        if not isinstance(ref_raw, dict):
            raise ValueError(f"{context}scene_references[{i}] must be an object")
        ref_data = cast(dict[str, Any], ref_raw)
        name = str(ref_data.get("name", "")).strip()
        if not name:
            raise ValueError(f"{context}scene_references[{i}].name is required")
        ref: dict[str, Any] = {
            "name": name,
            "timestamp": _coerce_float(
                ref_data.get("timestamp"),
                f"scene_references[{i}].timestamp",
                required=True,
                context=context,
            ),
        }
        if "threshold" in ref_data and ref_data.get("threshold") is not None:
            ref["threshold"] = _coerce_float(
                ref_data.get("threshold"),
                f"scene_references[{i}].threshold",
                context=context,
            )
        validated_refs.append(ref)
    return validated_refs


def _sanitize_floats(obj: Any) -> Any:
    """Replace non-finite floats (inf, nan) with None for JSON safety."""
    if isinstance(obj, float) and not math.isfinite(obj):
        return None
    if isinstance(obj, dict):
        return {k: _sanitize_floats(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize_floats(v) for v in obj]
    return obj


def _clean_task(task: dict[str, Any]) -> dict[str, Any]:
    """Remove internal fields from a task dict for API responses."""
    cleaned = {k: v for k, v in task.items() if not k.startswith("_")}
    if "parameters" in cleaned:
        cleaned["parameters"] = {
            k: v for k, v in cleaned["parameters"].items() if k not in _TASK_BINARY_KEYS
        }
        # Strip binary data and internal coords from multitool step parameters
        if "steps" in cleaned["parameters"]:
            _step_strip_keys = _TASK_BINARY_KEYS + ("region_coords",)
            cleaned["parameters"]["steps"] = [
                {k: v for k, v in s.items() if k not in _step_strip_keys}
                for s in cleaned["parameters"]["steps"]
            ]
    return _sanitize_floats(cleaned)


# ---- State initialization ----


def _init_screenspace_state(
    sheet_context: Any = None,
    participant_list: list[str] | None = None,
) -> None:
    """Initialize module-level state for Screenspace routes.

    Loads manifest, resolves participant video paths, and starts the
    background worker thread.
    """
    import screenspace

    global _manifest, _worker, _output_dir, _participants

    _output_dir = str(utils.get_effective_output_dir())
    _manifest = screenspace.load_screenspace_manifest()

    _participants = []
    study_name = ""
    if sheet_context is not None:
        study_name = getattr(sheet_context, "study_name", "")

    if participant_list:
        for pid in participant_list:
            filename = files.get_source_video_filename(study_name, pid)
            video_path = utils.resolve_input_path(filename)
            _participants.append(
                {
                    "id": pid,
                    "video_path": str(video_path),
                    "has_video": video_path.is_file(),
                }
            )
    else:
        _discover_participant_videos(study_name)

    _worker = screenspace.ScreenspaceWorker()
    _worker.on_task_complete = _persist_manifest
    _worker.on_progress_update = lambda: _notify_sse_clients("progress")
    _worker.restore_tasks(_manifest.get("tasks", []))
    _worker.start()


def _discover_participant_videos(study_name: str) -> None:
    """Scan input directory for source video files and populate _participants."""
    global _participants  # noqa: PLW0603
    _participants = utils.discover_participant_videos(study_name)


def _do_persist(*, drain_events: bool = True) -> None:
    """Persist manifest to disk — caller must hold _manifest_lock."""
    import screenspace

    if _worker and drain_events:
        new_events = _worker.drain_new_events()
        if new_events:
            _manifest.setdefault("events", []).extend(new_events)
    tasks = _worker.get_all_tasks() if _worker else _manifest.get("tasks", [])
    _manifest["tasks"] = tasks
    screenspace.save_screenspace_manifest(
        _manifest.get("regions", {}),
        tasks,
        _manifest.get("events", []),
        stashes=_manifest.get("stashes", []),
    )


def _persist_manifest(*, drain_events: bool = True) -> None:
    """Persist the current manifest state through a single synchronized path."""
    with _manifest_lock:
        _do_persist(drain_events=drain_events)
