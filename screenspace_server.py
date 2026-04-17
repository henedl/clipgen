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
  GET  /api/preview/<participant>/<ts>      – PNG of the active tool's preprocessed view
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
from typing import TYPE_CHECKING, Any, cast

if TYPE_CHECKING:
    import screenspace

from flask import Blueprint, Response, jsonify, request, send_file

import files
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


@screenspace_bp.route("/api/preview/<participant>/<timestamp>")
def api_preview(participant: str, timestamp: str) -> FlaskResponse:
    """Render what the selected tool's CV pipeline sees at ``timestamp``.

    Returns a PNG composite image (grayscale crop, diff mask, edge map, flow
    vectors, pHash bit grid, etc.) tailored to the active tool.  Optional query
    params:

      tool=<name>          one of the 11 screenspace tool types
      region=x,y,w,h       normalized 0–1 region coordinates (required for most tools)
      prev=<seconds>       prior timestamp for change/flow (defaults to ts-1s)
      ref=<seconds>        reference timestamp for similarity (full-frame region crop)
      noise=<int>          change tool's noise_threshold override
      h,s,v=<int>          color tool's target HSV override
      magnitude=<float>    flow tool's magnitude threshold override
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


def _denormalize_region(
    region: dict[str, Any], target_w: int, target_h: int
) -> dict[str, int]:
    """Convert normalized region to pixel coordinates for a target resolution.

    Legacy regions (without ``source_width``) pass through unchanged.
    """
    if "source_width" not in region:
        return {"x": region["x"], "y": region["y"], "w": region["w"], "h": region["h"]}
    return {
        "x": int(round(region["x"] * target_w)),
        "y": int(round(region["y"] * target_h)),
        "w": int(round(region["w"] * target_w)),
        "h": int(round(region["h"] * target_h)),
    }


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
    region: dict[str, Any]
    if (
        isinstance(canvas_w, (int, float))
        and isinstance(canvas_h, (int, float))
        and canvas_w > 0
        and canvas_h > 0
    ):
        region = _normalize_region(
            data["x"], data["y"], data["w"], data["h"], int(canvas_w), int(canvas_h)
        )
    else:
        region = {
            "x": int(data["x"]),
            "y": int(data["y"]),
            "w": int(data["w"]),
            "h": int(data["h"]),
        }
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


def _validate_task_request(
    data: dict[str, Any],
) -> tuple[str, str, str, dict[str, Any], dict[str, Any]] | FlaskResponse:
    """Validate the task creation request body.

    Returns (task_type, participant, region_name, parameters, all_known_regions)
    on success, or a Flask error response on failure.
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
    if not region_name and not has_uploaded_template and task_type != "multitool":
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

    # Build combined region lookup dict (active + stashes)
    all_known_regions: dict[str, Any] = dict(_manifest.get("regions", {}))
    for stash in _manifest.get("stashes", []):
        all_known_regions.update(stash.get("regions", {}))

    # Validate regions
    if task_type == "multitool":
        mt_steps_early: list[dict[str, Any]] = parameters.get("steps", [])
        for i, step in enumerate(mt_steps_early):
            step_region = (step.get("region") or "").strip()
            if not step_region:
                return jsonify(
                    {"ok": False, "error": f"Step {i}: region is required"}
                ), 400
            if step_region not in all_known_regions:
                return jsonify(
                    {
                        "ok": False,
                        "error": f"Step {i}: region '{step_region}' not found",
                    }
                ), 400
    else:
        if region_name and region_name not in all_known_regions:
            return jsonify(
                {"ok": False, "error": f"Region '{region_name}' not found"}
            ), 400

    return task_type, participant, region_name, parameters, all_known_regions


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
    import screenspace

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
        import base64

        import cv2
        import numpy as np

        upload_b64 = parameters.pop("template_image_data", None)
        if upload_b64:
            try:
                img_bytes = base64.b64decode(upload_b64)
                img_arr = np.frombuffer(img_bytes, dtype=np.uint8)
                img = cv2.imdecode(img_arr, cv2.IMREAD_UNCHANGED)
            except (ValueError, binascii.Error):
                return jsonify(
                    {"ok": False, "error": "Could not decode uploaded image"}
                ), 400
            if img is None:
                return jsonify({"ok": False, "error": "Invalid image data"}), 400
            if len(img.shape) == 2:
                img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
            if img.shape[2] == 4:
                parameters["template_mask"] = img[:, :, 3]
                parameters["template_image"] = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
            else:
                parameters["template_image"] = img
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
    import base64

    import cv2
    import numpy as np
    import screenspace

    steps = parameters.get("steps", [])
    for i, step in enumerate(steps):
        stype = step.get("type", "")

        # Resolve this step's region to pixel coords
        step_region_name = (step.get("region") or "").strip()
        if step_region_name and step_region_name in all_known_regions:
            step["region_coords"] = resolve_region_fn(step_region_name)
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
                    img_bytes = base64.b64decode(upload_b64)
                    img_arr = np.frombuffer(img_bytes, dtype=np.uint8)
                    img = cv2.imdecode(img_arr, cv2.IMREAD_UNCHANGED)
                except (ValueError, binascii.Error):
                    return jsonify(
                        {
                            "ok": False,
                            "error": f"Step {i}: could not decode uploaded image",
                        }
                    ), 400
                if img is None:
                    return jsonify(
                        {"ok": False, "error": f"Step {i}: invalid image data"}
                    ), 400
                if len(img.shape) == 2:
                    img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
                if img.shape[2] == 4:
                    step["template_mask"] = img[:, :, 3]
                    step["template_image"] = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                else:
                    step["template_image"] = img
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
    assert isinstance(validated, tuple) and len(validated) == 5  # success tuple
    task_type, participant, region_name, parameters, all_known_regions = cast(
        tuple[str, str, str, dict[str, Any], dict[str, Any]], validated
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

    def _resolve_region_coords(name: str) -> dict[str, Any]:
        """Convert a named region to pixel coordinates."""
        rd = all_known_regions[name]
        if props and props.get("width") and props.get("height"):
            return _denormalize_region(rd, props["width"], props["height"])
        return {k: rd[k] for k in ("x", "y", "w", "h") if k in rd}

    if region_name and region_name in all_known_regions:
        region_coords = _resolve_region_coords(region_name)
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

    import screenspace

    task = screenspace.create_task(
        task_type=task_type,
        participant=participant,
        source_video=source_video,
        video_path=video_path,
        region_name=region_name,
        region_coords=region_coords,
        parameters=parameters,
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
