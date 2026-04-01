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
  GET  /api/video/info/<participant>        – video metadata (duration, resolution, fps)
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

from __future__ import annotations

import copy
import math
import threading
import uuid
from collections import OrderedDict
from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Tuple, Union

if TYPE_CHECKING:
    import cv2
    import screenspace

from flask import Blueprint, Response, jsonify, request, send_from_directory

import config
import files
import utils
import video

FlaskResponse = Union[Response, Tuple[Response, int]]

# ---- Module-level state (set once by _init_screenspace_state) ----

_manifest: Dict[str, Any] = {}
_worker: Optional["screenspace.ScreenspaceWorker"] = None
_output_dir: str = ""
_participants: List[Dict[str, Any]] = []
_video_cap_cache: OrderedDict[str, Any] = OrderedDict()
_VIDEO_CAP_MAX = 3
_video_cap_lock = threading.Lock()
_video_metadata_cache: Dict[str, Dict[str, Any]] = {}

_assets_dir = utils.get_bundled_assets_root() / "assets" / "web"

# ---- Blueprint ----

screenspace_bp = Blueprint("screenspace", __name__)


# ---- Static file serving ----


@screenspace_bp.route("/")
def serve_index() -> FlaskResponse:
    return send_from_directory(_assets_dir, "screenspace.html")


@screenspace_bp.route("/<path:filename>")
def serve_static(filename: str) -> FlaskResponse:
    return send_from_directory(_assets_dir, filename)


@screenspace_bp.route("/icons/<path:filename>")
def serve_icons(filename: str) -> FlaskResponse:
    icons_dir = utils.get_bundled_assets_root() / "assets" / "icons"
    return send_from_directory(icons_dir, filename)


@screenspace_bp.route("/media/<path:filename>")
def serve_media(filename: str) -> FlaskResponse:
    if not _output_dir:
        return jsonify({"ok": False, "error": "Output directory not configured"}), 500
    return send_from_directory(_output_dir, filename)


# ---- Participants ----


@screenspace_bp.route("/api/participants")
def api_participants() -> FlaskResponse:
    """List participants with source video availability."""
    return jsonify({"ok": True, "participants": _participants})


# ---- Video frame extraction ----


def _find_participant_video(participant_id: str) -> Optional[str]:
    """Resolve the video path for a participant."""
    for p in _participants:
        if p["id"] == participant_id:
            if p["has_video"]:
                return p["video_path"]
            return None
    return None


def _get_video_cap(video_path: str) -> Optional["cv2.VideoCapture"]:
    """Return a cached VideoCapture for *video_path*, opening a new one if needed."""
    import cv2

    if video_path in _video_cap_cache:
        cap = _video_cap_cache[video_path]
        if cap.isOpened():
            _video_cap_cache.move_to_end(video_path)
            return cap
        # Stale entry
        cap.release()
        del _video_cap_cache[video_path]
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return None
    _video_cap_cache[video_path] = cap
    if len(_video_cap_cache) > _VIDEO_CAP_MAX:
        _, old_cap = _video_cap_cache.popitem(last=False)
        old_cap.release()
    return cap


@screenspace_bp.route("/api/video/frame/<participant>/<timestamp>")
def api_video_frame(participant: str, timestamp: str) -> FlaskResponse:
    """Extract and return a single JPEG frame at the given timestamp."""
    try:
        ts = float(timestamp)
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "Invalid timestamp"}), 400

    video_path = _find_participant_video(participant)
    if video_path is None:
        return jsonify(
            {"ok": False, "error": f"No video for participant {participant}"}
        ), 404

    import cv2

    with _video_cap_lock:
        cap = _get_video_cap(video_path)
        if cap is None:
            return jsonify({"ok": False, "error": "Could not open video file"}), 500

        cap.set(cv2.CAP_PROP_POS_MSEC, ts * 1000.0)
        ret, frame = cap.read()

    if not ret:
        return jsonify({"ok": False, "error": "Could not read frame at timestamp"}), 400

    _, jpeg = cv2.imencode(".jpg", frame, [cv2.IMWRITE_JPEG_QUALITY, 85])
    return Response(
        jpeg.tobytes(),
        mimetype="image/jpeg",
        headers={"Cache-Control": "public, max-age=60"},
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

    import cv2

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return jsonify({"ok": False, "error": "Could not open video file"}), 500

    vid_fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
    duration = round(total_frames / vid_fps) if vid_fps > 0 else 0
    cap.release()

    info: Dict[str, Any] = {
        "participant": participant,
        "duration": duration,
        "fps": vid_fps,
        "width": width if width > 0 else None,
        "height": height if height > 0 else None,
    }
    _video_metadata_cache[participant] = info

    return jsonify({"ok": True, "info": info})


# ---- Region coordinate normalization ----


def _normalize_region(
    x: float,
    y: float,
    w: float,
    h: float,
    frame_w: int,
    frame_h: int,
) -> Dict[str, Any]:
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
    region: Dict[str, Any], target_w: int, target_h: int
) -> Dict[str, int]:
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
    region: Dict[str, Any]
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

    import screenspace

    _manifest.setdefault("regions", {})[name] = region
    screenspace.save_screenspace_manifest(
        _manifest.get("regions", {}),
        _manifest.get("tasks", []),
        _manifest.get("events", []),
        stashes=_manifest.get("stashes", []),
    )

    return jsonify({"ok": True, "region": region})


@screenspace_bp.route("/api/regions/<name>", methods=["DELETE"])
def api_regions_delete(name: str) -> FlaskResponse:
    """Delete a region definition."""
    regions = _manifest.get("regions", {})
    if name not in regions:
        return jsonify({"ok": False, "error": f"Region '{name}' not found"}), 404

    import screenspace

    del regions[name]
    screenspace.save_screenspace_manifest(
        regions,
        _manifest.get("tasks", []),
        _manifest.get("events", []),
        stashes=_manifest.get("stashes", []),
    )

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
    import screenspace

    _manifest.setdefault("stashes", []).append(stash)
    _manifest["regions"] = {}
    screenspace.save_screenspace_manifest(
        _manifest.get("regions", {}),
        _manifest.get("tasks", []),
        _manifest.get("events", []),
        stashes=_manifest.get("stashes", []),
    )
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

    import screenspace

    name = data.get("name", "").strip()
    if name:
        stash["name"] = name

    screenspace.save_screenspace_manifest(
        _manifest.get("regions", {}),
        _manifest.get("tasks", []),
        _manifest.get("events", []),
        stashes=stashes,
    )
    return jsonify({"ok": True, "stash": stash})


@screenspace_bp.route("/api/stashes/<stash_id>", methods=["DELETE"])
def api_stashes_delete(stash_id: str) -> FlaskResponse:
    """Dismiss a region stash."""
    stashes = _manifest.get("stashes", [])
    idx = next((i for i, s in enumerate(stashes) if s["id"] == stash_id), None)
    if idx is None:
        return jsonify({"ok": False, "error": "Stash not found"}), 404

    import screenspace

    stashes.pop(idx)
    screenspace.save_screenspace_manifest(
        _manifest.get("regions", {}),
        _manifest.get("tasks", []),
        _manifest.get("events", []),
        stashes=stashes,
    )
    return jsonify({"ok": True})


@screenspace_bp.route("/api/stashes/<stash_id>/restore", methods=["POST"])
def api_stashes_restore(stash_id: str) -> FlaskResponse:
    """Restore a stash: replace active regions with stashed ones, remove stash."""
    stashes = _manifest.get("stashes", [])
    idx = next((i for i, s in enumerate(stashes) if s["id"] == stash_id), None)
    if idx is None:
        return jsonify({"ok": False, "error": "Stash not found"}), 404

    import screenspace

    stash = stashes.pop(idx)
    _manifest["regions"] = copy.deepcopy(stash["regions"])
    screenspace.save_screenspace_manifest(
        _manifest["regions"],
        _manifest.get("tasks", []),
        _manifest.get("events", []),
        stashes=stashes,
    )
    return jsonify({"ok": True, "regions": _manifest["regions"]})


# ---- Tasks CRUD ----


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


@screenspace_bp.route("/api/tasks", methods=["POST"])
def api_tasks_create() -> FlaskResponse:
    """Enqueue a new analysis task."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500

    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "JSON body required"}), 400

    task_type = data.get("type", "").strip()
    valid_types = (
        "color",
        "change",
        "similarity",
        "text",
        "numbers",
        "timelapse",
        "template",
        "flow",
        "scene",
    )
    if task_type not in valid_types:
        return jsonify(
            {"ok": False, "error": f"type must be one of: {', '.join(valid_types)}"}
        ), 400

    participant = data.get("participant", "").strip()
    if not participant:
        return jsonify({"ok": False, "error": "participant is required"}), 400

    region_name = data.get("region", "").strip()

    # Template tasks with an uploaded image scan the full frame; no region needed
    has_uploaded_template = task_type == "template" and data.get("parameters", {}).get(
        "template_image_data"
    )

    if not region_name and not has_uploaded_template:
        return jsonify({"ok": False, "error": "region is required"}), 400

    regions: Dict[str, Any] = {}
    if region_name:
        regions = _manifest.get("regions", {})
        if region_name not in regions:
            for stash in _manifest.get("stashes", []):
                if region_name in stash.get("regions", {}):
                    regions = stash["regions"]
                    break
        if region_name not in regions:
            return jsonify(
                {"ok": False, "error": f"Region '{region_name}' not found"}
            ), 400

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

    if region_name:
        region_data = regions[region_name]
        props = video.probe_video_properties(video_path)
        if props and props.get("width") and props.get("height"):
            region_coords = _denormalize_region(
                region_data, props["width"], props["height"]
            )
        else:
            region_coords = {
                k: region_data[k] for k in ("x", "y", "w", "h") if k in region_data
            }
    else:
        # Full-frame template scan -- sentinel region
        region_name = "full_frame"
        region_coords = {"x": 0, "y": 0, "w": 0, "h": 0}
    parameters = data.get("parameters", {})

    import screenspace

    # Similarity tasks: extract reference frame from video at given timestamp
    if task_type == "similarity":
        import cv2

        ref_ts = parameters.get("reference_timestamp")
        if ref_ts is None:
            return jsonify(
                {"ok": False, "error": "Similarity scan requires reference_timestamp"}
            ), 400
        cap = cv2.VideoCapture(video_path)
        if not cap.isOpened():
            return jsonify(
                {"ok": False, "error": "Could not open video for reference frame"}
            ), 500
        cap.set(cv2.CAP_PROP_POS_MSEC, float(ref_ts) * 1000.0)
        ret, frame = cap.read()
        cap.release()
        if not ret:
            return jsonify(
                {"ok": False, "error": "Could not read reference frame"}
            ), 400
        ref_region = screenspace.extract_region(frame, region_coords)
        parameters["reference_frame"] = ref_region

    # Template tasks: use uploaded PNG or extract template region from video
    if task_type == "template":
        import base64

        import cv2
        import numpy as np

        upload_b64 = parameters.pop("template_image_data", None)
        if upload_b64:
            # Decode uploaded PNG (may have alpha channel for masking)
            try:
                img_bytes = base64.b64decode(upload_b64)
                img_arr = np.frombuffer(img_bytes, dtype=np.uint8)
                img = cv2.imdecode(img_arr, cv2.IMREAD_UNCHANGED)
            except Exception:
                return jsonify(
                    {"ok": False, "error": "Could not decode uploaded image"}
                ), 400
            if img is None:
                return jsonify({"ok": False, "error": "Invalid image data"}), 400
            if len(img.shape) == 2:
                # Grayscale → convert to BGR
                img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
            if img.shape[2] == 4:
                # Extract alpha as mask, convert to BGR for template
                parameters["template_mask"] = img[:, :, 3]
                parameters["template_image"] = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
            else:
                parameters["template_image"] = img
        else:
            ref_ts = parameters.get("reference_timestamp")
            if ref_ts is None:
                return jsonify(
                    {
                        "ok": False,
                        "error": "Template scan requires reference_timestamp or uploaded image",
                    }
                ), 400
            cap = cv2.VideoCapture(video_path)
            if not cap.isOpened():
                return jsonify(
                    {"ok": False, "error": "Could not open video for template capture"}
                ), 500
            cap.set(cv2.CAP_PROP_POS_MSEC, float(ref_ts) * 1000.0)
            ret, frame = cap.read()
            cap.release()
            if not ret:
                return jsonify(
                    {"ok": False, "error": "Could not read template frame"}
                ), 400
            parameters["template_image"] = screenspace.extract_region(
                frame, region_coords
            )

    # Scene tasks: extract reference frame for each scene type
    if task_type == "scene":
        import cv2

        scene_refs = parameters.get("scene_references")
        if not scene_refs or not isinstance(scene_refs, list):
            return jsonify(
                {"ok": False, "error": "Scene scan requires scene_references list"}
            ), 400
        cap = cv2.VideoCapture(video_path)
        if not cap.isOpened():
            return jsonify(
                {"ok": False, "error": "Could not open video for scene references"}
            ), 500
        reference_scenes = []
        for ref in scene_refs:
            cap.set(cv2.CAP_PROP_POS_MSEC, float(ref["timestamp"]) * 1000.0)
            ret, frame = cap.read()
            if not ret:
                cap.release()
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
        cap.release()
        parameters["reference_scenes"] = reference_scenes

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
    _manifest.setdefault("tasks", []).append(task)

    return jsonify({"ok": True, "task": _clean_task(task)})


@screenspace_bp.route("/api/tasks/<task_id>", methods=["DELETE"])
def api_tasks_cancel(task_id: str) -> FlaskResponse:
    """Cancel or dismiss a task.  ?dismiss=true fully removes the task."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500
    if request.args.get("dismiss") == "true":
        if not _worker.remove_task(task_id):
            return jsonify({"ok": False, "error": "Task not found"}), 404
        _manifest["tasks"] = [
            t for t in _manifest.get("tasks", []) if t["id"] != task_id
        ]
        _persist_manifest()
        return jsonify({"ok": True})
    if _worker.cancel(task_id):
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
    return jsonify({"ok": True})


@screenspace_bp.route("/api/tasks/pause", methods=["POST"])
def api_tasks_pause() -> FlaskResponse:
    """Pause the task queue."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500
    _worker.pause()
    return jsonify({"ok": True, "paused": True})


@screenspace_bp.route("/api/tasks/resume", methods=["POST"])
def api_tasks_resume() -> FlaskResponse:
    """Resume the task queue."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500
    _worker.resume()
    return jsonify({"ok": True, "paused": False})


@screenspace_bp.route("/api/tasks/<task_id>/results")
def api_tasks_results(task_id: str) -> FlaskResponse:
    """Get task results (timestamps, artifacts)."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500
    task = _worker.get_task(task_id)
    if task is None:
        return jsonify({"ok": False, "error": "Task not found"}), 404
    return jsonify({"ok": True, "results": task.get("result")})


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
    return jsonify({"ok": True, "events": events})


@screenspace_bp.route("/api/events/<event_id>/exclude", methods=["PUT"])
def api_event_exclude(event_id: str) -> FlaskResponse:
    """Set an event as excluded."""
    for e in _manifest.get("events", []):
        if e["id"] == event_id:
            e["excluded"] = True
            _persist_manifest()
            return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Event not found"}), 404


@screenspace_bp.route("/api/events/<event_id>/include", methods=["PUT"])
def api_event_include(event_id: str) -> FlaskResponse:
    """Set an event as included."""
    for e in _manifest.get("events", []):
        if e["id"] == event_id:
            e["excluded"] = False
            _persist_manifest()
            return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Event not found"}), 404


@screenspace_bp.route("/api/events/bulk-exclude", methods=["PUT"])
def api_events_bulk_exclude() -> FlaskResponse:
    """Bulk-exclude events by ID list."""
    data = request.get_json(silent=True) or {}
    ids = set(data.get("ids", []))
    if not ids:
        return jsonify({"ok": False, "error": "ids list required"}), 400
    count = 0
    for e in _manifest.get("events", []):
        if e["id"] in ids:
            e["excluded"] = True
            count += 1
    _persist_manifest()
    return jsonify({"ok": True, "updated": count})


@screenspace_bp.route("/api/events/bulk-include", methods=["PUT"])
def api_events_bulk_include() -> FlaskResponse:
    """Bulk-include events by ID list."""
    data = request.get_json(silent=True) or {}
    ids = set(data.get("ids", []))
    if not ids:
        return jsonify({"ok": False, "error": "ids list required"}), 400
    count = 0
    for e in _manifest.get("events", []):
        if e["id"] in ids:
            e["excluded"] = False
            count += 1
    _persist_manifest()
    return jsonify({"ok": True, "updated": count})


# ---- Helpers ----


def _sanitize_floats(obj: Any) -> Any:
    """Replace non-finite floats (inf, nan) with None for JSON safety."""
    if isinstance(obj, float) and not math.isfinite(obj):
        return None
    if isinstance(obj, dict):
        return {k: _sanitize_floats(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize_floats(v) for v in obj]
    return obj


def _clean_task(task: Dict[str, Any]) -> Dict[str, Any]:
    """Remove internal fields from a task dict for API responses."""
    cleaned = {k: v for k, v in task.items() if not k.startswith("_")}
    if "parameters" in cleaned:
        cleaned["parameters"] = {
            k: v
            for k, v in cleaned["parameters"].items()
            if k
            not in (
                "reference_frame",
                "template_image",
                "template_mask",
                "reference_scenes",
            )
        }
    return _sanitize_floats(cleaned)


# ---- State initialization ----


def _init_screenspace_state(
    sheet_context: Any = None,
    participant_list: Optional[List[str]] = None,
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
    _worker.restore_tasks(_manifest.get("tasks", []))
    _worker.start()


def _discover_participant_videos(study_name: str) -> None:
    """Scan input directory for source video files and populate _participants."""
    global _participants
    input_dir = Path(utils.get_effective_input_dir())
    if not input_dir.is_dir():
        return
    for path in sorted(input_dir.glob(f"*{config.FILEFORMAT}")):
        name = path.stem
        parts = name.rsplit("_", 1)
        if len(parts) == 2:
            pid = parts[1]
            if pid and pid[0] in config.PARTICIPANT_PREFIXES:
                _participants.append(
                    {
                        "id": pid,
                        "video_path": str(path),
                        "has_video": True,
                    }
                )


def _persist_manifest() -> None:
    """Save manifest after a task completes."""
    import screenspace

    if _worker:
        new_events = _worker.drain_new_events()
        if new_events:
            _manifest.setdefault("events", []).extend(new_events)
        tasks = _worker.get_all_tasks()
        screenspace.save_screenspace_manifest(
            _manifest.get("regions", {}),
            tasks,
            _manifest.get("events", []),
            stashes=_manifest.get("stashes", []),
        )
