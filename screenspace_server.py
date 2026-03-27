# -*- coding: utf-8 -*-
"""Screenspace web server for clipgen.

Serves the Screenspace front-end and exposes REST endpoints for
region management, task execution, video frame extraction, and media serving.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import cv2
from flask import Blueprint, Response, jsonify, request, send_from_directory

import config
import files
import screenspace
import utils
import video

FlaskResponse = Union[Response, Tuple[Response, int]]

# ---- Module-level state (set once by _init_screenspace_state) ----

_manifest: Dict[str, Any] = {}
_worker: Optional[screenspace.ScreenspaceWorker] = None
_output_dir: str = ""
_participants: List[Dict[str, Any]] = []

_assets_dir = Path(__file__).resolve().parent / "assets" / "web"

# ---- Blueprint ----

screenspace_bp = Blueprint("screenspace", __name__)


# ---- Static file serving ----


@screenspace_bp.route("/")
def serve_index() -> FlaskResponse:
    return send_from_directory(_assets_dir, "screenspace.html")


@screenspace_bp.route("/<path:filename>")
def serve_static(filename: str) -> FlaskResponse:
    return send_from_directory(_assets_dir, filename)


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


@screenspace_bp.route("/api/video/frame/<participant>/<float:timestamp>")
def api_video_frame(participant: str, timestamp: float) -> FlaskResponse:
    """Extract and return a single JPEG frame at the given timestamp."""
    video_path = _find_participant_video(participant)
    if video_path is None:
        return jsonify(
            {"ok": False, "error": f"No video for participant {participant}"}
        ), 404

    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return jsonify({"ok": False, "error": "Could not open video file"}), 500

    cap.set(cv2.CAP_PROP_POS_MSEC, timestamp * 1000.0)
    ret, frame = cap.read()
    cap.release()

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
    video_path = _find_participant_video(participant)
    if video_path is None:
        return jsonify(
            {"ok": False, "error": f"No video for participant {participant}"}
        ), 404

    duration = video.get_file_duration(video_path)
    props = video.probe_video_properties(video_path)

    info: Dict[str, Any] = {"participant": participant, "duration": duration}
    if props:
        info["width"] = props.get("width")
        info["height"] = props.get("height")
        info["video_codec"] = props.get("video_codec")

    cap = cv2.VideoCapture(video_path)
    if cap.isOpened():
        info["fps"] = cap.get(cv2.CAP_PROP_FPS)
        cap.release()

    return jsonify({"ok": True, "info": info})


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

    region = {
        "x": int(data["x"]),
        "y": int(data["y"]),
        "w": int(data["w"]),
        "h": int(data["h"]),
    }
    if "description" in data:
        region["description"] = str(data["description"])

    _manifest.setdefault("regions", {})[name] = region
    screenspace.save_screenspace_manifest(
        _manifest.get("regions", {}), _manifest.get("tasks", [])
    )

    return jsonify({"ok": True, "region": region})


@screenspace_bp.route("/api/regions/<name>", methods=["DELETE"])
def api_regions_delete(name: str) -> FlaskResponse:
    """Delete a region definition."""
    regions = _manifest.get("regions", {})
    if name not in regions:
        return jsonify({"ok": False, "error": f"Region '{name}' not found"}), 404

    del regions[name]
    screenspace.save_screenspace_manifest(regions, _manifest.get("tasks", []))

    return jsonify({"ok": True})


# ---- Tasks CRUD ----


@screenspace_bp.route("/api/tasks")
def api_tasks_list() -> FlaskResponse:
    """List all tasks with status and progress."""
    tasks = _worker.get_all_tasks() if _worker else []
    clean = [_clean_task(t) for t in tasks]
    return jsonify({"ok": True, "tasks": clean})


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
    valid_types = ("color", "change", "similarity", "text", "timelapse")
    if task_type not in valid_types:
        return jsonify(
            {"ok": False, "error": f"type must be one of: {', '.join(valid_types)}"}
        ), 400

    participant = data.get("participant", "").strip()
    if not participant:
        return jsonify({"ok": False, "error": "participant is required"}), 400

    region_name = data.get("region", "").strip()
    if not region_name:
        return jsonify({"ok": False, "error": "region is required"}), 400

    regions = _manifest.get("regions", {})
    if region_name not in regions:
        return jsonify({"ok": False, "error": f"Region '{region_name}' not found"}), 400

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

    region_coords = regions[region_name]
    parameters = data.get("parameters", {})

    # Similarity tasks: extract reference frame from video at given timestamp
    if task_type == "similarity":
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
        parameters.pop("reference_timestamp", None)

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
    """Cancel a queued or running task."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500
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


@screenspace_bp.route("/api/tasks/<task_id>/results")
def api_tasks_results(task_id: str) -> FlaskResponse:
    """Get task results (timestamps, artifacts)."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500
    task = _worker.get_task(task_id)
    if task is None:
        return jsonify({"ok": False, "error": "Task not found"}), 404
    return jsonify({"ok": True, "results": task.get("result")})


# ---- Helpers ----


def _clean_task(task: Dict[str, Any]) -> Dict[str, Any]:
    """Remove internal fields from a task dict for API responses."""
    cleaned = {k: v for k, v in task.items() if not k.startswith("_")}
    if "parameters" in cleaned:
        cleaned["parameters"] = {
            k: v for k, v in cleaned["parameters"].items() if k != "reference_frame"
        }
    return cleaned


# ---- State initialization ----


def _init_screenspace_state(
    sheet_context: Any = None,
    participant_list: Optional[List[str]] = None,
) -> None:
    """Initialize module-level state for Screenspace routes.

    Loads manifest, resolves participant video paths, and starts the
    background worker thread.
    """
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
    if _worker:
        tasks = _worker.get_all_tasks()
        screenspace.save_screenspace_manifest(_manifest.get("regions", {}), tasks)
