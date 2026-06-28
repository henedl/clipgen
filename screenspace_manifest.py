# -*- coding: utf-8 -*-
"""Screenspace task + manifest helpers.

Task-status constants, task construction, manifest load/save, result-time
offsetting for multi-video scans, and event generation from raw results.
Imports the confidence extractor from screenspace_tools.
"""

import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import config
import utils
from screenspace_tools import _extract_confidence


# ---------------------------------------------------------------------------
# Task queue and worker
# ---------------------------------------------------------------------------

TASK_STATUS_QUEUED = "queued"
TASK_STATUS_RUNNING = "running"
TASK_STATUS_COMPLETED = "completed"
TASK_STATUS_FAILED = "failed"
TASK_STATUS_CANCELLED = "cancelled"
TASK_STATUS_PAUSED = "paused"

_SENTINEL = object()


def create_task(
    task_type: str,
    participant: str,
    source_video: str,
    video_paths: list[str],
    region_name: str,
    region_coords: dict[str, int],
    parameters: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Build a new task dict with all fields initialized.

    *video_paths* is the participant's ordered source video(s): one entry for a
    normal participant, several for a multi-video participant whose session spans
    files (analyzed as one continuous timeline; the worker maps the task's global
    time range into the owning sub-video per part).
    """
    return {
        "id": f"ss_{uuid.uuid4().hex[:8]}",
        "type": task_type,
        "participant": participant,
        "source_video": source_video,
        "video_paths": video_paths,
        "region": region_name,
        "region_coords": region_coords,
        "parameters": parameters or {},
        "status": TASK_STATUS_QUEUED,
        "progress": 0.0,
        "priority": 100,
        "result": None,
        "error": None,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "completed_at": None,
        "_cancelled": False,
    }


# ---------------------------------------------------------------------------
# Heatmap generation
# ---------------------------------------------------------------------------


def _empty_screenspace_manifest() -> dict[str, Any]:
    return {
        "regions": {},
        "tasks": [],
        "events": [],
        "stashes": [],
        "per_participant": {},
        "pins": {},
    }


def load_screenspace_manifest() -> dict[str, Any]:
    """Load the screenspace manifest from the output directory."""
    return utils.load_json_manifest(
        config.SCREENSPACE_MANIFEST_FILENAME, default=_empty_screenspace_manifest()
    )


def _is_empty_screenspace_manifest(payload: dict[str, Any]) -> bool:
    """True when no regions, tasks, events, stashes, per-participant data, or pins
    exist — i.e. nothing worth writing into the output dir."""
    return not (
        payload.get("regions")
        or payload.get("tasks")
        or payload.get("events")
        or payload.get("stashes")
        or payload.get("per_participant")
        or payload.get("pins")
    )


def save_screenspace_manifest(
    regions: dict[str, dict[str, Any]],
    tasks: list[dict[str, Any]],
    events: list[dict[str, Any]] | None = None,
    stashes: list[dict[str, Any]] | None = None,
    per_participant: dict[str, dict[str, Any]] | None = None,
    pins: dict[str, list[dict[str, Any]]] | None = None,
) -> Path | None:
    """Write the screenspace manifest to disk.

    Strips internal fields (prefixed with ``_``) from tasks before writing.
    Returns the manifest path on success, or ``None`` on failure.
    """
    clean_tasks = []
    for task in tasks:
        ct = {k: v for k, v in task.items() if not k.startswith("_")}
        if "parameters" in ct:
            _binary_keys = (
                "reference_frame",
                "template_image",
                "template_mask",
                "reference_scenes",
            )
            ct["parameters"] = {
                k: v for k, v in ct["parameters"].items() if k not in _binary_keys
            }
            # Strip binary data and internal coords from multitool step parameters
            if "steps" in ct["parameters"]:
                _step_strip_keys = _binary_keys + ("region_coords",)
                ct["parameters"]["steps"] = [
                    {k: v for k, v in s.items() if k not in _step_strip_keys}
                    for s in ct["parameters"]["steps"]
                ]
        # Strip large per-frame heatmap grids from results (not needed on disk)
        if isinstance(ct.get("result"), list):
            _grid_keys = ("flow_grid", "change_grid")
            ct["result"] = [
                {k: v for k, v in r.items() if k not in _grid_keys}
                for r in ct["result"]
            ]
        clean_tasks.append(ct)
    if pins is None:
        existing = load_screenspace_manifest()
        existing_pins = existing.get("pins")
        pins_payload = existing_pins if isinstance(existing_pins, dict) else {}
    else:
        pins_payload = pins

    payload = utils.sanitize_floats(
        {
            "regions": regions,
            "tasks": clean_tasks,
            "events": events or [],
            "stashes": stashes or [],
            "per_participant": per_participant or {},
            "pins": pins_payload,
        }
    )
    if _is_empty_screenspace_manifest(payload):
        utils.remove_json_manifest(config.SCREENSPACE_MANIFEST_FILENAME)
        return None
    return utils.save_json_manifest(
        config.SCREENSPACE_MANIFEST_FILENAME,
        payload,
        warn_label="screenspace manifest",
    )


def _offset_result_times(result: dict[str, Any], offset: int) -> None:
    """Shift a scan result's time fields by *offset* seconds (in place).

    Maps a sub-video's local result times back onto the participant's global
    timeline for multi-video scans. Covers point events (``timestamp``) and span
    events (``start``/``end``, e.g. inactivity). A no-op when offset is 0.
    """
    if not offset:
        return
    for key in ("timestamp", "start", "end"):
        value = result.get(key)
        if isinstance(value, (int, float)):
            result[key] = value + offset


def generate_events_from_results(
    task: dict[str, Any], raw_results: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    """Convert raw task results into ScreenspaceEvent records."""
    task_type = task.get("type", "")
    if task_type == "timelapse":
        return []
    events: list[dict[str, Any]] = []
    for r in raw_results:
        ts = r.get("timestamp", r.get("start", 0.0))
        confidence = _extract_confidence(task_type, r)
        metadata: dict[str, Any] = {}
        if task_type == "change":
            metadata["magnitude"] = r.get("magnitude", 0.0)
        elif task_type == "similarity":
            metadata["score"] = r.get("score", 0.0)
        elif task_type == "text":
            metadata["text_found"] = r.get("text_found", "")
        elif task_type == "numbers":
            metadata["value"] = r.get("number_found", 0)
        elif task_type == "template":
            metadata["match_count"] = r.get("match_count", 0)
            metadata["best_score"] = r.get("best_score", 0.0)
        elif task_type == "flow":
            metadata["magnitude"] = r.get("magnitude", 0.0)
            metadata["angle"] = r.get("angle", 0.0)
        elif task_type == "scene":
            metadata["scene_name"] = r.get("scene_name", "")
            metadata["score"] = r.get("score", 0.0)
        elif task_type == "multitool":
            metadata["tool_types"] = r.get("tool_types", [])
            metadata["steps"] = r.get("steps", [])
        elif task_type == "inactivity":
            metadata["duration"] = r.get("duration", 0.0)
            metadata["avg_distance"] = r.get("avg_distance", 0.0)
        elif task_type == "boundary":
            metadata["distance"] = r.get("distance", 0.0)
            # Scene/hybrid metrics emit the period each boundary opens; absent
            # for the phash metric. Carried so Studio/Viewer can later render
            # segments instead of bare ticks.
            if "period_start" in r:
                metadata["period_start"] = r.get("period_start")
            if "period_end" in r:
                metadata["period_end"] = r.get("period_end")
            if "scene_label" in r:
                metadata["scene_label"] = r.get("scene_label")
        ev = create_event(task, ts, confidence, metadata)
        # Multi-video scans tag each result with the sub-video it came from.
        source_override = r.get("_source_video")
        if source_override:
            ev["source_video"] = source_override
        if task_type == "inactivity" and "end" in r:
            ev["time_out"] = round(r["end"], 2)
        if task_type == "boundary":
            # Boundaries are for orientation, not clip candidacy. The
            # navigational flag lets Studio intake hide them by default.
            ev["navigational"] = True
        events.append(ev)
    return events


def create_event(
    task: dict[str, Any],
    timestamp: float,
    confidence: float,
    metadata: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Create a ScreenspaceEvent from a task result entry."""
    event_label = task.get("parameters", {}).get("event_label", "")
    if not event_label:
        event_label = task["type"] + ": " + task.get("region", "")
    return {
        "id": f"ev_{uuid.uuid4().hex[:8]}",
        "source_video": task.get("source_video", ""),
        "participant": task.get("participant", ""),
        "detector": task["type"],
        "event_type": event_label,
        "time_in": round(timestamp, 2),
        "time_out": round(timestamp, 2),
        "confidence": round(max(0.0, min(1.0, confidence)), 4),
        "metadata": metadata or {},
        "excluded": False,
        "task_id": task["id"],
        "region": task.get("region", ""),
    }
