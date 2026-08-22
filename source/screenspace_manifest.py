"""Screenspace task + manifest helpers."""

import uuid
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import utils
from screenspace_tools import _extract_confidence


# ---------------------------------------------------------------------------
# Task construction
# ---------------------------------------------------------------------------

TASK_STATUS_QUEUED = "queued"
TASK_STATUS_RUNNING = "running"
TASK_STATUS_COMPLETED = "completed"
TASK_STATUS_FAILED = "failed"
TASK_STATUS_CANCELLED = "cancelled"
TASK_STATUS_PAUSED = "paused"

# Task parameter keys carrying binary payloads (base64 frames/templates) that
# must never reach the manifest on disk or JSON API responses.
TASK_BINARY_KEYS = (
    "reference_frame",
    "template_image",
    "template_mask",
    "reference_scenes",
)

_SENTINEL = object()


def strip_task_param_binaries(params: dict[str, Any]) -> dict[str, Any]:
    """Copy of task ``parameters`` without ``TASK_BINARY_KEYS``, also stripping
    binaries + internal ``region_coords`` from multitool ``steps``. Shared by
    manifest writes and API responses (``screenspace_server._clean_task``)."""
    params = {k: v for k, v in params.items() if k not in TASK_BINARY_KEYS}
    if "steps" in params:
        step_strip_keys = TASK_BINARY_KEYS + ("region_coords",)
        params["steps"] = [
            {k: v for k, v in s.items() if k not in step_strip_keys}
            for s in params["steps"]
        ]
    return params


# OpenCV-style HSV hue buckets (h in 0-179, wraparound at 180) for color-task
# names. Each entry is (upper_bound_exclusive, name); red owns both ends.
_HUE_BUCKETS = [
    (10, "red"),
    (22, "orange"),
    (35, "yellow"),
    (78, "green"),
    (100, "cyan"),
    (128, "blue"),
    (145, "purple"),
    (160, "pink"),
    (180, "red"),
]


def _hue_bucket_name(color: dict[str, Any]) -> str:
    """Human color word for an OpenCV HSV dict (h 0-179, s/v 0-255)."""
    h = float(color.get("h", 0))
    s = float(color.get("s", 0))
    v = float(color.get("v", 0))
    if v < 46:
        return "black"
    if s < 40:
        return "white" if v > 200 else "gray"
    for upper, name in _HUE_BUCKETS:
        if h < upper:
            return name
    return "red"


def _fmt_num(value: Any) -> str:
    """Format a number without a trailing .0 (100.0 -> '100', 0.5 -> '0.5')."""
    return f"{float(value):g}"


def _describe(task_type: str, params: dict[str, Any]) -> str:
    """Tool label + distinguishing parameter, without the region suffix."""
    label = task_type.capitalize()
    if task_type == "color":
        target = params.get("target_color")
        if isinstance(target, dict):
            return f"{label}: {_hue_bucket_name(target)}"
    elif task_type == "change":
        if "threshold" in params:
            return f"{label} ≥{_fmt_num(round(float(params['threshold']) * 100, 1))}%"
    elif task_type == "similarity":
        if params.get("reference_timestamp") is not None:
            ts = utils.seconds_to_timestamp(float(params["reference_timestamp"]))
            return f"{label} to {ts}"
    elif task_type == "text":
        search = str(params.get("search_string", "")).strip()
        if search:
            if len(search) > 24:
                search = search[:24] + "…"
            return f'{label} "{search}"'
    elif task_type == "numbers":
        op = params.get("operator", "")
        if op == "range":
            if (
                params.get("range_min") is not None
                and params.get("range_max") is not None
            ):
                return f"{label} {_fmt_num(params['range_min'])}–{_fmt_num(params['range_max'])}"
        elif (
            op in ("eq", "gt", "lt", "gte", "lte")
            and params.get("target_value") is not None
        ):
            sym = {"eq": "=", "gt": ">", "lt": "<", "gte": "≥", "lte": "≤"}[op]
            return f"{label} {sym} {_fmt_num(params['target_value'])}"
    elif task_type == "template":
        if params.get("template_name"):
            return f"{label}: {params['template_name']}"
        if params.get("reference_timestamp") is not None:
            ts = utils.seconds_to_timestamp(float(params["reference_timestamp"]))
            return f"{label} @ {ts}"
    elif task_type == "flow":
        if "magnitude_threshold" in params:
            return f"{label} ≥{_fmt_num(params['magnitude_threshold'])}"
    elif task_type == "scene":
        names = [
            str(ref.get("name", "")).strip()
            for ref in params.get("scene_references", [])
            if isinstance(ref, dict) and str(ref.get("name", "")).strip()
        ]
        if names:
            shown = ", ".join(names[:2])
            extra = f" +{len(names) - 2}" if len(names) > 2 else ""
            return f"{label}: {shown}{extra}"
    elif task_type == "inactivity":
        if "min_duration" in params:
            return f"{label} ≥{_fmt_num(params['min_duration'])}s"
    elif task_type == "attention":
        if "shift_threshold" in params:
            return f"{label} Δ≥{_fmt_num(params['shift_threshold'])}"
    elif task_type == "timelapse":
        speedup = params.get("speedup_factor")
        if speedup is not None:
            fmt = str(params.get("output_format", "mp4")).upper()
            return f"{label} {_fmt_num(speedup)}× {fmt}"
    elif task_type == "multitool":
        step_labels = [
            str(step.get("type", "")).capitalize()
            for step in params.get("steps", [])
            if isinstance(step, dict) and step.get("type")
        ]
        if step_labels:
            return f"{label}: {' + '.join(step_labels)}"
    return label


def describe_task(task_type: str, region_name: str, parameters: dict[str, Any]) -> str:
    """Build a descriptive display name from a task's distinguishing params.

    e.g. 'Text "checkout" · header', 'Color: blue · HUD', 'Numbers > 100'. Total:
    malformed params degrade to the capitalized tool label rather than raising.
    A user-supplied event_label still overrides this wherever it is shown.
    """
    try:
        name = _describe(task_type, parameters)
    except Exception:
        name = task_type.capitalize()
    if region_name and region_name not in ("full_frame", "per_step"):
        name += f" · {region_name}"
    return name


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
        "name": describe_task(task_type, region_name, parameters or {}),
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
        "created_at": datetime.now(UTC).isoformat(),
        "completed_at": None,
        "_cancelled": False,
    }


# ---------------------------------------------------------------------------
# Manifest persistence
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
    return utils.load_manifest_section(
        "screenspace", default=_empty_screenspace_manifest()
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
            ct["parameters"] = strip_task_param_binaries(ct["parameters"])
        # Strip large per-frame heatmap grids from results (not needed on disk)
        if isinstance(ct.get("result"), list):
            _grid_keys = ("flow_grid", "change_grid", "saliency_grid")
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
        utils.save_manifest_section("screenspace", None)
        return None
    return utils.save_manifest_section("screenspace", payload)


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
    if task_type == "attention":
        # Events come from the shift-only on_result stream; this guards the
        # regeneration paths that could hand over the full per-sample list.
        raw_results = [r for r in raw_results if r.get("shift")]
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
        elif task_type == "attention":
            metadata["shift_distance"] = r.get("shift_distance", 0.0)
            metadata["from_x"] = r.get("from_x", 0.0)
            metadata["from_y"] = r.get("from_y", 0.0)
            metadata["to_x"] = r.get("to_x", r.get("peak_x", 0.0))
            metadata["to_y"] = r.get("to_y", r.get("peak_y", 0.0))
            metadata["peak_value"] = r.get("peak_value", 0.0)
        elif task_type == "boundary":
            metadata["distance"] = r.get("distance", 0.0)
            # Scene/hybrid metrics emit the period each boundary opens (absent for
            # phash), so Studio/Viewer can render segments instead of bare ticks.
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
            # Orientation, not clip candidacy — the flag lets Studio intake hide
            # boundaries by default.
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
        event_label = task.get("name") or task["type"] + ": " + task.get("region", "")
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
