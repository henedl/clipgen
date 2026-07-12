"""Composer Flask blueprint — source-video cutting timeline + local manifest.

Registered at ``/composer`` by ``server.build_combined_app`` (mutually exclusive
launch with the other web modes, but all blueprints are always mounted). Phase 1
ships the static page routes, participant/video discovery, and the composer
manifest CRUD: in/out **cut pairs** plus persisted UI state (marker-lane
toggles). Trims of existing source markers (P2) and visual annotations with
burn-in export (P3) extend this module later.

All Composer state lives in ``composer_manifest.json`` in the output dir
(load-on-startup, save-after-mutations). Composer never writes to the
spreadsheet — cut pairs feed clip generation through Studio's
``/api/generate-intake`` endpoint, which takes raw start/end seconds.

Times are source-video **global seconds**: a multi-part participant (a session
recorded across several files) is addressed on the stitched timeline, and the
frontend maps global time onto parts using the ``parts[].offset`` values from
``GET /api/participants``.

Module-level state (``_input_dir``, ``_sheet_context``, ``_manifest``) is
initialized by :func:`_init_composer_state`, mirroring the other blueprints.
Mutations hold ``_manifest_lock`` and persist via :func:`_persist_locked`.
"""

from __future__ import annotations

import copy
import json
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from flask import Blueprint, request

import config
import utils
import video
from server_utils import err, ok

# ---- Module state (initialized by _init_composer_state) ----

_input_dir: str = ""
_sheet_context: Any = None
_manifest: dict[str, Any] = {}
_manifest_lock = threading.Lock()

# Shortest allowed cut pair. Anything under this is a misclick, not a clip.
MIN_CUT_SECONDS = 0.2

# ---- Blueprint ----

composer_bp = Blueprint("composer", __name__)

utils.register_static_routes(
    composer_bp,
    "composer.html",
    media_dir_getter=lambda: _input_dir,
    media_error="Input directory not configured",
    icons=True,
)


# ---- Manifest I/O ----


def _manifest_path() -> Path:
    return Path(utils.get_effective_output_dir()) / config.COMPOSER_MANIFEST_FILENAME


def _empty_manifest() -> dict[str, Any]:
    return {
        "cuts": [],
        "ui": {"markerSources": {src: True for src in config.CONVERGENCE_SOURCES}},
    }


def _load_manifest() -> dict[str, Any]:
    path = _manifest_path()
    if path.is_file():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except (OSError, json.JSONDecodeError):
            utils.warning_print(
                f"Could not read {path.name}; starting with an empty composer manifest."
            )
    return _empty_manifest()


def _persist_locked() -> None:
    """Write the composer manifest to disk. Caller must hold ``_manifest_lock``."""
    path = _manifest_path()
    try:
        path.write_text(json.dumps(_manifest, indent=2), encoding="utf-8")
    except OSError as exc:
        utils.error_print(f"Could not write {path.name}: {exc}")


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
    for p in utils.discover_participant_videos():
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
    for p in utils.discover_participant_videos():
        if not p.get("has_video"):
            continue
        parts = _participant_parts(p["video_paths"])
        participants.append(
            {
                "id": p["id"],
                "has_video": True,
                "parts": parts or [{"name": Path(vp).name} for vp in p["video_paths"]],
                "total_duration": (
                    sum(part["duration"] for part in parts) if parts else None
                ),
            }
        )
    return ok(participants=participants, config=utils.get_frontend_config())


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
        "createdAt": datetime.now(timezone.utc).isoformat(),
    }
    with _manifest_lock:
        _manifest.setdefault("cuts", []).append(cut)
        _persist_locked()
    return ok(cut=cut)


@composer_bp.route("/api/cuts/<cut_id>", methods=["PATCH"])
def api_cut_update(cut_id: str) -> Any:
    data = request.get_json(silent=True) or {}
    with _manifest_lock:
        cut = next(
            (c for c in _manifest.get("cuts", []) if c.get("id") == cut_id), None
        )
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
        cuts = _manifest.get("cuts", [])
        remaining = [c for c in cuts if c.get("id") != cut_id]
        if len(remaining) == len(cuts):
            return err(f"No cut {cut_id}", 404)
        _manifest["cuts"] = remaining
        _persist_locked()
    return ok()


@composer_bp.route("/api/ui", methods=["PUT"])
def api_ui_update() -> Any:
    """Persist UI state (lane toggles + fold states) so it survives a restart.

    Body may carry ``markerSources`` (lane visibility) and/or ``laneFolds``
    (True = the lane renders as one compact row); at least one is required.
    """
    data = request.get_json(silent=True) or {}
    sources = data.get("markerSources")
    folds = data.get("laneFolds")
    if not isinstance(sources, dict) and not isinstance(folds, dict):
        return err("markerSources or laneFolds object is required")
    response: dict[str, Any] = {}
    with _manifest_lock:
        ui = _manifest.setdefault("ui", {})
        if isinstance(sources, dict):
            ui["markerSources"] = {
                src: bool(sources.get(src, True)) for src in config.CONVERGENCE_SOURCES
            }
            response["markerSources"] = ui["markerSources"]
        if isinstance(folds, dict):
            ui["laneFolds"] = {
                src: bool(folds.get(src, True)) for src in config.CONVERGENCE_SOURCES
            }
            response["laneFolds"] = ui["laneFolds"]
        _persist_locked()
    return ok(**response)


# ---- State init ----


def _init_composer_state(
    sheet_context: Any = None,
    participant_list: list[str] | None = None,
) -> None:
    """Initialize module-level state for Composer routes.

    ``participant_list`` is accepted for parity with the other blueprints' init
    signatures; participants are discovered from the input dir on demand.
    """
    global _input_dir, _sheet_context, _manifest  # noqa: PLW0603

    _input_dir = str(utils.get_effective_input_dir())
    _sheet_context = sheet_context
    _manifest = _load_manifest()
