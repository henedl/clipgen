"""Combined Flask server for clipgen Studio, Screenspace, Transcripts, and Workflows.

Entry point: start_combined_server(worksheet, port, default_page) registers
Studio, Screenspace, Transcripts, and Workflows blueprints on one app at config.SERVER_PORT (8089).
Module-level state: _worksheet, _sheet_context, _generated_artifacts, _generated_reels
(initialized by _init_studio_state()).

Studio API endpoints (studio_bp, mounted under /studio/):
  Media  GET  /api/thumbnail/<p>/<t>   – JPEG thumbnail frame from participant video
         GET  /api/sprite/<p>          – hover-scrubber sprite sheet for a clip
         GET  /api/clip-audio/<p>      – hover-scrubber audio snippet for a clip
  Sheet  GET  /api/sheet               – spreadsheet grid data (rows, participants, timestamps)
         POST /api/sheet/refresh       – re-fetch spreadsheet data from source (Google/Excel)
         GET  /api/sheet/baseline      – per-participant baseline timestamps for convergence
         GET/PUT /api/convergence/offsets – read/persist per-participant convergence offsets
  Build  POST /api/generate            – generate clip/screen/gif artifacts for cells
         POST /api/generate/cancel     – cancel an in-progress clip generation
         POST /api/generate-intake     – generate artifacts from an intake/screenspace manifest
         POST /api/generate-intake/cancel – cancel an in-progress intake generation
         POST /api/highlights-preview  – preview highlights reel selection without generating
         POST /api/reel                – build a reel from specified cells
         POST /api/reel-direct         – build a reel from explicit clip paths
         POST /api/reel/cancel         – cancel an in-progress reel build
         GET  /api/job-status          – poll the status of a background build job
  Export POST /api/viewer              – generate timeline viewer from session artifacts
         POST /api/open-viewer         – open an already-generated viewer HTML
         POST /api/timeline-viewer     – batch-export all clips and generate timeline viewer
         POST /api/timeline-viewer/cancel – cancel a batch timeline-viewer export
         POST /api/gallery             – generate gallery from a video file
         POST /api/gallery/cancel      – cancel a gallery build
  State  GET/POST /api/manifest        – read or write the cumulative artifact manifest
         POST /api/regenerate          – regenerate all media from saved manifest
         GET/POST /api/stashes         – reel stash CRUD
         GET/POST /api/artifact-stashes – artifact stash CRUD
  Cards  GET  /api/titlecards          – list title/end card background options
         GET  /api/titlecards/default/<kind> – default title/end card image
         GET/DELETE /api/titlecards/image/<path:name> – fetch or remove an uploaded card
         POST /api/titlecards/upload   – upload a title/end card background
  Config GET/PUT /api/settings         – read or update config settings

Combined app-level routes (registered by start_combined_server, not under /studio/):
  GET  /                        – Start overlay / active-tool landing page
  GET  /api/status              – which interfaces are active (studio/screenspace/transcripts)
  GET  /api/export/status, POST /api/export – analysis-ready JSON/CSV export
  GET/POST /api/dirs            – input/output directory picker
  GET  /api/spreadsheets/excel|google – spreadsheet discovery
  POST /api/spreadsheets/google/auth, /api/spreadsheets/open – Google auth + open
"""

import concurrent.futures
import copy
import hashlib
import json
import os
import queue
import re
import socket
import sys
import threading
import time
import traceback
import webbrowser
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from collections.abc import Callable

from flask import (
    Blueprint,
    Flask,
    Response,
    g,
    jsonify,
    redirect,
    request,
    send_file,
    send_from_directory,
)
from flask.json.provider import DefaultJSONProvider
from werkzeug.serving import ThreadedWSGIServer, WSGIRequestHandler, make_server

import config
import files
import spreadsheet
import pipeline
import profiling
import start_settings
import titlecards
import utils
import video
import viewer
from server_utils import (
    MediaCache,
    clip_media_response,
    err,
    json_endpoint,
    mtime_or_zero,
    ok,
    parse_number_arg,
    profiled_stream,
)
from datetime import UTC

FlaskResponse = Response | tuple[Response, int]

# ---- Module-level state (set once by _init_studio_state) ----

_worksheet: Any = None
_sheet_context: spreadsheet.SheetContext | None = None
_sheet_payload_cache: tuple[Any, dict[str, Any]] | None = None
_sheet_payload_cache_lock = threading.Lock()
# Drive's spreadsheet listing, cached as (monotonic stamp, metas). One picker
# flow asks for it three times (list → worksheet dropdown → open-by-name); see
# _cached_spreadsheet_meta. Cleared when a new Google client authenticates.
_google_sheet_list_cache: tuple[float, list[dict[str, str]]] | None = None
_google_sheet_list_lock = threading.Lock()
_GOOGLE_SHEET_LIST_TTL_SEC = 300.0
# Active spreadsheet metadata, read by /api/status so the Start overlay can
# pre-select the right tab (Google/Excel) and re-highlight the right item. Only
# set when opened via the runtime picker; CLI-loaded sheets leave it None.
_active_sheet_meta: dict[str, str] | None = None
# The parsed MindNode document, when the session opened from a mind map instead
# of (or alongside) a spreadsheet. Mind-map teams run Studio with no sheet at all,
# so it is an independent source: never cleared by a sheet swap, and read only by
# Studio's MindNode Intake tab.
_mindnode_doc: dict[str, Any] | None = None
_mindnode_lock = threading.Lock()
# What this session's recent-projects entry is keyed by (the last descriptor given
# to `record_project_session`). The Start overlay's "current session" highlight
# compares against exactly this, so it can't be re-derived from _active_sheet_meta
# / _mindnode_doc: with both open, only the one opened *last* was recorded.
_active_project_source: dict[str, str] | None = None
_generated_artifacts: list[dict[str, Any]] = []
# Index by (cellRow, cellCol, type) for O(1) lookup in /api/generate Phase 1.
# Mutated under _generated_output_lock together with _generated_artifacts.
_generated_artifacts_index: dict[tuple[int, int, str], list[dict[str, Any]]] = {}
_generated_reels: list[dict[str, Any]] = []
# Bounded LRU: entries are JPEG bytes (tens of KB), so a few hundred is plenty.
# _MediaCache below adds single-flight semantics, so concurrent identical misses
# don't each spawn ffmpeg (Flask's dev server serves api_* routes threaded).
_THUMBNAIL_CACHE_MAX = 256
# Card-scrubber assets (opt-in hover preview). Sprite sheets are small JPEGs;
# audio segments are PCM WAV (~1 MB per short clip), so cap audio far lower.
_SPRITE_CACHE_MAX = 256
_AUDIO_CACHE_MAX = 32
# A second concurrent /api/generate or /api/reel call would clobber the
# shared cancel event; reject with 409 instead while one is in flight.
_reel_cancel_event = threading.Event()
_generate_cancel_event = threading.Event()
# Independent cancel event for /api/generate-intake: the sheet and intake branches
# run concurrently from one Studio Generate click on separate streams, so Cancel
# posts to both /api/generate/cancel and /api/generate-intake/cancel.
_intake_cancel_event = threading.Event()
# Cancel events for the two long-running viewer builds (/api/timeline-viewer
# re-cuts every clip, /api/gallery extracts frames/GIFs). Both run synchronously in
# the request thread, but Flask is threaded, so the matching /cancel endpoint can
# set the event mid-build and the cancel_flag checks short-circuit it.
_timeline_viewer_cancel_event = threading.Event()
_gallery_cancel_event = threading.Event()
_busy_lock = threading.Lock()
# Single-job slots for those builds: each shares one module-level cancel event, so
# a second concurrent build (a second Studio tab) must be rejected rather than
# allowed to clobber the other's signal.
_busy_slots: dict[str, bool] = {
    "generate": False,
    "reel": False,
    "timeline_viewer": False,
    "gallery": False,
}
# Count of in-flight /api/generate-intake streams. Intake has no single-job
# slot (it must run alongside /api/generate for mixed queues), but a sheet
# swap still needs to know whether any intake work is active.
_intake_active = 0
# Serializes load → mutate → save for the stash manifests so concurrent
# stash CRUD requests don't drop each other's writes.
_stash_lock = threading.Lock()
# Serializes mutations to in-memory generated lists and quiet manifest saves.
_generated_output_lock = threading.Lock()
# Latest progress per in-flight job, exposed by /api/job-status so Studio can
# re-attach (progress + Cancel) after navigating away mid-build and back.
_job_state_lock = threading.Lock()
# `started_at` is a wall-clock epoch (seconds) stamped when a build begins, so a
# Studio reattach can show accurate elapsed time after a page reload.
_reel_job_state: dict[str, Any] = {
    "total_clips": 0,
    "clips_done": 0,
    "concat_progress": 0.0,
    "phase": None,
    "endpoint": None,  # "reel" or "reel-direct"
    "started_at": None,
}
_generate_job_state: dict[str, Any] = {
    "total": 0,
    "done": 0,
    "started_at": None,
}
_intake_job_state: dict[str, Any] = {
    "total": 0,
    "done": 0,
    "started_at": None,
}


# Cached gspread client + threaded-auth state for the Start overlay's Google
# Sheets picker. Populated lazily by /api/spreadsheets/google/auth.
@dataclass
class _GoogleAuthState:
    client: Any = None
    in_flight: bool = False
    error: str = ""
    lock: threading.Lock = field(default_factory=threading.Lock)


_google_auth = _GoogleAuthState()

# Set by the boot build when a window-first `-s` launch could not open the
# requested spreadsheet (no cached Google token, bad name, network failure):
# ``{"message": ..., "source_type": "google"|"excel"}``. Surfaced on
# /api/status so the Start overlay — the recovery surface for a sheetless
# boot — can say why nothing is loaded and land on the failed source's tab;
# cleared when a sheet opens.
_startup_notice: dict[str, str] | None = None

# Snapshot config defaults before any settings file is loaded.
# Deep-copied so dict-valued defaults are not aliased to live config state.
_settings_defaults: dict[str, Any] = {
    name: copy.deepcopy(getattr(config, name))
    for name in getattr(config, "STUDIO_SETTINGS", {})
}

_HEX_COLOR_RE = re.compile(r"^#[0-9a-fA-F]{6}$")
_MARK_KEY_RE = re.compile(r"^[a-z0-9_]+$")

# Hotkey overrides: structural validation only. The action catalog lives in
# assets/web/hotkeys.js, so unknown ids are stored as-is and never dispatch.
_HOTKEY_ID_RE = re.compile(r"^[a-z][a-zA-Z0-9]*(\.[a-zA-Z0-9]+)+$")
_HOTKEY_COMBO_RE = re.compile(
    r"^((Mod|Ctrl|Alt|Shift)\+)*([\x21-\x7E]|[A-Za-z][A-Za-z0-9]+)$"
)
_HOTKEY_RESERVED_KEYS = ("Escape", "Tab")


def _coerce_hotkey_overrides(value: Any) -> dict[str, str] | None:
    """Validate and normalize a hotkey-overrides payload.

    Keys are dot-namespaced action ids; values are ``""`` (shortcut disabled)
    or whitespace-separated combo tokens like ``Mod+Shift+Z``. Escape and Tab
    are reserved for modal/focus semantics and rejected as final keys.
    Returns the cleaned dict, or None if the payload is structurally invalid.
    """
    if not isinstance(value, dict):
        return None
    cleaned: dict[str, str] = {}
    for raw_key, raw_combo in value.items():
        if not isinstance(raw_key, str) or not _HOTKEY_ID_RE.match(raw_key):
            return None
        if not isinstance(raw_combo, str):
            return None
        combo = raw_combo.strip()
        for token in combo.split():
            if not _HOTKEY_COMBO_RE.match(token):
                return None
            if token.split("+")[-1] in _HOTKEY_RESERVED_KEYS:
                return None
        cleaned[raw_key] = combo
    return cleaned


def _coerce_mark_categories(value: Any) -> dict[str, dict[str, str]] | None:
    """Validate and normalize a mark_categories payload.

    Returns the cleaned dict, or None if the payload is structurally invalid.
    """
    if not isinstance(value, dict):
        return None
    cleaned: dict[str, dict[str, str]] = {}
    for raw_key, raw_entry in value.items():
        if not isinstance(raw_key, str):
            return None
        key = raw_key.strip()
        if not key or not _MARK_KEY_RE.match(key):
            return None
        if not isinstance(raw_entry, dict):
            return None
        label = str(raw_entry.get("label", "")).strip()
        color = str(raw_entry.get("color", "")).strip()
        if not label or not _HEX_COLOR_RE.match(color):
            return None
        cleaned[key] = {"label": label, "color": color}
    return cleaned


# Moved to server_utils.MediaCache (shared with the Composer scrubber routes);
# the alias keeps the historical `server._MediaCache` name for tests.
_MediaCache = MediaCache

_thumbnail_cache = _MediaCache(_THUMBNAIL_CACHE_MAX)
_sprite_cache = _MediaCache(_SPRITE_CACHE_MAX)
_audio_cache = _MediaCache(_AUDIO_CACHE_MAX)


def _try_claim_busy(slot: str) -> bool:
    """Atomically reserve the single-job slot for *slot*.

    Valid slots: the ``_busy_slots`` keys. Returns True on success (caller must
    call ``_release_busy`` when done) or False if another request is already
    holding the slot (or the slot name is unknown).
    """
    with _busy_lock:
        if slot not in _busy_slots or _busy_slots[slot]:
            return False
        _busy_slots[slot] = True
        return True


def _parse_titlecard_request(
    data: dict[str, Any],
) -> tuple[bool | None, int | None]:
    """Parse optional titlecard fields from a Studio JSON request body."""
    enabled: bool | None = None
    if data.get("titlecards_enabled") is not None:
        enabled = bool(data["titlecards_enabled"])
    duration: int | None = None
    raw_duration = data.get("titlecard_duration")
    if raw_duration is not None:
        try:
            val = int(raw_duration)
            if val > 0:
                duration = val
        except (ValueError, TypeError):
            pass
    return enabled, duration


def _index_artifact(a: dict[str, Any]) -> None:
    """Insert *a* into _generated_artifacts_index. Caller must hold the lock."""
    row = a.get("cellRow")
    col = a.get("cellCol")
    type_ = a.get("type")
    if row is None or col is None or type_ is None:
        return
    _generated_artifacts_index.setdefault((row, col, type_), []).append(a)


def _rebuild_artifact_index() -> None:
    """Clear and repopulate the index from _generated_artifacts. Caller holds lock."""
    _generated_artifacts_index.clear()
    for a in _generated_artifacts:
        _index_artifact(a)


def _extend_generated_artifacts(artifacts: list[dict[str, Any]]) -> None:
    if not artifacts:
        return
    with _generated_output_lock:
        _generated_artifacts.extend(artifacts)
        for a in artifacts:
            _index_artifact(a)


def _append_generated_artifact(artifact: dict[str, Any]) -> None:
    with _generated_output_lock:
        _generated_artifacts.append(artifact)
        _index_artifact(artifact)


def _extend_generated_reels(reels: list[dict[str, Any]]) -> None:
    if not reels:
        return
    with _generated_output_lock:
        _generated_reels.extend(reels)


def _append_generated_reel(reel: dict[str, Any]) -> None:
    with _generated_output_lock:
        _generated_reels.append(reel)


def _release_busy(slot: str) -> None:
    """Release the single-job slot for *slot* (unknown slots are a no-op)."""
    with _busy_lock:
        if slot in _busy_slots:
            _busy_slots[slot] = False


def _reset_reel_job_state(endpoint: str) -> None:
    """Clear the reel progress snapshot at the start of a new build."""
    with _job_state_lock:
        _reel_job_state["total_clips"] = 0
        _reel_job_state["clips_done"] = 0
        _reel_job_state["concat_progress"] = 0.0
        _reel_job_state["phase"] = "starting"
        _reel_job_state["endpoint"] = endpoint
        _reel_job_state["started_at"] = time.time()


def _record_reel_event(event: dict[str, Any]) -> None:
    """Mirror a reel progress event into the snapshot so /api/job-status can
    report it. Called from the worker thread alongside event_queue.put."""
    phase = event.get("phase")
    with _job_state_lock:
        if phase:
            _reel_job_state["phase"] = phase
        if phase == "start":
            total = event.get("total_clips")
            if isinstance(total, int) and total >= 0:
                _reel_job_state["total_clips"] = total
                _reel_job_state["clips_done"] = 0
                _reel_job_state["concat_progress"] = 0.0
        elif phase == "clip_done":
            idx = event.get("clip_index")
            if isinstance(idx, int):
                _reel_job_state["clips_done"] = max(
                    _reel_job_state["clips_done"], idx + 1
                )
        elif phase == "concat":
            prog = event.get("progress")
            if isinstance(prog, (int, float)):
                _reel_job_state["concat_progress"] = float(prog)
        elif phase == "done":
            _reel_job_state["concat_progress"] = 1.0


def _reset_generate_job_state(total: int) -> None:
    """Initialize the generate progress snapshot when /api/generate starts.

    *total* counts artifacts (one per timestamp segment), matching the Studio
    queue's card count — not cells, and not yielded NDJSON lines.
    """
    with _job_state_lock:
        _generate_job_state["total"] = max(0, int(total))
        _generate_job_state["done"] = 0
        _generate_job_state["started_at"] = time.time()


def _increment_generate_done(n: int = 1) -> None:
    """Advance the generate-job 'done' counter by n artifacts.

    One yielded line covers a whole cell, so callers pass that cell's segment
    count rather than 1.
    """
    with _job_state_lock:
        _generate_job_state["done"] += n


def _reset_intake_job_state(total: int) -> None:
    """Initialize the intake progress snapshot when /api/generate-intake starts."""
    with _job_state_lock:
        _intake_job_state["total"] = max(0, int(total))
        _intake_job_state["done"] = 0
        _intake_job_state["started_at"] = time.time()


def _increment_intake_done(n: int = 1) -> None:
    """Advance the intake-job 'done' counter by n (one per yielded line)."""
    with _job_state_lock:
        _intake_job_state["done"] += n


def _mark_intake_active(active: bool) -> None:
    """Increment/decrement the in-flight intake-stream counter."""
    global _intake_active
    with _busy_lock:
        _intake_active += 1 if active else -1


def _generation_busy() -> bool:
    """Return True while any clip, reel, timeline-viewer, gallery, or intake
    generation is in flight.

    Consulted before a sheet swap so generated lists/manifest are not rebound
    under an active build (e.g. a timeline-viewer build would otherwise append
    old-sheet artifacts into the new sheet's freshly-rebound list/manifest).
    """
    with _busy_lock:
        return any(_busy_slots.values()) or _intake_active > 0


@contextmanager
def _override_config(**overrides: Any) -> Iterator[None]:
    """Temporarily override config attributes, restoring originals on exit."""
    saved = {name: getattr(config, name) for name in overrides}
    for name, value in overrides.items():
        setattr(config, name, value)
    try:
        yield
    finally:
        for name, value in saved.items():
            setattr(config, name, value)


# ---- Blueprint ----

studio_bp = Blueprint("studio", __name__)

# media_dir_getter resolves per request, not as a snapshot, so /studio/media/<file>
# keeps serving artifacts after POST /api/dirs moves config.OUTPUT_DIR mid-session.
# The Overview Reports tab's clip strip plays from here.
utils.register_static_routes(
    studio_bp,
    "studio.html",
    icons=True,
    media_dir_getter=lambda: str(utils.get_effective_output_dir()),
)


# ---- Helpers ----


def _resolve_participant_sources(participant: str) -> list[Path]:
    """Return a participant's ordered source-video paths (one continuous timeline).

    Mirrors the pipeline's resolution (override ``+`` list, else plain file, else
    on-disk numbered parts) but without fuzzy-match prompts (the studio runs
    non-interactively). Returns [] for an unknown participant; otherwise returns
    at least one path (which may not exist on disk — callers check ``is_file``).
    """
    if _sheet_context is None:
        return []
    ctx = _sheet_context
    participants = spreadsheet.get_participant_list(
        ctx.header_row, ctx.id_cell, ctx.num_participants
    )
    if participant not in participants:
        return []
    override = spreadsheet.participant_filename_overrides(ctx).get(participant)
    return files.resolve_source_video_paths(
        ctx.study_name, participant, override, utils.get_effective_input_dir()
    )


# ---- API endpoints ----


@studio_bp.route("/api/thumbnail/<participant>/<start_seconds>")
@json_endpoint
def api_thumbnail(participant: str, start_seconds: str) -> FlaskResponse:
    if _sheet_context is None:
        return err("No spreadsheet loaded", 404)

    # Thumbnails are second-granular: int_only floors "12.5" and max(0, ...) clamps
    # negatives rather than rejecting, matching the other media routes.
    start_sec = max(0, parse_number_arg(start_seconds, "timestamp", int_only=True))
    sources = _resolve_participant_sources(participant)
    if not sources or not sources[0].is_file():
        return err("Source video not found", 404)

    # Multi-video participant: map the global second into the owning sub-video so
    # the hover thumbnail comes from the right file at the right local offset.
    cut_sec = start_sec
    video_path = sources[0]
    if len(sources) >= 2:
        timeline = video.build_source_timeline([str(p) for p in sources])
        if timeline is None:
            return err("Source video not found", 404)
        mapped = utils.resolve_timeline_segment(timeline, start_sec)
        if mapped is None:
            return err("Timestamp beyond recording", 404)
        video_path = Path(mapped[0])
        cut_sec = int(mapped[1])

    # Include mtime so replacing a source file on disk invalidates stale thumbnails.
    cache_key = (str(video_path), cut_sec, mtime_or_zero(video_path))
    jpeg_bytes = _thumbnail_cache.get_or_compute(
        cache_key,
        lambda: video.extract_thumbnail_bytes(
            str(video_path), cut_sec, width=config.STUDIO_THUMBNAIL_WIDTH
        ),
    )
    if jpeg_bytes is None:
        return err("Thumbnail extraction failed", 404)

    return Response(
        jpeg_bytes,
        mimetype="image/jpeg",
        headers={"Cache-Control": "public, max-age=86400"},
    )


def _resolve_clip_media_source(
    participant: str, start_sec: float
) -> tuple[Path, float] | None:
    """Resolve ``(video_path, local_start_seconds)`` for a participant timestamp.

    Mirrors api_thumbnail's source resolution: maps a global second into the
    owning sub-video for multi-video participants. Returns ``None`` when no
    source video exists or the timestamp is beyond the recording.
    """
    sources = _resolve_participant_sources(participant)
    if not sources or not sources[0].is_file():
        return None
    if len(sources) >= 2:
        timeline = video.build_source_timeline([str(p) for p in sources])
        if timeline is None:
            return None
        mapped = utils.resolve_timeline_segment(timeline, int(start_sec))
        if mapped is None:
            return None
        return Path(mapped[0]), float(mapped[1])
    return sources[0], max(0.0, start_sec)


def _resolve_clip_media_paths(
    participant: str, start_sec: float, duration: float
) -> tuple[str, float, float] | None:
    """clip_media_response-shaped adapter over :func:`_resolve_clip_media_source`."""
    resolved = _resolve_clip_media_source(participant, start_sec)
    if resolved is None:
        return None
    video_path, local_start = resolved
    return str(video_path), local_start, duration


@studio_bp.route("/api/sprite/<participant>")
def api_sprite(participant: str) -> FlaskResponse:
    """Tiled JPEG sprite sheet of a clip for the opt-in hover card scrubber.

    A clip straddling a multi-video boundary maps its *start* into the owning
    sub-video and the sprite samples that file's tail (hover preview only).
    """
    if _sheet_context is None:
        return err("No spreadsheet loaded", 404)
    cols = config.STUDIO_SCRUBBER_SPRITE_COLS
    rows = config.STUDIO_SCRUBBER_SPRITE_ROWS
    return clip_media_response(
        cache=_sprite_cache,
        resolve=lambda start, dur: _resolve_clip_media_paths(participant, start, dur),
        produce=lambda path, local_start, dur: video.extract_sprite_sheet_bytes(
            path, local_start, dur, cols, rows
        ),
        mimetype="image/jpeg",
        kind_label="Sprite",
        key_extras=(cols, rows),
    )


@studio_bp.route("/api/clip-audio/<participant>")
def api_clip_audio(participant: str) -> FlaskResponse:
    """Mono WAV of a clip's audio for the opt-in hover card scrubber.

    PCM WAV (not the source's compressed audio) so the browser's WebAudio
    ``decodeAudioData`` decodes it reliably. Boundary-straddling clips map their
    *start* into the owning sub-video (hover preview only).
    """
    if _sheet_context is None:
        return err("No spreadsheet loaded", 404)
    return clip_media_response(
        cache=_audio_cache,
        resolve=lambda start, dur: _resolve_clip_media_paths(participant, start, dur),
        produce=video.extract_audio_segment_bytes,
        mimetype="audio/wav",
        kind_label="Audio",
    )


def _build_sheet_payload(ctx: spreadsheet.SheetContext) -> dict[str, Any]:
    participants = spreadsheet.get_participant_list(
        ctx.header_row, ctx.id_cell, ctx.num_participants
    )

    rows: list[dict[str, Any]] = []
    for row_idx in range(ctx.first_data_row_idx, len(ctx.sheet_data)):
        if ctx.baseline_row_idx is not None and row_idx == ctx.baseline_row_idx:
            continue
        if ctx.filename_row_idx is not None and row_idx == ctx.filename_row_idx:
            continue

        row_data = ctx.sheet_data[row_idx]
        obs_col = ctx.observation_cell.col - 1
        cat_col = ctx.category_cell.col - 1

        observation = row_data[obs_col] if obs_col < len(row_data) else ""
        category = row_data[cat_col] if cat_col < len(row_data) else ""

        severity = ""
        if ctx.severity_cell:
            sev_col = ctx.severity_cell.col - 1
            if sev_col < len(row_data) and row_data[sev_col].strip():
                severity = utils.normalize_severity(row_data[sev_col])

        cells: dict[str, dict[str, Any]] = {}
        row_keywords: set[str] = set()
        for p_idx, pid in enumerate(participants):
            col_idx = ctx.id_cell.col + p_idx
            value = row_data[col_idx] if col_idx < len(row_data) else ""
            has_text = bool(value.strip())
            valid = False
            cell_keywords: list[str] = []
            if has_text:
                cleaned, _, cell_annotations = utils.parse_cell_annotations(value)
                parsed = utils.parse_timestamps(cleaned)
                valid = bool(parsed)
                cell_keywords = sorted(cell_annotations)
                row_keywords.update(cell_annotations)
            cells[pid] = {
                "value": value.strip(),
                "valid": valid,
                "hasText": has_text,
                "keywords": cell_keywords,
            }

        rows.append(
            {
                "rowNum": row_idx + 1,
                "observation": observation,
                "category": category,
                "severity": severity,
                "keywords": sorted(row_keywords),
                "cells": cells,
            }
        )

    return {"participants": participants, "rows": rows}


def _get_sheet_payload(ctx: spreadsheet.SheetContext) -> dict[str, Any]:
    global _sheet_payload_cache

    with _sheet_payload_cache_lock:
        if _sheet_payload_cache is not None:
            cached_ctx, cached_payload = _sheet_payload_cache
            if cached_ctx is ctx:
                return cached_payload
        payload = _build_sheet_payload(ctx)
        _sheet_payload_cache = (ctx, payload)
        return payload


def _set_sheet_context(ctx: spreadsheet.SheetContext | None) -> None:
    """Replace the active sheet context and clear derived row payloads atomically."""
    global _sheet_context, _sheet_payload_cache

    with _sheet_payload_cache_lock:
        _sheet_context = ctx
        _sheet_payload_cache = None


def _sheet_observation_rows() -> list[dict[str, Any]]:
    """Reduced per-participant sheet-observation records.

    One record per (sheet row x participant) cell that marks the observation
    as applying to that participant — valid timestamps or plain cell text:
    ``{"participant", "category", "severity", "text", "timestamps": N,
    "seconds": [start_seconds, ...]}`` (text-only cells carry ``timestamps: 0``
    and empty ``seconds``). Injected into thinking_agents via
    ``thinking_agents.configure()`` (the report agent reads ``text``), so it
    works off live sheet state — no artifact generation required.
    """
    if _sheet_context is None:
        return []
    payload = _get_sheet_payload(_sheet_context)
    records: list[dict[str, Any]] = []
    for row in payload["rows"]:
        for pid, cell in row["cells"].items():
            if not (cell.get("valid") or cell.get("hasText")):
                continue
            pairs: list[tuple[str, str]] = []
            if cell.get("valid"):
                cleaned, _, _ = utils.parse_cell_annotations(cell["value"])
                pairs = utils.parse_timestamps(cleaned)
            seconds = [
                s
                for s in (utils.timestamp_to_seconds(start) for start, _ in pairs)
                if s is not None
            ]
            records.append(
                {
                    "participant": pid,
                    "category": row["category"],
                    "severity": row["severity"],
                    "text": row["observation"],
                    "timestamps": len(pairs),
                    "seconds": seconds,
                }
            )
    return records


def _sheet_common_fields() -> dict[str, Any]:
    """Payload fields shared by both branches of :func:`api_sheet`."""
    return {
        "version": utils.get_version(),
        "highlightsDuration": config.HIGHLIGHTS_REEL_DURATION_SECONDS,
        "titlecardsEnabled": config.TITLECARDS_ENABLED,
        "titlecardDuration": config.TITLECARD_DURATION_SECONDS,
        "cellExpandHover": config.STUDIO_CELL_EXPAND_HOVER,
        "cardScrubberEnabled": config.STUDIO_CARD_SCRUBBER,
        "metadataClusterScreenspace": config.STUDIO_METADATA_CLUSTER_SCREENSPACE,
        "config": utils.get_frontend_config(),
    }


@studio_bp.route("/api/sheet")
def api_sheet() -> FlaskResponse:
    if _sheet_context is None:
        # A mind-map-only session still has a study and participants, but they ride
        # their own keys: every consumer reads `participants` as *sheet columns*
        # paired with `rows`, so filling it would invent a cohort with no rows.
        mn = _mindnode_doc or {}
        return jsonify(
            {
                "ok": True,
                "sheet_loaded": False,
                "study": str(mn.get("study", "")),
                "mindnodeParticipants": list(mn.get("participants", [])),
                **_sheet_common_fields(),
                "participants": [],
                "rows": [],
            }
        )

    ctx = _sheet_context
    sheet_payload = _get_sheet_payload(ctx)

    return jsonify(
        {
            "ok": True,
            "sheet_loaded": True,
            "study": ctx.study_name,
            **_sheet_common_fields(),
            "defaultDuration": config.DEFAULT_DURATION_SECONDS,
            "participants": sheet_payload["participants"],
            "rows": sheet_payload["rows"],
        }
    )


@studio_bp.route("/api/mindnode")
def api_mindnode() -> FlaskResponse:
    """Return the active MindNode document for the intake tab.

    Response: ``{ok, mindnode_loaded, document}`` where *document* is the
    :func:`mindnode.parse_document` payload (study, participants, notes) or
    ``None`` when no mind map is open. Re-parsed on request rather than cached
    so editing the map in MindNode and hitting Refresh shows the new notes.
    """
    global _mindnode_doc
    with _mindnode_lock:
        doc = _mindnode_doc
    if doc is None:
        return jsonify({"ok": True, "mindnode_loaded": False, "document": None})

    import mindnode

    try:
        fresh = mindnode.parse_document(doc["path"])
    except ValueError as exc:
        # The bundle moved or was corrupted since it was opened. Report it
        # rather than serving a stale tree the researcher can no longer see.
        return err(str(exc), 404)
    with _mindnode_lock:
        # Re-check under the lock: the parse above runs with it released (slow, and
        # re-run per request so an edited map shows new notes), so a close landing
        # in that window would otherwise be undone here — leaving the map open
        # server-side while the UI believed it shut.
        if _mindnode_doc is not doc:
            return jsonify(
                {
                    "ok": True,
                    "mindnode_loaded": _mindnode_doc is not None,
                    "document": _mindnode_doc,
                }
            )
        _mindnode_doc = fresh
    return jsonify({"ok": True, "mindnode_loaded": True, "document": fresh})


@studio_bp.route("/api/sheet/baseline")
def api_sheet_baseline() -> FlaskResponse:
    """Return per-participant baseline offsets in seconds for convergence.

    Response: {"ok": true, "baselines": {"P01": 33120, "P02": 0, ...}}
    Values are integers (seconds) parsed via _clock_to_seconds which
    correctly treats "22:00" as HH:MM (22 hours) rather than MM:SS.
    Empty baselines dict when no baseline row exists.
    """
    if _sheet_context is None:
        return ok(sheet_loaded=False, baselines={})

    ctx = _sheet_context
    if ctx.baseline_row_idx is None:
        return ok(baselines={})

    participants = spreadsheet.get_participant_list(
        ctx.header_row, ctx.id_cell, ctx.num_participants
    )
    baselines: dict[str, int] = {}
    for p_idx, pid in enumerate(participants):
        col_idx = ctx.id_cell.col + p_idx
        value = ""
        if 0 <= ctx.baseline_row_idx < len(ctx.sheet_data) and col_idx < len(
            ctx.sheet_data[ctx.baseline_row_idx]
        ):
            value = ctx.sheet_data[ctx.baseline_row_idx][col_idx].strip()
        baselines[pid] = utils._clock_to_seconds(value) or 0

    return ok(baselines=baselines)


@studio_bp.route("/api/sheet/refresh", methods=["POST"])
def api_sheet_refresh() -> FlaskResponse:
    if _worksheet is None:
        return err("No spreadsheet loaded — pick one from the Start panel.")
    new_context = spreadsheet.build_sheet_context(_worksheet)
    if new_context is None:
        return err("Failed to refresh sheet data", 500)
    _set_sheet_context(new_context)
    return ok()


def _save_manifest_quiet() -> None:
    """Save manifest after generate/reel.

    Narrow exception handling: filesystem and serialization errors are logged
    (the manifest writer already handles atomicity), but other exceptions
    bubble up so they aren't lost silently — those are real bugs we want to
    see, not transient I/O issues.
    """
    # Snapshot the shared lists before serializing so a concurrent intake or
    # generate-worker extend can't surface as a partially-saved manifest.
    with _generated_output_lock:
        artifacts = list(_generated_artifacts)
        reels = list(_generated_reels)
    if not artifacts and not reels:
        return
    try:
        study = ""
        if artifacts:
            study = artifacts[0].get("study", "")
        elif reels:
            study = reels[0].get("study", "")
        viewer.save_manifest(
            artifacts,
            new_reels=reels or None,
            study=study,
            worksheet_title=getattr(_worksheet, "title", ""),
            is_excel=pipeline.is_excel_worksheet(_worksheet) if _worksheet else False,
            mode="studio",
        )
    except (OSError, TypeError, ValueError) as e:
        utils.warning_print(f"Failed to save manifest: {e}")


def _resolve_intake_video_paths(participant: str, source: str = "") -> list[str]:
    """Resolve the ordered source-video path(s) for an intake participant.

    Reads the live input directory (and the open sheet, when there is one)
    rather than the Screenspace/Transcripts ``_participants`` caches. Those
    only refresh on ``/api/participants``; ``POST /api/dirs`` and a MindNode-only
    session never hit that route, so a generate would look at the boot-time
    scan and report "No video for P01". ``source`` is kept for call-site
    compatibility — both tool lists scan the same files.
    """
    p = files.find_participant_record(_sheet_context, participant)
    return list(p["video_paths"]) if p is not None and p.get("has_video") else []


def _effective_study() -> str:
    """The study name to stamp on generated artifacts.

    Falls back to the open mind map when there is no spreadsheet — a mind-map
    session has no ``SheetContext``, and an empty study would land every clip
    under a nameless study and break ``{study}_{participant}`` lookups.
    """
    if _sheet_context is not None:
        return _sheet_context.study_name
    return str((_mindnode_doc or {}).get("study") or "")


def _process_intake_item(
    item: dict[str, Any],
    output_format: str,
    study: str,
    index: int = 0,
    *,
    cancel_flag: Any = None,
) -> dict[str, Any]:
    """Process a single intake item; returns one dict with _ok/_error keys.

    *index* is the item's position in the request batch — folded into the
    artifact id so two intake events covering the same participant span do not
    hash to the same id and get collapsed by manifest dedup.

    *cancel_flag* is an optional callable that returns True when the request
    has been cancelled. Checked before ffmpeg is launched and forwarded into
    ``video.run_ffmpeg`` so an in-flight encode can be terminated.
    """
    participant = item.get("participant", "")
    start = float(item.get("start", 0))
    end = float(item.get("end", 0))
    event_type = item.get("event_type", "")
    event_ids = item.get("event_ids", [])
    source = item.get("source", "screenspace")
    mark_ids = item.get("mark_ids", [])
    # A MindNode document can hold several detached trees, each with its own
    # root title, so an item may name a study of its own.
    study = str(item.get("study") or "") or study

    video_paths = _resolve_intake_video_paths(participant, source)

    if not video_paths:
        return {"_ok": False, "_error": f"No video for {participant}"}
    timeline = video.timeline_or_none(video_paths)

    out_path: str | None = None

    span_hash = hashlib.md5(f"{participant}_{start}_{end}".encode()).hexdigest()[:8]
    # Two intake events can cover the same participant span (e.g. distinct
    # Screenspace events at the same timestamp). Fold source metadata and the
    # batch index into the artifact id so manifest dedup does not silently
    # collapse them onto one record.
    id_basis = "|".join(
        [
            participant,
            str(start),
            str(end),
            source,
            ",".join(str(e) for e in event_ids),
            ",".join(str(m) for m in mark_ids),
            str(index),
        ]
    )
    id_hash = hashlib.md5(id_basis.encode()).hexdigest()[:8]
    safe_event_type = utils.sanitize_filename(event_type) if event_type else ""
    desc_part = f"{safe_event_type} " if safe_event_type else ""
    out_name = (
        f"{study} {participant} {desc_part}intake {span_hash}{config.FILEFORMAT}"
        if study
        else f"intake_{span_hash}{config.FILEFORMAT}"
    )
    out_path = files.get_unique_filename(out_name)

    if cancel_flag and cancel_flag():
        files.release_reservation(out_path)
        return {"_ok": False, "_error": "cancelled", "_cancelled": True}

    # Map the global span into the participant's source video(s), stitching across
    # a recording boundary when multi-video; single-video is a plain cut. Release
    # the reserved placeholder on any failure (None return *or* exception) so no
    # 0-byte file is left behind.
    try:
        source_fields = pipeline.cut_global_range(
            timeline,
            video_paths[0],
            start,
            end,
            out_path,
            reencode=config.REENCODING,
            cancel_flag=cancel_flag,
        )
    except Exception:
        files.release_reservation(out_path)
        raise

    if source_fields is None:
        files.release_reservation(out_path)
        return {"_ok": False, "_error": "ffmpeg failed"}

    # Enforce the size cap on the finished clip (intake has no titlecard wrap, so
    # this is the only gate). Screenshots/GIFs are never compressed.
    if output_format == "clip":
        video.enforce_filesize_limit(out_path, cancel_flag=cancel_flag)

    default_desc = {
        "transcript": "Transcript intake",
        "mindnode": "MindNode intake",
    }.get(source, "Screenspace intake")
    item_text = str(item.get("text") or "").strip()
    item_label = str(item.get("label") or "").strip()
    description = event_type or default_desc
    if source == "transcript":
        # Prefer the user's label, then a truncated transcript excerpt,
        # then the category — so cards aren't all titled "Transcript intake".
        if item_label:
            description = item_label
        elif item_text:
            description = item_text if len(item_text) <= 80 else item_text[:77] + "…"
        else:
            description = event_type or default_desc
    artifact: dict[str, Any] = {
        "id": f"intake_{id_hash}_s0",
        "type": output_format,
        "file": Path(out_path).name,
        "start": start,
        "end": end,
        "thumbnail": "",
        "study": study,
        "participant": participant,
        # Sheet-backed intake sources carry no category; a mind map's question
        # branch is one, so honour it when the item supplies it.
        "category": str(item.get("category") or ""),
        "severity": "",
        "description": description,
        "cellRow": None,
        "cellCol": None,
        "cellA1": "",
        "annotations": [],
        "source": source,
        "event_ids": event_ids,
        "mark_ids": mark_ids,
        "intake_label": event_type,
        "_ok": True,
        "_error": "",
    }
    # sourceVideo + localStart/localEnd (+ parts for a stitched span) drive
    # regeneration; for single-video they equal the global start/end.
    artifact.update(source_fields)
    if source == "transcript":
        import transcripts_server

        with transcripts_server._manifest_lock:
            src_entry = transcripts_server._manifest.get("source_transcripts", {}).get(
                participant, {}
            )
            transcript_text = item_text
            if not transcript_text and mark_ids:
                # Fallback: look up segment text by id from the manifest.
                wanted = set(mark_ids)
                parts: list[str] = []
                for seg in src_entry.get("segments", []) or []:
                    if seg.get("id") in wanted:
                        t = (seg.get("text") or "").strip()
                        if t:
                            parts.append(t)
                transcript_text = " ".join(parts)
        artifact["transcript_version"] = src_entry.get("transcribed_at", "")
        artifact["transcriptText"] = transcript_text
        if item_label:
            artifact["transcriptLabel"] = item_label
    return artifact


def _generate_intake_clips(
    items: list[dict[str, Any]],
    output_format: str = "clip",
    study: str = "",
    *,
    cancel_flag: Any = None,
) -> list[dict[str, Any]]:
    """Generate clips from intake items (Screenspace or Transcript).

    Each returned dict has an extra ``"_ok"`` key (True on success, False if
    the video was missing or ffmpeg failed) plus an ``"_error"`` string so
    callers can report per-item results without duplicating the loop.

    *cancel_flag* is forwarded to each ``_process_intake_item`` so a cancelled
    timeline-viewer build stops spawning ffmpeg for the remaining items.
    """
    workers = pipeline._resolve_clip_workers()
    if workers < 2 or len(items) < 2:
        return [
            _process_intake_item(
                item, output_format, study, index=idx, cancel_flag=cancel_flag
            )
            for idx, item in enumerate(items)
        ]

    results: list[dict[str, Any]] = [{} for _ in items]
    # See pipeline._parallel_map_ordered for what this label pair buys.
    _worker = profiling.timed("pipeline.clip")(_process_intake_item)
    with (
        profiling.span("pipeline.pool_wall"),
        concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool,
    ):
        future_to_idx = {
            pool.submit(
                _worker,
                item,
                output_format,
                study,
                index=idx,
                cancel_flag=cancel_flag,
            ): idx
            for idx, item in enumerate(items)
        }
        for future in concurrent.futures.as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                results[idx] = future.result()
            except Exception as exc:
                results[idx] = {"_ok": False, "_error": str(exc)}
    return results


# Reel and artifact stashes share the manifest's "stashes" section.
def _load_stash_list(kind: str) -> list[dict[str, Any]]:
    data = utils.load_manifest_section("stashes", default={})
    items = data.get(kind) if isinstance(data, dict) else None
    return items if isinstance(items, list) else []


def _save_stash_list(kind: str, stashes: list[dict[str, Any]]) -> Path | None:
    data = utils.load_manifest_section("stashes", default={})
    if not isinstance(data, dict):
        data = {}
    data[kind] = stashes
    if not any(data.values()):
        return utils.save_manifest_section("stashes", None)
    return utils.save_manifest_section("stashes", data)


def _load_stashes() -> list[dict[str, Any]]:
    return _load_stash_list("reels")


def _save_stashes(stashes: list[dict[str, Any]]) -> Path | None:
    return _save_stash_list("reels", stashes)


def _load_artifact_stashes() -> list[dict[str, Any]]:
    return _load_stash_list("artifacts")


def _save_artifact_stashes(stashes: list[dict[str, Any]]) -> Path | None:
    return _save_stash_list("artifacts", stashes)


def _coerce_studio_setting(name: str, value: Any) -> tuple[bool, Any, str | None]:
    """Coerce/validate one studio setting against its STUDIO_SETTINGS meta.

    The single source of truth for the settings type ladder, shared by the load
    path (_load_studio_settings) and the PUT path (_apply_settings_payload) so a
    new setting type is added once, not twice. Returns (ok, coerced, error):

      - ok=True                     -> ``coerced`` is the value to setattr on config.
      - ok=False, error is not None -> hard validation failure (bad mark_categories /
                                       card_picker / prompt / color). The PUT path surfaces
                                       ``error``; the load path skips the key (keeps the
                                       default), so a stale/tampered studio_settings.json can
                                       never apply a value the PUT path would reject.
      - ok=False, error is None     -> soft coercion failure (bad int/float cast). Both skip.

    Callers must have already confirmed ``name in config.STUDIO_SETTINGS``.
    """
    meta = config.STUDIO_SETTINGS[name]
    default = _settings_defaults.get(name)
    stype = meta.get("type")

    if stype == "mark_categories":
        cleaned = _coerce_mark_categories(value)
        if cleaned is None:
            return False, None, f"Invalid {name} payload"
        return True, cleaned, None
    if stype == "hotkeys":
        cleaned_hotkeys = _coerce_hotkey_overrides(value)
        if cleaned_hotkeys is None:
            return False, None, f"Invalid {name} payload"
        return True, cleaned_hotkeys, None
    if stype == "card_picker":
        cleaned = _coerce_card_image(value, str(meta.get("kind", "title")))
        if cleaned is None:
            return False, None, f"Invalid {name} payload"
        return True, cleaned, None
    if stype == "prompt":
        text = str(value)
        err = utils.validate_prompt(text, meta.get("placeholders") or [])
        if err is not None:
            return False, None, f"Invalid {name}: {err}"
        return True, text, None
    if name in ("TITLECARD_COLOR", "ENDCARD_COLOR"):
        color = str(value)
        if not _HEX_COLOR_RE.match(color):
            return False, None, f"Invalid {name}: expected a #rrggbb hex color"
        return True, color, None
    if name == "SOURCE_FILENAME_PATTERN":
        text = str(value).strip()
        err = utils.validate_source_filename_pattern(text)
        if err is not None:
            return False, None, f"Invalid {name}: {err}"
        return True, text, None

    expected_type = type(default) if default is not None else str
    try:
        if expected_type is bool:
            coerced: Any = (
                value
                if isinstance(value, bool)
                else str(value).lower() in ("true", "1", "yes", "on")
            )
        elif expected_type is int:
            coerced = int(value)
        elif expected_type is float:
            coerced = float(value)
        else:
            coerced = str(value)
    except (ValueError, TypeError):
        return False, None, None  # soft skip in both paths
    return True, coerced, None


def _load_studio_settings() -> dict[str, Any]:
    """Load studio_settings.json and apply non-default values to config module.

    The file lives in the per-user config dir beside ``start.json``, not in the
    output dir: these are application preferences, so switching projects must
    not switch them.
    """
    data = start_settings.load_config_json(config.STUDIO_SETTINGS_FILENAME, default={})

    applied: dict[str, Any] = {}
    for name, value in data.items():
        if name not in config.STUDIO_SETTINGS:
            continue
        ok, coerced, _ = _coerce_studio_setting(name, value)
        if not ok:
            continue
        setattr(config, name, coerced)
        applied[name] = coerced
    return applied


def _revert_unsupported_formats() -> None:
    """Re-run the ffmpeg capability guards against the loaded studio settings.

    ``cli.main`` validates webp/vp9/drawtext support before this module applies
    ``studio_settings.json``, so a persisted format can name an encoder this
    ffmpeg build lacks — or re-enable titlecards the CLI guard just disabled
    for missing drawtext. By now the server is booting, so warn and revert to
    the defaults instead of exiting. The ``video.check_*`` probes are cached,
    so the repeat costs nothing when the CLI already ran them.
    """
    webp_names = [
        name
        for name in ("SCREENSHOT_FORMAT", "GIF_FORMAT")
        if str(getattr(config, name)).lower() == ".webp"
    ]
    if webp_names and not video.check_webp_support():
        for name in webp_names:
            setattr(config, name, _settings_defaults[name])
        utils.warning_print(
            f"{', '.join(webp_names)} set to .webp but ffmpeg lacks libwebp; "
            f"reverting to the default format."
        )
    if config.GIF_FORMAT.lower() == ".webm" and not video.check_vp9_support():
        config.GIF_FORMAT = _settings_defaults["GIF_FORMAT"]
        utils.warning_print(
            "GIF_FORMAT set to .webm but ffmpeg lacks libvpx-vp9; "
            "reverting to the default format."
        )
    if config.TITLECARDS_ENABLED and not video.check_drawtext_support():
        config.TITLECARDS_ENABLED = False
        utils.warning_print(
            "Titlecards are enabled but ffmpeg lacks the drawtext filter; "
            "disabling titlecards for this run."
        )


def _studio_settings_path() -> Path:
    """Where the settings modal's values are persisted."""
    return start_settings.config_json_path(config.STUDIO_SETTINGS_FILENAME)


def _save_studio_settings(overrides: dict[str, Any]) -> Path | None:
    """Write only non-default settings to the config dir's studio_settings.json."""
    to_save = {}
    for name, value in overrides.items():
        if name in _settings_defaults and value != _settings_defaults[name]:
            to_save[name] = value
    if not to_save:
        start_settings.remove_config_json(config.STUDIO_SETTINGS_FILENAME)
        return None
    return start_settings.save_config_json(config.STUDIO_SETTINGS_FILENAME, to_save)


def _find_existing_artifacts(
    cell_row: int,
    cell_col: int,
    artifact_type: str,
    existence_cache: dict[str, bool] | None = None,
) -> list[dict[str, Any]]:
    """Return cached artifact records for a cell+type whose files still exist on disk.

    *existence_cache* memoizes path → is_file() within a single caller scope
    (e.g. one /api/generate Phase 1 pass), so the same artifact path doesn't
    hit the disk twice. Pass None for one-shot callers.
    """
    with _generated_output_lock:
        matches = list(
            _generated_artifacts_index.get((cell_row, cell_col, artifact_type), [])
        )
    if not matches:
        return []
    results: list[dict[str, Any]] = []
    for a in matches:
        resolved = str(utils.resolve_output_path(a["file"]))
        if existence_cache is not None:
            cached = existence_cache.get(resolved)
            if cached is None:
                cached = Path(resolved).is_file()
                existence_cache[resolved] = cached
            exists = cached
        else:
            exists = Path(resolved).is_file()
        if exists:
            results.append(a)
    return results


def _stream_reel_job(
    work: Callable[[Callable[[dict[str, Any]], None]], None],
    *,
    slot: str = "reel",
    on_cleanup: Callable[[], None] | None = None,
) -> Iterator[str]:
    """Run a reel build on a worker thread and stream its events as NDJSON.

    Shared scaffold for the reel-building endpoints (``/api/reel`` and
    ``/api/reel-direct``). *work* receives an ``emit_event`` callback: every
    call mirrors the event into the reel job-status snapshot and enqueues it
    for the NDJSON stream, and *work* reports its terminal result by emitting
    a final ``{"ok": ...}`` line the same way.

    The worker thread owns busy-slot release (in its ``finally``), so a client
    disconnect (e.g. browser navigation) does not abort the encode or orphan
    the reel from the manifest. The generator can die at any point; the worker
    keeps running until ffmpeg completes, then runs *on_cleanup* (if given) and
    frees the slot. *on_cleanup* is for per-request teardown such as purging
    temp files or a titlecard cache.
    """
    event_queue: queue.Queue[dict[str, Any]] = queue.Queue()
    sentinel: dict[str, Any] = {"__sentinel__": True}

    def emit_event(event: dict[str, Any]) -> None:
        _record_reel_event(event)
        event_queue.put(event)

    def worker() -> None:
        try:
            work(emit_event)
        except Exception as exc:
            emit_event({"ok": False, "error": str(exc)})
        finally:
            if on_cleanup is not None:
                on_cleanup()
            event_queue.put(sentinel)
            _release_busy(slot)

    try:
        threading.Thread(target=worker, daemon=True).start()
    except BaseException:
        # Worker never ran, so its finally won't release the slot.
        _release_busy(slot)
        raise

    while True:
        event = event_queue.get()
        if event is sentinel:
            return
        yield json.dumps(event) + "\n"


def _stream_process_reel(
    clips: list[Any],
    cancel_flag: Any,
    *,
    titlecards_enabled: bool | None = None,
    titlecard_duration_seconds: int | None = None,
) -> Iterator[str]:
    """Run pipeline.process_reel on a worker thread and yield its progress events
    as NDJSON lines, finishing with a final result/error line."""

    def work(emit_event: Callable[[dict[str, Any]], None]) -> None:
        generated, reel_records = pipeline.process_reel(
            clips,
            cancel_flag=cancel_flag,
            progress_cb=emit_event,
            titlecards_enabled=titlecards_enabled,
            titlecard_duration_seconds=titlecard_duration_seconds,
        )
        if cancel_flag and cancel_flag():
            emit_event(
                {
                    "ok": False,
                    "cancelled": True,
                    "error": "Reel generation cancelled",
                }
            )
            return
        _extend_generated_reels(reel_records)
        _save_manifest_quiet()
        emit_event({"ok": True, "generated": generated, "reels": reel_records})

    yield from _stream_reel_job(work)


def _apply_time_overrides(clips: list[Any], overrides: dict[str, Any]) -> None:
    """Replace a clip's parsed timestamps with frontend-edited in/out points.

    ``overrides`` maps a ``"participant.row"`` cell key to the complete,
    segment-ordered list of ``[start_seconds, end_seconds]`` pairs currently
    shown for that cell in the Studio queue (the user dragged/typed new in/out
    points on the duration badge). Setting ``clip["times"]`` here makes
    ``files.prepare_clip()`` take its pre-parsed fast path and skip the cell
    re-parse, so the edited durations win over the spreadsheet values.

    Must run *before* the caller prepares the clips (``/api/generate`` and
    ``/api/reel`` both prepare in the route so they can count segments), or the
    sheet parse would land first and the overrides would be ignored.
    """
    if not overrides:
        return
    # Folded lookup: the client keys these on the ref it posted, while the clip
    # carries the sheet header the server resolved it to, and the two only match
    # case-insensitively (spreadsheet.find_participant_column). An exact-match
    # lookup would silently drop the user's edited in/out points.
    by_key = {str(k).lower(): v for k, v in overrides.items()}
    for clip in clips:
        key = (clip["participant"] + "." + str(clip["cell"].row)).lower()
        seg_times = by_key.get(key)
        if not seg_times:
            continue
        new_times: list[tuple[str, str]] = []
        for pair in seg_times:
            try:
                start_sec = float(pair[0])
                end_sec = float(pair[1])
            except (TypeError, ValueError, IndexError):
                continue
            if end_sec <= start_sec:
                continue
            # Force hours on both ends when either crosses the hour mark: a mixed
            # M:SS / H:MM:SS pair breaks downstream duration parsing.
            needs_hours = start_sec >= 3600 or end_sec >= 3600
            new_times.append(
                (
                    utils.seconds_to_timestamp(
                        round(start_sec), force_hours=needs_hours
                    ),
                    utils.seconds_to_timestamp(round(end_sec), force_hours=needs_hours),
                )
            )
        if new_times:
            clip["times"] = new_times


@studio_bp.route("/api/generate", methods=["POST"])
def api_generate() -> FlaskResponse:
    if _worksheet is None:
        return err("No spreadsheet loaded — pick one from the Start panel.")

    data = request.get_json(silent=True) or {}
    cell_strings = data.get("cells", [])
    output_format = data.get("format", "clip")
    overrides: dict[str, Any] = data.get("overrides") or {}
    # Folded for the same reason _apply_time_overrides folds its lookup: the
    # client keys on the ref it posted, the clip carries the sheet header.
    override_keys = {str(k).lower() for k in overrides}
    titlecards_enabled, titlecard_duration_seconds = _parse_titlecard_request(data)

    if not cell_strings:
        return err("No cells specified")

    if output_format not in ("clip", "screen", "gif"):
        return err(f"Invalid format: {output_format}")

    if not _try_claim_busy("generate"):
        return err("A clip generation is already in progress", 409)

    try:
        cell_input = ", ".join(cell_strings)
        cell_specs = spreadsheet.parse_cell_specifications(cell_input)
        if not cell_specs:
            _release_busy("generate")
            return err("Could not parse cell specifications")

        clips = spreadsheet.generate_list(
            _worksheet,
            "cell",
            ctx=_sheet_context,
            cell_specs=cell_specs,
            skip_prompts=True,
        )
        _apply_time_overrides(clips, overrides)
    except Exception as e:
        _release_busy("generate")
        return err(str(e), 500)

    # Parse timestamps up front so progress is counted in artifacts (one per
    # segment) rather than cells — a cell holding "1:20-1:35 4:02-4:20" produces
    # two files and two queue cards, and the readout must agree with both.
    # prepare_clip is pure string work (~3 ms for 500 cells) and its pre-parsed
    # fast path makes the pipeline's later call a no-op; /api/reel does the same.
    # Per-clip guard keeps a malformed cell failing on its own line inside
    # _generate_and_persist instead of 500-ing the whole request.
    for clip in clips:
        try:
            files.prepare_clip(clip)
        except Exception as exc:
            utils.debug_print(f"prepare_clip failed during progress count: {exc}")
    total_artifacts = sum(len(clip.get("times") or []) for clip in clips)

    def stream() -> Any:
        _generate_cancel_event.clear()
        # Fall back to the cell count if nothing parsed, so the readout still
        # shows a denominator instead of hiding itself.
        _reset_generate_job_state(total_artifacts or len(cell_strings))
        cancel_flag = _generate_cancel_event.is_set
        clip_cells: set[str] = set()
        req_cards, req_dur = pipeline._resolve_titlecard_options(
            titlecards_enabled, titlecard_duration_seconds
        )
        req_title_img, req_end_img = pipeline._resolve_titlecard_images(req_cards)

        # Pass 1: yield already-existing artifacts, collect clips that need generation
        to_generate: list[tuple[Any, str]] = []
        existence_cache: dict[str, bool] = {}
        for clip in clips:
            cell_str = clip["participant"] + "." + str(clip["cell"].row)
            # Folded: cell_str is the *sheet header* the ref resolved to, and the
            # match is case-insensitive, so the trailing "No clip found" sweep
            # below must not treat a differently-spelled ref as unresolved.
            clip_cells.add(cell_str.lower())

            existing = _find_existing_artifacts(
                clip["cell"].row,
                clip["cell"].col,
                output_format,
                existence_cache=existence_cache,
            )
            # A cached clip is reusable only when its recorded titlecard state
            # matches the request, so toggling Titlecards (or its duration) takes
            # effect on the next Generate. An overridden cell is always stale:
            # artifacts are keyed by cell row/col/format only, so an edited in/out
            # point would otherwise be reused at the old duration.
            cell_overridden = cell_str.lower() in override_keys
            fresh: list[dict[str, Any]] = []
            stale: list[dict[str, Any]] = []
            for a in existing:
                matches = not cell_overridden and (
                    output_format != "clip"
                    or (
                        bool(a.get("titlecards", False)) == req_cards
                        and (
                            not req_cards
                            or (
                                a.get("titlecardDuration") == req_dur
                                and a.get("titlecardImage", "") == req_title_img
                                and a.get("endcardImage", "") == req_end_img
                            )
                        )
                    )
                )
                (fresh if matches else stale).append(a)

            if fresh:
                # Advance by the cell's segment count, not len(fresh): a
                # 3-segment cell with 2 cached artifacts still retires 3 queue
                # cards on the client, and the two counters must stay in step.
                _increment_generate_done(len(clip.get("times") or []))
                yield (
                    json.dumps(
                        {
                            "cell": cell_str,
                            "ok": True,
                            "generated": len(fresh),
                            "artifacts": fresh,
                            "skipped": True,
                        }
                    )
                    + "\n"
                )
            else:
                # Drop stale records + files so regeneration reuses the same
                # filename/id and the new record cleanly replaces the old one.
                if stale:
                    stale_ids = {a.get("id") for a in stale}
                    with _generated_output_lock:
                        _generated_artifacts[:] = [
                            a
                            for a in _generated_artifacts
                            if a.get("id") not in stale_ids
                        ]
                        _rebuild_artifact_index()
                    for a in stale:
                        resolved = str(utils.resolve_output_path(a["file"]))
                        existence_cache[resolved] = False
                        try:
                            Path(resolved).unlink(missing_ok=True)
                        except OSError:
                            pass
                to_generate.append((clip, cell_str))

        # Pass 2: generate in parallel, yielding as each completes. The per-clip
        # worker self-persists via _extend_generated_artifacts, so results landing
        # while the client disconnects still reach the manifest.
        def _generate_and_persist(
            clip: Any,
        ) -> tuple[int, list[dict[str, Any]]]:
            generated, artifacts = pipeline.process_clips(
                [clip],
                output_format=output_format,
                cancel_flag=cancel_flag,
                titlecards_enabled=titlecards_enabled,
                titlecard_duration_seconds=titlecard_duration_seconds,
                clear_titlecard_cache=False,
            )
            # A future finishing concurrently with the cancel signal still returns
            # (generated, artifacts), but the streaming contract forbids appending
            # those post-cancel — drop the files too, so no orphan media.
            if cancel_flag():
                for a in artifacts:
                    try:
                        Path(utils.resolve_output_path(a["file"])).unlink(
                            missing_ok=True
                        )
                    except OSError:
                        pass
                return 0, []
            if generated > 0:
                _extend_generated_artifacts(artifacts)
            return generated, artifacts

        if to_generate:
            workers = pipeline._resolve_clip_workers()
            # Same pipeline.clip / pipeline.pool_wall pair as
            # pipeline._parallel_map_ordered — CLIP_PARALLEL_WORKERS drives this
            # pool too, so CLI-only coverage would leave half the knob dark.
            _worker = profiling.timed("pipeline.clip")(_generate_and_persist)
            if workers >= 2 and len(to_generate) >= 2:
                with (
                    profiling.span("pipeline.pool_wall"),
                    concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool,
                ):
                    future_to_cell: dict[concurrent.futures.Future, tuple[Any, str]] = {
                        pool.submit(_worker, clip): (clip, cell_str)
                        for clip, cell_str in to_generate
                    }
                    for future in concurrent.futures.as_completed(future_to_cell):
                        if cancel_flag():
                            for f in future_to_cell:
                                f.cancel()
                            break
                        clip, cell_str = future_to_cell[future]
                        _increment_generate_done(len(clip.get("times") or []))
                        try:
                            generated, artifacts = future.result()
                            yield (
                                json.dumps(
                                    {
                                        "cell": cell_str,
                                        "ok": generated > 0,
                                        "generated": generated,
                                        "artifacts": artifacts,
                                    }
                                )
                                + "\n"
                            )
                        except Exception as e:
                            yield (
                                json.dumps(
                                    {"cell": cell_str, "ok": False, "error": str(e)}
                                )
                                + "\n"
                            )
            else:
                for clip, cell_str in to_generate:
                    if cancel_flag():
                        break
                    _increment_generate_done(len(clip.get("times") or []))
                    try:
                        generated, artifacts = _generate_and_persist(clip)
                        yield (
                            json.dumps(
                                {
                                    "cell": cell_str,
                                    "ok": generated > 0,
                                    "generated": generated,
                                    "artifacts": artifacts,
                                }
                            )
                            + "\n"
                        )
                    except Exception as e:
                        yield (
                            json.dumps({"cell": cell_str, "ok": False, "error": str(e)})
                            + "\n"
                        )

        for cs in cell_strings:
            if str(cs).lower() not in clip_cells:
                # No clip means no segments, so this ref contributed nothing to
                # total_artifacts — advancing here would overshoot the total.
                yield (
                    json.dumps({"cell": cs, "ok": False, "error": "No clip found"})
                    + "\n"
                )
        if cancel_flag():
            yield json.dumps({"cancelled": True}) + "\n"

    def stream_with_busy_release() -> Any:
        try:
            yield from stream()
        finally:
            # Persist + purge the endcard cache even on a mid-stream disconnect,
            # so artifacts never sit on disk without manifest records and temp
            # files don't leak. Per-cell process_clips() runs with
            # clear_titlecard_cache=False, so the purge happens once, here.
            titlecards.clear_endcard_cache()
            _save_manifest_quiet()
            _release_busy("generate")

    return Response(
        profiled_stream(stream_with_busy_release()),
        mimetype="application/x-ndjson",
        headers={"X-Accel-Buffering": "no"},
    )


@studio_bp.route("/api/highlights-preview", methods=["POST"])
def api_highlights_preview() -> FlaskResponse:
    if _worksheet is None:
        return err("No spreadsheet loaded — pick one from the Start panel.")

    data = request.get_json(silent=True) or {}
    highlights_duration = data.get("highlights_duration")

    overrides: dict[str, Any] = {}
    if highlights_duration is not None:
        try:
            val = int(highlights_duration)
            if val > 0:
                overrides["HIGHLIGHTS_REEL_DURATION_SECONDS"] = val
        except (ValueError, TypeError):
            pass

    with _override_config(**overrides):
        clips = spreadsheet.generate_list(
            _worksheet,
            "reel",
            ctx=_sheet_context,
            reel_input="highlights, batch",
            skip_prompts=True,
        )

    if not clips:
        return err("No clips found for highlights selection")

    result = []
    for clip in clips:
        cell = clip.get("cell")
        result.append(
            {
                "participant": clip.get("participant", ""),
                "row": cell.row if cell else 0,
                "desc": clip.get("desc", ""),
                "timestamp": str(cell.value) if cell else "",
            }
        )

    return ok(clips=result)


@studio_bp.route("/api/reel", methods=["POST"])
def api_reel() -> FlaskResponse:
    """Build a reel from spreadsheet cell refs (spreadsheet-only queues).

    Reel panel order and per-segment card removals are ignored: unique cell
    refs are re-resolved from the sheet and sorted by row then column, and
    every timestamp in those cells is included. Intake or mixed queues go
    through ``/api/reel-direct`` instead.
    """
    if _worksheet is None:
        return err("No spreadsheet loaded — pick one from the Start panel.")

    data = request.get_json(silent=True) or {}
    cell_strings = data.get("cells", [])
    highlights_duration = data.get("highlights_duration")
    reel_overrides: dict[str, Any] = data.get("overrides") or {}
    titlecards_enabled, titlecard_duration_seconds = _parse_titlecard_request(data)

    if not cell_strings:
        return err("No cells specified")

    highlights_overrides: dict[str, Any] = {}
    if highlights_duration is not None:
        try:
            val = int(highlights_duration)
            if val > 0:
                highlights_overrides["HIGHLIGHTS_REEL_DURATION_SECONDS"] = val
        except (ValueError, TypeError):
            pass

    if not _try_claim_busy("reel"):
        return err("A reel build is already in progress", 409)

    def stream() -> Any:
        # _stream_process_reel's worker owns the busy slot once control passes to
        # it; until then every exit path releases, so the slot can't leak when the
        # route returns without starting work.
        worker_started = False
        try:
            with _override_config(**highlights_overrides):
                reel_input = ", ".join(cell_strings)

                clips = spreadsheet.generate_list(
                    _worksheet,
                    "reel",
                    ctx=_sheet_context,
                    reel_input=reel_input,
                    skip_prompts=True,
                )
                _apply_time_overrides(clips, reel_overrides)

                if not clips:
                    yield (
                        json.dumps(
                            {
                                "ok": False,
                                "error": "No clips found for the specified cells",
                            }
                        )
                        + "\n"
                    )
                    return

                # Check if an identical reel already exists. compute_reel_id only
                # hashes cellRow/cellCol/start/end, so the components built here
                # don't need an accurate sourceVideo — they are throwaway.
                components: list[dict[str, Any]] = []
                for clip in clips:
                    files.prepare_clip(clip)
                    for start_str, end_str in clip.get("times", []):
                        components.append(
                            utils.build_reel_component(clip, "", start_str, end_str)
                        )
                req_cards, req_dur = pipeline._resolve_titlecard_options(
                    titlecards_enabled, titlecard_duration_seconds
                )
                req_title_img, req_end_img = pipeline._resolve_titlecard_images(
                    req_cards
                )
                if components:
                    expected_id = pipeline.compute_reel_id(components)
                    with _generated_output_lock:
                        reels_snapshot = list(_generated_reels)
                    for reel in reels_snapshot:
                        if reel.get("id") != expected_id:
                            continue
                        reel_path = Path(utils.resolve_output_path(reel["file"]))
                        if not reel_path.is_file():
                            continue
                        matches = bool(reel.get("titlecards", False)) == req_cards and (
                            not req_cards
                            or (
                                reel.get("titlecardDuration") == req_dur
                                and reel.get("titlecardImage", "") == req_title_img
                                and reel.get("endcardImage", "") == req_end_img
                            )
                        )
                        if matches:
                            yield (
                                json.dumps(
                                    {
                                        "ok": True,
                                        "generated": 1,
                                        "reels": [reel],
                                        "skipped": True,
                                    }
                                )
                                + "\n"
                            )
                            return
                        # Stale reel (e.g. titlecards toggled): drop record + file.
                        with _generated_output_lock:
                            stale_id = reel.get("id")
                            _generated_reels[:] = [
                                r for r in _generated_reels if r.get("id") != stale_id
                            ]
                        try:
                            reel_path.unlink(missing_ok=True)
                        except OSError:
                            pass

                _reel_cancel_event.clear()
                _reset_reel_job_state("reel")
                cancel_flag = _reel_cancel_event.is_set
                worker_started = True
                yield from _stream_process_reel(
                    clips,
                    cancel_flag,
                    titlecards_enabled=titlecards_enabled,
                    titlecard_duration_seconds=titlecard_duration_seconds,
                )
        except Exception as e:
            yield json.dumps({"ok": False, "error": str(e)}) + "\n"
        finally:
            if not worker_started:
                _release_busy("reel")

    return Response(
        profiled_stream(stream()),
        mimetype="application/x-ndjson",
        headers={"X-Accel-Buffering": "no"},
    )


@studio_bp.route("/api/viewer", methods=["POST"])
def api_viewer() -> FlaskResponse:
    with _generated_output_lock:
        artifacts = list(_generated_artifacts)
    if not artifacts:
        return err("No artifacts to build viewer from. Generate artifacts first.")

    try:
        study = artifacts[0].get("study", "")
        participant = artifacts[0].get("participant", "")
        ss_events = viewer.load_screenspace_events_for_viewer()
        data = viewer.finalize_timeline_data(
            artifacts,
            study=study,
            participant=participant,
            mode="studio",
            screenspace_events=ss_events or None,
        )
        viewer_path = viewer.generate_timeline_viewer(data)
        if viewer_path:
            return ok(file=str(viewer_path))
        return err("Failed to generate viewer", 500)

    except Exception as e:
        return err(str(e), 500)


def _discard_artifact_files(artifacts: list[dict[str, Any]]) -> None:
    """Unlink the on-disk media for *artifacts* so a cancelled viewer/gallery
    build leaves no orphan clips or captures: the manifest is never published in
    that case, but ffmpeg may have already written files before the cancel."""
    for art in artifacts:
        name = art.get("file", "")
        if not name:
            continue
        try:
            Path(utils.resolve_output_path(name)).unlink(missing_ok=True)
        except OSError:
            pass


@studio_bp.route("/api/timeline-viewer", methods=["POST"])
def api_timeline_viewer() -> FlaskResponse:
    if _worksheet is None:
        return err("No spreadsheet loaded — pick one from the Start panel.")

    if not _try_claim_busy("timeline_viewer"):
        return err("A timeline viewer build is already in progress.", 409)

    try:
        _timeline_viewer_cancel_event.clear()
        req = request.get_json(silent=True) or {}
        include_intake = req.get("include_intake", False)
        intake_items = req.get("intake_items", [])

        clips_list = spreadsheet.generate_list(
            _worksheet, "batch", ctx=_sheet_context, skip_prompts=True
        )
        if not clips_list:
            return err("No clips found in sheet")

        generated, artifacts = pipeline.process_clips(
            clips_list,
            output_format="clip",
            cancel_flag=_timeline_viewer_cancel_event.is_set,
        )
        if _timeline_viewer_cancel_event.is_set():
            _discard_artifact_files(artifacts)
            return jsonify({"ok": False, "cancelled": True})
        if not artifacts:
            return err("No artifacts were generated")

        # Generate intake clips if requested. Accumulate (but don't publish)
        # alongside the sheet clips so a cancel anywhere below discards them all.
        intake_artifacts: list[dict[str, Any]] = []
        if include_intake and intake_items:
            raw = _generate_intake_clips(
                intake_items, cancel_flag=_timeline_viewer_cancel_event.is_set
            )
            for r in raw:
                if r.pop("_ok", False):
                    r.pop("_error", None)
                    intake_artifacts.append(r)
            artifacts = artifacts + intake_artifacts

        # Cancel gate before any publish. Nothing has entered _generated_artifacts
        # yet, so discard every clip already written to disk (discard-on-cancel,
        # like generate) and skip writing the viewer HTML entirely.
        if _timeline_viewer_cancel_event.is_set():
            _discard_artifact_files(artifacts)
            return jsonify({"ok": False, "cancelled": True})

        _extend_generated_artifacts(artifacts)

        study = artifacts[0].get("study", "")
        ss_events = viewer.load_screenspace_events_for_viewer()
        viewer_data = viewer.finalize_timeline_data(
            artifacts,
            study=study,
            worksheet_title=getattr(_worksheet, "title", ""),
            is_excel=pipeline.is_excel_worksheet(_worksheet),
            mode="timeline-viewer",
            screenspace_events=ss_events or None,
        )
        viewer_path = viewer.generate_timeline_viewer(
            viewer_data,
            template_name="timeline-viewer.html",
            output_basename="timeline_viewer.html",
        )
        if viewer_path:
            _save_manifest_quiet()
            return ok(
                file=str(viewer_path),
                generated=generated + len(intake_artifacts),
            )
        return err("Failed to generate timeline viewer", 500)

    except Exception as e:
        return err(str(e), 500)
    finally:
        _release_busy("timeline_viewer")


@studio_bp.route("/api/gallery", methods=["POST"])
def api_gallery() -> FlaskResponse:
    if _sheet_context is None:
        return err("No spreadsheet loaded — pick one from the Start panel.")

    data = request.get_json(silent=True) or {}
    participant = data.get("participant", "")
    output_format = data.get("format", "screen")
    interval = data.get("interval", config.GALLERY_INTERVAL_SECONDS)
    bundle = bool(data.get("bundle", config.GALLERY_BUNDLE_ENABLED))

    if not participant:
        return err("No participant specified")

    if output_format not in ("screen", "gif"):
        return err(f"Invalid format: {output_format}")

    if not _try_claim_busy("gallery"):
        return err("A gallery build is already in progress.", 409)

    try:
        try:
            interval = int(interval)
            if interval < 1:
                interval = config.GALLERY_INTERVAL_SECONDS
        except (ValueError, TypeError):
            interval = config.GALLERY_INTERVAL_SECONDS

        sources = _resolve_participant_sources(participant)
        if not sources or not sources[0].is_file():
            return err(f"Source video not found for {participant}", 404)

        _gallery_cancel_event.clear()
        # Multi-video participants form one continuous timeline: capture each part
        # and shift its timestamps by the part's cumulative start so the gallery
        # spans the whole recording with global times. Single-video is unchanged.
        timeline = video.timeline_or_none([str(p) for p in sources])
        if timeline is None:
            artifacts = video.generate_interval_captures(
                str(sources[0]),
                interval_seconds=interval,
                output_format=output_format,
                gif_duration_seconds=config.GALLERY_GIF_DURATION_SECONDS,
                cancel_flag=_gallery_cancel_event.is_set,
            )
            duration = video.get_file_duration(str(sources[0])) or 0
            source_name = sources[0].name
        else:
            artifacts = []
            for part_path, dur, cumulative in timeline:
                if _gallery_cancel_event.is_set():
                    break
                # Align each part's local grid to the global interval so spacing
                # stays even across part boundaries; a part whose duration isn't
                # a multiple of the interval would otherwise shift the next
                # part's grid off the global cadence.
                first = (interval - cumulative % interval) % interval
                local_ts = list(range(first, dur, interval))
                part_artifacts = video.generate_interval_captures(
                    part_path,
                    interval_seconds=interval,
                    output_format=output_format,
                    gif_duration_seconds=config.GALLERY_GIF_DURATION_SECONDS,
                    timestamps=local_ts,
                    cancel_flag=_gallery_cancel_event.is_set,
                )
                for a in part_artifacts:
                    a["timestamp"] = a["timestamp"] + cumulative
                    a["timestamp_formatted"] = utils.seconds_to_timestamp(
                        int(a["timestamp"])
                    )
                    artifacts.append(a)
            duration = timeline[-1][1] + timeline[-1][2]
            source_name = " + ".join(Path(p).name for p, _d, _c in timeline)

        if _gallery_cancel_event.is_set():
            _discard_artifact_files(artifacts)
            return jsonify({"ok": False, "cancelled": True})
        if not artifacts:
            return err("No captures generated", 500)

        gallery_data = viewer.finalize_gallery_data(
            artifacts,
            source_video=source_name,
            video_duration=duration,
            output_format=output_format,
            interval=interval,
            bundle=bundle,
        )
        # Final cancel gate before writing the gallery HTML, closing the window
        # where capture extraction finished but the user clicks Cancel during
        # the duration probe / finalize.
        if _gallery_cancel_event.is_set():
            _discard_artifact_files(artifacts)
            return jsonify({"ok": False, "cancelled": True})
        gallery_path = viewer.generate_gallery_viewer(gallery_data)
        if gallery_path:
            return ok(file=str(gallery_path))
        return err("Failed to generate gallery viewer", 500)

    except Exception as e:
        return err(str(e), 500)
    finally:
        _release_busy("gallery")


@studio_bp.route("/api/open-viewer", methods=["POST"])
def api_open_viewer() -> FlaskResponse:
    """Open a generated viewer HTML file in the default browser."""
    data = request.get_json(silent=True) or {}
    file_path = data.get("file", "")
    if not file_path:
        return err("No file specified")

    p = Path(file_path).resolve()
    output_dir = Path(utils.get_effective_output_dir()).resolve()

    if p.suffix != ".html" or not p.is_relative_to(output_dir):
        return err("Invalid file path", 403)

    if not p.is_file():
        return err("File not found", 404)

    webbrowser.open(p.as_uri())
    return ok()


@studio_bp.route("/api/manifest", methods=["GET", "POST"])
def api_manifest() -> FlaskResponse:
    if request.method == "GET":
        artifacts, reels = viewer.load_manifest_both()
        return ok(artifacts=artifacts, reels=reels)

    # Snapshot the shared lists so a worker thread extending mid-export
    # can't produce a partial/aliased manifest snapshot.
    with _generated_output_lock:
        artifacts = list(_generated_artifacts)
        reels = list(_generated_reels)
    if not artifacts and not reels:
        return err("No artifacts to export. Generate artifacts first.")

    try:
        study = ""
        if artifacts:
            study = artifacts[0].get("study", "")
        elif reels:
            study = reels[0].get("study", "")

        manifest_path = viewer.save_manifest(
            artifacts,
            new_reels=reels or None,
            study=study,
            worksheet_title=getattr(_worksheet, "title", ""),
            is_excel=pipeline.is_excel_worksheet(_worksheet) if _worksheet else False,
            mode="studio",
        )
        if manifest_path:
            return ok(file=str(manifest_path))
        return err("Failed to write manifest", 500)

    except Exception as e:
        return err(str(e), 500)


@studio_bp.route("/api/regenerate", methods=["POST"])
def api_regenerate() -> FlaskResponse:
    try:
        artifacts, reels = viewer.load_manifest_both()
        if not artifacts and not reels:
            return err("No manifest found on disk. Export a manifest first.")

        media_count = sum(1 for a in artifacts if a.get("type") != "transcript")
        reel_count = len(reels)
        total = media_count + reel_count

        regenerated = pipeline.regenerate_from_manifest(artifacts, reels=reels)
        return ok(regenerated=regenerated, total=total)

    except Exception as e:
        return err(str(e), 500)


def _handle_stash_crud(load_fn: Any, save_fn: Any, id_prefix: str) -> FlaskResponse:
    """Shared create/update/delete logic for stash endpoints."""
    import uuid
    from datetime import datetime

    data = request.get_json(silent=True) or {}
    action = data.get("action", "create")

    # Serialize the load → mutate → save cycle so two concurrent stash
    # POSTs can't both read the same list, both append, and have the
    # second save overwrite the first.
    with _stash_lock:
        stashes = load_fn()

        if action == "create":
            items = data.get("items", [])
            if not items:
                return err("No items to stash")
            name = data.get("name", "")
            total_duration = data.get("totalDuration") or sum(
                item.get("segDuration", 0) for item in items
            )
            stash = {
                "id": f"{id_prefix}_{uuid.uuid4().hex[:8]}",
                "name": name or f"Stash {len(stashes) + 1}",
                "items": items,
                "count": len(items),
                "totalDuration": total_duration,
                "createdAt": datetime.now(UTC).isoformat(),
            }
            stashes.append(stash)
            save_fn(stashes)
            return ok(stash=stash)

        if action == "update":
            stash_id = data.get("id")
            if not stash_id:
                return err("No stash ID")
            for s in stashes:
                if s["id"] == stash_id:
                    name = data.get("name")
                    if name is not None:
                        s["name"] = name
                    save_fn(stashes)
                    return ok(stash=s)
            return err("Stash not found", 404)

        if action == "delete":
            stash_id = data.get("id")
            if not stash_id:
                return err("No stash ID")
            for i, s in enumerate(stashes):
                if s["id"] == stash_id:
                    stashes.pop(i)
                    save_fn(stashes)
                    return ok()
            return err("Stash not found", 404)

        return err(f"Unknown action: {action}")


@studio_bp.route("/api/stashes", methods=["GET"])
def api_stashes_get() -> FlaskResponse:
    return ok(stashes=_load_stashes())


@studio_bp.route("/api/stashes", methods=["POST"])
def api_stashes_post() -> FlaskResponse:
    return _handle_stash_crud(_load_stashes, _save_stashes, "stash")


@studio_bp.route("/api/artifact-stashes", methods=["GET"])
def api_artifact_stashes_get() -> FlaskResponse:
    return ok(stashes=_load_artifact_stashes())


@studio_bp.route("/api/artifact-stashes", methods=["POST"])
def api_artifact_stashes_post() -> FlaskResponse:
    return _handle_stash_crud(_load_artifact_stashes, _save_artifact_stashes, "astash")


def _settings_records() -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for name, meta in config.STUDIO_SETTINGS.items():
        records.append(
            {
                "name": name,
                "value": getattr(config, name),
                "default": _settings_defaults.get(name),
                "description": config.SETTINGS_DESCRIPTIONS.get(name, ""),
                "tab": meta.get("tab", "General"),
                "group": meta.get("group", ""),
                "type": meta.get("type", "str"),
                "options": meta.get("options"),
                "min": meta.get("min"),
                "max": meta.get("max"),
                "step": meta.get("step"),
                "provider": meta.get("provider"),
                "kind": meta.get("kind"),
                "emptyLabel": meta.get("emptyLabel"),
                "placeholders": meta.get("placeholders"),
            }
        )
    return records


def _apply_settings_payload(data: dict[str, Any]) -> tuple[dict[str, Any], str | None]:
    """Apply a settings PUT payload (direct values or a reset directive).

    Returns (applied, error). `error` is None on success.
    """
    reset = data.get("reset")
    if reset is not None:
        if reset == "all":
            target_names = list(config.STUDIO_SETTINGS.keys())
        elif isinstance(reset, str) and reset.startswith("tab:"):
            tab_name = reset[len("tab:") :]
            target_names = [
                name
                for name, meta in config.STUDIO_SETTINGS.items()
                if meta.get("tab", "General") == tab_name
            ]
        else:
            return {}, f"Invalid reset directive: {reset!r}"

        applied: dict[str, Any] = {}
        for name in target_names:
            default = copy.deepcopy(_settings_defaults.get(name))
            setattr(config, name, default)
            applied[name] = default

        # Preserve other non-default overrides already on disk: merge the reset
        # subset (now at defaults) with the current non-default values of
        # everything else, and let _save_studio_settings drop any keys that
        # equal their default (including the ones we just reset).
        merged = {name: getattr(config, name) for name in config.STUDIO_SETTINGS}
        _save_studio_settings(merged)
        return applied, None

    settings_data = data.get("settings")
    if not isinstance(settings_data, dict):
        return {}, "Invalid settings payload"

    for format_name in ("SCREENSHOT_FORMAT", "GIF_FORMAT"):
        if format_name not in settings_data:
            continue
        new_value = str(settings_data[format_name]).lower()
        current_value = str(getattr(config, format_name, "")).lower()
        # Only validate when the user is *changing* the value; an unchanged
        # value already on disk shouldn't block edits to other fields.
        if new_value == current_value:
            continue
        if new_value == ".webp" and not video.check_webp_support():
            return {}, (
                f"WebP not available: ffmpeg has no libwebp encoder. "
                f"Install an ffmpeg build with libwebp to set {format_name} to .webp."
            )
        if new_value == ".webm" and not video.check_vp9_support():
            return {}, (
                f"WebM not available: ffmpeg has no libvpx-vp9 encoder. "
                f"Install an ffmpeg build with libvpx to set {format_name} to .webm."
            )

    if "TITLECARDS_ENABLED" in settings_data:
        raw_value = settings_data["TITLECARDS_ENABLED"]
        new_enabled = (
            raw_value
            if isinstance(raw_value, bool)
            else str(raw_value).lower() in ("true", "1", "yes", "on")
        )
        if (
            new_enabled
            and not getattr(config, "TITLECARDS_ENABLED", False)
            and not video.check_drawtext_support()
        ):
            return {}, (
                "Titlecards not available: ffmpeg lacks the drawtext filter "
                "(requires libfreetype). Install an ffmpeg build with libfreetype to enable titlecards."
            )

    applied: dict[str, Any] = {}
    for name, value in settings_data.items():
        if name not in config.STUDIO_SETTINGS:
            continue
        ok, coerced, error = _coerce_studio_setting(name, value)
        if not ok:
            if error is not None:
                return {}, error
            continue
        setattr(config, name, coerced)
        applied[name] = coerced

    # Persist the full current settings state, not just the submitted keys: a
    # partial PUT (e.g. only the inline titlecard toggle) must not drop other
    # non-default settings already on disk. Every submitted key is now on
    # config, so snapshot all of STUDIO_SETTINGS and let _save_studio_settings
    # drop the ones equal to their default. Mirrors the reset path above.
    merged = {name: getattr(config, name) for name in config.STUDIO_SETTINGS}
    _save_studio_settings(merged)
    return applied, None


@studio_bp.route("/api/settings", methods=["GET"])
def api_settings_get() -> FlaskResponse:
    return ok(
        settings=_settings_records(),
        path=str(_studio_settings_path()),
        # `desktop` tells the modal whether to offer the reveal button: a
        # browser tab can open the path itself, a native window cannot.
        desktop=utils.GUI_LAUNCH,
    )


@studio_bp.route("/api/settings", methods=["PUT"])
def api_settings_put() -> FlaskResponse:
    data = request.get_json(silent=True) or {}
    applied, error = _apply_settings_payload(data)
    if error is not None:
        return err(error)
    return ok(applied=applied)


# ── Titlecard / endcard background picker ────────────────────────────────
_ALLOWED_CARD_EXTS: set[str] = {".png", ".jpg", ".jpeg", ".webp"}
_MAX_CARD_UPLOAD_BYTES: int = 10 * 1024 * 1024  # 10 MB
# URL-reserved / unsafe ASCII characters that sanitize_filename leaves intact.
# Card images are served at /api/titlecards/image/<name>, so a stem containing
# e.g. '#' would have its tail dropped by the browser as a URL fragment. These
# are replaced with '_' at upload time; unicode is preserved (only these ASCII
# characters are touched, matching sanitize_filename's unicode policy).
_URL_UNSAFE_CARD_CHARS: str = "#%&+=;@$,!*()[]{}^~` "


def _titlecard_images_dir() -> Path:
    """Directory holding user-uploaded card backgrounds (shared by both cards)."""
    return Path(utils.get_effective_output_dir()) / config.TITLECARD_IMAGES_DIRNAME


def _list_uploaded_titlecards() -> list[str]:
    """Return uploaded card filenames (sorted, case-insensitive)."""
    images_dir = _titlecard_images_dir()
    if not images_dir.is_dir():
        return []
    names = [
        p.name
        for p in images_dir.iterdir()
        if p.is_file() and p.suffix.lower() in _ALLOWED_CARD_EXTS
    ]
    return sorted(names, key=str.lower)


def _coerce_card_image(value: Any, kind: str) -> str | None:
    """Validate a card-picker selection against the upload pool.

    Returns the cleaned selection id, or None when the value is invalid. Accepts
    the sentinel ids (empty = bundled default, CARD_IMAGE_COLOR = solid color,
    and — for endcards only — CARD_IMAGE_NONE) or the basename of an existing
    uploaded image. Rejects path separators / traversal so a setting can never
    point outside the titlecard_images pool (see titlecards.resolve_card_background).
    """
    if not isinstance(value, str):
        return None
    if value in ("", config.CARD_IMAGE_COLOR):
        return value
    if kind == "end" and value == config.CARD_IMAGE_NONE:
        return value
    # Otherwise it must be a real uploaded file inside the pool: a bare basename
    # with an allowed extension that exists on disk.
    if Path(value).name != value:
        return None
    if Path(value).suffix.lower() not in _ALLOWED_CARD_EXTS:
        return None
    if not (_titlecard_images_dir() / value).is_file():
        return None
    return value


def _card_picker_payload(kind: str) -> dict[str, Any]:
    """Build the {selected, items} payload for one card kind ('title' or 'end')."""
    selected = config.TITLECARD_IMAGE if kind == "title" else config.ENDCARD_IMAGE
    items: list[dict[str, Any]] = [
        {
            "id": "",
            "label": "Default",
            "kind": "default",
            "url": f"/api/titlecards/default/{kind}",
            "deletable": False,
        }
    ]
    if kind == "end":
        items.append(
            {
                "id": config.CARD_IMAGE_NONE,
                "label": "None",
                "kind": "none",
                "url": None,
                "deletable": False,
            }
        )
    items.append(
        {
            "id": config.CARD_IMAGE_COLOR,
            "label": "Solid color",
            "kind": "color",
            "url": None,
            "deletable": False,
        }
    )
    for name in _list_uploaded_titlecards():
        items.append(
            {
                "id": name,
                "label": name,
                "kind": "upload",
                "url": f"/api/titlecards/image/{name}",
                "deletable": True,
            }
        )
    return {"selected": selected, "items": items}


@studio_bp.route("/api/titlecards", methods=["GET"])
def api_titlecards_list() -> FlaskResponse:
    """List background choices (default, color, none, uploads) for both cards."""
    return ok(
        title=_card_picker_payload("title"),
        end=_card_picker_payload("end"),
    )


@studio_bp.route("/api/titlecards/default/<kind>", methods=["GET"])
def api_titlecard_default(kind: str) -> FlaskResponse:
    """Serve the bundled default titlecard/endcard image for previews."""
    if kind not in ("title", "end"):
        return err("Invalid kind")
    asset = "titlecard.png" if kind == "title" else "endcard.png"
    path = utils.get_bundled_assets_root() / "assets" / asset
    if not path.is_file():
        return err("No default image", 404)
    return send_file(str(path))


@studio_bp.route("/api/titlecards/image/<path:name>", methods=["GET"])
def api_titlecard_image(name: str) -> FlaskResponse:
    """Serve an uploaded card background by filename (used for previews)."""
    safe = Path(name).name
    if safe != name or Path(safe).suffix.lower() not in _ALLOWED_CARD_EXTS:
        return err("Invalid filename")
    if not (_titlecard_images_dir() / safe).is_file():
        return err("Not found", 404)
    return send_from_directory(str(_titlecard_images_dir()), safe)


@studio_bp.route("/api/titlecards/upload", methods=["POST"])
def api_titlecard_upload() -> FlaskResponse:
    """Accept a card background image upload (PNG/JPG/WebP) into the upload pool."""
    file = request.files.get("file")
    filename = file.filename if file is not None else None
    if file is None or not filename:
        return err("No file provided")
    ext = Path(filename).suffix.lower()
    if ext not in _ALLOWED_CARD_EXTS:
        return err(f"Unsupported file type '{ext}'. Use PNG, JPG, or WebP.")
    file.seek(0, os.SEEK_END)
    size = file.tell()
    file.seek(0)
    if size > _MAX_CARD_UPLOAD_BYTES:
        return err("File too large (max 10 MB).")

    # sanitize_filename strips the dot from extensions, so clean the stem only,
    # then replace URL-reserved chars so the served image URL isn't truncated.
    stem = utils.sanitize_filename(Path(filename).stem).strip()
    for ch in _URL_UNSAFE_CARD_CHARS:
        stem = stem.replace(ch, "_")
    stem = stem.strip("_") or "titlecard"
    images_dir = _titlecard_images_dir()
    images_dir.mkdir(parents=True, exist_ok=True)
    candidate = f"{stem}{ext}"
    counter = 2
    while (images_dir / candidate).exists():
        candidate = f"{stem}_{counter}{ext}"
        counter += 1
    try:
        file.save(str(images_dir / candidate))
    except OSError as error:
        return err(str(error), 500)
    return ok(
        item={
            "id": candidate,
            "label": candidate,
            "kind": "upload",
            "url": f"/api/titlecards/image/{candidate}",
            "deletable": True,
        },
    )


@studio_bp.route("/api/titlecards/image/<path:name>", methods=["DELETE"])
def api_titlecard_delete(name: str) -> FlaskResponse:
    """Delete an uploaded card background; reset any selection that used it."""
    safe = Path(name).name
    if safe != name:
        return err("Invalid filename")
    target = _titlecard_images_dir() / safe
    if not target.is_file():
        return err("Not found", 404)
    try:
        target.unlink()
    except OSError as error:
        return err(str(error), 500)
    reset: dict[str, str] = {}
    for setting in ("TITLECARD_IMAGE", "ENDCARD_IMAGE"):
        if getattr(config, setting, "") == safe:
            setattr(config, setting, "")
            reset[setting] = ""
    if reset:
        merged = {n: getattr(config, n) for n in config.STUDIO_SETTINGS}
        _save_studio_settings(merged)
    return ok(reset=reset)


# ---- Screenspace Intake ----


@studio_bp.route("/api/generate-intake", methods=["POST"])
def api_generate_intake() -> FlaskResponse:
    """Generate clips from intake spans (Screenspace or Transcript).

    Streams NDJSON: one ``{"index", "ok", "artifact"|"error"}`` line per item
    so the client can paint each queue card as soon as its ffmpeg call
    finishes, rather than waiting for the whole batch.
    """
    data = request.get_json(silent=True) or {}
    items = data.get("items", [])
    if not items:
        return err("No intake items specified")

    output_format = data.get("format", "clip")
    study = _effective_study()

    def stream() -> Iterator[str]:
        _intake_cancel_event.clear()
        cancel_flag = _intake_cancel_event.is_set

        def _emit(idx: int, result: dict[str, Any]) -> str:
            ok = result.pop("_ok", False)
            error = result.pop("_error", "")
            result.pop("_cancelled", None)
            # Drop artifacts that finished after cancel was set: don't append
            # to the manifest and unlink the produced file so cancelled
            # intake work leaves no orphan media.
            if ok and cancel_flag():
                try:
                    Path(utils.resolve_output_path(result.get("file", ""))).unlink(
                        missing_ok=True
                    )
                except OSError:
                    pass
                return (
                    json.dumps({"index": idx, "ok": False, "error": "cancelled"}) + "\n"
                )
            if ok:
                _append_generated_artifact(result)
                return json.dumps({"index": idx, "ok": True, "artifact": result}) + "\n"
            return json.dumps({"index": idx, "ok": False, "error": error}) + "\n"

        _mark_intake_active(True)
        _reset_intake_job_state(len(items))
        try:
            workers = pipeline._resolve_clip_workers()
            # Same pipeline.clip / pipeline.pool_wall pair as
            # pipeline._parallel_map_ordered — CLIP_PARALLEL_WORKERS drives this
            # pool too, so CLI-only coverage would leave half the knob dark.
            _worker = profiling.timed("pipeline.clip")(_process_intake_item)
            if workers >= 2 and len(items) >= 2:
                with (
                    profiling.span("pipeline.pool_wall"),
                    concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool,
                ):
                    future_to_idx = {
                        pool.submit(
                            _worker,
                            item,
                            output_format,
                            study,
                            index=idx,
                            cancel_flag=cancel_flag,
                        ): idx
                        for idx, item in enumerate(items)
                    }
                    for future in concurrent.futures.as_completed(future_to_idx):
                        idx = future_to_idx[future]
                        try:
                            result = future.result()
                        except Exception as exc:
                            yield (
                                json.dumps(
                                    {"index": idx, "ok": False, "error": str(exc)}
                                )
                                + "\n"
                            )
                            continue
                        _increment_intake_done()
                        yield _emit(idx, result)
                        if cancel_flag():
                            # Drain remaining submitted futures so workers can
                            # observe cancel_flag and terminate ffmpeg, but
                            # short-circuit yielding once the event is set.
                            for f in future_to_idx:
                                f.cancel()
                            break
            else:
                for idx, item in enumerate(items):
                    if cancel_flag():
                        break
                    result = _process_intake_item(
                        item,
                        output_format,
                        study,
                        index=idx,
                        cancel_flag=cancel_flag,
                    )
                    _increment_intake_done()
                    yield _emit(idx, result)
            if cancel_flag():
                yield json.dumps({"cancelled": True}) + "\n"
        finally:
            # Persist whatever completed even if the client disconnects or a
            # later item raises mid-stream.
            _save_manifest_quiet()
            _mark_intake_active(False)

    return Response(
        profiled_stream(stream()),
        mimetype="application/x-ndjson",
        headers={"X-Accel-Buffering": "no"},
    )


@studio_bp.route("/api/reel-direct", methods=["POST"])
def api_reel_direct() -> FlaskResponse:
    """Build a reel from direct timestamp segments (for intake / mixed queues).

    Concatenates in panel order from explicit start/end pairs. Titlecards and
    highlights are not applied on this path; spreadsheet-only queues use
    ``/api/reel``, which resolves cells from the sheet instead.
    """
    import tempfile

    data = request.get_json(silent=True) or {}
    segments = data.get("segments", [])
    if not segments:
        return err("No segments specified")

    if not _try_claim_busy("reel"):
        return err("A reel build is already in progress", 409)

    _reel_cancel_event.clear()
    _reset_reel_job_state("reel-direct")

    titlecards_enabled, titlecard_duration_seconds = _parse_titlecard_request(data)
    cards_enabled, card_duration = pipeline._resolve_titlecard_options(
        titlecards_enabled, titlecard_duration_seconds
    )

    def stream() -> Iterator[str]:
        # temp_clips is hoisted so cleanup() can purge it after the worker
        # finishes, even on a mid-build failure.
        temp_clips: list[str] = []

        def work(emit_event: Callable[[dict[str, Any]], None]) -> None:
            output_dir = Path(utils.get_effective_output_dir())
            clip_paths: list[str] = []
            # A reel is one deliverable: a segment that fails to cut must abort the
            # build, not silently shrink it (see pipeline.process_reel).
            failed_segments: list[str] = []
            all_cards_applied = True
            # Throttle concat progress emissions to ~5 Hz, same as before.
            concat_last_emit = [0.0]

            def on_concat_progress(fraction: float) -> None:
                import time as _time

                now = _time.monotonic()
                if (
                    concat_last_emit[0] != 0.0
                    and fraction < 0.99
                    and now - concat_last_emit[0] < 0.2
                ):
                    return
                concat_last_emit[0] = now
                emit_event({"phase": "concat", "progress": fraction})

            total = len(segments)
            emit_event({"phase": "start", "total_clips": total})

            completed = 0
            for seg in segments:
                if _reel_cancel_event.is_set():
                    break

                participant = seg.get("participant", "")
                start = float(seg.get("start", 0))
                end = float(seg.get("end", 0))
                if end <= start:
                    completed += 1
                    emit_event(
                        {
                            "phase": "clip_done",
                            "clip_index": completed - 1,
                            "total_clips": total,
                        }
                    )
                    continue

                source = seg.get("source", "screenspace")
                video_paths = _resolve_intake_video_paths(participant, source)

                if not video_paths:
                    completed += 1
                    emit_event(
                        {
                            "phase": "clip_done",
                            "clip_index": completed - 1,
                            "total_clips": total,
                        }
                    )
                    continue
                timeline = video.timeline_or_none(video_paths)

                fd, tmp_path = tempfile.mkstemp(
                    prefix=config.TEMP_ARTIFACT_PREFIX,
                    suffix=config.FILEFORMAT,
                    dir=str(output_dir),
                )
                # Track for cleanup BEFORE os.close(fd) or any other call
                # that could raise; otherwise the tmp file is on disk but
                # not in temp_clips, so cleanup() won't unlink it.
                temp_clips.append(tmp_path)
                os.close(fd)

                # Map the global span into the participant's source video(s)
                # (stitching across a recording boundary); single-video is a plain cut.
                ok = (
                    pipeline.cut_global_range(
                        timeline,
                        video_paths[0],
                        start,
                        end,
                        tmp_path,
                        reencode=config.REENCODING,
                        cancel_flag=_reel_cancel_event.is_set,
                    )
                    is not None
                )
                if ok and cards_enabled:
                    # Wrap at the cut clip's own resolution (probed inside
                    # wrap_clip_with_cards). A global span may be cut from a
                    # later source part whose resolution differs from the
                    # first; trusting the clip avoids a concat mismatch.
                    wrap_clip: utils.ClipRecord = {
                        "desc": seg.get("event_type") or seg.get("desc") or "",
                    }
                    ok, seg_cards_applied = titlecards.wrap_clip_with_cards(
                        wrap_clip,
                        tmp_path,
                        cancel_flag=_reel_cancel_event.is_set,
                        titlecards_enabled=cards_enabled,
                        titlecard_duration_seconds=card_duration,
                    )
                    if ok and not seg_cards_applied:
                        all_cards_applied = False
                if ok:
                    clip_paths.append(tmp_path)
                else:
                    failed_segments.append(
                        (seg.get("event_type") or seg.get("desc") or "").strip()
                        or f"segment {completed + 1}"
                    )
                completed += 1
                emit_event(
                    {
                        "phase": "clip_done",
                        "clip_index": completed - 1,
                        "total_clips": total,
                    }
                )

            if _reel_cancel_event.is_set():
                emit_event(
                    {
                        "ok": False,
                        "error": "Reel generation cancelled",
                        "cancelled": True,
                    }
                )
                return

            if not clip_paths:
                emit_event({"ok": False, "error": "No clips could be generated"})
                return

            if failed_segments:
                # Abort rather than ship a reel that looks complete but omits
                # marked moments. The temp cuts are dropped by cleanup(), so a
                # re-run starts clean once the sources are fixed.
                emit_event(
                    {
                        "ok": False,
                        "error": (
                            f"Reel aborted: {len(failed_segments)} of {total} "
                            "segment(s) could not be generated"
                        ),
                        "failedSegments": failed_segments,
                    }
                )
                return

            reel_study = _effective_study()
            reel_base = f"{reel_study} intake reel" if reel_study else "intake_reel"
            reel_name = files.get_unique_filename(f"{reel_base}{config.FILEFORMAT}")

            try:
                concat_ok = video.concatenate_clips(
                    clip_paths,
                    reel_name,
                    reencode_on_fail=True,
                    cancel_flag=_reel_cancel_event.is_set,
                    on_progress=on_concat_progress,
                )
            except Exception as exc:
                utils.error_print(f"Concat failed: {exc}")
                concat_ok = False

            if concat_ok:
                # Record the cards that actually landed, not the ones requested —
                # a lying manifest makes the generate-cache skip this reel forever.
                reel_carded = cards_enabled and all_cards_applied
                direct_title_img, direct_end_img = pipeline._resolve_titlecard_images(
                    reel_carded
                )
                reel_record: dict[str, Any] = {
                    "id": f"reel_intake_{hashlib.md5(reel_name.encode()).hexdigest()[:8]}",
                    "file": Path(reel_name).name,
                    "source": "intake",
                    "description": f"Intake reel ({len(clip_paths)} segments)",
                    "titlecards": reel_carded,
                    "titlecardDuration": card_duration if reel_carded else 0,
                    "titlecardImage": direct_title_img,
                    "endcardImage": direct_end_img,
                }
                _append_generated_reel(reel_record)
                _save_manifest_quiet()
                emit_event({"ok": True, "generated": 1, "reels": [reel_record]})
            else:
                files.release_reservation(reel_name)
                emit_event({"ok": False, "error": "Reel concatenation failed"})

        def cleanup() -> None:
            # Endcard temp files are cached per-process across all wrap calls;
            # purge them here so per-request cards do not leak between
            # consecutive reel builds.
            titlecards.clear_endcard_cache()
            for tmp in temp_clips:
                try:
                    Path(tmp).unlink(missing_ok=True)
                except OSError:
                    pass

        yield from _stream_reel_job(work, on_cleanup=cleanup)

    return Response(
        profiled_stream(stream()),
        mimetype="application/x-ndjson",
        headers={"X-Accel-Buffering": "no"},
    )


@studio_bp.route("/api/job-status", methods=["GET"])
def api_job_status() -> FlaskResponse:
    """Snapshot of in-flight generate/reel jobs so Studio can re-attach.

    A user who navigates from /studio/ to /screenspace/ mid-build aborts the
    original streaming fetch but the worker keeps running (see
    _stream_process_reel + /api/reel-direct's worker). On return, Studio polls
    this endpoint to restore the progress bar and Cancel button. Cancel still
    works because /api/reel/cancel and /api/generate/cancel just set their
    shared cancel events, which the workers continue to honor.
    """
    with _busy_lock:
        reel_busy = _busy_slots["reel"]
        generate_busy = _busy_slots["generate"]
        intake_busy = _intake_active > 0
    with _job_state_lock:
        reel_snapshot = dict(_reel_job_state)
        generate_snapshot = dict(_generate_job_state)
        intake_snapshot = dict(_intake_job_state)
    return ok(
        reel={
            "in_progress": reel_busy,
            "cancelling": reel_busy and _reel_cancel_event.is_set(),
            **reel_snapshot,
        },
        generate={
            "in_progress": generate_busy,
            "cancelling": generate_busy and _generate_cancel_event.is_set(),
            **generate_snapshot,
        },
        intake={
            "in_progress": intake_busy,
            "cancelling": intake_busy and _intake_cancel_event.is_set(),
            **intake_snapshot,
        },
    )


@studio_bp.route("/api/reel/cancel", methods=["POST"])
def api_reel_cancel() -> FlaskResponse:
    """Signal cancellation for the in-progress reel build."""
    _reel_cancel_event.set()
    return ok()


@studio_bp.route("/api/generate/cancel", methods=["POST"])
def api_generate_cancel() -> FlaskResponse:
    """Signal cancellation for the in-progress clip generation."""
    _generate_cancel_event.set()
    return ok()


@studio_bp.route("/api/generate-intake/cancel", methods=["POST"])
def api_generate_intake_cancel() -> FlaskResponse:
    """Signal cancellation for the in-progress intake generation.

    Sheet and intake branches run in parallel from a single Studio Generate
    click and have independent streams, so the Cancel button must hit both
    endpoints to stop the full set of in-flight ffmpeg subprocesses.
    """
    _intake_cancel_event.set()
    return ok()


@studio_bp.route("/api/timeline-viewer/cancel", methods=["POST"])
def api_timeline_viewer_cancel() -> FlaskResponse:
    """Signal cancellation for the in-progress timeline-viewer build."""
    _timeline_viewer_cancel_event.set()
    return ok()


@studio_bp.route("/api/gallery/cancel", methods=["POST"])
def api_gallery_cancel() -> FlaskResponse:
    """Signal cancellation for the in-progress gallery build."""
    _gallery_cancel_event.set()
    return ok()


# ---- State initialization ----


def _init_studio_state(worksheet: Any) -> None:
    """Initialize module-level state for Studio routes.

    *worksheet* may be ``None`` — Studio's blueprint still registers and serves
    the HTML, but spreadsheet-dependent routes report ``sheet_loaded: false``
    until a sheet is opened via ``POST /api/spreadsheets/open``.
    """
    global _worksheet, _generated_artifacts, _generated_reels

    _load_studio_settings()
    _revert_unsupported_formats()
    _worksheet = worksheet
    new_context = None
    if worksheet is not None:
        new_context = spreadsheet.build_sheet_context(worksheet)
        if new_context is None:
            utils.error_print("Could not load spreadsheet data for Studio.")
            sys.exit(1)
    _set_sheet_context(new_context)
    # Rebind the shared generated lists under their lock so a streaming
    # generate/intake append can't run against a half-swapped reference.
    with _generated_output_lock:
        _generated_artifacts, _generated_reels = viewer.load_manifest_both()
        _rebuild_artifact_index()
    _thumbnail_cache.clear()
    _sprite_cache.clear()
    _audio_cache.clear()


# ---- Entry point ----


def _derive_sheet_meta(worksheet: Any) -> dict[str, str] | None:
    """Return ``{type, id_or_path, label}`` for a CLI-loaded worksheet.

    Used so the Start overlay's spreadsheet picker can show the currently
    loaded sheet when the overlay is opened on a session that was launched
    from the CLI (not via the runtime ``/api/spreadsheets/open`` endpoint).
    """
    return files.derive_sheet_meta(worksheet)


def _spreadsheet_label() -> str:
    """Human-readable identifier for the currently loaded spreadsheet."""
    if _worksheet is None:
        return ""
    parent = getattr(_worksheet, "spreadsheet", None)
    parent_title = getattr(parent, "title", "") if parent is not None else ""
    sheet_title = getattr(_worksheet, "title", "")
    if parent_title and sheet_title:
        return f"{parent_title} ({sheet_title})"
    return parent_title or sheet_title or ""


def _swap_worksheet(new_worksheet: Any) -> None:
    """Replace the active worksheet and refresh every blueprint's sheet state.

    All five that ``build_combined_app`` initializes, not just the three that
    hold a full init: Workflows and Composer were left holding whatever sheet
    the *process* started with, which on a desktop launch is none — so a
    spreadsheet opened from the Start overlay never reached a workflow run's
    NodeContext or Composer's participant list.

    Those two re-pin through a sheet-only ``repin_sheet_state`` rather than
    their init: re-running the full init would reload their manifests, and
    Workflows' would also reseed the watch-dir baseline and restart the trigger
    daemon on every spreadsheet the user opens.

    Atomic: if any step fails, the prior state is restored so the blueprints
    don't end up pointing at different sheets. Used by
    ``/api/spreadsheets/open`` and ``/api/spreadsheets/close``.
    """
    import composer_server
    import screenspace_server
    import transcripts_server
    import workflows_server

    global _worksheet, _generated_artifacts, _generated_reels
    prev_worksheet = _worksheet
    prev_sheet_context = _sheet_context
    prev_artifacts = _generated_artifacts
    prev_reels = _generated_reels
    try:
        _init_studio_state(new_worksheet)
        screenspace_server._init_screenspace_state(
            sheet_context=_sheet_context,
        )
        transcripts_server._init_transcripts_state(
            sheet_context=_sheet_context,
        )
        workflows_server.repin_sheet_state(
            sheet_context=_sheet_context,
            worksheet=new_worksheet,
        )
        composer_server.repin_sheet_state(sheet_context=_sheet_context)
    except Exception:
        _worksheet = prev_worksheet
        _set_sheet_context(prev_sheet_context)
        with _generated_output_lock:
            _generated_artifacts = prev_artifacts
            _generated_reels = prev_reels
            _rebuild_artifact_index()
        # Best-effort: re-pin the sister blueprints to the restored state.
        # If these themselves throw, swallow — the studio state is already
        # consistent and the original exception is what we want to surface.
        try:
            screenspace_server._init_screenspace_state(
                sheet_context=_sheet_context,
            )
        except Exception:  # noqa: S110 - deliberate, see the comment above
            pass
        try:
            transcripts_server._init_transcripts_state(
                sheet_context=_sheet_context,
            )
        except Exception:  # noqa: S110 - deliberate, see the comment above
            pass
        try:
            workflows_server.repin_sheet_state(
                sheet_context=_sheet_context,
                worksheet=prev_worksheet,
            )
            composer_server.repin_sheet_state(sheet_context=_sheet_context)
        except Exception:  # noqa: S110 - deliberate, see the comment above
            pass
        raise


def _cached_spreadsheet_meta(*, force: bool = False) -> list[dict[str, str]]:
    """Return Drive's spreadsheet listing, cached for ``_GOOGLE_SHEET_LIST_TTL_SEC``.

    One "pick a sheet" flow asked Drive for the same ``files.list`` three times
    (picker list, worksheet dropdown, open-by-name), each one retried with 2/4/8s
    backoff on a 429 — so the redundancy cost the most exactly when Drive was
    rate-limiting. ``force=True`` (the picker's Refresh button, or a name missing
    from the cached list) re-lists.

    The lock is deliberately held across the fetch: concurrent misses should queue
    behind one Drive call rather than each spend a round-trip, and a waiter
    re-checks freshness after acquiring. Nothing else takes this lock, so a slow
    listing can only block another listing. Only successful fetches are stored, so
    a rate-limited call never poisons the cache.
    """
    global _google_sheet_list_cache
    import google_api

    with _google_sheet_list_lock:
        cached = _google_sheet_list_cache
        if (
            not force
            and cached is not None
            and time.monotonic() - cached[0] < _GOOGLE_SHEET_LIST_TTL_SEC
        ):
            return list(cached[1])
        metas = google_api.get_all_spreadsheet_meta(_google_auth.client)
        _google_sheet_list_cache = (time.monotonic(), metas)
        # Hand back a copy so a caller can't mutate what the next reader sees.
        return list(metas)


def _warm_spreadsheet_meta_cache() -> None:
    """Best-effort prime of the Drive listing cache (boot's warm thread).

    Failures are the overlay's problem to report on its own request — a warm
    miss must never surface as a boot error.
    """
    try:
        _cached_spreadsheet_meta()
    except Exception as exc:
        utils.verbose_print(f"Drive listing warm-up failed: {exc}")


def _invalidate_spreadsheet_meta() -> None:
    """Drop the cached Drive listing (a different account may have signed in)."""
    global _google_sheet_list_cache

    with _google_sheet_list_lock:
        _google_sheet_list_cache = None


def _spreadsheet_names_for(name: str) -> list[str]:
    """Cached Drive spreadsheet names, re-listed once when ``name`` isn't among them.

    Keeps a spreadsheet created *after* the cache filled openable instead of
    failing for the rest of the TTL. Uses the same fuzzy matcher the open path
    does, so a "did you mean" hit doesn't trigger a needless re-list.
    """
    import google_api

    names = [m["name"] for m in _cached_spreadsheet_meta()]
    if google_api.find_spreadsheet_by_name(name, names) < 0:
        names = [m["name"] for m in _cached_spreadsheet_meta(force=True)]
    return names


def _open_worksheet_for(
    type_: str, id_or_path: str, worksheet: str | None
) -> tuple[Any, str]:
    """Open a worksheet and return ``(worksheet, label)`` without swapping state.

    The read-only half of ``POST /api/spreadsheets/open``: resolves an Excel path
    or a Google spreadsheet (by URL, else by name against the Drive listing) and
    hands back the worksheet plus its display label. Callers decide what to do
    with it — ``/api/spreadsheets/open`` swaps it in via :func:`_swap_worksheet`,
    ``/api/spreadsheets/preview`` only reads from it. Returns ``(None, "")`` when
    the spreadsheet can't be resolved; raises whatever the backing library does.
    """
    if type_ == "excel":
        import excel_io

        return (
            excel_io.open_excel_workbook(id_or_path, worksheet_name=worksheet),
            Path(id_or_path).name,
        )

    if _google_auth.client is None:
        raise RuntimeError(
            "Not authenticated with Google — click 'Connect Google' first."
        )
    import app as _app

    if id_or_path.startswith(("http://", "https://")):
        new_ws = _app.open_spreadsheet_by_url(
            _google_auth.client, id_or_path, worksheet_name=worksheet
        )
    else:
        doc_list = _spreadsheet_names_for(id_or_path)
        new_ws = _app.open_spreadsheet_by_name(
            _google_auth.client, doc_list, id_or_path, worksheet_name=worksheet
        )
    label = ""
    if new_ws is not None:
        parent = getattr(new_ws, "spreadsheet", None)
        label = getattr(parent, "title", "") or id_or_path
    return new_ws, label


def _invalidate_participant_caches() -> None:
    """Force the Transcripts / Screenspace participant caches to rebuild.

    Both only rebuild when the input directory's ``st_mtime_ns`` moves (see
    ``server_utils.make_participant_cache``). A filename override changes which
    files the same directory resolves to without touching the directory at all,
    so the mtime gate has to be poked by hand or those two pages keep serving
    the previous paths. ``dir = ""`` is the same force-a-build sentinel their
    initialisers use; it can never equal a real directory string.

    Composer and Remux resolve per request and need nothing here.
    """
    import screenspace_server
    import transcripts_server

    for module in (screenspace_server, transcripts_server):
        source = module._participant_source
        if source is not None:
            source["dir"] = ""


def _seed_filename_overrides(source: dict[str, str] | None) -> None:
    """Point ``config.FILENAME_OVERRIDES`` at the source that is now open.

    The user's per-participant filename overrides are stored per spreadsheet /
    mind map in ``start.json``, but every consumer reads them off ``config`` —
    a ``SheetContext`` carries no spreadsheet identity. So the identity is
    resolved exactly here, whenever the open source changes (open, worksheet
    swap, close, override edit, CLI launch). ``None`` clears the map, which is
    what a session with no identifiable source must run with.
    """
    overrides = (
        start_settings.filename_overrides(
            source.get("type", ""),
            source.get("id_or_path", ""),
            source.get("worksheet", ""),
        )
        if source
        else {}
    )
    if overrides == config.FILENAME_OVERRIDES:
        return  # a needless rebuild costs a seekability probe per participant
    config.FILENAME_OVERRIDES = overrides
    _invalidate_participant_caches()


def _preview_source_rows(
    study: str,
    participants: list[str],
    sheet_overrides: dict[str, str | None],
    user_overrides: dict[str, str],
    base_dir: Path,
) -> tuple[list[dict[str, Any]], list[str]]:
    """Build the Start overlay's source-video preview rows for one source.

    One row per participant — the filenames clipgen will look for, whether they
    are all on disk, the sheet's own ``Filename``-row value and the user's own
    override (empty unless they set one here, which is what enables the row's
    Restore button). The user's override wins, the same precedence
    ``spreadsheet.participant_filename_overrides`` applies.

    Also returns the video files in *base_dir* that no participant claims: the
    datalist the override input offers, i.e. exactly the footage that is sitting
    there unused while some participant reads as missing.
    """
    rows: list[dict[str, Any]] = []
    claimed: set[str] = set()
    for pid in participants:
        sheet_value = sheet_overrides.get(pid) or ""
        user_value = user_overrides.get(pid) or ""
        override = user_value or sheet_value or None
        paths = files.resolve_source_video_paths(study, pid, override, base_dir)
        for path in paths:
            claimed.add(path.name.lower())
        rows.append(
            {
                "id": pid,
                "filenames": [p.name for p in paths],
                "found": all(p.is_file() for p in paths),
                "override": bool(override),
                "override_value": user_value,
                "sheet_value": sheet_value,
            }
        )
    unmatched: list[str] = []
    try:
        for path in sorted(base_dir.glob(f"*{config.FILEFORMAT}")):
            if path.is_file() and path.name.lower() not in claimed:
                unmatched.append(path.name)
    except OSError:
        pass  # unreadable folder: no suggestions is fine, an error here is not
    return rows, unmatched


def _mindnode_source() -> dict[str, str] | None:
    """The open mind map as a recent-projects source descriptor, if any."""
    if _mindnode_doc is None:
        return None
    return {
        "type": "mindnode",
        "id_or_path": str(_mindnode_doc.get("path", "")),
        "label": str(_mindnode_doc.get("name", "")),
        "worksheet": "",
    }


def _open_mindnode(id_or_path: str, project_name: str | None) -> FlaskResponse:
    """Open a ``.mindnode`` document as the session's observation source.

    A mind map is not a worksheet, so this deliberately skips
    :func:`_swap_worksheet` — Studio runs with ``sheet_loaded: false`` and the
    MindNode Intake tab supplies the notes. Any spreadsheet already open is left
    alone; the two sources coexist.
    """
    import mindnode

    global _mindnode_doc, _active_project_source
    try:
        doc = mindnode.parse_document(id_or_path)
    except ValueError as exc:
        return err(str(exc), 400)

    with _mindnode_lock:
        _mindnode_doc = doc
    label = Path(id_or_path).name
    source = {
        "type": "mindnode",
        "id_or_path": id_or_path,
        "label": label,
        "worksheet": "",
    }
    _seed_filename_overrides(source)
    start_settings.record_recent_spreadsheet("mindnode", id_or_path, label, "")
    start_settings.record_project_session(
        str(utils.get_effective_input_dir()),
        str(utils.get_effective_output_dir()),
        source,
        name=project_name,
    )
    _active_project_source = source
    return jsonify(
        {
            "ok": True,
            "sheet_loaded": _worksheet is not None,
            "mindnode_loaded": True,
            "spreadsheet_label": _spreadsheet_label(),
            "mindnode_label": label,
            "study": doc["study"],
            "notes": len(doc["notes"]),
        }
    )


def _profile_request_start() -> None:
    """before_request hook: stamp the start time when profiling is on."""
    if config.PROFILING:
        g._prof_t0 = time.perf_counter()


def _profile_request_end(response):
    """after_request hook: accumulate per-route wall time under ``route <rule>``.

    Labeling by ``url_rule.rule`` (not path) bounds label cardinality to the
    app's ~200 rules and aggregates poll endpoints instead of spamming one
    line per hit — the totals view is the useful one for polls anyway.
    """
    t0 = getattr(g, "_prof_t0", None)
    if t0 is not None and request.url_rule is not None:
        profiling.add(f"route {request.url_rule.rule}", time.perf_counter() - t0)
    return response


def _set_cache_headers(response):
    """after_request hook: apply content-type-aware caching to static assets.

    ``send_from_directory`` stamps a bare ``Cache-Control: no-cache`` by default
    (Werkzeug, when ``max_age`` is unset), which carries no real caching intent —
    normalize it per content type. Deliberate cache headers set by a route
    (thumbnails, SSE streams, video previews) are anything other than a bare
    ``no-cache`` and are preserved untouched.

    HTML/JS/CSS deliberately get ``no-cache`` (= revalidate each request,
    answered with cheap localhost 304s via the ETag/Last-Modified that
    ``send_from_directory`` already sets) rather than a TTL: a max-age on JS
    once let a browser run hour-old page scripts against fresh no-cache HTML
    after an update, which presents as "the fix didn't work" bug reports.
    Only SVG icons keep a real TTL — content-stable, requested in bulk.
    """
    existing = response.headers.get("Cache-Control")
    if existing and existing != "no-cache":
        return response
    ct = response.content_type or ""
    if ct.startswith(
        ("text/html", "text/css", "application/javascript", "text/javascript")
    ):
        response.headers["Cache-Control"] = "no-cache"
    elif ct.startswith("image/svg"):
        response.headers["Cache-Control"] = "public, max-age=86400"
    return response


def _init_combined_state(
    worksheet: Any = None,
    gspread_client: Any = None,
) -> None:
    """Wire every blueprint's process state for a combined-app launch.

    Split out of :func:`build_combined_app` so the state wiring can be re-run on
    its own, without also recompiling the app's ~200 Werkzeug URL rules — that
    compilation is ~90 % of an app build, and the tests reuse one app across a
    module while re-initialising state per test.

    Order matters: ``_init_studio_state`` populates the ``_sheet_context`` /
    ``_worksheet`` globals that every sibling initialiser below reads.
    """
    import composer_server
    import screenspace_server
    import transcripts_server
    import workflows_server

    _init_studio_state(worksheet)
    # Seed the meta + recent-projects rail for CLI launches that already
    # have a worksheet — the runtime /api/spreadsheets/open path handles
    # both of these itself.
    global _active_sheet_meta
    _active_sheet_meta = _derive_sheet_meta(worksheet)
    # Before the sibling initialisers below, which resolve participants.
    _seed_filename_overrides(_active_sheet_meta)
    if gspread_client is not None:
        _google_auth.client = gspread_client
    if _active_sheet_meta is not None:
        start_settings.record_project_session(
            str(utils.get_effective_input_dir()),
            str(utils.get_effective_output_dir()),
            _active_sheet_meta,
        )
        global _active_project_source
        _active_project_source = _active_sheet_meta

    screenspace_server._init_screenspace_state(
        sheet_context=_sheet_context,
    )
    transcripts_server._init_transcripts_state(
        sheet_context=_sheet_context,
    )
    workflows_server._init_workflows_state(
        sheet_context=_sheet_context,
        worksheet=_worksheet,
    )
    composer_server._init_composer_state(
        sheet_context=_sheet_context,
    )

    # The report thinking agent reads sheet observations and transcript marks,
    # both of which live in other modules' process state — inject them (import
    # cycles rule out direct imports from thinking_agents).
    import thinking_agents

    thinking_agents.configure(
        observation_rows_getter=_sheet_observation_rows,
        participant_marks_getter=transcripts_server.marks_for_participant,
    )

    # Last, so an armed trigger firing on the daemon's immediate first poll
    # can't start a workflow run against half-initialized sibling state (or a
    # report agent with the getters above still unset).
    workflows_server._start_watch_thread()


def build_combined_app(
    worksheet: Any = None,
    default_page: str = "studio",
    gspread_client: Any = None,
) -> Flask:
    """Build the combined Studio + Screenspace + Transcripts + Workflows Flask app.

    Same setup as :func:`start_combined_server` but stops short of
    ``app.run`` so tests (and any embedding caller) can hold the live
    ``Flask`` instance and exercise routes via ``app.test_client()``.
    """
    import composer_server
    import overview
    import screenspace_server
    import transcripts_server
    import workflows_server

    combined = Flask(__name__, static_folder=None)
    # Preserve insertion order in JSON responses. Flask defaults to sorting object
    # keys alphabetically, which would clobber manifest-ordered data such as the
    # region list (drag-to-reorder relies on GET /api/regions echoing manifest order).
    assert isinstance(combined.json, DefaultJSONProvider)  # Flask's stock provider
    combined.json.sort_keys = False

    _init_combined_state(worksheet, gspread_client)

    combined.register_blueprint(studio_bp, url_prefix="/studio")
    combined.register_blueprint(
        screenspace_server.screenspace_bp, url_prefix="/screenspace"
    )
    combined.register_blueprint(
        transcripts_server.transcripts_bp, url_prefix="/transcripts"
    )
    combined.register_blueprint(workflows_server.workflows_bp, url_prefix="/workflows")
    combined.register_blueprint(composer_server.composer_bp, url_prefix="/composer")
    combined.register_blueprint(overview.overview_bp, url_prefix="/overview")

    combined.after_request(_set_cache_headers)
    combined.before_request(_profile_request_start)
    combined.after_request(_profile_request_end)

    @combined.route("/")
    def root():
        return redirect(f"/{default_page}/")

    @combined.route("/api/profile")
    def api_profile() -> FlaskResponse:
        """Profiling snapshot for agents (``?reset=1`` brackets a window).

        404 when profiling is off so a plain launch exposes nothing — the
        endpoint mirrors the ``--profile`` opt-in rather than adding its own.
        """
        if not config.PROFILING:
            return err("profiling is off (launch with --profile)", 404)
        snap = profiling.snapshot()
        if request.args.get("reset") == "1":
            profiling.reset()
        # peak_rss rides alongside the label map rather than inside it: it is not
        # a label (no seconds/count), it is monotonic, and ?reset=1 cannot clear
        # it. The knobs it exists for (SCREENSPACE_OCR_POOL_SIZE,
        # WORKFLOWS_BATCH_WORKERS) are live-server knobs, so it has to be
        # reachable here and not only from the atexit report.
        return ok(
            profile=snap,
            peak_rss_mb=profiling.peak_rss_mb(),
            # Like peak_rss: not a label, records once per process, and
            # ?reset=1 cannot clear it.
            startup=profiling.startup_snapshot(),
        )

    @combined.route("/api/status")
    def status() -> Response:
        meta = _active_sheet_meta if _worksheet is not None else None
        return jsonify(
            {
                "studio": True,
                "screenspace": True,
                "transcripts": True,
                "workflows": True,
                "composer": True,
                "overview": True,
                "sheet_loaded": _worksheet is not None,
                "startup_notice": (_startup_notice or {}).get("message", ""),
                "startup_notice_source": (_startup_notice or {}).get("source_type", ""),
                # What record_project_session last stored, so the overlay's
                # current-session key matches its recent-projects key.
                "active_source": _active_project_source,
                "mindnode_loaded": _mindnode_doc is not None,
                "mindnode_label": (_mindnode_doc or {}).get("name", ""),
                "mindnode_path": (_mindnode_doc or {}).get("path", ""),
                "spreadsheet_label": _spreadsheet_label(),
                "spreadsheet_type": (meta or {}).get("type", ""),
                "spreadsheet_id_or_path": (meta or {}).get("id_or_path", ""),
                "spreadsheet_worksheet": (meta or {}).get("worksheet", ""),
                "input_dir": str(utils.get_effective_input_dir()),
                "output_dir": str(utils.get_effective_output_dir()),
                "videos_in_input": len(utils.discover_participant_videos()),
                "version": utils.get_version(),
                "author": "Henrik Edlund",
                "license": "MIT",
                "repo_url": config.REPO_URL,
            }
        )

    @combined.route("/api/export/status")
    def api_export_status() -> Response:
        """Report which exportable manifest sections are present.

        Used by the frontend to gate the Export quick action — if no
        sections exist there is nothing for ``write_export_bundle`` to write.
        """
        present = utils.manifest_sections()
        screenspace = "screenspace" in present
        transcripts = "transcripts" in present
        return ok(
            screenspace=screenspace,
            transcripts=transcripts,
            any=screenspace or transcripts,
        )

    @combined.route("/api/export", methods=["POST"])
    def api_export() -> FlaskResponse:
        """Write the same JSON+CSV bundle the ``--export`` CLI flag produces."""
        import data_export

        output_dir = Path(utils.get_effective_output_dir())
        try:
            written = data_export.write_export_bundle()
        except Exception as exc:
            return err(str(exc), 500)
        if not written:
            return err(
                f"Nothing to export: no Screenspace or Transcript manifest in {output_dir}. "
                "Run a Screenspace scan or transcribe a video first.",
                404,
            )
        return ok(
            written=[p.name for p in written],
            output_dir=str(output_dir),
        )

    # ---- Start overlay: directories, spreadsheet picker, persistence ----

    @combined.route("/api/dirs", methods=["GET"])
    def api_dirs_get() -> Response:
        s = start_settings.load_start_settings()
        return ok(
            input=str(utils.get_effective_input_dir()),
            output=str(utils.get_effective_output_dir()),
            recent_inputs=s.get("recent_inputs", []),
            recent_outputs=s.get("recent_outputs", []),
        )

    @combined.route("/api/dirs", methods=["POST"])
    def api_dirs_post() -> FlaskResponse:
        data = request.get_json(silent=True) or {}
        new_input = data.get("input")
        new_output = data.get("output")
        errors: dict[str, str] = {}

        if new_input is not None:
            p = Path(str(new_input)).expanduser()
            if not p.is_dir():
                errors["input"] = f"Input directory does not exist: {p}"
            else:
                config.INPUT_DIR = str(p)
                start_settings.record_recent_input(str(p))

        if new_output is not None:
            p = Path(str(new_output)).expanduser()
            try:
                p.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                errors["output"] = f"Could not create output directory: {exc}"
            else:
                config.OUTPUT_DIR = str(p)
                start_settings.record_recent_output(str(p))

        if errors:
            return jsonify({"ok": False, "errors": errors}), 400
        return api_dirs_get()

    @combined.route("/api/spreadsheets/excel", methods=["GET"])
    def api_spreadsheets_excel() -> Response:
        input_dir = Path(utils.get_effective_input_dir())
        files_list: list[dict[str, Any]] = []
        if input_dir.is_dir():
            for p in sorted(input_dir.glob("*.xlsx")):
                if p.name.startswith("~$"):
                    continue
                try:
                    modified = p.stat().st_mtime
                except OSError:
                    modified = 0.0
                files_list.append(
                    {"path": str(p), "name": p.name, "modified": modified}
                )
        return ok(input_dir=str(input_dir), files=files_list)

    @combined.route("/api/spreadsheets/mindnode", methods=["GET"])
    def api_spreadsheets_mindnode() -> Response:
        """List ``.mindnode`` bundles in the input dir for the Start overlay."""
        import mindnode

        input_dir = Path(utils.get_effective_input_dir())
        return ok(input_dir=str(input_dir), files=mindnode.find_documents(input_dir))

    @combined.route("/api/spreadsheets/mindnode/preview", methods=["GET"])
    def api_spreadsheets_mindnode_preview() -> FlaskResponse:
        """Summarize a ``.mindnode`` document before it is opened.

        Read-only counterpart to ``/api/spreadsheets/preview``: parses the
        bundle without touching module state so the Start overlay can show what
        the map holds while the user is still choosing. Carries the same
        editable ``sources`` rows and ``unmatched`` datalist as the spreadsheet
        preview — a mind map has no Filename row, so an override set here is the
        *only* way to point a participant at differently-named footage.
        """
        import mindnode

        path = (request.args.get("path") or "").strip()
        input_dir = (request.args.get("input_dir") or "").strip()
        if not path:
            return err("Required: path")
        try:
            doc = mindnode.parse_document(path)
        except ValueError as exc:
            return err(str(exc), 400)
        base_dir = (
            Path(input_dir).expanduser()
            if input_dir
            else utils.get_effective_input_dir()
        )
        rows, unmatched = _preview_source_rows(
            doc["study"],
            list(doc["participants"]),
            {},
            start_settings.filename_overrides("mindnode", path, ""),
            base_dir,
        )
        return ok(
            study=doc["study"],
            roots=doc["roots"],
            participants=doc["participants"],
            categories=sorted({n["category"] for n in doc["notes"] if n["category"]}),
            notes=len(doc["notes"]),
            with_times=doc["with_times"],
            without_times=doc["without_times"],
            has_preview=(Path(path) / mindnode.PREVIEW_RELPATH).is_file(),
            sources=rows,
            unmatched=unmatched,
        )

    @combined.route("/api/spreadsheets/mindnode/thumb", methods=["GET"])
    def api_spreadsheets_mindnode_thumb() -> FlaskResponse:
        """Serve a ``.mindnode`` bundle's own QuickLook render of the map."""
        import mindnode

        path = (request.args.get("path") or "").strip()
        thumb = Path(path) / mindnode.PREVIEW_RELPATH if path else None
        if thumb is None or not thumb.is_file():
            return err("No preview in this document", 404)
        return send_file(str(thumb), mimetype="image/jpeg")

    @combined.route("/api/spreadsheets/google", methods=["GET"])
    def api_spreadsheets_google() -> Response:
        """List the account's spreadsheets for the Start overlay picker.

        Served from the ``_cached_spreadsheet_meta`` TTL cache; ``?refresh=true``
        (the picker's Refresh button) re-lists from Drive so a spreadsheet created
        mid-session shows up without waiting out the TTL.
        """
        if _google_auth.client is None:
            import cli as _cli

            # The searched paths and the setup link ride along on the
            # unauthenticated response only: they answer "what is
            # credentials.json and where does it go", which is the one question
            # the overlay could never answer. cli already computes the list —
            # it just had nowhere to go but a stdout no windowed launch has.
            return ok(
                authenticated=False,
                auth_in_flight=_google_auth.in_flight,
                auth_error=_google_auth.error,
                sheets=[],
                credentials_filename=_cli.CREDENTIALS_FILENAME,
                credentials_paths=[str(p) for p in _cli.credentials_search_paths()],
                credentials_found=str(_cli.resolve_credentials_path() or ""),
                credentials_guide_url="https://docs.gspread.org/en/latest/oauth2.html",
            )
        try:
            metas = _cached_spreadsheet_meta(
                force=request.args.get("refresh") == "true"
            )
        except Exception as exc:
            return ok(
                authenticated=True,
                auth_in_flight=False,
                auth_error=str(exc),
                sheets=[],
            )
        # id stays the name so the open-by-name path is unchanged; modifiedTime
        # (Drive ISO-8601) powers the "Edited …" sub-line in the picker.
        return ok(
            authenticated=True,
            auth_in_flight=False,
            auth_error="",
            sheets=[
                {
                    "name": m["name"],
                    "id": m["name"],
                    "modifiedTime": m.get("modifiedTime", ""),
                }
                for m in metas
            ],
        )

    @combined.route("/api/spreadsheets/worksheets", methods=["GET"])
    def api_spreadsheets_worksheets() -> Response:
        """List a spreadsheet's worksheet titles for the Start overlay dropdown.

        Query params: ``type`` ('google'|'excel') and ``id_or_path``. Returns
        ``{worksheets: [titles…], recommended: "<title>"}``. ``recommended`` is
        the priority auto-pick the open path would use when no tab is chosen.
        Fetched once per selection; the client caches by ``type|id_or_path``.
        """
        type_ = (request.args.get("type") or "").strip()
        id_or_path = (request.args.get("id_or_path") or "").strip()
        if type_ not in ("google", "excel") or not id_or_path:
            return err("Required: type ('google'|'excel') and id_or_path")

        titles: list[str] = []
        recommended = ""
        try:
            if type_ == "excel":
                import excel_io

                titles, recommended = excel_io.list_worksheet_titles(id_or_path)
            else:
                if _google_auth.client is None:
                    return err("Not authenticated with Google.")
                import app as _app

                # A URL resolves without a Drive listing (see 35a7a606); only
                # name lookups need the cached doc list handed down.
                doc_list = (
                    None
                    if id_or_path.startswith(("http://", "https://"))
                    else _spreadsheet_names_for(id_or_path)
                )
                titles, recommended = _app.list_worksheet_titles(
                    _google_auth.client, id_or_path, doc_list=doc_list
                )
        except Exception as exc:
            return err(str(exc), 500)
        return ok(worksheets=titles, recommended=recommended)

    @combined.route("/api/spreadsheets/preview", methods=["GET"])
    def api_spreadsheets_preview() -> Response:
        """Preview the source-video filenames a spreadsheet will expect.

        Query params: ``type`` ('google'|'excel'), ``id_or_path``, optional
        ``worksheet``, and optional ``input_dir`` (the Start overlay's *typed*
        folder, which is not yet the server's effective input dir). Returns
        ``{study, worksheet, unmatched, participants: [{id, filenames, found,
        override, override_value, sheet_value}]}`` — the names
        ``files.resolve_source_video_paths`` will look for, and whether they are
        on disk, so a naming mismatch surfaces before the workspace is opened.
        Each row is editable: see ``/api/spreadsheets/preview/override``.

        Read-only: builds a throwaway :class:`SheetContext` and never calls
        ``_swap_worksheet``, so the active sheet is untouched.

        Cost: one spreadsheet read (``get_all_values``) per (sheet, worksheet)
        pair, duplicating the read ``/api/spreadsheets/open`` does moments later.
        That duplication is deliberate — handing the parsed context off to the
        open path would mean threading it through ``_swap_worksheet`` /
        ``_init_studio_state`` (atomic swap, rollback contract) and opening a
        staleness window between preview and open. The client caches per
        ``type|id_or_path|worksheet|input_dir`` and must never poll this route.
        """
        type_ = (request.args.get("type") or "").strip()
        id_or_path = (request.args.get("id_or_path") or "").strip()
        worksheet = (request.args.get("worksheet") or "").strip() or None
        input_dir = (request.args.get("input_dir") or "").strip()
        if type_ not in ("google", "excel") or not id_or_path:
            return err("Required: type ('google'|'excel') and id_or_path")
        if type_ == "google" and _google_auth.client is None:
            return err("Not authenticated with Google.")

        try:
            ws, _label = _open_worksheet_for(type_, id_or_path, worksheet)
            if ws is None:
                return err("Could not open spreadsheet", 404)
            ctx = spreadsheet.build_sheet_context(ws)
        except Exception as exc:
            return err(str(exc), 500)
        if ctx is None:
            return err(
                "Could not read participants from this worksheet — check that it "
                "has ID, Observation and Category headers and P/G participant "
                "columns."
            )

        participants = spreadsheet.get_participant_list(
            ctx.header_row, ctx.id_cell, ctx.num_participants
        )
        # The sheet's own Filename row only: the user overrides belong to the
        # *previewed* identity, which is usually not the open one, so they are
        # read here rather than taken from config.FILENAME_OVERRIDES.
        sheet_overrides = spreadsheet.participant_filename_overrides(ctx, {})
        loaded_worksheet = getattr(ws, "title", "") or (worksheet or "")
        user_overrides = start_settings.filename_overrides(
            type_, id_or_path, loaded_worksheet
        )
        base_dir = (
            Path(input_dir).expanduser()
            if input_dir
            else utils.get_effective_input_dir()
        )
        rows, unmatched = _preview_source_rows(
            ctx.study_name, participants, sheet_overrides, user_overrides, base_dir
        )
        return ok(
            study=ctx.study_name,
            worksheet=loaded_worksheet,
            participants=rows,
            unmatched=unmatched,
        )

    @combined.route("/api/spreadsheets/preview/override", methods=["POST"])
    def api_spreadsheets_preview_override() -> FlaskResponse:
        """Set (or clear) one participant's source-video filename override.

        Body: ``{type, id_or_path, worksheet, participant, filename, study,
        sheet_value, input_dir}``. An empty *filename* clears the override and
        the participant falls back to *sheet_value* (the sheet's Filename row)
        or ``{study}_{participant}.mp4``.

        Deliberately does not re-read the spreadsheet: the preview it belongs to
        costs a ``get_all_values`` (rate-limited on Google) while re-resolving a
        path is pure disk work, so *study* and *sheet_value* are echoed back
        from the preview payload the client already holds. They only affect what
        this route reports — the authoritative resolution happens against the
        real sheet when the workspace opens.
        """
        data = request.get_json(silent=True) or {}
        type_ = (data.get("type") or "").strip()
        id_or_path = (data.get("id_or_path") or "").strip()
        worksheet = (data.get("worksheet") or "").strip()
        participant = (data.get("participant") or "").strip()
        filename = (data.get("filename") or "").strip()
        study = (data.get("study") or "").strip()
        sheet_value = (data.get("sheet_value") or "").strip()
        input_dir = (data.get("input_dir") or "").strip()
        if type_ not in ("google", "excel", "mindnode") or not id_or_path:
            return err("Required: type ('google'|'excel'|'mindnode') and id_or_path")
        if not participant:
            return err("Required: participant")

        user_overrides = start_settings.set_filename_override(
            type_, id_or_path, worksheet, participant, filename
        )
        # If this is the source the session already has open, the change has to
        # take effect without reopening it — including in the two blueprints
        # that cache their participant list, hence re-seeding rather than
        # assigning the map here.
        active = _active_project_source
        if (
            active
            and active.get("type") == type_
            and active.get("id_or_path") == id_or_path
            and (active.get("worksheet") or "") == worksheet
        ):
            _seed_filename_overrides(active)

        base_dir = (
            Path(input_dir).expanduser()
            if input_dir
            else utils.get_effective_input_dir()
        )
        # Only this participant's row is recomputed. The datalist the client
        # holds may now list a file this row claims; a suggestion list that is
        # one entry stale is not worth a second full preview read.
        rows, _unmatched = _preview_source_rows(
            study,
            [participant],
            {participant: sheet_value or None},
            user_overrides,
            base_dir,
        )
        return ok(row=rows[0])

    @combined.route("/api/spreadsheets/google/auth", methods=["POST"])
    def api_spreadsheets_google_auth() -> FlaskResponse:
        with _google_auth.lock:
            if _google_auth.in_flight:
                return ok(started=False, in_flight=True)
            if _google_auth.client is not None:
                return ok(started=False, authenticated=True)
            _google_auth.in_flight = True
            _google_auth.error = ""

        def _run_auth() -> None:
            try:
                import cli as _cli

                client = _cli.authenticate_google()
                if client is None:
                    # "check credentials.json" is unhelpful when there is no
                    # such file to check — that is a setup step the user has
                    # not done yet, not a broken file. The overlay expands on
                    # whichever of the two this is.
                    _google_auth.error = (
                        "No credentials.json found."
                        if _cli.resolve_credentials_path() is None
                        else "Google sign-in failed — credentials.json was found "
                        "but could not be used."
                    )
                else:
                    _google_auth.client = client
                    # A different account may have signed in; the previous
                    # account's spreadsheet listing must not survive.
                    _invalidate_spreadsheet_meta()
            except Exception as exc:
                # Daemon-thread exceptions otherwise vanish; surface to logs
                # so a misconfigured credentials.json is debuggable.
                utils.error_print(f"Google auth thread failed: {exc}")
                _google_auth.error = str(exc)
            finally:
                _google_auth.in_flight = False

        threading.Thread(target=_run_auth, daemon=True).start()
        return jsonify({"ok": True, "started": True, "in_flight": True}), 202

    @combined.route("/api/spreadsheets/open", methods=["POST"])
    def api_spreadsheets_open() -> FlaskResponse:
        data = request.get_json(silent=True) or {}
        type_ = data.get("type", "")
        id_or_path = (data.get("id_or_path") or "").strip()
        worksheet = (data.get("worksheet") or "").strip() or None
        if type_ not in ("google", "excel", "mindnode") or not id_or_path:
            return err("Required: type ('google'|'excel'|'mindnode') and id_or_path")
        # Absent → None → record_project_session keeps any stored name. Only the
        # Start overlay sends it; Studio's runtime sheet switch never does.
        project_name = data.get("project_name")
        if project_name is not None and not isinstance(project_name, str):
            return err("project_name must be a string")

        if _generation_busy():
            return err(
                "Generation is in progress — wait for it to finish "
                "before switching spreadsheets.",
                409,
            )
        if type_ == "mindnode":
            return _open_mindnode(id_or_path, project_name)
        if type_ == "google" and _google_auth.client is None:
            return err("Not authenticated with Google — click 'Connect Google' first.")

        new_ws: Any = None
        label = ""
        try:
            new_ws, label = _open_worksheet_for(type_, id_or_path, worksheet)
        except Exception as exc:
            return err(str(exc), 500)

        if new_ws is None:
            return err("Could not open spreadsheet", 404)

        # Record the worksheet actually loaded (title after auto-pick fallback)
        # so recent-projects can restore the exact tab, not just the request.
        loaded_worksheet = getattr(new_ws, "title", "") or (worksheet or "")
        source = {
            "type": type_,
            "id_or_path": id_or_path,
            "label": label,
            "worksheet": loaded_worksheet,
        }
        # Before the swap: the blueprint re-inits it runs resolve participants,
        # and they must already see this sheet's overrides. Restored with the
        # rest of the state if the swap fails.
        prev_overrides = config.FILENAME_OVERRIDES
        _seed_filename_overrides(source)
        try:
            _swap_worksheet(new_ws)
        except Exception:
            config.FILENAME_OVERRIDES = prev_overrides
            raise
        if _sheet_context is None:
            return err("Could not parse the spreadsheet", 500)
        start_settings.record_recent_spreadsheet(
            type_, id_or_path, label, loaded_worksheet
        )
        start_settings.record_project_session(
            str(utils.get_effective_input_dir()),
            str(utils.get_effective_output_dir()),
            source,
            name=project_name,
        )
        global _active_sheet_meta, _active_project_source, _startup_notice
        _active_sheet_meta = dict(source)
        _active_project_source = source
        # A sheet is open now — whatever the boot build failed to open is moot.
        _startup_notice = None
        return ok(
            sheet_loaded=True,
            spreadsheet_label=_spreadsheet_label(),
        )

    @combined.route("/api/spreadsheets/close", methods=["POST"])
    def api_spreadsheets_close() -> FlaskResponse:
        global _active_sheet_meta, _mindnode_doc, _active_project_source
        if _generation_busy():
            return err("Generation is in progress — wait for it to finish.", 409)
        # A mind map is an independent source, so closing "the spreadsheet" from
        # the Start overlay must drop whichever one is actually open.
        if (request.get_json(silent=True) or {}).get("type") == "mindnode":
            with _mindnode_lock:
                _mindnode_doc = None
            # Whatever is still open becomes the session's source again.
            _active_project_source = _active_sheet_meta
            _seed_filename_overrides(_active_project_source)
            return ok(sheet_loaded=_worksheet is not None, mindnode_loaded=False)
        _swap_worksheet(None)
        _active_sheet_meta = None
        _active_project_source = _mindnode_source()
        _seed_filename_overrides(_active_project_source)
        return ok(sheet_loaded=False, mindnode_loaded=_mindnode_doc is not None)

    @combined.route("/api/folder-picker", methods=["POST"])
    def api_folder_picker() -> Response:
        """Open the host OS's native folder picker and return the chosen path.

        Called by the Start overlay's Browse buttons. Returns
        ``{ok: True, path: "/…"}`` on confirm, ``{ok: True, path: None}`` when
        the user cancels or the platform has no dialog available.
        """
        data = request.get_json(silent=True) or {}
        initial = (data.get("initial") or "").strip()
        path = utils.open_native_folder_picker(initial)
        return ok(path=path)

    @combined.route("/api/sessions/record", methods=["POST"])
    def api_sessions_record() -> FlaskResponse:
        """Record an "Open workspace" session — used by the no-spreadsheet path.

        The Google/Excel paths already record via api_spreadsheets_open after a
        successful sheet open; this endpoint covers the case where the user
        clicks "Open workspace" on the "No spreadsheet" tab.

        An omitted ``name`` leaves any stored project name alone; ``""`` clears
        it. See :func:`start_settings.record_project_session`.
        """
        data = request.get_json(silent=True) or {}
        input_raw = data.get("input")
        output_raw = data.get("output")
        if input_raw is not None and not isinstance(input_raw, str):
            return err("input must be a string")
        if output_raw is not None and not isinstance(output_raw, str):
            return err("output must be a string")
        input_dir = (input_raw or "").strip()
        output_dir = (output_raw or "").strip()

        name_raw = data.get("name")
        if name_raw is not None and not isinstance(name_raw, str):
            return err("name must be a string")

        spreadsheet_payload = data.get("spreadsheet")
        spreadsheet_dict: dict[str, Any] | None = None
        if spreadsheet_payload is not None:
            if not isinstance(spreadsheet_payload, dict):
                return err("spreadsheet must be an object or null")
            ss_type = (spreadsheet_payload.get("type") or "").strip()
            ss_id = (spreadsheet_payload.get("id_or_path") or "").strip()
            ss_label = (spreadsheet_payload.get("label") or "").strip()
            if ss_type not in ("google", "excel"):
                return err("spreadsheet.type must be 'google' or 'excel'")
            if not ss_id:
                return err("spreadsheet.id_or_path is required")
            spreadsheet_dict = {
                "type": ss_type,
                "id_or_path": ss_id,
                "label": ss_label or ss_id,
            }
        start_settings.record_project_session(
            input_dir or str(utils.get_effective_input_dir()),
            output_dir or str(utils.get_effective_output_dir()),
            spreadsheet_dict,
            name=name_raw,
        )
        return ok()

    @combined.route("/api/changelog")
    def api_changelog() -> Response:
        import changelog

        return ok(entries=changelog.load_entries())

    @combined.route("/api/licenses")
    def api_licenses() -> Response:
        # The SUMMARY table only — the Start overlay's About tab lists what is
        # bundled and links out for the full texts. `--licenses` still prints
        # the whole ~100 KB notice.
        import licenses

        return ok(components=licenses.load_components())

    @combined.route("/api/start-settings", methods=["GET"])
    def api_start_settings_get() -> Response:
        # `desktop` tells the overlay whether to offer the window-rect toggle:
        # in a browser tab there is no window for clipgen to remember.
        return ok(
            settings=start_settings.load_start_settings(),
            desktop=utils.GUI_LAUNCH,
        )

    @combined.route("/api/start-settings", methods=["POST"])
    def api_start_settings_post() -> FlaskResponse:
        data = request.get_json(silent=True) or {}
        if "persist_enabled" in data:
            start_settings.set_persist_enabled(bool(data["persist_enabled"]))
        if "remember_window" in data:
            start_settings.set_remember_window(bool(data["remember_window"]))
        return ok(settings=start_settings.load_start_settings())

    # ---- Shared settings (available from any page) ----

    @combined.route("/api/settings", methods=["GET"])
    def combined_settings_get() -> FlaskResponse:
        return ok(
            settings=_settings_records(),
            path=str(_studio_settings_path()),
            desktop=utils.GUI_LAUNCH,
        )

    @combined.route("/api/settings", methods=["PUT"])
    def combined_settings_put() -> FlaskResponse:
        data = request.get_json(silent=True) or {}
        applied, error = _apply_settings_payload(data)
        if error is not None:
            return err(error)
        return ok(applied=applied)

    @combined.route("/api/settings/reveal", methods=["POST"])
    def combined_settings_reveal() -> FlaskResponse:
        """Show the settings file in the OS file browser.

        Takes no arguments — the path is server-side only. When every setting is
        at its default the file does not exist, so the config dir is revealed
        instead (created first, or there would be nothing to open).
        """
        path = _studio_settings_path()
        if not path.is_file():
            path = path.parent
            try:
                path.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                return err(f"Could not open the folder: {exc}")
        if not utils.reveal_in_file_manager(path):
            return err("Could not open the folder")
        return ok(path=str(path))

    @combined.route("/api/models/llm/reveal", methods=["POST"])
    def combined_llm_reveal() -> FlaskResponse:
        """Show a downloaded model's GGUF in the OS file browser.

        Ollama-installed models have no file in the models dir until they are
        selected, so those reveal their blob in Ollama's own store.
        """
        import llm_client

        name = str((request.get_json(silent=True) or {}).get("model", "")).strip()
        if not name:
            return err("Missing model")
        path = llm_client.model_path(name)
        if path is None:
            return err("Model not found", 404)
        if not utils.reveal_in_file_manager(path):
            return err("Could not open the folder")
        return ok(path=str(path))

    # ---- Titlecard / endcard background picker (shared settings modal) ----
    combined.add_url_rule(
        "/api/titlecards", "combined_titlecards_list", api_titlecards_list
    )
    combined.add_url_rule(
        "/api/titlecards/default/<kind>",
        "combined_titlecard_default",
        api_titlecard_default,
    )
    combined.add_url_rule(
        "/api/titlecards/image/<path:name>",
        "combined_titlecard_image",
        api_titlecard_image,
    )
    combined.add_url_rule(
        "/api/titlecards/upload",
        "combined_titlecard_upload",
        api_titlecard_upload,
        methods=["POST"],
    )
    combined.add_url_rule(
        "/api/titlecards/image/<path:name>",
        "combined_titlecard_delete",
        api_titlecard_delete,
        methods=["DELETE"],
    )

    # The settings modal opens from every page, so its model delete calls the
    # combined root; the rule itself lives on the transcripts blueprint.
    combined.add_url_rule(
        "/api/models/llm/<name>",
        "combined_llm_delete",
        transcripts_server.api_llm_delete,
        methods=["DELETE"],
    )

    # ---- Model discovery ----

    @combined.route("/api/models")
    def api_models() -> Response:
        import llm_client
        import thinking_agents
        import transcripts

        whisper_models = [
            {
                "name": m["name"],
                "size_mb": m["size_mb"],
                "description": m["description"],
                "selected": m["name"] == config.TRANSCRIBE_MODEL,
                "cached": transcripts.is_whisper_model_cached(m["name"]),
            }
            for m in transcripts.WHISPER_MODELS
        ]

        # A filesystem scan, so the catalog answers even with the server down.
        raw = llm_client.list_models()
        # Models the router already refused once. Discovery cannot predict this
        # — an Ollama-converted GGUF looks fine until llama.cpp tries to read
        # it — so the picker reports what the last attempt learned.
        failures = llm_client.load_failures()
        llm_models = [
            {
                "name": m["name"],
                "size_mb": round(m["size_bytes"] / (1024 * 1024)),
                "unusable": failures.get(m["name"], ""),
            }
            for m in raw
        ]

        # Per thinking-agent model + install status, so the Transcripts UI can
        # confirm a download before running an agent against a missing model.
        llm_agents = []
        for a in thinking_agents.AGENTS:
            model = thinking_agents.resolve_model(a)
            llm_agents.append(
                {
                    "key": a["key"],
                    "model": model,
                    "installed": llm_client.is_model_installed(model, raw),
                    "unusable": failures.get(llm_client.model_name(model), ""),
                }
            )

        return ok(
            whisper={"models": whisper_models},
            llm={
                "available": llm_client.is_available(),
                # "not installed" and "installed but not running" need opposite
                # advice, and `available` alone cannot tell them apart — the
                # pages used to tell a user who had never installed the runtime
                # to "start it". install_hint carries the platform-specific
                # commands. Frozen builds bundle llama-server, so installed is
                # effectively always True there.
                "installed": llm_client.is_installed(),
                "install_hint": llm_client.install_guidance_lines(),
                "models": llm_models,
                "agents": llm_agents,
                "base_url": config.LLM_BASE_URL,
            },
        )

    return combined


# Successful (GET + 2xx) hits on these exact paths are suppressed from the
# Werkzeug access log because Studio polls them every 5–10s and the noise
# drowns out real activity. 4xx/5xx still surface, so a silently-failing
# poll endpoint won't be hidden. With VERBOSITY >= VERBOSE everything logs.
_QUIET_POLL_PATHS: frozenset[str] = frozenset(
    {
        "/screenspace/api/tasks",
        "/screenspace/api/events",
        "/screenspace/api/intake-poll",
        "/transcripts/api/intake-poll",
        "/transcripts/api/marks",
        "/transcripts/api/transcribe/status",
        "/transcripts/api/transcribe/model-status",
        "/transcripts/api/participants",
        "/studio/api/job-status",
    }
)


class QuietWSGIRequestHandler(WSGIRequestHandler):
    def log_request(self, code: int | str = "-", size: int | str = "-") -> None:
        if config.VERBOSITY < config.VERBOSE:
            try:
                status = int(code)
            except (TypeError, ValueError):
                status = 0
            if (
                self.command == "GET"
                and 200 <= status < 300
                and self.path.split("?", 1)[0] in _QUIET_POLL_PATHS
            ):
                return
        super().log_request(code, size)


@dataclass
class LiveServer:
    """Handle for a combined server running on a background thread.

    Returned by :func:`serve_combined_app`; must be handed back to
    :func:`stop_combined_app` to shut everything down.
    """

    origin: str  # scheme + host + port, no path
    url: str  # origin plus the landing page's prefix
    port: int
    srv: ThreadedWSGIServer
    thread: threading.Thread
    boot: dict[str, Any]  # dispatcher/build shared state; see serve_combined_app
    ready: threading.Event  # set once the build thread finished (success or not)


# Boot narration, keyed by phase id. The boot page renders `message` verbatim
# (no Python↔JS constant mirroring), and the same strings are echoed to the
# console so a terminal launch narrates the cold start too.
_BOOT_MESSAGES = {
    "starting": "Starting clipgen…",
    "vision_libs": "Loading computer-vision libraries…",
    "workspace": "Preparing workspace…",
    "sheet": "Connecting to your spreadsheet…",
    "interface": "Building the interface…",
    "ready": "Ready",
}


def _boot_wsgi_response(
    start_response: Callable,
    status: str,
    content_type: str,
    body: bytes,
) -> list[bytes]:
    """Emit a dispatcher-owned response (the real app's after_request never runs).

    no-store matters: a cached boot page served after the swap would poll and
    reload forever.
    """
    start_response(
        status,
        [
            ("Content-Type", content_type),
            ("Content-Length", str(len(body))),
            ("Cache-Control", "no-store"),
        ],
    )
    return [body]


def _boot_page_html() -> str:
    """Read the boot page, splicing in the desktop-chrome head when hosted.

    Read per request rather than memoized: the page is served a handful of
    times per launch, ever. utils.DESKTOP_CHROME is set by desktop.py before
    the window exists, so it is already correct on the first GET.
    """
    html = (utils.get_bundled_assets_root() / "assets" / "web" / "boot.html").read_text(
        encoding="utf-8"
    )
    chrome = utils.DESKTOP_CHROME
    if chrome:
        html = html.replace(
            "<!-- CLIPGEN_BOOT_CHROME -->", utils._desktop_chrome_head(chrome)
        )
    return html


def _make_boot_dispatcher(boot_state: dict[str, Any]) -> Callable:
    """WSGI dispatcher that serves a boot page until the real app is built.

    Binding the socket takes milliseconds; building the combined app takes tens
    of seconds on a cold machine (the av/cv2 native libraries alone are ~18s of
    disk I/O). The dispatcher lets the socket go live immediately: while
    ``boot_state["app"]`` is None, page requests get ``assets/web/boot.html``
    (which polls ``/api/boot-status`` and reloads once ready) and API requests
    get a 503 in the standard envelope; afterwards everything is delegated to
    the installed Flask app. ``/api/boot-status`` stays dispatcher-owned
    forever so a poll landing just after the swap cannot 404 against the real
    app.
    """

    def dispatcher(environ: dict[str, Any], start_response: Callable) -> Any:
        path = environ.get("PATH_INFO", "")
        if path == "/api/boot-status":
            if not boot_state.get("first_poll_seen"):
                # First poll proves the boot page's JS is executing — i.e. the
                # window has painted content. One-shot; races at worst record
                # two marks a few ms apart.
                boot_state["first_poll_seen"] = True
                profiling.mark("startup.boot_page_alive")
            body = json.dumps(
                {
                    "ready": boot_state["ready"],
                    "phase": boot_state["phase"],
                    "message": boot_state["message"],
                    "error": boot_state["error"],
                }
            ).encode("utf-8")
            return _boot_wsgi_response(
                start_response, "200 OK", "application/json", body
            )
        # Single read → the swap is atomic per request; a request that read the
        # boot handler finishes against it harmlessly.
        app = boot_state["app"]
        if app is not None:
            return app(environ, start_response)
        if "/api/" in path:
            # Stale tabs and satellites poll during boot; the standard envelope
            # lets their createPoller loops degrade quietly instead of parsing
            # HTML as JSON.
            body = json.dumps({"ok": False, "error": "Server is starting up"}).encode(
                "utf-8"
            )
            return _boot_wsgi_response(
                start_response, "503 Service Unavailable", "application/json", body
            )
        return _boot_wsgi_response(
            start_response,
            "200 OK",
            "text/html; charset=utf-8",
            _boot_page_html().encode("utf-8"),
        )

    return dispatcher


def _port_available(port: int) -> bool:
    """Report whether *port* can be bound on loopback right now.

    Probing first keeps werkzeug's own "Port N is in use" message (and its
    ``sys.exit(1)``) off the console in the ordinary second-instance case.
    """
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        # Match werkzeug's bind semantics (HTTPServer sets allow_reuse_address):
        # without SO_REUSEADDR the probe reports a port still draining TIME_WAIT
        # connections as taken, and a quick restart silently relocates off 8089
        # even though the real bind would have succeeded.
        probe.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            probe.bind(("127.0.0.1", port))
        except OSError:
            return False
    return True


def serve_combined_app(
    worksheet: Any = None,
    port: int | None = None,
    default_page: str = "studio",
    gspread_client: Any = None,
    gspread_client_factory: Any = None,
    worksheet_factory: Any = None,
    block_until_ready: bool = False,
) -> LiveServer:
    """Serve the combined app on a background thread and return once listening.

    Unlike :func:`start_combined_server` this does not block and does not open a
    browser, so a desktop host can point a native window at ``LiveServer.url``
    and tear the server down again when the window closes.

    The socket goes live in well under a second serving a boot page; the heavy
    build (av/cv2 preload + blueprint imports, tens of seconds on a cold
    machine) runs on a background thread and is swapped in when done. Pass
    ``block_until_ready=True`` to instead wait for the real app — tests and
    scripted callers want the built app, not the boot shell.

    *gspread_client* is a client the caller already authenticated;
    *gspread_client_factory* is the deferred form — called on the build thread
    (so the gspread import stays off the caller's window-paint path) and only
    when no ready client was passed. Building the client off the request
    threads is the established pattern here: the runtime "Connect Google"
    route does the same on a daemon thread.

    *worksheet_factory* is the same deferral for a `-s` launch's worksheet:
    ``factory(client) -> (worksheet, notice)``, run on the build thread under
    ``NO_INPUT_MODE`` (prompts degrade to failure). On failure the build
    continues sheetless and *notice* lands in ``_startup_notice`` — the Start
    overlay, not a dead boot-error page, is the recovery surface.
    """
    # The web server has no interactive console: every request/run/background
    # task executes on a Flask/daemon thread with no attached stdin. Force
    # non-interactive resolution so a missing source video (or any pipeline
    # prompt) is skipped-and-reported instead of blocking the thread forever on
    # ``input()`` — this previously hung Studio generate and watch-dir-triggered
    # workflow runs alike. Set synchronously: requests can arrive before the
    # build thread runs.
    utils.NO_INPUT_MODE = True

    # A fresh serve must not inherit a prior serve's notice (tests spin up
    # several servers per process).
    global _startup_notice
    _startup_notice = None

    boot_state: dict[str, Any] = {
        "ready": False,
        "phase": "starting",
        "message": _BOOT_MESSAGES["starting"],
        "error": None,
        "app": None,
    }
    ready = threading.Event()
    dispatcher = _make_boot_dispatcher(boot_state)

    # Not `port or ...`: 0 is a meaningful value here (bind an ephemeral port).
    requested = port if port is not None else config.SERVER_PORT
    if requested and not _port_available(requested):
        # Usually a second clipgen instance already holding the port. Fall back
        # to an ephemeral one rather than refusing to start at all.
        utils.warning_print(
            f"Port {requested} is already in use — starting on a free port instead."
        )
        requested = 0
    try:
        srv = make_server(
            "127.0.0.1",
            requested,
            dispatcher,
            threaded=True,
            request_handler=QuietWSGIRequestHandler,
        )
    except (OSError, SystemExit):
        # Belt and braces for the race between the probe above and this bind.
        # werkzeug turns EADDRINUSE into sys.exit(1), so SystemExit — which is a
        # BaseException — has to be caught explicitly alongside OSError.
        if requested == 0:
            raise
        srv = make_server(
            "127.0.0.1",
            0,
            dispatcher,
            threaded=True,
            request_handler=QuietWSGIRequestHandler,
        )

    # `threaded=True` guarantees this, but make_server's return type is the base
    # class and only the ThreadingMixIn subclass has block_on_close.
    assert isinstance(srv, ThreadedWSGIServer)
    # socketserver defaults block_on_close to True, so server_close() joins every
    # tracked connection thread. A single open SSE stream at teardown would hang
    # shutdown forever — the nastiest failure mode in this path.
    srv.block_on_close = False

    thread = threading.Thread(
        target=srv.serve_forever, daemon=True, name="clipgen-server"
    )
    thread.start()

    utils.info_print(
        "Starting clipgen — the first launch after a restart can take up to a "
        "minute while video libraries load."
    )
    started = time.monotonic()

    def set_phase(phase: str) -> None:
        boot_state["phase"] = phase
        boot_state["message"] = _BOOT_MESSAGES[phase]
        profiling.mark(f"startup.phase_{phase}")
        utils.info_print(_BOOT_MESSAGES[phase])

    def build() -> None:
        try:
            # Preload first, under the stderr suppressor, before the blueprint
            # imports below pull cv2 themselves — see preload_vision_libs_quietly
            # for why order matters.
            utils.preload_vision_libs_quietly(
                on_phase=lambda lib: set_phase("vision_libs")
            )
            set_phase("workspace")
            # Reclaim orphaned scratch files (atomic-write .tmp siblings, reel
            # temp-clips) a prior hard kill may have left in the output dir.
            # Must stay before the build: it unlinks *.json.tmp regardless of
            # age, and build_combined_app starts workers that atomic-write.
            utils.sweep_stale_temp_artifacts()
            client = gspread_client
            if client is None and gspread_client_factory is not None:
                client = gspread_client_factory()
                profiling.mark("startup.silent_google_auth")
            ws = worksheet
            if ws is None and worksheet_factory is not None:
                set_phase("sheet")
                ws, notice = worksheet_factory(client)
                if notice:
                    global _startup_notice
                    _startup_notice = notice
                profiling.mark("startup.worksheet_opened")
            set_phase("interface")
            combined = build_combined_app(
                worksheet=ws,
                default_page=default_page,
                gspread_client=client,
            )
            boot_state["app"] = combined
            set_phase("ready")
            boot_state["ready"] = True
            utils.info_print(f"clipgen ready in {time.monotonic() - started:.1f}s")
            if config.PROFILING:
                # A live desktop session should not have to exit (atexit
                # report) to see the startup attribution.
                profiling.report_startup()
            if _google_auth.client is not None:
                # Warm the 300s-TTL Drive listing cache so the auto-opened
                # Start overlay's Google panel answers instantly. Not an extra
                # API call in the common flow: the overlay would issue the
                # same files.list moments later, and the single-flight cache
                # dedupes the two.
                threading.Thread(
                    target=_warm_spreadsheet_meta_cache,
                    daemon=True,
                    name="clipgen-drive-warm",
                ).start()
        except BaseException as exc:
            # BaseException on purpose: _init_studio_state exits via sys.exit(1)
            # on a bad worksheet, and a SystemExit swallowed by a daemon thread
            # would leave the boot page spinning forever.
            boot_state["error"] = f"{type(exc).__name__}: {exc}"
            utils.error_print(
                "clipgen failed to start.",
                details=traceback.format_exc().strip().splitlines(),
            )
        finally:
            ready.set()

    threading.Thread(target=build, daemon=True, name="clipgen-boot-build").start()

    if block_until_ready:
        finished = ready.wait(timeout=300)
        if not finished or boot_state["error"] is not None:
            srv.shutdown()
            srv.server_close()
            raise RuntimeError(boot_state["error"] or "clipgen server build timed out")

    origin = f"http://127.0.0.1:{srv.server_port}"
    return LiveServer(
        origin=origin,
        url=f"{origin}/{default_page}/",
        port=srv.server_port,
        srv=srv,
        thread=thread,
        boot=boot_state,
        ready=ready,
    )


def stop_combined_app(live: LiveServer) -> None:
    """Close the socket, then stop every thread ``build_combined_app`` started.

    The Screenspace and Transcripts workers and the Workflows watch-dir thread
    are *module* globals rather than app-scoped, so without this they outlive
    the Flask app and keep polling whatever ``config.INPUT_DIR`` points at next.
    """
    live.srv.shutdown()
    live.srv.server_close()
    live.thread.join(timeout=10)

    # A close during boot: give an almost-finished build a moment to settle,
    # then skip the worker teardown if the app never got installed — the
    # blueprint modules were never imported, there are no workers to stop, and
    # importing screenspace_server cold just to find None would itself cost
    # seconds. A build still in flight is a daemon thread; it dies with the
    # process, which is where every mid-boot close leads anyway.
    live.ready.wait(timeout=1)
    if not live.boot.get("ready"):
        return

    # Imported lazily for the same reason build_combined_app does: keeping cv2
    # and onnxruntime off the import path of callers that never serve Screenspace.
    import screenspace_server
    import transcripts_server
    import workflows_server

    if screenspace_server._worker is not None:
        screenspace_server._worker.stop()
        screenspace_server._worker = None
    if transcripts_server._worker is not None:
        transcripts_server._worker.stop()
        transcripts_server._worker = None
    workflows_server._watch_stop.set()
    if workflows_server._watch_thread is not None:
        workflows_server._watch_thread.join(timeout=5)
        workflows_server._watch_thread = None
    workflows_server._watch_stop.clear()


def start_combined_server(
    worksheet: Any = None,
    port: int | None = None,
    default_page: str = "studio",
    gspread_client: Any = None,
    gspread_client_factory: Any = None,
    worksheet_factory: Any = None,
) -> None:
    """Start a combined Studio + Screenspace + Transcripts server on one port.

    All three blueprints are always registered. When *worksheet* is ``None``,
    sheet-dependent Studio routes return ``sheet_loaded: false`` placeholder
    responses; the frontend's Start overlay lets the user pick a spreadsheet
    via ``POST /api/spreadsheets/open``.

    If *gspread_client* is supplied (the CLI's auth already happened upstream),
    the Google Sheets list endpoint reports authenticated and skips the
    "Connect Google" CTA in the Start overlay.

    Opens the user's default browser once the socket is bound — the tab shows
    the boot page until the background build swaps the real app in — and blocks
    until interrupted. Desktop launches use :func:`serve_combined_app` directly.
    """
    live = serve_combined_app(
        worksheet=worksheet,
        port=port or config.SERVER_PORT,
        default_page=default_page,
        gspread_client=gspread_client,
        gspread_client_factory=gspread_client_factory,
        worksheet_factory=worksheet_factory,
    )
    utils.info_print(f"clipgen server running at {live.origin}")
    webbrowser.open(live.url)
    try:
        # Join in slices rather than blocking forever: on Windows a bare join
        # swallows Ctrl+C until the thread exits.
        while live.thread.is_alive():
            live.thread.join(timeout=1)
    except KeyboardInterrupt:
        pass
    finally:
        stop_combined_app(live)
