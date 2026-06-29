# -*- coding: utf-8 -*-
"""Combined Flask server for clipgen Studio, Screenspace, Transcripts, and Workflows.

Entry point: start_combined_server(worksheet, port, default_page) registers
Studio, Screenspace, Transcripts, and Workflows blueprints on one app at config.SERVER_PORT (8089).
Module-level state: _worksheet, _sheet_context, _generated_artifacts, _generated_reels
(initialized by _init_studio_state()).

Studio API endpoints (all under /studio/):
  GET  /api/sheet              – spreadsheet grid data (rows, participants, timestamps)
  POST /api/sheet/refresh      – re-fetch spreadsheet data from source (Google/Excel)
  GET  /api/thumbnail/<p>/<t>  – JPEG thumbnail frame from participant video
  POST /api/generate           – generate clip/screen/gif artifacts for specified cells
  POST /api/generate/cancel    – cancel an in-progress clip generation
  POST /api/highlights-preview – preview highlights reel selection without generating
  POST /api/reel               – build a reel from specified cells
  POST /api/reel-direct        – build a reel from explicit clip paths
  POST /api/reel/cancel        – cancel an in-progress reel build
  POST /api/viewer             – generate timeline viewer from session artifacts
  POST /api/timeline-viewer    – batch-export all clips and generate timeline viewer
  POST /api/gallery            – generate gallery from a video file
  GET/POST /api/manifest       – read or write the cumulative artifact manifest
  POST /api/regenerate         – regenerate all media from saved manifest
  GET/POST /api/stashes        – reel stash CRUD
  GET/POST /api/artifact-stashes – artifact stash CRUD
  POST /api/generate-intake    – generate artifacts from an intake/screenspace manifest
  GET  /api/sheet/baseline      – per-participant baseline timestamps for convergence
  GET  /api/convergence/offsets – per-participant convergence display offsets
  PUT  /api/convergence/offsets – persist per-participant convergence display offsets
  GET  /api/settings           – read current config settings
  PUT  /api/settings           – update config settings
  GET  /api/status             – reports which interfaces are active (studio/screenspace/transcripts)
"""

import concurrent.futures
import copy
import hashlib
import json
import math
import os
import queue
import re
import string
import sys
import threading
import time
import webbrowser
from collections import OrderedDict
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from collections.abc import Iterator
from typing import Any, cast

from flask import (
    Blueprint,
    Flask,
    Response,
    jsonify,
    redirect,
    request,
    send_file,
    send_from_directory,
)
from flask.json.provider import DefaultJSONProvider
from werkzeug.serving import WSGIRequestHandler

import config
import files
import spreadsheet
import pipeline
import titlecards
import utils
import video
import viewer

FlaskResponse = Response | tuple[Response, int]

# ---- Module-level state (set once by _init_studio_state) ----

_worksheet: Any = None
_sheet_context: spreadsheet.SheetContext | None = None
# Metadata for the active spreadsheet — used by /api/status so the Start
# overlay can pre-select the right tab (Google/Excel) and re-highlight the
# right item. Only set when the spreadsheet was opened via the runtime
# picker; CLI-loaded sheets leave this None.
_active_sheet_meta: dict[str, str] | None = None
_generated_artifacts: list[dict[str, Any]] = []
# Index by (cellRow, cellCol, type) for O(1) lookup in /api/generate Phase 1.
# Mutated under _generated_output_lock together with _generated_artifacts.
_generated_artifacts_index: dict[tuple[int, int, str], list[dict[str, Any]]] = {}
_generated_reels: list[dict[str, Any]] = []
# Bounded LRU so long sessions don't grow unbounded; entries are JPEG bytes
# (tens of KB each), so a few hundred is plenty.
_THUMBNAIL_CACHE_MAX = 256
_thumbnail_cache: "OrderedDict[tuple, bytes]" = OrderedDict()
# Guards every thumbnail cache access (get / move_to_end / insert / evict);
# api_thumbnail is served concurrently by Flask's threaded dev server.
_thumbnail_cache_lock = threading.Lock()
# Card-scrubber assets (opt-in hover preview). Sprite sheets are small JPEGs;
# audio segments are PCM WAV (~1 MB per short clip), so cap audio far lower.
_SPRITE_CACHE_MAX = 256
_sprite_cache: "OrderedDict[tuple, bytes]" = OrderedDict()
_sprite_cache_lock = threading.Lock()
_AUDIO_CACHE_MAX = 32
_audio_cache: "OrderedDict[tuple, bytes]" = OrderedDict()
_audio_cache_lock = threading.Lock()
# A second concurrent /api/generate or /api/reel call would clobber the
# shared cancel event; reject with 409 instead while one is in flight.
_reel_cancel_event = threading.Event()
_generate_cancel_event = threading.Event()
# Independent cancel event for /api/generate-intake — sheet and intake
# branches run concurrently from one Studio Generate click but have
# separate streams, so the Cancel button posts to both /api/generate/cancel
# and /api/generate-intake/cancel.
_intake_cancel_event = threading.Event()
# Cancel events for the two long-running viewer builds. /api/timeline-viewer
# re-cuts every clip (process_clips) plus optional intake clips; /api/gallery
# extracts frames/GIFs (generate_interval_captures). Both run synchronously in
# the request thread, but Flask is threaded so the matching /cancel endpoint
# can set the event mid-build and the cancel_flag checks short-circuit it.
_timeline_viewer_cancel_event = threading.Event()
_gallery_cancel_event = threading.Event()
_busy_lock = threading.Lock()
_generate_in_progress = False
_reel_in_progress = False
# Single-job slots for the two long-running viewer builds. Each shares one
# module-level cancel event, so a second concurrent build (e.g. a second Studio
# tab) must be rejected rather than allowed to clear/clobber the other's signal.
_timeline_viewer_in_progress = False
_gallery_in_progress = False
# Count of in-flight /api/generate-intake streams. Intake has no single-job
# slot (it must run alongside /api/generate for mixed queues), but a sheet
# swap still needs to know whether any intake work is active.
_intake_active = 0
# Serializes load → mutate → save for the stash manifests so concurrent
# stash CRUD requests don't drop each other's writes.
_stash_lock = threading.Lock()
# Serializes mutations to in-memory generated lists and quiet manifest saves.
_generated_output_lock = threading.Lock()
# Latest progress snapshots for in-flight jobs, exposed by /api/job-status so
# the Studio UI can re-attach (show progress + Cancel) after the user
# navigates away to /screenspace/ or /transcripts/ mid-build and comes back.
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

# Snapshot config defaults before any settings file is loaded.
# Deep-copied so dict-valued defaults are not aliased to live config state.
_settings_defaults: dict[str, Any] = {
    name: copy.deepcopy(getattr(config, name))
    for name in getattr(config, "STUDIO_SETTINGS", {})
}

_HEX_COLOR_RE = re.compile(r"^#[0-9a-fA-F]{6}$")
_MARK_KEY_RE = re.compile(r"^[a-z0-9_]+$")


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


def _thumbnail_cache_put(key: tuple, value: bytes) -> None:
    """Insert into the thumbnail cache with simple LRU eviction."""
    with _thumbnail_cache_lock:
        _thumbnail_cache[key] = value
        while len(_thumbnail_cache) > _THUMBNAIL_CACHE_MAX:
            _thumbnail_cache.popitem(last=False)


def _sprite_cache_put(key: tuple, value: bytes) -> None:
    """Insert into the sprite-sheet cache with simple LRU eviction."""
    with _sprite_cache_lock:
        _sprite_cache[key] = value
        while len(_sprite_cache) > _SPRITE_CACHE_MAX:
            _sprite_cache.popitem(last=False)


def _audio_cache_put(key: tuple, value: bytes) -> None:
    """Insert into the clip-audio cache with simple LRU eviction."""
    with _audio_cache_lock:
        _audio_cache[key] = value
        while len(_audio_cache) > _AUDIO_CACHE_MAX:
            _audio_cache.popitem(last=False)


def _try_claim_busy(slot: str) -> bool:
    """Atomically reserve the single-job slot for *slot*.

    Valid slots: ``'generate'``, ``'reel'``, ``'timeline_viewer'``, ``'gallery'``.
    Returns True on success (caller must call ``_release_busy`` when done) or
    False if another request is already holding the slot.
    """
    global _generate_in_progress, _reel_in_progress, _timeline_viewer_in_progress, _gallery_in_progress  # noqa: PLW0603
    with _busy_lock:
        if slot == "generate":
            if _generate_in_progress:
                return False
            _generate_in_progress = True
            return True
        if slot == "reel":
            if _reel_in_progress:
                return False
            _reel_in_progress = True
            return True
        if slot == "timeline_viewer":
            if _timeline_viewer_in_progress:
                return False
            _timeline_viewer_in_progress = True
            return True
        if slot == "gallery":
            if _gallery_in_progress:
                return False
            _gallery_in_progress = True
            return True
    return False


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
    """Release the single-job slot for *slot*.

    Valid slots: ``'generate'``, ``'reel'``, ``'timeline_viewer'``, ``'gallery'``.
    """
    global _generate_in_progress, _reel_in_progress, _timeline_viewer_in_progress, _gallery_in_progress  # noqa: PLW0603
    with _busy_lock:
        if slot == "generate":
            _generate_in_progress = False
        elif slot == "reel":
            _reel_in_progress = False
        elif slot == "timeline_viewer":
            _timeline_viewer_in_progress = False
        elif slot == "gallery":
            _gallery_in_progress = False


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
    """Initialize the generate progress snapshot when /api/generate starts."""
    with _job_state_lock:
        _generate_job_state["total"] = max(0, int(total))
        _generate_job_state["done"] = 0
        _generate_job_state["started_at"] = time.time()


def _increment_generate_done(n: int = 1) -> None:
    """Advance the generate-job 'done' counter by n (one per yielded line)."""
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
    global _intake_active  # noqa: PLW0603
    with _busy_lock:
        _intake_active += 1 if active else -1


def _generation_busy() -> bool:
    """Return True while any clip, reel, or intake generation is in flight.

    Consulted before a sheet swap so generated lists are not rebound under
    an active stream.
    """
    with _busy_lock:
        return _generate_in_progress or _reel_in_progress or _intake_active > 0


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

utils.register_static_routes(studio_bp, "studio.html", icons=True)


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
    p_idx = participants.index(participant)
    col_idx = ctx.id_cell.col + p_idx

    override = None
    if ctx.filename_row_idx is not None:
        row_data = ctx.sheet_data[ctx.filename_row_idx]
        if col_idx < len(row_data) and row_data[col_idx].strip():
            override = row_data[col_idx].strip()

    return files.resolve_source_video_paths(
        ctx.study_name, participant, override, utils.get_effective_input_dir()
    )


def _resolve_source_video(participant: str) -> Path | None:
    """Return the first source-video path for a participant, or None."""
    sources = _resolve_participant_sources(participant)
    return sources[0] if sources else None


# ---- API endpoints ----


@studio_bp.route("/api/thumbnail/<participant>/<start_seconds>")
def api_thumbnail(participant: str, start_seconds: str) -> FlaskResponse:
    if _sheet_context is None:
        return jsonify({"ok": False, "error": "No spreadsheet loaded"}), 404

    try:
        # Parse via float so a fractional second (thumbnails are second-granular,
        # so floor it) does not 400 the way bare int("12.5") would; matches the
        # float-tolerant parsing the project's other media routes use.
        start_sec = max(0, int(float(start_seconds)))
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "Invalid timestamp"}), 400
    sources = _resolve_participant_sources(participant)
    if not sources or not sources[0].is_file():
        return jsonify({"ok": False, "error": "Source video not found"}), 404

    # Multi-video participant: map the global second into the owning sub-video so
    # the hover thumbnail comes from the right file at the right local offset.
    cut_sec = start_sec
    video_path = sources[0]
    if len(sources) >= 2:
        timeline = video.build_source_timeline([str(p) for p in sources])
        if timeline is None:
            return jsonify({"ok": False, "error": "Source video not found"}), 404
        mapped = utils.resolve_timeline_segment(timeline, start_sec)
        if mapped is None:
            return jsonify({"ok": False, "error": "Timestamp beyond recording"}), 404
        video_path = Path(mapped[0])
        cut_sec = int(mapped[1])

    try:
        mtime = video_path.stat().st_mtime
    except OSError:
        mtime = 0.0
    # Include mtime so replacing a source file on disk invalidates stale thumbnails.
    cache_key = (str(video_path), cut_sec, mtime)
    with _thumbnail_cache_lock:
        cached = _thumbnail_cache.get(cache_key)
        if cached is not None:
            _thumbnail_cache.move_to_end(cache_key)
    if cached is not None:
        return Response(
            cached,
            mimetype="image/jpeg",
            headers={"Cache-Control": "public, max-age=86400"},
        )

    jpeg_bytes = video.extract_thumbnail_bytes(
        str(video_path), cut_sec, width=config.STUDIO_THUMBNAIL_WIDTH
    )
    if jpeg_bytes is None:
        return jsonify({"ok": False, "error": "Thumbnail extraction failed"}), 404

    _thumbnail_cache_put(cache_key, jpeg_bytes)
    return Response(
        jpeg_bytes,
        mimetype="image/jpeg",
        headers={"Cache-Control": "public, max-age=86400"},
    )


def _parse_clip_window() -> tuple[float, float] | None:
    """Parse + validate the ``?start=&end=`` seconds shared by the scrubber
    media routes. Returns ``(start_seconds, duration_seconds)`` or ``None`` when
    the params are missing/non-numeric or the range is empty."""
    try:
        start_sec = max(0.0, float(request.args.get("start", "")))
        end_sec = float(request.args.get("end", ""))
    except (ValueError, TypeError):
        return None
    duration = end_sec - start_sec
    if duration <= 0:
        return None
    return start_sec, duration


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


@studio_bp.route("/api/sprite/<participant>")
def api_sprite(participant: str) -> FlaskResponse:
    """Tiled JPEG sprite sheet of a clip for the opt-in hover card scrubber.

    A clip straddling a multi-video boundary maps its *start* into the owning
    sub-video and the sprite samples that file's tail (hover preview only).
    """
    if _sheet_context is None:
        return jsonify({"ok": False, "error": "No spreadsheet loaded"}), 404
    window = _parse_clip_window()
    if window is None:
        return jsonify({"ok": False, "error": "Invalid clip range"}), 400
    start_sec, duration = window

    resolved = _resolve_clip_media_source(participant, start_sec)
    if resolved is None:
        return jsonify({"ok": False, "error": "Source video not found"}), 404
    video_path, local_start = resolved
    cols = config.STUDIO_SCRUBBER_SPRITE_COLS
    rows = config.STUDIO_SCRUBBER_SPRITE_ROWS

    try:
        mtime = video_path.stat().st_mtime
    except OSError:
        mtime = 0.0
    cache_key = (
        str(video_path),
        round(local_start, 3),
        round(duration, 3),
        cols,
        rows,
        mtime,
    )
    with _sprite_cache_lock:
        cached = _sprite_cache.get(cache_key)
        if cached is not None:
            _sprite_cache.move_to_end(cache_key)
    if cached is not None:
        return Response(
            cached,
            mimetype="image/jpeg",
            headers={"Cache-Control": "public, max-age=86400"},
        )

    sprite_bytes = video.extract_sprite_sheet_bytes(
        str(video_path), local_start, duration, cols, rows
    )
    if sprite_bytes is None:
        return jsonify({"ok": False, "error": "Sprite extraction failed"}), 404

    _sprite_cache_put(cache_key, sprite_bytes)
    return Response(
        sprite_bytes,
        mimetype="image/jpeg",
        headers={"Cache-Control": "public, max-age=86400"},
    )


@studio_bp.route("/api/clip-audio/<participant>")
def api_clip_audio(participant: str) -> FlaskResponse:
    """Mono WAV of a clip's audio for the opt-in hover card scrubber.

    PCM WAV (not the source's compressed audio) so the browser's WebAudio
    ``decodeAudioData`` decodes it reliably. Boundary-straddling clips map their
    *start* into the owning sub-video (hover preview only).
    """
    if _sheet_context is None:
        return jsonify({"ok": False, "error": "No spreadsheet loaded"}), 404
    window = _parse_clip_window()
    if window is None:
        return jsonify({"ok": False, "error": "Invalid clip range"}), 400
    start_sec, duration = window

    resolved = _resolve_clip_media_source(participant, start_sec)
    if resolved is None:
        return jsonify({"ok": False, "error": "Source video not found"}), 404
    video_path, local_start = resolved

    try:
        mtime = video_path.stat().st_mtime
    except OSError:
        mtime = 0.0
    cache_key = (str(video_path), round(local_start, 3), round(duration, 3), mtime)
    with _audio_cache_lock:
        cached = _audio_cache.get(cache_key)
        if cached is not None:
            _audio_cache.move_to_end(cache_key)
    if cached is not None:
        return Response(
            cached,
            mimetype="audio/wav",
            headers={"Cache-Control": "public, max-age=86400"},
        )

    wav_bytes = video.extract_audio_segment_bytes(
        str(video_path), local_start, duration
    )
    if wav_bytes is None:
        return jsonify({"ok": False, "error": "Audio extraction failed"}), 404

    _audio_cache_put(cache_key, wav_bytes)
    return Response(
        wav_bytes,
        mimetype="audio/wav",
        headers={"Cache-Control": "public, max-age=86400"},
    )


@studio_bp.route("/api/sheet")
def api_sheet() -> FlaskResponse:
    if _sheet_context is None:
        return jsonify(
            {
                "ok": True,
                "sheet_loaded": False,
                "study": "",
                "version": utils.get_version(),
                "highlightsDuration": config.HIGHLIGHTS_REEL_DURATION_SECONDS,
                "titlecardsEnabled": config.TITLECARDS_ENABLED,
                "titlecardDuration": config.TITLECARD_DURATION_SECONDS,
                "cellExpandHover": config.STUDIO_CELL_EXPAND_HOVER,
                "cardScrubberEnabled": config.STUDIO_CARD_SCRUBBER,
                "config": utils.get_frontend_config(),
                "participants": [],
                "rows": [],
            }
        )

    ctx = _sheet_context
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

    return jsonify(
        {
            "ok": True,
            "sheet_loaded": True,
            "study": ctx.study_name,
            "version": utils.get_version(),
            "highlightsDuration": config.HIGHLIGHTS_REEL_DURATION_SECONDS,
            "titlecardsEnabled": config.TITLECARDS_ENABLED,
            "titlecardDuration": config.TITLECARD_DURATION_SECONDS,
            "cellExpandHover": config.STUDIO_CELL_EXPAND_HOVER,
            "cardScrubberEnabled": config.STUDIO_CARD_SCRUBBER,
            "defaultDuration": config.DEFAULT_DURATION_SECONDS,
            "config": utils.get_frontend_config(),
            "participants": participants,
            "rows": rows,
        }
    )


@studio_bp.route("/api/sheet/baseline")
def api_sheet_baseline() -> FlaskResponse:
    """Return per-participant baseline offsets in seconds for convergence.

    Response: {"ok": true, "baselines": {"P01": 33120, "P02": 0, ...}}
    Values are integers (seconds) parsed via _clock_to_seconds which
    correctly treats "22:00" as HH:MM (22 hours) rather than MM:SS.
    Empty baselines dict when no baseline row exists.
    """
    if _sheet_context is None:
        return jsonify({"ok": True, "sheet_loaded": False, "baselines": {}})

    ctx = _sheet_context
    if ctx.baseline_row_idx is None:
        return jsonify({"ok": True, "baselines": {}})

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

    return jsonify({"ok": True, "baselines": baselines})


def _clean_convergence_offsets(raw: object) -> dict[str, dict[str, float]]:
    """Normalize nested per-lane convergence offsets to {pid: {source: float}}.

    Drops: non-string/empty participant ids, non-dict participant values,
    unknown source keys (outside config.CONVERGENCE_SOURCES), non-numeric /
    non-finite / zero lane values, and participants left with no lanes.
    """
    cleaned: dict[str, dict[str, float]] = {}
    if not isinstance(raw, dict):
        return cleaned
    raw_map = cast(dict[str, Any], raw)
    for pid, lanes in raw_map.items():
        if not isinstance(pid, str) or not pid or not isinstance(lanes, dict):
            continue
        lane_map = cast(dict[str, Any], lanes)
        clean_lanes: dict[str, float] = {}
        for source, value in lane_map.items():
            if source not in config.CONVERGENCE_SOURCES:
                continue
            try:
                num = float(value)
            except (TypeError, ValueError):
                continue
            if math.isfinite(num) and num != 0:
                clean_lanes[source] = num
        if clean_lanes:
            cleaned[pid] = clean_lanes
    return cleaned


@studio_bp.route("/api/convergence/offsets")
def api_convergence_offsets_get() -> FlaskResponse:
    """Return persisted per-lane convergence display offsets (seconds, signed).

    Independent from /api/sheet/baseline: baselines convert sheet wall-clock
    to video-time (sheet-only). Offsets shift a participant's events per data
    source so misaligned recording start times — or a single drifting source
    such as spreadsheet timestamps — can be nudged until lanes line up visually
    in the Convergence Browser.

    Response: {"ok": true, "offsets": {"P01": {"sheet": 12.5, "screenspace": 12.5}}}
    """
    data = utils.load_json_manifest(
        config.CONVERGENCE_OFFSETS_FILENAME, default={"offsets": {}}
    )
    raw = data.get("offsets") if isinstance(data, dict) else None
    return jsonify({"ok": True, "offsets": _clean_convergence_offsets(raw)})


@studio_bp.route("/api/convergence/offsets", methods=["PUT"])
def api_convergence_offsets_put() -> FlaskResponse:
    """Persist per-lane convergence display offsets.

    Body: {"offsets": {"P01": {"sheet": 12.5, ...}, ...}}. Unknown sources,
    zeros, and non-finite values are dropped per lane; participants left with
    no lanes are dropped. When the cleaned dict is empty, the manifest file is
    deleted so a clean output dir has no leftover empty manifest.
    """
    data = request.get_json(silent=True) or {}
    raw = data.get("offsets")
    if not isinstance(raw, dict):
        return jsonify({"ok": False, "error": "Invalid offsets payload"}), 400

    cleaned = _clean_convergence_offsets(raw)

    settings_path = (
        Path(utils.get_effective_output_dir()) / config.CONVERGENCE_OFFSETS_FILENAME
    )
    if not cleaned:
        if settings_path.is_file():
            try:
                settings_path.unlink()
            except OSError:
                pass
    else:
        utils.save_json_manifest(
            config.CONVERGENCE_OFFSETS_FILENAME, {"offsets": cleaned}
        )

    return jsonify({"ok": True, "offsets": cleaned})


@studio_bp.route("/api/sheet/refresh", methods=["POST"])
def api_sheet_refresh() -> FlaskResponse:
    global _sheet_context
    if _worksheet is None:
        return jsonify(
            {
                "ok": False,
                "error": "No spreadsheet loaded — pick one from the Start panel.",
            }
        ), 400
    new_context = spreadsheet.build_sheet_context(_worksheet)
    if new_context is None:
        return jsonify({"ok": False, "error": "Failed to refresh sheet data"}), 500
    _sheet_context = new_context
    return jsonify({"ok": True})


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

    Tries the source-specific participant list first, then falls back to the
    other.  Both lists are populated from the same source videos so the
    fallback is a safety net. Returns [] when the participant has no video.
    A multi-video participant returns all parts (one continuous timeline).
    """
    import screenspace_server
    import transcripts_server

    lists = (
        [transcripts_server._participants, screenspace_server._participants]
        if source == "transcript"
        else [screenspace_server._participants, transcripts_server._participants]
    )
    for plist in lists:
        for p in plist:
            if p["id"] == participant and p.get("has_video"):
                return list(p["video_paths"])
    return []


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

    # Map the global span into the participant's source video(s) (stitching across
    # a recording boundary for multi-video participants); single-video is a plain cut.
    # Release the reserved placeholder on any failure — a None return *or* an
    # exception — so we never leave a 0-byte file behind.
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

    default_desc = (
        "Transcript intake" if source == "transcript" else "Screenspace intake"
    )
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
        "category": "",
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
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        future_to_idx = {
            pool.submit(
                _process_intake_item,
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


def _load_stashes() -> list[dict[str, Any]]:
    return utils.load_json_manifest(config.STASHES_MANIFEST_FILENAME, default=[])


def _save_stashes(stashes: list[dict[str, Any]]) -> Path | None:
    return utils.save_json_manifest(config.STASHES_MANIFEST_FILENAME, stashes)


def _load_artifact_stashes() -> list[dict[str, Any]]:
    return utils.load_json_manifest(
        config.ARTIFACT_STASHES_MANIFEST_FILENAME, default=[]
    )


def _save_artifact_stashes(stashes: list[dict[str, Any]]) -> Path | None:
    return utils.save_json_manifest(config.ARTIFACT_STASHES_MANIFEST_FILENAME, stashes)


def _load_studio_settings() -> dict[str, Any]:
    """Load studio_settings.json and apply non-default values to config module."""
    data = utils.load_json_manifest(config.STUDIO_SETTINGS_FILENAME, default={})

    applied: dict[str, Any] = {}
    for name, value in data.items():
        if name not in config.STUDIO_SETTINGS:
            continue
        meta = config.STUDIO_SETTINGS[name]
        default = _settings_defaults.get(name)
        if meta.get("type") == "mark_categories":
            cleaned = _coerce_mark_categories(value)
            if cleaned is None:
                continue
            setattr(config, name, cleaned)
            applied[name] = cleaned
            continue
        if meta.get("type") == "card_picker":
            # Validate persisted selections too, so a stale studio_settings.json
            # (traversal, a deleted upload, or __none__ on a titlecard) doesn't
            # apply a value the PUT path would reject; leave config at its default.
            cleaned = _coerce_card_image(value, str(meta.get("kind", "title")))
            if cleaned is None:
                continue
            setattr(config, name, cleaned)
            applied[name] = cleaned
            continue
        if name in ("TITLECARD_COLOR", "ENDCARD_COLOR"):
            color = str(value)
            if not _HEX_COLOR_RE.match(color):
                continue  # ignore a tampered/invalid value, keep the default
            setattr(config, name, color)
            applied[name] = color
            continue
        expected_type = type(default) if default is not None else str
        try:
            if expected_type is bool:
                coerced = (
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
            continue
        setattr(config, name, coerced)
        applied[name] = coerced
    return applied


def _save_studio_settings(overrides: dict[str, Any]) -> Path | None:
    """Write only non-default settings to studio_settings.json."""
    to_save = {}
    for name, value in overrides.items():
        if name in _settings_defaults and value != _settings_defaults[name]:
            to_save[name] = value
    if not to_save:
        utils.remove_json_manifest(config.STUDIO_SETTINGS_FILENAME)
        return None
    return utils.save_json_manifest(config.STUDIO_SETTINGS_FILENAME, to_save)


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


def _stream_process_reel(
    clips: list[Any],
    cancel_flag: Any,
    *,
    titlecards_enabled: bool | None = None,
    titlecard_duration_seconds: int | None = None,
) -> Iterator[str]:
    """Run pipeline.process_reel on a worker thread and yield its progress events
    as NDJSON lines, finishing with a final result/error line.

    The worker thread also owns result persistence and busy-slot release, so
    that a client disconnect (e.g. browser navigation) does not abort the
    encode or orphan the reel from the manifest. The generator can die at any
    point; the worker keeps running until ffmpeg completes, then persists and
    frees the slot in its finally.
    """
    event_queue: queue.Queue[dict[str, Any]] = queue.Queue()
    sentinel: dict[str, Any] = {"__sentinel__": True}

    def emit_event(event: dict[str, Any]) -> None:
        _record_reel_event(event)
        event_queue.put(event)

    def worker() -> None:
        try:
            generated, reel_records = pipeline.process_reel(
                clips,
                cancel_flag=cancel_flag,
                progress_cb=emit_event,
                titlecards_enabled=titlecards_enabled,
                titlecard_duration_seconds=titlecard_duration_seconds,
            )
            if cancel_flag and cancel_flag():
                event_queue.put(
                    {
                        "ok": False,
                        "cancelled": True,
                        "error": "Reel generation cancelled",
                    }
                )
                return
            _extend_generated_reels(reel_records)
            _save_manifest_quiet()
            event_queue.put({"ok": True, "generated": generated, "reels": reel_records})
        except Exception as exc:
            event_queue.put({"ok": False, "error": str(exc)})
        finally:
            event_queue.put(sentinel)
            _release_busy("reel")

    try:
        threading.Thread(target=worker, daemon=True).start()
    except BaseException:
        # Worker never ran, so its finally won't release the slot.
        _release_busy("reel")
        raise

    while True:
        event = event_queue.get()
        if event is sentinel:
            return
        yield json.dumps(event) + "\n"


def _apply_time_overrides(clips: list[Any], overrides: dict[str, Any]) -> None:
    """Replace a clip's parsed timestamps with frontend-edited in/out points.

    ``overrides`` maps a ``"participant.row"`` cell key to the complete,
    segment-ordered list of ``[start_seconds, end_seconds]`` pairs currently
    shown for that cell in the Studio queue (the user dragged/typed new in/out
    points on the duration badge). Setting ``clip["times"]`` here makes
    ``files.prepare_clip()`` take its pre-parsed fast path and skip the cell
    re-parse, so the edited durations win over the spreadsheet values.
    """
    if not overrides:
        return
    for clip in clips:
        key = clip["participant"] + "." + str(clip["cell"].row)
        seg_times = overrides.get(key)
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
            # Force hours on both ends when either crosses the hour mark, so we
            # never emit a mixed M:SS / H:MM:SS pair (which breaks downstream
            # duration parsing — see AGENTS.md timestamp gotcha).
            needs_hours = start_sec >= 3600 or end_sec >= 3600
            new_times.append(
                (
                    utils.seconds_to_timestamp(
                        int(round(start_sec)), force_hours=needs_hours
                    ),
                    utils.seconds_to_timestamp(
                        int(round(end_sec)), force_hours=needs_hours
                    ),
                )
            )
        if new_times:
            clip["times"] = new_times


@studio_bp.route("/api/generate", methods=["POST"])
def api_generate() -> FlaskResponse:
    if _worksheet is None:
        return jsonify(
            {
                "ok": False,
                "error": "No spreadsheet loaded — pick one from the Start panel.",
            }
        ), 400

    data = request.get_json(silent=True) or {}
    cell_strings = data.get("cells", [])
    output_format = data.get("format", "clip")
    overrides: dict[str, Any] = data.get("overrides") or {}
    titlecards_enabled, titlecard_duration_seconds = _parse_titlecard_request(data)

    if not cell_strings:
        return jsonify({"ok": False, "error": "No cells specified"}), 400

    if output_format not in ("clip", "screen", "gif"):
        return jsonify({"ok": False, "error": f"Invalid format: {output_format}"}), 400

    if not _try_claim_busy("generate"):
        return jsonify(
            {"ok": False, "error": "A clip generation is already in progress"}
        ), 409

    try:
        cell_input = ", ".join(cell_strings)
        cell_specs = spreadsheet.parse_cell_specifications(cell_input)
        if not cell_specs:
            _release_busy("generate")
            return jsonify(
                {"ok": False, "error": "Could not parse cell specifications"}
            ), 400

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
        return jsonify({"ok": False, "error": str(e)}), 500

    def stream() -> Any:
        _generate_cancel_event.clear()
        _reset_generate_job_state(len(cell_strings))
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
            clip_cells.add(cell_str)

            existing = _find_existing_artifacts(
                clip["cell"].row,
                clip["cell"].col,
                output_format,
                existence_cache=existence_cache,
            )
            # A cached clip is only reusable when its recorded titlecard state
            # matches the request; mismatched ones are discarded and regenerated
            # so toggling Titlecards on/off (or changing the duration) takes
            # effect on the next Generate.
            # An overridden cell carries edited in/out points, but existing
            # artifacts are keyed only by cell row/col/format — they'd be reused
            # at the old duration. Treat them all as stale so the clip always
            # regenerates with the new times.
            cell_overridden = cell_str in overrides
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
                _increment_generate_done()
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

        # Pass 2: generate in parallel and yield as each completes. The
        # per-clip worker self-persists via _extend_generated_artifacts so
        # that results landing while the client is disconnecting (shutdown
        # waits for in-flight futures) still reach the manifest.
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
            # A future that happened to finish concurrently with the cancel
            # signal still produces (generated, artifacts); per the streaming
            # contract we must not append those to the manifest after cancel.
            # Drop the files on disk too so the user does not see orphan media.
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
            if workers >= 2 and len(to_generate) >= 2:
                with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
                    future_to_cell: dict[concurrent.futures.Future, tuple[Any, str]] = {
                        pool.submit(_generate_and_persist, clip): (clip, cell_str)
                        for clip, cell_str in to_generate
                    }
                    for future in concurrent.futures.as_completed(future_to_cell):
                        if cancel_flag():
                            for f in future_to_cell:
                                f.cancel()
                            break
                        clip, cell_str = future_to_cell[future]
                        _increment_generate_done()
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
                    _increment_generate_done()
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
            if cs not in clip_cells:
                _increment_generate_done()
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
            # Persist + purge the per-request endcard cache even when the
            # client disconnects mid-stream, so generated artifacts are not
            # left on disk without manifest records and endcard temp files
            # do not leak. Per-cell process_clips() calls run with
            # clear_titlecard_cache=False, so the cache is purged once here.
            titlecards.clear_endcard_cache()
            _save_manifest_quiet()
            _release_busy("generate")

    return Response(
        stream_with_busy_release(),
        mimetype="application/x-ndjson",
        headers={"X-Accel-Buffering": "no"},
    )


@studio_bp.route("/api/highlights-preview", methods=["POST"])
def api_highlights_preview() -> FlaskResponse:
    if _worksheet is None:
        return jsonify(
            {
                "ok": False,
                "error": "No spreadsheet loaded — pick one from the Start panel.",
            }
        ), 400

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
        return jsonify(
            {"ok": False, "error": "No clips found for highlights selection"}
        ), 400

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

    return jsonify({"ok": True, "clips": result})


@studio_bp.route("/api/reel", methods=["POST"])
def api_reel() -> FlaskResponse:
    if _worksheet is None:
        return jsonify(
            {
                "ok": False,
                "error": "No spreadsheet loaded — pick one from the Start panel.",
            }
        ), 400

    data = request.get_json(silent=True) or {}
    cell_strings = data.get("cells", [])
    highlights_duration = data.get("highlights_duration")
    reel_overrides: dict[str, Any] = data.get("overrides") or {}
    titlecards_enabled, titlecard_duration_seconds = _parse_titlecard_request(data)

    if not cell_strings:
        return jsonify({"ok": False, "error": "No cells specified"}), 400

    highlights_overrides: dict[str, Any] = {}
    if highlights_duration is not None:
        try:
            val = int(highlights_duration)
            if val > 0:
                highlights_overrides["HIGHLIGHTS_REEL_DURATION_SECONDS"] = val
        except (ValueError, TypeError):
            pass

    if not _try_claim_busy("reel"):
        return jsonify(
            {"ok": False, "error": "A reel build is already in progress"}
        ), 409

    def stream() -> Any:
        # _stream_process_reel's worker takes ownership of the busy slot once
        # we hand control to it; until then we release on every exit path so
        # the slot doesn't leak when the route returns without starting work.
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
        stream(),
        mimetype="application/x-ndjson",
        headers={"X-Accel-Buffering": "no"},
    )


@studio_bp.route("/api/viewer", methods=["POST"])
def api_viewer() -> FlaskResponse:
    with _generated_output_lock:
        artifacts = list(_generated_artifacts)
    if not artifacts:
        return jsonify(
            {
                "ok": False,
                "error": "No artifacts to build viewer from. Generate artifacts first.",
            }
        ), 400

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
            return jsonify({"ok": True, "file": str(viewer_path)})
        return jsonify({"ok": False, "error": "Failed to generate viewer"}), 500

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


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
        return jsonify(
            {
                "ok": False,
                "error": "No spreadsheet loaded — pick one from the Start panel.",
            }
        ), 400

    if not _try_claim_busy("timeline_viewer"):
        return jsonify(
            {"ok": False, "error": "A timeline viewer build is already in progress."}
        ), 409

    try:
        _timeline_viewer_cancel_event.clear()
        req = request.get_json(silent=True) or {}
        include_intake = req.get("include_intake", False)
        intake_items = req.get("intake_items", [])

        clips_list = spreadsheet.generate_list(
            _worksheet, "batch", ctx=_sheet_context, skip_prompts=True
        )
        if not clips_list:
            return jsonify({"ok": False, "error": "No clips found in sheet"}), 400

        generated, artifacts = pipeline.process_clips(
            clips_list,
            output_format="clip",
            cancel_flag=_timeline_viewer_cancel_event.is_set,
        )
        if _timeline_viewer_cancel_event.is_set():
            _discard_artifact_files(artifacts)
            return jsonify({"ok": False, "cancelled": True})
        if not artifacts:
            return jsonify({"ok": False, "error": "No artifacts were generated"}), 400

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
            output_format="clip",
            screenspace_events=ss_events or None,
        )
        viewer_path = viewer.generate_timeline_viewer(
            viewer_data,
            template_name="timeline-viewer.html",
            output_basename="timeline_viewer.html",
        )
        if viewer_path:
            _save_manifest_quiet()
            return jsonify(
                {
                    "ok": True,
                    "file": str(viewer_path),
                    "generated": generated + len(intake_artifacts),
                }
            )
        return jsonify(
            {"ok": False, "error": "Failed to generate timeline viewer"}
        ), 500

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        _release_busy("timeline_viewer")


@studio_bp.route("/api/gallery", methods=["POST"])
def api_gallery() -> FlaskResponse:
    if _sheet_context is None:
        return jsonify(
            {
                "ok": False,
                "error": "No spreadsheet loaded — pick one from the Start panel.",
            }
        ), 400

    data = request.get_json(silent=True) or {}
    participant = data.get("participant", "")
    output_format = data.get("format", "screen")
    interval = data.get("interval", config.GALLERY_INTERVAL_SECONDS)
    bundle = bool(data.get("bundle", config.GALLERY_BUNDLE_ENABLED))

    if not participant:
        return jsonify({"ok": False, "error": "No participant specified"}), 400

    if output_format not in ("screen", "gif"):
        return jsonify({"ok": False, "error": f"Invalid format: {output_format}"}), 400

    if not _try_claim_busy("gallery"):
        return jsonify(
            {"ok": False, "error": "A gallery build is already in progress."}
        ), 409

    try:
        try:
            interval = int(interval)
            if interval < 1:
                interval = config.GALLERY_INTERVAL_SECONDS
        except (ValueError, TypeError):
            interval = config.GALLERY_INTERVAL_SECONDS

        sources = _resolve_participant_sources(participant)
        if not sources or not sources[0].is_file():
            return jsonify(
                {"ok": False, "error": f"Source video not found for {participant}"}
            ), 404

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
            return jsonify({"ok": False, "error": "No captures generated"}), 500

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
            return jsonify({"ok": True, "file": str(gallery_path)})
        return jsonify({"ok": False, "error": "Failed to generate gallery viewer"}), 500

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        _release_busy("gallery")


@studio_bp.route("/api/open-viewer", methods=["POST"])
def api_open_viewer() -> FlaskResponse:
    """Open a generated viewer HTML file in the default browser."""
    data = request.get_json(silent=True) or {}
    file_path = data.get("file", "")
    if not file_path:
        return jsonify({"ok": False, "error": "No file specified"}), 400

    p = Path(file_path).resolve()
    output_dir = Path(utils.get_effective_output_dir()).resolve()

    if p.suffix != ".html" or not p.is_relative_to(output_dir):
        return jsonify({"ok": False, "error": "Invalid file path"}), 403

    if not p.is_file():
        return jsonify({"ok": False, "error": "File not found"}), 404

    webbrowser.open(p.as_uri())
    return jsonify({"ok": True})


@studio_bp.route("/api/manifest", methods=["GET", "POST"])
def api_manifest() -> FlaskResponse:
    if request.method == "GET":
        artifacts, reels = viewer._load_manifest_both()
        return jsonify({"ok": True, "artifacts": artifacts, "reels": reels})

    # Snapshot the shared lists so a worker thread extending mid-export
    # can't produce a partial/aliased manifest snapshot.
    with _generated_output_lock:
        artifacts = list(_generated_artifacts)
        reels = list(_generated_reels)
    if not artifacts and not reels:
        return jsonify(
            {
                "ok": False,
                "error": "No artifacts to export. Generate artifacts first.",
            }
        ), 400

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
            return jsonify({"ok": True, "file": str(manifest_path)})
        return jsonify({"ok": False, "error": "Failed to write manifest"}), 500

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@studio_bp.route("/api/regenerate", methods=["POST"])
def api_regenerate() -> FlaskResponse:
    try:
        artifacts, reels = viewer._load_manifest_both()
        if not artifacts and not reels:
            return jsonify(
                {
                    "ok": False,
                    "error": "No manifest found on disk. Export a manifest first.",
                }
            ), 400

        media_count = sum(1 for a in artifacts if a.get("type") != "transcript")
        reel_count = len(reels)
        total = media_count + reel_count

        regenerated = pipeline.regenerate_from_manifest(artifacts, reels=reels)
        return jsonify({"ok": True, "regenerated": regenerated, "total": total})

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


def _handle_stash_crud(load_fn: Any, save_fn: Any, id_prefix: str) -> FlaskResponse:
    """Shared create/update/delete logic for stash endpoints."""
    import uuid
    from datetime import datetime, timezone

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
                return jsonify({"ok": False, "error": "No items to stash"}), 400
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
                "createdAt": datetime.now(timezone.utc).isoformat(),
            }
            stashes.append(stash)
            save_fn(stashes)
            return jsonify({"ok": True, "stash": stash})

        if action == "update":
            stash_id = data.get("id")
            if not stash_id:
                return jsonify({"ok": False, "error": "No stash ID"}), 400
            for s in stashes:
                if s["id"] == stash_id:
                    name = data.get("name")
                    if name is not None:
                        s["name"] = name
                    save_fn(stashes)
                    return jsonify({"ok": True, "stash": s})
            return jsonify({"ok": False, "error": "Stash not found"}), 404

        if action == "delete":
            stash_id = data.get("id")
            if not stash_id:
                return jsonify({"ok": False, "error": "No stash ID"}), 400
            for i, s in enumerate(stashes):
                if s["id"] == stash_id:
                    stashes.pop(i)
                    save_fn(stashes)
                    return jsonify({"ok": True})
            return jsonify({"ok": False, "error": "Stash not found"}), 404

        return jsonify({"ok": False, "error": f"Unknown action: {action}"}), 400


@studio_bp.route("/api/stashes", methods=["GET"])
def api_stashes_get() -> FlaskResponse:
    return jsonify({"ok": True, "stashes": _load_stashes()})


@studio_bp.route("/api/stashes", methods=["POST"])
def api_stashes_post() -> FlaskResponse:
    return _handle_stash_crud(_load_stashes, _save_stashes, "stash")


@studio_bp.route("/api/artifact-stashes", methods=["GET"])
def api_artifact_stashes_get() -> FlaskResponse:
    return jsonify({"ok": True, "stashes": _load_artifact_stashes()})


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


def _validate_prompt(text: str, placeholders: list[str]) -> str | None:
    """Validate a user-edited thinking-agent prompt.

    Returns an error string, or ``None`` if the prompt is safe to save. Prompts
    that declare *placeholders* are ``.format()``-ed at runtime, so every
    required placeholder must be present and the prompt must format cleanly with
    exactly those keys. Prompts with no placeholders (the ``*_SYSTEM`` strings)
    are sent to the model verbatim and accept any text.
    """
    if not placeholders:
        return None
    allowed = set(placeholders)
    used: set[str] = set()
    try:
        for _literal, field_name, _spec, _conv in string.Formatter().parse(text):
            if field_name:
                # Strip any attribute/index access, e.g. "{a.b}" / "{a[0]}".
                used.add(re.split(r"[.\[]", field_name, maxsplit=1)[0])
    except (ValueError, IndexError):
        return "unbalanced { } braces — escape literal braces as {{ and }}"
    missing = allowed - used
    if missing:
        return "missing required placeholder(s): " + ", ".join(
            "{" + p + "}" for p in sorted(missing)
        )
    # Ground truth: the prompt must .format() cleanly with exactly these keys.
    # Catches unknown placeholders ({foo}), stray positional {}, and nested
    # format-spec references the parse() scan above does not surface.
    try:
        text.format(**{p: "" for p in placeholders})
    except (KeyError, IndexError, ValueError) as exc:
        bad = exc.args[0] if exc.args else exc
        return f"references an unknown placeholder ({bad}); allowed: " + ", ".join(
            "{" + p + "}" for p in placeholders
        )
    return None


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
        merged = {name: getattr(config, name) for name in config.STUDIO_SETTINGS.keys()}
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

    applied = {}
    for name, value in settings_data.items():
        if name not in config.STUDIO_SETTINGS:
            continue
        meta = config.STUDIO_SETTINGS[name]
        default = _settings_defaults.get(name)
        if meta.get("type") == "mark_categories":
            cleaned = _coerce_mark_categories(value)
            if cleaned is None:
                return {}, f"Invalid {name} payload"
            setattr(config, name, cleaned)
            applied[name] = cleaned
            continue
        if meta.get("type") == "card_picker":
            cleaned = _coerce_card_image(value, str(meta.get("kind", "title")))
            if cleaned is None:
                return {}, f"Invalid {name} payload"
            setattr(config, name, cleaned)
            applied[name] = cleaned
            continue
        if meta.get("type") == "prompt":
            text = str(value)
            err = _validate_prompt(text, meta.get("placeholders") or [])
            if err is not None:
                return {}, f"Invalid {name}: {err}"
            setattr(config, name, text)
            applied[name] = text
            continue
        if name in ("TITLECARD_COLOR", "ENDCARD_COLOR"):
            color = str(value)
            if not _HEX_COLOR_RE.match(color):
                return {}, f"Invalid {name}: expected a #rrggbb hex color"
            setattr(config, name, color)
            applied[name] = color
            continue
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
    return jsonify({"ok": True, "settings": _settings_records()})


@studio_bp.route("/api/settings", methods=["PUT"])
def api_settings_put() -> FlaskResponse:
    data = request.get_json(silent=True) or {}
    applied, error = _apply_settings_payload(data)
    if error is not None:
        return jsonify({"ok": False, "error": error}), 400
    return jsonify({"ok": True, "applied": applied})


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
    return jsonify(
        {
            "ok": True,
            "title": _card_picker_payload("title"),
            "end": _card_picker_payload("end"),
        }
    )


@studio_bp.route("/api/titlecards/default/<kind>", methods=["GET"])
def api_titlecard_default(kind: str) -> FlaskResponse:
    """Serve the bundled default titlecard/endcard image for previews."""
    if kind not in ("title", "end"):
        return jsonify({"ok": False, "error": "Invalid kind"}), 400
    asset = "titlecard.png" if kind == "title" else "endcard.png"
    path = utils.get_bundled_assets_root() / "assets" / asset
    if not path.is_file():
        return jsonify({"ok": False, "error": "No default image"}), 404
    return send_file(str(path))


@studio_bp.route("/api/titlecards/image/<path:name>", methods=["GET"])
def api_titlecard_image(name: str) -> FlaskResponse:
    """Serve an uploaded card background by filename (used for previews)."""
    safe = Path(name).name
    if safe != name or Path(safe).suffix.lower() not in _ALLOWED_CARD_EXTS:
        return jsonify({"ok": False, "error": "Invalid filename"}), 400
    if not (_titlecard_images_dir() / safe).is_file():
        return jsonify({"ok": False, "error": "Not found"}), 404
    return send_from_directory(str(_titlecard_images_dir()), safe)


@studio_bp.route("/api/titlecards/upload", methods=["POST"])
def api_titlecard_upload() -> FlaskResponse:
    """Accept a card background image upload (PNG/JPG/WebP) into the upload pool."""
    file = request.files.get("file")
    filename = file.filename if file is not None else None
    if file is None or not filename:
        return jsonify({"ok": False, "error": "No file provided"}), 400
    ext = Path(filename).suffix.lower()
    if ext not in _ALLOWED_CARD_EXTS:
        return jsonify(
            {
                "ok": False,
                "error": f"Unsupported file type '{ext}'. Use PNG, JPG, or WebP.",
            }
        ), 400
    file.seek(0, os.SEEK_END)
    size = file.tell()
    file.seek(0)
    if size > _MAX_CARD_UPLOAD_BYTES:
        return jsonify({"ok": False, "error": "File too large (max 10 MB)."}), 400

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
        return jsonify({"ok": False, "error": str(error)}), 500
    return jsonify(
        {
            "ok": True,
            "item": {
                "id": candidate,
                "label": candidate,
                "kind": "upload",
                "url": f"/api/titlecards/image/{candidate}",
                "deletable": True,
            },
        }
    )


@studio_bp.route("/api/titlecards/image/<path:name>", methods=["DELETE"])
def api_titlecard_delete(name: str) -> FlaskResponse:
    """Delete an uploaded card background; reset any selection that used it."""
    safe = Path(name).name
    if safe != name:
        return jsonify({"ok": False, "error": "Invalid filename"}), 400
    target = _titlecard_images_dir() / safe
    if not target.is_file():
        return jsonify({"ok": False, "error": "Not found"}), 404
    try:
        target.unlink()
    except OSError as error:
        return jsonify({"ok": False, "error": str(error)}), 500
    reset: dict[str, str] = {}
    for setting in ("TITLECARD_IMAGE", "ENDCARD_IMAGE"):
        if getattr(config, setting, "") == safe:
            setattr(config, setting, "")
            reset[setting] = ""
    if reset:
        merged = {n: getattr(config, n) for n in config.STUDIO_SETTINGS.keys()}
        _save_studio_settings(merged)
    return jsonify({"ok": True, "reset": reset})


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
        return jsonify({"ok": False, "error": "No intake items specified"}), 400

    output_format = data.get("format", "clip")
    study = _sheet_context.study_name if _sheet_context else ""

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
            if workers >= 2 and len(items) >= 2:
                with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
                    future_to_idx = {
                        pool.submit(
                            _process_intake_item,
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
        stream(),
        mimetype="application/x-ndjson",
        headers={"X-Accel-Buffering": "no"},
    )


@studio_bp.route("/api/reel-direct", methods=["POST"])
def api_reel_direct() -> FlaskResponse:
    """Build a reel from direct timestamp segments (for intake / mixed queues)."""
    import tempfile

    data = request.get_json(silent=True) or {}
    segments = data.get("segments", [])
    if not segments:
        return jsonify({"ok": False, "error": "No segments specified"}), 400

    if not _try_claim_busy("reel"):
        return jsonify(
            {"ok": False, "error": "A reel build is already in progress"}
        ), 409

    _reel_cancel_event.clear()
    _reset_reel_job_state("reel-direct")

    titlecards_enabled, titlecard_duration_seconds = _parse_titlecard_request(data)
    cards_enabled, card_duration = pipeline._resolve_titlecard_options(
        titlecards_enabled, titlecard_duration_seconds
    )

    def stream() -> Iterator[str]:
        # The worker thread does the actual ffmpeg work and owns the busy
        # slot. The generator just drains its event queue so that a client
        # disconnect (e.g. browser navigation to a sibling frontend) does not
        # abort the encode or orphan the reel from the manifest.
        event_queue: queue.Queue[dict[str, Any]] = queue.Queue()
        sentinel: dict[str, Any] = {"__sentinel__": True}

        def emit_event(event: dict[str, Any]) -> None:
            _record_reel_event(event)
            event_queue.put(event)

        def worker() -> None:
            output_dir = Path(utils.get_effective_output_dir())
            clip_paths: list[str] = []
            temp_clips: list[str] = []
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

            try:
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
                    # not in temp_clips, so the finally block won't unlink it.
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
                        ok = titlecards.wrap_clip_with_cards(
                            wrap_clip,
                            tmp_path,
                            cancel_flag=_reel_cancel_event.is_set,
                            titlecards_enabled=cards_enabled,
                            titlecard_duration_seconds=card_duration,
                        )
                    if ok:
                        clip_paths.append(tmp_path)
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

                reel_study = _sheet_context.study_name if _sheet_context else ""
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
                    direct_title_img, direct_end_img = (
                        pipeline._resolve_titlecard_images(cards_enabled)
                    )
                    reel_record: dict[str, Any] = {
                        "id": f"reel_intake_{hashlib.md5(reel_name.encode()).hexdigest()[:8]}",
                        "file": Path(reel_name).name,
                        "source": "intake",
                        "description": f"Intake reel ({len(clip_paths)} segments)",
                        "titlecards": cards_enabled,
                        "titlecardDuration": card_duration if cards_enabled else 0,
                        "titlecardImage": direct_title_img,
                        "endcardImage": direct_end_img,
                    }
                    _append_generated_reel(reel_record)
                    _save_manifest_quiet()
                    emit_event({"ok": True, "generated": 1, "reels": [reel_record]})
                else:
                    files.release_reservation(reel_name)
                    emit_event({"ok": False, "error": "Reel concatenation failed"})
            except Exception as exc:
                emit_event({"ok": False, "error": str(exc)})
            finally:
                # Endcard temp files are cached per-process across all wrap
                # calls; purge them here so per-request cards do not leak
                # between consecutive reel builds.
                titlecards.clear_endcard_cache()
                for tmp in temp_clips:
                    try:
                        Path(tmp).unlink(missing_ok=True)
                    except OSError:
                        pass
                event_queue.put(sentinel)
                _release_busy("reel")

        try:
            threading.Thread(target=worker, daemon=True).start()
        except BaseException:
            # Worker never ran, so its finally won't release the slot.
            _release_busy("reel")
            raise

        while True:
            event = event_queue.get()
            if event is sentinel:
                return
            yield json.dumps(event) + "\n"

    return Response(
        stream(),
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
        reel_busy = _reel_in_progress
        generate_busy = _generate_in_progress
        intake_busy = _intake_active > 0
    with _job_state_lock:
        reel_snapshot = dict(_reel_job_state)
        generate_snapshot = dict(_generate_job_state)
        intake_snapshot = dict(_intake_job_state)
    return jsonify(
        {
            "ok": True,
            "reel": {
                "in_progress": reel_busy,
                "cancelling": reel_busy and _reel_cancel_event.is_set(),
                **reel_snapshot,
            },
            "generate": {
                "in_progress": generate_busy,
                "cancelling": generate_busy and _generate_cancel_event.is_set(),
                **generate_snapshot,
            },
            "intake": {
                "in_progress": intake_busy,
                "cancelling": intake_busy and _intake_cancel_event.is_set(),
                **intake_snapshot,
            },
        }
    )


@studio_bp.route("/api/reel/cancel", methods=["POST"])
def api_reel_cancel() -> FlaskResponse:
    """Signal cancellation for the in-progress reel build."""
    _reel_cancel_event.set()
    return jsonify({"ok": True})


@studio_bp.route("/api/generate/cancel", methods=["POST"])
def api_generate_cancel() -> FlaskResponse:
    """Signal cancellation for the in-progress clip generation."""
    _generate_cancel_event.set()
    return jsonify({"ok": True})


@studio_bp.route("/api/generate-intake/cancel", methods=["POST"])
def api_generate_intake_cancel() -> FlaskResponse:
    """Signal cancellation for the in-progress intake generation.

    Sheet and intake branches run in parallel from a single Studio Generate
    click and have independent streams, so the Cancel button must hit both
    endpoints to stop the full set of in-flight ffmpeg subprocesses.
    """
    _intake_cancel_event.set()
    return jsonify({"ok": True})


@studio_bp.route("/api/timeline-viewer/cancel", methods=["POST"])
def api_timeline_viewer_cancel() -> FlaskResponse:
    """Signal cancellation for the in-progress timeline-viewer build."""
    _timeline_viewer_cancel_event.set()
    return jsonify({"ok": True})


@studio_bp.route("/api/gallery/cancel", methods=["POST"])
def api_gallery_cancel() -> FlaskResponse:
    """Signal cancellation for the in-progress gallery build."""
    _gallery_cancel_event.set()
    return jsonify({"ok": True})


# ---- State initialization ----


def _init_studio_state(worksheet: Any) -> None:
    """Initialize module-level state for Studio routes.

    *worksheet* may be ``None`` — Studio's blueprint still registers and serves
    the HTML, but spreadsheet-dependent routes report ``sheet_loaded: false``
    until a sheet is opened via ``POST /api/spreadsheets/open``.
    """
    global \
        _worksheet, \
        _sheet_context, \
        _generated_artifacts, \
        _generated_reels, \
        _thumbnail_cache, \
        _sprite_cache, \
        _audio_cache

    _load_studio_settings()
    _worksheet = worksheet
    if worksheet is not None:
        _sheet_context = spreadsheet.build_sheet_context(worksheet)
        if _sheet_context is None:
            utils.error_print("Could not load spreadsheet data for Studio.")
            sys.exit(1)
    else:
        _sheet_context = None
    # Rebind the shared generated lists under their lock so a streaming
    # generate/intake append can't run against a half-swapped reference.
    with _generated_output_lock:
        _generated_artifacts, _generated_reels = viewer._load_manifest_both()
        _rebuild_artifact_index()
    with _thumbnail_cache_lock:
        _thumbnail_cache = OrderedDict()
    with _sprite_cache_lock:
        _sprite_cache = OrderedDict()
    with _audio_cache_lock:
        _audio_cache = OrderedDict()


# ---- Entry point ----


def _resolve_participants() -> list[str] | None:
    """Extract participant IDs from the loaded sheet context."""
    if _sheet_context is None:
        return None
    return spreadsheet.get_participant_list(
        _sheet_context.header_row,
        _sheet_context.id_cell,
        _sheet_context.num_participants,
    )


def _derive_sheet_meta(worksheet: Any) -> dict[str, str] | None:
    """Return ``{type, id_or_path, label}`` for a CLI-loaded worksheet.

    Used so the Start overlay's spreadsheet picker can show the currently
    loaded sheet when the overlay is opened on a session that was launched
    from the CLI (not via the runtime ``/api/spreadsheets/open`` endpoint).
    """
    if worksheet is None:
        return None
    try:
        import excel_io

        if isinstance(worksheet, excel_io.ExcelSheetAdapter):
            path = getattr(worksheet, "_workbook_path", "") or ""
            if not path:
                return None
            return {
                "type": "excel",
                "id_or_path": path,
                "label": Path(path).name,
            }
    except Exception:
        pass
    # gspread Worksheet (or anything quacking like one): use the parent
    # spreadsheet title as both the identifier and the label.
    parent = getattr(worksheet, "spreadsheet", None)
    title = getattr(parent, "title", "") if parent is not None else ""
    if not title:
        return None
    return {"type": "google", "id_or_path": title, "label": title}


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
    """Replace the active worksheet and refresh all three blueprints' state.

    Atomic: if any of the three init steps fails, the prior state is restored
    so Studio / Screenspace / Transcripts don't end up pointing at different
    sheets. Used by ``/api/spreadsheets/open`` and ``/api/spreadsheets/close``.
    """
    import screenspace_server
    import transcripts_server

    global _worksheet, _sheet_context, _generated_artifacts, _generated_reels
    prev_worksheet = _worksheet
    prev_sheet_context = _sheet_context
    prev_artifacts = _generated_artifacts
    prev_reels = _generated_reels
    try:
        _init_studio_state(new_worksheet)
        screenspace_server._init_screenspace_state(
            sheet_context=_sheet_context,
            participant_list=_resolve_participants(),
        )
        transcripts_server._init_transcripts_state(
            sheet_context=_sheet_context,
            participant_list=_resolve_participants(),
        )
    except Exception:
        _worksheet = prev_worksheet
        _sheet_context = prev_sheet_context
        with _generated_output_lock:
            _generated_artifacts = prev_artifacts
            _generated_reels = prev_reels
            _rebuild_artifact_index()
        # Best-effort: re-pin the two sister blueprints to the restored state.
        # If these themselves throw, swallow — the studio state is already
        # consistent and the original exception is what we want to surface.
        try:
            screenspace_server._init_screenspace_state(
                sheet_context=_sheet_context,
                participant_list=_resolve_participants(),
            )
        except Exception:
            pass
        try:
            transcripts_server._init_transcripts_state(
                sheet_context=_sheet_context,
                participant_list=_resolve_participants(),
            )
        except Exception:
            pass
        raise


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
    import screenspace_server
    import start_settings
    import transcripts_server
    import workflows_server

    combined = Flask(__name__, static_folder=None)
    # Preserve insertion order in JSON responses. Flask defaults to sorting object
    # keys alphabetically, which would clobber manifest-ordered data such as the
    # region list (drag-to-reorder relies on GET /api/regions echoing manifest order).
    assert isinstance(combined.json, DefaultJSONProvider)  # Flask's stock provider
    combined.json.sort_keys = False

    _init_studio_state(worksheet)
    # Seed the meta + recent-projects rail for CLI launches that already
    # have a worksheet — the runtime /api/spreadsheets/open path handles
    # both of these itself.
    global _active_sheet_meta
    _active_sheet_meta = _derive_sheet_meta(worksheet)
    if gspread_client is not None:
        _google_auth.client = gspread_client
    if _active_sheet_meta is not None:
        start_settings.record_project_session(
            str(utils.get_effective_input_dir()),
            str(utils.get_effective_output_dir()),
            _active_sheet_meta,
        )
    combined.register_blueprint(studio_bp, url_prefix="/studio")

    screenspace_server._init_screenspace_state(
        sheet_context=_sheet_context,
        participant_list=_resolve_participants(),
    )
    combined.register_blueprint(
        screenspace_server.screenspace_bp, url_prefix="/screenspace"
    )

    transcripts_server._init_transcripts_state(
        sheet_context=_sheet_context,
        participant_list=_resolve_participants(),
    )
    combined.register_blueprint(
        transcripts_server.transcripts_bp, url_prefix="/transcripts"
    )

    workflows_server._init_workflows_state(
        sheet_context=_sheet_context,
        participant_list=_resolve_participants(),
        worksheet=_worksheet,
    )
    combined.register_blueprint(workflows_server.workflows_bp, url_prefix="/workflows")

    @combined.after_request
    def _set_cache_headers(response):
        # Skip if a route already set Cache-Control (e.g. thumbnails)
        if "Cache-Control" in response.headers:
            return response
        ct = response.content_type or ""
        if ct.startswith("text/html"):
            response.headers["Cache-Control"] = "no-cache"
        elif ct.startswith(("text/css", "application/javascript", "text/javascript")):
            response.headers["Cache-Control"] = "public, max-age=3600"
        elif ct.startswith("image/svg"):
            response.headers["Cache-Control"] = "public, max-age=86400"
        return response

    @combined.route("/")
    def root():
        return redirect(f"/{default_page}/")

    @combined.route("/api/status")
    def status() -> Response:
        meta = _active_sheet_meta if _worksheet is not None else None
        return jsonify(
            {
                "studio": True,
                "screenspace": True,
                "transcripts": True,
                "workflows": True,
                "sheet_loaded": _worksheet is not None,
                "spreadsheet_label": _spreadsheet_label(),
                "spreadsheet_type": (meta or {}).get("type", ""),
                "spreadsheet_id_or_path": (meta or {}).get("id_or_path", ""),
                "input_dir": str(utils.get_effective_input_dir()),
                "output_dir": str(utils.get_effective_output_dir()),
                "videos_in_input": len(utils.discover_participant_videos()),
                "version": utils.get_version(),
                "author": "Henrik Edlund",
                "license": "MIT",
                "repo_url": "https://github.com/henedl/clipgen",
            }
        )

    @combined.route("/api/export/status")
    def api_export_status() -> Response:
        """Report which surface manifests are present in the output directory.

        Used by the frontend to gate the Export quick action — if no
        manifests exist there is nothing for ``write_export_bundle`` to write.
        """
        output_dir = Path(utils.get_effective_output_dir())
        screenspace = (output_dir / config.SCREENSPACE_MANIFEST_FILENAME).is_file()
        transcripts = (output_dir / config.TRANSCRIPTS_MANIFEST_FILENAME).is_file()
        return jsonify(
            {
                "ok": True,
                "screenspace": screenspace,
                "transcripts": transcripts,
                "any": screenspace or transcripts,
            }
        )

    @combined.route("/api/export", methods=["POST"])
    def api_export() -> FlaskResponse:
        """Write the same JSON+CSV bundle the ``--export`` CLI flag produces."""
        import data_export

        output_dir = Path(utils.get_effective_output_dir())
        try:
            written = data_export.write_export_bundle(output_dir)
        except Exception as err:
            return jsonify({"ok": False, "error": str(err)}), 500
        if not written:
            return jsonify(
                {
                    "ok": False,
                    "error": (
                        "No manifests in output directory. Expected one of "
                        f"{config.SCREENSPACE_MANIFEST_FILENAME} or "
                        f"{config.TRANSCRIPTS_MANIFEST_FILENAME}."
                    ),
                }
            ), 404
        return jsonify(
            {
                "ok": True,
                "written": [p.name for p in written],
                "output_dir": str(output_dir),
            }
        )

    # ---- Start overlay: directories, spreadsheet picker, persistence ----

    @combined.route("/api/dirs", methods=["GET"])
    def api_dirs_get() -> Response:
        import start_settings

        s = start_settings.load_start_settings()
        return jsonify(
            {
                "ok": True,
                "input": str(utils.get_effective_input_dir()),
                "output": str(utils.get_effective_output_dir()),
                "recent_inputs": s.get("recent_inputs", []),
                "recent_outputs": s.get("recent_outputs", []),
            }
        )

    @combined.route("/api/dirs", methods=["POST"])
    def api_dirs_post() -> FlaskResponse:
        import start_settings

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
        files_list: list[dict[str, str]] = []
        if input_dir.is_dir():
            for p in sorted(input_dir.glob("*.xlsx")):
                if p.name.startswith("~$"):
                    continue
                files_list.append({"path": str(p), "name": p.name})
        return jsonify({"ok": True, "input_dir": str(input_dir), "files": files_list})

    @combined.route("/api/spreadsheets/google", methods=["GET"])
    def api_spreadsheets_google() -> Response:
        if _google_auth.client is None:
            return jsonify(
                {
                    "ok": True,
                    "authenticated": False,
                    "auth_in_flight": _google_auth.in_flight,
                    "auth_error": _google_auth.error,
                    "sheets": [],
                }
            )
        import google_api

        try:
            names = google_api.get_all_spreadsheets(_google_auth.client)
        except Exception as exc:
            return jsonify(
                {
                    "ok": True,
                    "authenticated": True,
                    "auth_in_flight": False,
                    "auth_error": str(exc),
                    "sheets": [],
                }
            )
        return jsonify(
            {
                "ok": True,
                "authenticated": True,
                "auth_in_flight": False,
                "auth_error": "",
                "sheets": [{"name": n, "id": n} for n in names],
            }
        )

    @combined.route("/api/spreadsheets/google/auth", methods=["POST"])
    def api_spreadsheets_google_auth() -> FlaskResponse:
        with _google_auth.lock:
            if _google_auth.in_flight:
                return jsonify({"ok": True, "started": False, "in_flight": True})
            if _google_auth.client is not None:
                return jsonify({"ok": True, "started": False, "authenticated": True})
            _google_auth.in_flight = True
            _google_auth.error = ""

        def _run_auth() -> None:
            try:
                import cli as _cli

                client = _cli.authenticate_google()
                if client is None:
                    _google_auth.error = (
                        "Google authentication failed — check credentials.json."
                    )
                else:
                    _google_auth.client = client
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
        import start_settings

        data = request.get_json(silent=True) or {}
        type_ = data.get("type", "")
        id_or_path = (data.get("id_or_path") or "").strip()
        if type_ not in ("google", "excel") or not id_or_path:
            return jsonify(
                {
                    "ok": False,
                    "error": "Required: type ('google'|'excel') and id_or_path",
                }
            ), 400

        if _generation_busy():
            return jsonify(
                {
                    "ok": False,
                    "error": (
                        "Generation is in progress — wait for it to finish "
                        "before switching spreadsheets."
                    ),
                }
            ), 409

        new_ws: Any = None
        label = ""
        try:
            if type_ == "excel":
                import excel_io

                new_ws = excel_io.open_excel_workbook(id_or_path)
                label = Path(id_or_path).name
            else:
                if _google_auth.client is None:
                    return jsonify(
                        {
                            "ok": False,
                            "error": (
                                "Not authenticated with Google — "
                                "click 'Connect Google' first."
                            ),
                        }
                    ), 400
                import clipgen as _clipgen
                import google_api

                if id_or_path.startswith("http://") or id_or_path.startswith(
                    "https://"
                ):
                    new_ws = _clipgen.open_spreadsheet_by_url(
                        _google_auth.client, id_or_path
                    )
                else:
                    doc_list = google_api.get_all_spreadsheets(_google_auth.client)
                    new_ws = _clipgen.open_spreadsheet_by_name(
                        _google_auth.client, doc_list, id_or_path
                    )
                if new_ws is not None:
                    parent = getattr(new_ws, "spreadsheet", None)
                    label = getattr(parent, "title", "") or id_or_path
        except Exception as exc:
            return jsonify({"ok": False, "error": str(exc)}), 500

        if new_ws is None:
            return jsonify({"ok": False, "error": "Could not open spreadsheet"}), 404

        _swap_worksheet(new_ws)
        if _sheet_context is None:
            return jsonify(
                {"ok": False, "error": "Could not parse the spreadsheet"}
            ), 500

        start_settings.record_recent_spreadsheet(type_, id_or_path, label)
        start_settings.record_project_session(
            str(utils.get_effective_input_dir()),
            str(utils.get_effective_output_dir()),
            {"type": type_, "id_or_path": id_or_path, "label": label},
        )
        global _active_sheet_meta
        _active_sheet_meta = {
            "type": type_,
            "id_or_path": id_or_path,
            "label": label,
        }
        return jsonify(
            {
                "ok": True,
                "sheet_loaded": True,
                "spreadsheet_label": _spreadsheet_label(),
            }
        )

    @combined.route("/api/spreadsheets/close", methods=["POST"])
    def api_spreadsheets_close() -> FlaskResponse:
        global _active_sheet_meta
        if _generation_busy():
            return jsonify(
                {
                    "ok": False,
                    "error": "Generation is in progress — wait for it to finish.",
                }
            ), 409
        _swap_worksheet(None)
        _active_sheet_meta = None
        return jsonify({"ok": True, "sheet_loaded": False})

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
        return jsonify({"ok": True, "path": path})

    @combined.route("/api/sessions/record", methods=["POST"])
    def api_sessions_record() -> FlaskResponse:
        """Record an "Open workspace" session — used by the no-spreadsheet path.

        The Google/Excel paths already record via api_spreadsheets_open after a
        successful sheet open; this endpoint covers the case where the user
        clicks "Open workspace" on the "No spreadsheet" tab.
        """
        import start_settings

        data = request.get_json(silent=True) or {}
        input_raw = data.get("input")
        output_raw = data.get("output")
        if input_raw is not None and not isinstance(input_raw, str):
            return jsonify({"ok": False, "error": "input must be a string"}), 400
        if output_raw is not None and not isinstance(output_raw, str):
            return jsonify({"ok": False, "error": "output must be a string"}), 400
        input_dir = (input_raw or "").strip()
        output_dir = (output_raw or "").strip()

        spreadsheet_payload = data.get("spreadsheet")
        spreadsheet_dict: dict[str, Any] | None = None
        if spreadsheet_payload is not None:
            if not isinstance(spreadsheet_payload, dict):
                return jsonify(
                    {"ok": False, "error": "spreadsheet must be an object or null"}
                ), 400
            ss_type = (spreadsheet_payload.get("type") or "").strip()
            ss_id = (spreadsheet_payload.get("id_or_path") or "").strip()
            ss_label = (spreadsheet_payload.get("label") or "").strip()
            if ss_type not in ("google", "excel"):
                return jsonify(
                    {
                        "ok": False,
                        "error": "spreadsheet.type must be 'google' or 'excel'",
                    }
                ), 400
            if not ss_id:
                return jsonify(
                    {"ok": False, "error": "spreadsheet.id_or_path is required"}
                ), 400
            spreadsheet_dict = {
                "type": ss_type,
                "id_or_path": ss_id,
                "label": ss_label or ss_id,
            }
        start_settings.record_project_session(
            input_dir or str(utils.get_effective_input_dir()),
            output_dir or str(utils.get_effective_output_dir()),
            spreadsheet_dict,
        )
        return jsonify({"ok": True})

    @combined.route("/api/changelog")
    def api_changelog() -> Response:
        import changelog

        return jsonify({"ok": True, "entries": changelog.load_entries()})

    @combined.route("/api/start-settings", methods=["GET"])
    def api_start_settings_get() -> Response:
        import start_settings

        return jsonify({"ok": True, "settings": start_settings.load_start_settings()})

    @combined.route("/api/start-settings", methods=["POST"])
    def api_start_settings_post() -> FlaskResponse:
        import start_settings

        data = request.get_json(silent=True) or {}
        if "persist_enabled" in data:
            start_settings.set_persist_enabled(bool(data["persist_enabled"]))
        return jsonify({"ok": True, "settings": start_settings.load_start_settings()})

    # ---- Shared settings (available from any page) ----

    @combined.route("/api/settings", methods=["GET"])
    def combined_settings_get() -> FlaskResponse:
        return jsonify({"ok": True, "settings": _settings_records()})

    @combined.route("/api/settings", methods=["PUT"])
    def combined_settings_put() -> FlaskResponse:
        data = request.get_json(silent=True) or {}
        applied, error = _apply_settings_payload(data)
        if error is not None:
            return jsonify({"ok": False, "error": error}), 400
        return jsonify({"ok": True, "applied": applied})

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

    # ---- Model discovery ----

    @combined.route("/api/models")
    def api_models() -> Response:
        import ollama_client
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

        ollama_models: list[dict[str, Any]] = []
        ollama_available = False
        raw = ollama_client.list_models()
        if raw is not None:
            ollama_available = True
            for m in raw:
                size_mb = round(m["size_bytes"] / (1024 * 1024))
                ollama_models.append(
                    {
                        "name": m["name"],
                        "size_mb": size_mb,
                        "parameter_size": m["parameter_size"],
                        "family": m["family"],
                    }
                )

        # Per thinking-agent model + install status, so the Transcripts UI can
        # confirm a download before running an agent against a missing model.
        ollama_agents = []
        for a in thinking_agents.AGENTS:
            model = thinking_agents.resolve_model(a)
            ollama_agents.append(
                {
                    "key": a["key"],
                    "model": model,
                    "installed": ollama_client.is_model_installed(model, raw or []),
                }
            )

        return jsonify(
            {
                "ok": True,
                "whisper": {"models": whisper_models},
                "ollama": {
                    "available": ollama_available,
                    "models": ollama_models,
                    "agents": ollama_agents,
                    "base_url": config.OLLAMA_BASE_URL,
                },
            }
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


def start_combined_server(
    worksheet: Any = None,
    port: int | None = None,
    default_page: str = "studio",
    gspread_client: Any = None,
) -> None:
    """Start a combined Studio + Screenspace + Transcripts server on one port.

    All three blueprints are always registered. When *worksheet* is ``None``,
    sheet-dependent Studio routes return ``sheet_loaded: false`` placeholder
    responses; the frontend's Start overlay lets the user pick a spreadsheet
    via ``POST /api/spreadsheets/open``.

    If *gspread_client* is supplied (the CLI's auth already happened upstream),
    the Google Sheets list endpoint reports authenticated and skips the
    "Connect Google" CTA in the Start overlay.
    """
    # The web server has no interactive console: every request/run/background
    # task executes on a Flask/daemon thread with no attached stdin. Force
    # non-interactive resolution so a missing source video (or any pipeline
    # prompt) is skipped-and-reported instead of blocking the thread forever on
    # ``input()`` — this previously hung Studio generate and watch-dir-triggered
    # workflow runs alike.
    utils.NO_INPUT_MODE = True

    # Reclaim orphaned scratch files (atomic-write .tmp siblings, reel temp-clips)
    # a prior hard kill may have left in the output dir, before workers spin up.
    utils.sweep_stale_temp_artifacts()

    combined = build_combined_app(
        worksheet=worksheet,
        default_page=default_page,
        gspread_client=gspread_client,
    )
    port = port or config.SERVER_PORT
    url = f"http://127.0.0.1:{port}/{default_page}/"

    utils.info_print(f"clipgen server running at http://127.0.0.1:{port}")
    webbrowser.open(url)

    combined.run(
        host="127.0.0.1",
        port=port,
        debug=False,
        threaded=True,
        request_handler=QuietWSGIRequestHandler,
    )
