"""Screenspace Flask blueprint for clipgen.

Registered at /screenspace/ by start_combined_server(). Works with or without a
spreadsheet; auto-discovers participant videos from the input directory.
Module-level state: _manifest, _worker, _participants (initialized by
_init_screenspace_state()). The output directory is deliberately *not* cached
here — /media/ resolves it per request via utils.get_effective_output_dir().

API endpoints (all under /screenspace/):
  GET  /media/<filename>                    – serve artifact media files
  GET  /api/participants                    – list discovered participant videos
  GET  /api/participants/<pid>/notes        – get persisted free-form notes for a participant
  PUT  /api/participants/<pid>/notes        – persist free-form notes (max 64 KB)
  GET  /api/participants/<pid>/issues       – top severity-ranked Sheet rows for a participant
  GET  /api/participants/<pid>/marks        – resolved transcript marks tagged to a participant
  GET  /api/pins/<participant>              – list calibration pins (with stale flag)
  POST /api/pins/<participant>              – pin a frame as a positive/negative anchor
  PUT  /api/pins/<pin_id>                   – update a pin (polarity toggle / label edit)
  DELETE /api/pins/<pin_id>                 – remove a pin by id
  DELETE /api/pins/<participant>/all        – remove every pin for a participant
  POST /api/calibrate                       – score a participant's pins vs a tool synchronously
  GET  /api/video/frame/<participant>/<ts>  – extract a JPEG frame at timestamp
  GET  /api/heatmap-sprite/<gif>            – hover-scrub sprite sheet tiled from a heatmap GIF
  GET|POST /api/preview/<participant>/<ts>   – PNG of the active tool's preprocessed view (?layer=<id> for single-layer overlay)
  GET  /api/preview/layers                   – JSON catalog of overlay-eligible layers per tool
  GET  /api/video/info/<participant>        – video metadata (duration, resolution, fps)
  GET  /api/video/stream/<participant>     – stream source video (mp4, range-aware)
  GET  /api/regions                         – list regions
  POST /api/regions                         – create or update a region
  PUT  /api/regions/reorder                – reorder active regions by name
  DELETE /api/regions/<name>               – delete a region
  DELETE /api/regions                      – delete every active region (stashes untouched)
  GET/POST /api/stashes                    – stash CRUD (save/restore named region sets)
  PUT  /api/stashes/<id>                   – update stash
  DELETE /api/stashes/<id>                 – delete stash
  POST /api/stashes/<id>/restore           – restore a stash
  POST /api/stashes/<id>/regions           – copy one active region into a stash
  GET  /api/tasks                          – list task queue
  GET  /api/tasks/stream                   – SSE stream of live task updates (replaces polling)
  GET  /api/tasks/<task_id>               – get single task
  POST /api/tasks                          – create and enqueue a new task
  DELETE /api/tasks/<task_id>             – cancel/remove a task
  PUT  /api/tasks/reorder                  – reorder task queue by priority
  POST /api/tasks/pause                    – pause the worker
  POST /api/tasks/resume                   – resume the worker
  GET  /api/tasks/<task_id>/results        – analysis results (timestamps, artifact paths)
  GET  /api/events                         – list result events across all tasks
  GET  /api/intake-poll                     – Studio-intake poll: task-status booleans + filtered events
  GET  /api/export/events                  – export events as analysis-ready JSON (default) or CSV (?format=csv)
  PUT  /api/events/<event_id>/exclude      – mark an event excluded
  PUT  /api/events/<event_id>/include      – mark an event included
  PUT  /api/events/bulk-exclude            – bulk exclude events by task/time range
  PUT  /api/events/bulk-include            – bulk include events by task/time range
"""

import atexit
import binascii
import copy
import json
import math
import sys
import threading
import uuid
from collections import OrderedDict
from collections.abc import Callable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, TypeGuard, cast

from flask import Blueprint, Response, jsonify, request, send_file

import config
import files
import profiling
import remux_server
import screenspace
import spreadsheet
import utils
import video
from server_utils import (
    MediaCache,
    err,
    err_no_video,
    find_by_id,
    json_endpoint,
    make_debounced_persist,
    make_participant_cache,
    make_sse_channel,
    ok,
    opt_number,
    parse_number_arg,
    remove_by_id,
)

# Per-tool optional float overrides api_preview reads straight into params.
_PREVIEW_FLOAT_ARGS: dict[str, tuple[str, ...]] = {
    "color": ("h", "s", "v"),
    "shape": (
        "threshold",
        "scale_min",
        "scale_max",
        "scale_steps",
        "scale_y_min",
        "scale_y_max",
        "scale_y_steps",
    ),
    "attention": (
        # Channel-weight / center-bias overrides so the Model view tunes the
        # same math the scan runs (saliency_kwargs_from_params on both paths).
        "weight_spectral",
        "weight_contrast",
        "weight_motion",
        "weight_face",
        "center_bias",
    ),
}


def _preview_ref_rect(
    region_coords: dict[str, Any] | None, frame_w: int, frame_h: int
) -> dict[str, Any] | None:
    """Pixel rect the preview cuts its template/shape sample from.

    ``ref_region`` (normalized x,y,w,h in the query string) is the capture
    region; without it the run region doubles as the sample rect — the
    CLI/workflows single-region semantics.
    """
    ref_region_str = request.args.get("ref_region", "").strip()
    if ref_region_str:
        rr_parts = ref_region_str.split(",")
        if len(rr_parts) == 4:
            try:
                rrx, rry, rrw, rrh = (float(p) for p in rr_parts)
            except ValueError:
                pass
            else:
                return {
                    "x": round(rrx * frame_w),
                    "y": round(rry * frame_h),
                    "w": round(rrw * frame_w),
                    "h": round(rrh * frame_h),
                }
    return region_coords


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
    "shape",
    "flow",
    "scene",
    "inactivity",
    "boundary",
    "attention",
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
_participants: list[dict[str, Any]] = []
# What _participants was built from: {"sheet_context", "dir", "mtime"}, or None
# before _init_screenspace_state has run (the same "not configured yet" state
# _worker = None expresses). While None, _refresh_participants() is a no-op, so a
# directly-assigned _participants survives.
_participant_source: dict[str, Any] | None = None
_participants_lock = threading.Lock()
# Source timeline for multi-video participants (one continuous recording across
# several files), keyed by participant id as (parts_mtimes, timeline) so a
# re-encoded part invalidates the offsets. Single-video is never cached — mapping
# is a no-op.
_participant_timeline_cache: dict[
    str, tuple[tuple[int, ...], list[tuple[str, int, int]] | None]
] = {}
_participant_timeline_lock = threading.Lock()
# Values are (mtime_ns, info) so a stale file is re-probed automatically.
_video_metadata_cache: dict[str, tuple[int, dict[str, Any]]] = {}
_video_metadata_cache_lock = threading.Lock()
# Bounded LRU: entries are JPEG bytes (tens of KB), so a few hundred is plenty.
# The key includes ``mtime_ns``, so a re-encoded source is a distinct entry rather
# than stale bytes.
_FRAME_CACHE_MAX = 256
_frame_cache: "OrderedDict[tuple[str, int, float, int], bytes]" = OrderedDict()
_frame_cache_lock = threading.Lock()

# Decoded-frame cache for the synchronous CV helpers: calibration and preview
# re-run on every parameter nudge, so BGR frames memoize separately from the JPEG
# cache above. Keyed with ``mtime_ns``, so re-encodes invalidate naturally.
_DECODED_FRAME_CACHE_MAX = max(8, 2 * config.SCREENSPACE_MAX_PINS)
_decoded_frame_cache: "OrderedDict[tuple[str, int, float], Any]" = OrderedDict()
_decoded_frame_cache_lock = threading.Lock()
_PIN_OCR_CACHE_MAX = 64
_pin_ocr_cache: "OrderedDict[tuple[Any, ...], list[Any]]" = OrderedDict()
_pin_ocr_cache_lock = threading.Lock()

# Heatmap hover-scrub sprite sheets, re-tiled from the GIFs on demand and never
# written to the output dir — sprites are a derived view everywhere in clipgen.
# A handful of small PNGs at most; keyed with ``mtime_ns`` so a regenerated GIF
# invalidates naturally.
_heatmap_sprite_cache = MediaCache(32)

# ---- SSE (Server-Sent Events) client registry ----

# Broadcast channel (all task-stream clients registered under the None key);
# ``_sse_clients`` is the channel's live registry list. See make_sse_channel.
_sse_notify, _sse_stream, _sse_clients = make_sse_channel()
_manifest_lock = threading.Lock()

# Bumped under _manifest_lock on every _manifest["events"] mutation. The poll
# routes echo it, so an unchanged tick short-circuits the deep-copy + sanitize.
_events_version = 0


def _bump_events_version() -> None:
    """Mark manifest events changed. Caller must hold _manifest_lock."""
    global _events_version
    _events_version += 1


def _notify_sse_clients(event_type: str = "update") -> None:
    """Broadcast a task-state change to every connected SSE client.

    Thin adapter over the shared channel's coalescing ``notify`` (see
    make_sse_channel): on a full client queue it discards one stale entry and
    re-pushes a marker, so a lagging client still re-emits fresh task state once
    it catches up.
    """
    _sse_notify(marker=event_type)


def _sse_task_payload() -> str:
    """Build an SSE data line with current task state."""
    # Slim ticks: no result lists (clients pull tails via /api/tasks/<id>/results).
    tasks = _worker.get_all_tasks(include_results=False) if _worker else []
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
    # Resolved per request, never snapshotted: POST /api/dirs moves
    # config.OUTPUT_DIR mid-session without re-running _init_screenspace_state (the
    # Start overlay's "no spreadsheet" path closes without a reload), and a snapshot
    # then served heatmaps/timelapses from the old directory — every artifact URL a
    # dead link until something re-inited the blueprint.
    media_dir_getter=lambda: str(utils.get_effective_output_dir()),
    media_error="Output directory not configured",
    icons=True,
)

remux_server.register_remux_routes(
    screenspace_bp,
    lambda: _participant_source["sheet_context"] if _participant_source else None,
)


# ---- Participants ----


_NOTES_MAX_BYTES = 64 * 1024


# The refresh/find pair over this module's _participants globals; the factory
# reads them as module attributes so _init_screenspace_state and tests that
# monkeypatch them keep working. See server_utils.make_participant_cache.
_refresh_participants, _find_participant_record = make_participant_cache(
    sys.modules[__name__],
    input_dir_getter=utils.get_effective_input_dir,
    resolve=files.resolve_participant_videos,
)


def _participant_exists(pid: str) -> bool:
    return _find_participant_record(pid) is not None


@screenspace_bp.route("/api/participants")
def api_participants() -> FlaskResponse:
    """List participants with source video availability.

    Each entry includes a ``version`` field (``st_mtime_ns`` of the source
    file) so the frontend can cache-bust the preloaded frame-0 URL from the
    first paint instead of waiting on a follow-up ``/api/video/info`` call.
    """
    _refresh_participants()
    payload: list[dict[str, Any]] = []
    for p in _participants:
        entry = dict(p)
        if p.get("has_video"):
            entry["version"] = _participant_video_version(p["id"])
            # Multi-video: expose the timeline so the frontend can switch <video>
            # source per part and seek the local offset. Omitted for a single video
            # (no probe), leaving the frontend on its one-file path.
            timeline = _participant_timeline(p["id"])
            if timeline is not None:
                entry["timeline"] = [
                    {
                        "filename": Path(path).name,
                        "duration": dur,
                        "cumulativeStart": cum,
                    }
                    for path, dur, cum in timeline
                ]
        payload.append(entry)
    # Bootstrap channel for shared frontend config (hotkey overrides etc.);
    # this page has no sheet-data fetch, so the config rides along here.
    # ``has_sheet`` gates the off-sheet badge: with no sheet every entry is
    # ``in_sheet: False``, and marking them all would be noise.
    return ok(
        participants=payload,
        has_sheet=bool(_participant_source and _participant_source["sheet_context"]),
        config=utils.get_frontend_config(),
    )


@screenspace_bp.route("/api/participants/<pid>/notes")
def api_participant_notes_get(pid: str) -> FlaskResponse:
    """Return persisted free-form notes for a participant."""
    if not _participant_exists(pid):
        return err(f"Unknown participant {pid}", 404)
    with _manifest_lock:
        entry = _manifest.get("per_participant", {}).get(pid, {})
        notes = entry.get("notes", "")
    return ok(notes=notes)


@screenspace_bp.route("/api/participants/<pid>/notes", methods=["PUT"])
def api_participant_notes_set(pid: str) -> FlaskResponse:
    """Persist free-form notes for a participant."""
    if not _participant_exists(pid):
        return err(f"Unknown participant {pid}", 404)
    data = request.get_json(silent=True) or {}
    notes = data.get("notes", "")
    if not isinstance(notes, str):
        return err("notes must be a string")
    if len(notes.encode("utf-8")) > _NOTES_MAX_BYTES:
        return err("notes too large", 413)

    with _manifest_lock:
        per_participant = _manifest.setdefault("per_participant", {})
        entry = per_participant.setdefault(pid, {})
        entry["notes"] = notes
        _do_persist(drain_events=False)
    return ok()


@screenspace_bp.route("/api/participants/<pid>/issues")
def api_participant_issues(pid: str) -> FlaskResponse:
    """Return up to five Sheet rows tagged to a participant, ranked by severity.

    Returns an empty list when Screenspace runs without a Sheet (no Studio).
    Mirrors the row construction in ``server.api_sheet`` so the participant
    column lookup, baseline/filename row skipping, and severity normalization
    match Studio's view.
    """
    if not _participant_exists(pid):
        return err(f"Unknown participant {pid}", 404)

    import server  # lazy: avoid module-level snapshot of _sheet_context

    ctx = getattr(server, "_sheet_context", None)
    if ctx is None:
        return ok(issues=[])

    participants = spreadsheet.get_participant_list(
        ctx.header_row, ctx.id_cell, ctx.num_participants
    )
    if pid not in participants:
        return ok(issues=[])
    p_idx = participants.index(pid)
    col_idx = ctx.id_cell.col + p_idx

    obs_col = ctx.observation_cell.col - 1
    sev_col = ctx.severity_cell.col - 1 if ctx.severity_cell else None

    candidates: list[dict[str, Any]] = []
    for row_idx in range(ctx.first_data_row_idx, len(ctx.sheet_data)):
        if ctx.baseline_row_idx is not None and row_idx == ctx.baseline_row_idx:
            continue
        if ctx.filename_row_idx is not None and row_idx == ctx.filename_row_idx:
            continue
        row_data = ctx.sheet_data[row_idx]
        if col_idx >= len(row_data) or not row_data[col_idx].strip():
            continue
        raw_cell = row_data[col_idx].strip()
        ts_pairs = utils.parse_timestamps(raw_cell)
        ts_seconds = utils.timestamp_to_seconds(ts_pairs[0][0]) if ts_pairs else None
        observation = row_data[obs_col] if obs_col < len(row_data) else ""
        severity = ""
        if (
            sev_col is not None
            and sev_col < len(row_data)
            and row_data[sev_col].strip()
        ):
            severity = utils.normalize_severity(row_data[sev_col])
        candidates.append(
            {
                "rowNum": row_idx + 1,
                "observation": observation,
                "severity": severity,
                "timestamp": ts_seconds,
                "_row_idx": row_idx,
            }
        )

    # Ascending severity score puts the most-negative (Critical = -4) first, with
    # positives and unranked rows at the end. Tie-break by row order for stability.
    candidates.sort(
        key=lambda c: (utils.severity_sort_key(c["severity"]), c["_row_idx"])
    )
    chosen = candidates[:5]
    issues = [{k: v for k, v in c.items() if not k.startswith("_")} for c in chosen]
    return ok(issues=issues)


@screenspace_bp.route("/api/participants/<pid>/marks")
def api_participant_marks(pid: str) -> FlaskResponse:
    """Return transcript marks tagged to a participant, sorted by start time.

    Reuses the in-memory transcripts manifest (loaded for every page by
    build_combined_app via _init_transcripts_state) — no extra disk read.
    Returns an empty list when the participant has no resolvable marks.
    """
    if not _participant_exists(pid):
        return err(f"Unknown participant {pid}", 404)

    import transcripts_server  # lazy: mirrors the `import server` pattern above

    marks = transcripts_server.marks_for_participant(pid)
    return ok(marks=marks, categories=config.MARK_CATEGORIES)


# ---- Calibration pins ----
#
# A pin is a tool- and region-agnostic marker that "this frame matters", carrying
# only a timestamp and a polarity (``positive`` = the condition is true here,
# ``negative`` = it must not fire here). Stored per participant under the manifest
# ``pins`` key, driving synchronous detector calibration. Frames are never stored
# — they're fetched through the frame API, so a re-encode invalidates via
# ``mtime_ns``.

_PIN_POLARITIES = ("positive", "negative")
_PIN_LABEL_MAX_CHARS = 120


def _pin_manifest() -> dict[str, list[dict[str, Any]]]:
    """Return the manifest's pin map, replacing malformed roots with empty state."""
    pins = _manifest.get("pins")
    if isinstance(pins, dict):
        return pins
    _manifest["pins"] = {}
    return _manifest["pins"]


def _participant_pin_list(
    participant_id: str, *, create: bool = False
) -> list[dict[str, Any]]:
    """Return a participant's pin list; malformed entries behave like no pins."""
    pins_by_participant = _pin_manifest()
    pins = pins_by_participant.get(participant_id)
    if isinstance(pins, list):
        return pins
    if create:
        pins_by_participant[participant_id] = []
        return pins_by_participant[participant_id]
    return []


def _annotate_pin_staleness(
    participant_id: str, pins: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    """Return copies of ``pins`` with a ``stale`` flag for out-of-range timestamps.

    A pin is stale when its timestamp lies beyond the current source video's
    duration (e.g. the footage was replaced with a shorter cut). Duration that
    cannot be determined leaves every pin non-stale.
    """
    duration = _participant_video_duration(participant_id)
    out: list[dict[str, Any]] = []
    for pin in pins:
        if not isinstance(pin, dict):
            continue
        entry = dict(pin)
        entry["stale"] = duration is not None and pin.get("timestamp", 0.0) > duration
        out.append(entry)
    return out


def _find_pin(pin_id: str) -> tuple[str, list[dict[str, Any]], int] | None:
    """Locate a pin by id across all participants.

    Returns ``(participant_id, pin_list, index)`` or ``None``. Caller holds
    ``_manifest_lock``.
    """
    for participant_id, pins in _pin_manifest().items():
        if not isinstance(pins, list):
            continue
        for idx, pin in enumerate(pins):
            if isinstance(pin, dict) and pin.get("id") == pin_id:
                return participant_id, pins, idx
    return None


@screenspace_bp.route("/api/pins/<participant>")
def api_pins_list(participant: str) -> FlaskResponse:
    """List calibration pins for a participant (annotated with a stale flag)."""
    if not _participant_exists(participant):
        return err(f"Unknown participant {participant}", 404)
    with _manifest_lock:
        pins = copy.deepcopy(_participant_pin_list(participant))
    return ok(
        pins=_annotate_pin_staleness(participant, pins),
        max_pins=config.SCREENSPACE_MAX_PINS,
    )


@screenspace_bp.route("/api/pins/<participant>", methods=["POST"])
@json_endpoint
def api_pins_create(participant: str) -> FlaskResponse:
    """Pin the given frame as a positive or negative calibration anchor."""
    if not _participant_exists(participant):
        return err(f"Unknown participant {participant}", 404)
    data = request.get_json(silent=True) or {}

    timestamp = parse_number_arg(
        data.get("timestamp"), "timestamp", min_=0, finite=True
    )

    polarity = data.get("polarity")
    if polarity not in _PIN_POLARITIES:
        return err("polarity must be 'positive' or 'negative'")

    label = data.get("label", "")
    if not isinstance(label, str):
        return err("label must be a string")
    label = label.strip()[:_PIN_LABEL_MAX_CHARS]

    pin = {
        "id": "pin_" + uuid.uuid4().hex[:8],
        "timestamp": round(timestamp, 3),
        "polarity": polarity,
        "label": label,
        "created_at": datetime.now(UTC).isoformat(),
    }
    with _manifest_lock:
        pins = _participant_pin_list(participant, create=True)
        if len(pins) >= config.SCREENSPACE_MAX_PINS:
            return err(f"Pin limit reached (max {config.SCREENSPACE_MAX_PINS})", 409)
        pins.append(pin)
        _do_persist(drain_events=False)
    return ok(pin=pin)


@screenspace_bp.route("/api/pins/<pin_id>", methods=["PUT"])
def api_pins_update(pin_id: str) -> FlaskResponse:
    """Update a pin's polarity and/or label (located by id across participants)."""
    data = request.get_json(silent=True) or {}
    with _manifest_lock:
        found = _find_pin(pin_id)
        if found is None:
            return err("Pin not found", 404)
        _participant_id, pins, idx = found
        pin = pins[idx]
        if "polarity" in data:
            if data["polarity"] not in _PIN_POLARITIES:
                return err("polarity must be 'positive' or 'negative'")
            pin["polarity"] = data["polarity"]
        if "label" in data:
            if not isinstance(data["label"], str):
                return err("label must be a string")
            pin["label"] = data["label"].strip()[:_PIN_LABEL_MAX_CHARS]
        _do_persist(drain_events=False)
    return ok(pin=pin)


@screenspace_bp.route("/api/pins/<pin_id>", methods=["DELETE"])
def api_pins_delete(pin_id: str) -> FlaskResponse:
    """Remove a pin by id (located across all participants)."""
    with _manifest_lock:
        found = _find_pin(pin_id)
        if found is None:
            return err("Pin not found", 404)
        participant_id, pins, idx = found
        pins.pop(idx)
        if not pins:
            _pin_manifest().pop(participant_id, None)
        _do_persist(drain_events=False)
    return ok()


@screenspace_bp.route("/api/pins/<participant>/all", methods=["DELETE"])
def api_pins_delete_all(participant: str) -> FlaskResponse:
    """Remove every calibration pin for a participant.

    A distinct two-segment path so it does not collide with the single-segment
    ``DELETE /api/pins/<pin_id>`` route above (both use the DELETE method).
    """
    if not _participant_exists(participant):
        return err(f"Unknown participant {participant}", 404)
    with _manifest_lock:
        _pin_manifest().pop(participant, None)
        _do_persist(drain_events=False)
    return ok()


# ---- Pin calibration (synchronous, off the task queue) ----


def _decoded_video_frame(video_path: str, mtime_ns: int, ts: float) -> "Any | None":
    """Return a decoded BGR frame at ``ts``, memoized per (video, mtime, ts).

    Backs calibration and preview so repeated parameter nudges don't re-decode
    the same current, companion, or reference frames. Returns ``None`` when
    extraction fails.
    """
    key = (video_path, mtime_ns, round(ts, 3))
    with _decoded_frame_cache_lock:
        cached = _decoded_frame_cache.get(key)
        if cached is not None:
            _decoded_frame_cache.move_to_end(key)
            profiling.count("screenspace.decoded_frame_cache.hit")
            return cached
    profiling.count("screenspace.decoded_frame_cache.miss")
    frame = video.extract_frame_at_timestamp(video_path, ts)
    if frame is None:
        return None
    with _decoded_frame_cache_lock:
        _decoded_frame_cache[key] = frame
        while len(_decoded_frame_cache) > _DECODED_FRAME_CACHE_MAX:
            _decoded_frame_cache.popitem(last=False)
    return frame


def _make_pin_ocr_reader(
    video_path: str, mtime_ns: int, ts: float, frame: "Any"
) -> "Callable[[str, dict[str, Any], dict[str, Any]], list[Any]]":
    """Build the cached OCR reader passed to the score functions for one pin.

    Memoizes raw OCR readings per (video, mtime, ts, region, langs, preprocess)
    so changing only the fuzzy/confidence threshold re-scores from cached
    readings without re-running OCR. Text and numbers pins on the same region
    share readings — integers_only and the numeric operators are applied at
    scoring time, which is exactly what this cache exists for.
    """

    def _reader(
        tool_type: str, region_coords: dict[str, Any], params: dict[str, Any]
    ) -> list[Any]:
        del tool_type  # readings are tool-independent; scoring applies the tool
        langs = tuple(params.get("languages") or ["en"])
        key = (
            video_path,
            mtime_ns,
            round(ts, 3),
            (
                region_coords.get("x"),
                region_coords.get("y"),
                region_coords.get("w"),
                region_coords.get("h"),
                # Shaped regions with the same bbox but different contours
                # must not share cached readings.
                screenspace.mask_points_key(region_coords.get("mask_points")),
            ),
            langs,
            bool(params.get("ocr_preprocess", False)),
        )
        with _pin_ocr_cache_lock:
            cached = _pin_ocr_cache.get(key)
            if cached is not None:
                _pin_ocr_cache.move_to_end(key)
                profiling.count("screenspace.pin_ocr_cache.hit")
                return cached
        profiling.count("screenspace.pin_ocr_cache.miss")
        readings = screenspace.run_calibration_ocr(frame, region_coords, params)
        with _pin_ocr_cache_lock:
            _pin_ocr_cache[key] = readings
            while len(_pin_ocr_cache) > _PIN_OCR_CACHE_MAX:
                _pin_ocr_cache.popitem(last=False)
        return readings

    return _reader


def _calibration_interval(task_type: str, parameters: dict[str, Any]) -> float:
    """Companion-frame interval for change/flow/inactivity (and multitool step 0)."""
    if task_type == "multitool":
        steps = parameters.get("steps", [])
        raw = steps[0].get("interval") if steps else None
    else:
        raw = parameters.get("interval")
    try:
        val = float(raw) if raw is not None else 0.0
    except (TypeError, ValueError):
        val = 0.0
    return val if val > 0 else float(config.SCREENSPACE_DEFAULT_INTERVAL)


def _calibratable_tool(tool: str) -> bool:
    """A tool is calibratable when it exposes a per-frame scalar (or is multitool)."""
    if tool == "timelapse" or tool not in _VALID_TASK_TYPES:
        return False
    if tool == "multitool":
        return True
    return bool(screenspace.TOOLS[tool].score_key)


@screenspace_bp.route("/api/calibrate", methods=["POST"])
def api_calibrate() -> FlaskResponse:
    """Score a participant's pins against a tool + parameters, synchronously.

    Body: ``{participant, tool, parameters, region_ref?, region?, pin_ids?}``.
    Returns one ``{pin_id, timestamp, polarity, status, score?, passed?, detail?}``
    per pin (multitool entries also carry ``steps`` and a chain ``passed``). This
    is per-frame only — temporal params (consecutive/interval) are not validated.
    """
    data = request.get_json(silent=True)
    if not data:
        return err("JSON body required")

    tool = (data.get("tool") or "").strip()
    if not _calibratable_tool(tool):
        return err(f"Tool '{tool}' is not calibratable")

    # Reuse task-creation validation by reshaping the body into a task request.
    validated = _validate_task_request(
        {
            "type": tool,
            "participant": data.get("participant", ""),
            "region": data.get("region", ""),
            "region_ref": data.get("region_ref"),
            "parameters": data.get("parameters"),
        }
    )
    if _is_flask_error_response(validated):
        return validated
    if not (isinstance(validated, tuple) and len(validated) == 6):
        return err("Invalid calibration request")
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

    resolved = _find_participant_video_with_mtime(participant)
    if resolved is None:
        return err_no_video(participant)
    video_path, _mtime_ns = resolved

    props = video.probe_video_properties(video_path)
    resolve_region = _region_coords_resolver(props, all_known_regions)

    if requested_region is not None:
        region_coords = resolve_region(region_name, requested_region)
    else:
        region_coords = {"x": 0, "y": 0, "w": 0, "h": 0}

    prepared = _prepare_task_media(
        task_type,
        participant,
        parameters,
        all_known_regions,
        region_coords,
        resolve_region,
    )
    if isinstance(prepared, tuple):
        return prepared
    parameters = cast(dict[str, Any], prepared)

    pin_ids = data.get("pin_ids")
    if pin_ids is not None and not isinstance(pin_ids, list):
        return err("pin_ids must be a list")

    with _manifest_lock:
        all_pins = copy.deepcopy(_participant_pin_list(participant))
    all_pins = [p for p in all_pins if isinstance(p, dict)]
    if isinstance(pin_ids, list):
        wanted = set(pin_ids)
        all_pins = [p for p in all_pins if p.get("id") in wanted]
    all_pins = all_pins[: config.SCREENSPACE_MAX_PINS]

    interval = _calibration_interval(task_type, parameters)
    steps = parameters.get("steps", [])
    needs_prev = task_type in ("change", "flow", "inactivity") or (
        task_type == "multitool"
        and any(s.get("type") in ("change", "flow", "inactivity") for s in steps)
    )

    results: list[dict[str, Any]] = []
    for pin in all_pins:
        ts = float(pin.get("timestamp", 0.0))
        entry: dict[str, Any] = {
            "pin_id": pin.get("id"),
            "timestamp": ts,
            "polarity": pin.get("polarity"),
        }
        try:
            # Map the pin's global timestamp into the owning sub-video.
            mapped = _map_participant_time(participant, ts)
            if mapped is None:
                entry["status"] = "not_evaluable"
                results.append(entry)
                continue
            sub_path, local_ts = mapped
            sub_mtime = _mtime_or_zero(sub_path)
            frame = _decoded_video_frame(sub_path, sub_mtime, local_ts)
            if frame is None:
                entry["status"] = "not_evaluable"
                results.append(entry)
                continue
            # Companion frame for temporal tools; ts < interval ⇒ no companion
            # (the score path then reports not_evaluable for that step/tool).
            prev_frame = None
            if needs_prev and ts >= interval:
                mapped_prev = _map_participant_time(participant, ts - interval)
                if mapped_prev is not None:
                    prev_frame = _decoded_video_frame(
                        mapped_prev[0], _mtime_or_zero(mapped_prev[0]), mapped_prev[1]
                    )
            ocr_reader = _make_pin_ocr_reader(sub_path, sub_mtime, local_ts, frame)
            if task_type == "multitool":
                score = screenspace.score_multitool_frame(
                    frame, prev_frame, steps, ocr_reader=ocr_reader
                )
            else:
                score = screenspace.score_frame_for_tool(
                    task_type,
                    frame,
                    prev_frame,
                    region_coords,
                    parameters,
                    ocr_reader=ocr_reader,
                )
            entry.update(score)
        except Exception as exc:  # one bad pin must not 500 the whole batch
            utils.warning_print(f"Calibration failed for pin {pin.get('id')}: {exc}")
            entry["status"] = "not_evaluable"
        results.append(entry)

    return jsonify(utils.sanitize_floats({"ok": True, "tool": tool, "pins": results}))


# ---- Video frame extraction ----


def _participant_video_paths(participant_id: str) -> list[str]:
    """Return a participant's ordered source video path(s), or [] if none.

    A multi-video participant (recording split across files) returns all parts
    in timeline order; a normal participant returns a single-element list.

    Refreshes first: the frame/stream/task routes are reachable directly (API
    use, automation) without ``/api/participants`` having run since the video
    landed, and without this they would 404 on a participant that is on disk.
    The refresh is one ``stat()`` in the steady state — nothing against the
    ffmpeg work these callers are about to do.
    """
    record = _find_participant_record(participant_id)
    return list(record["video_paths"]) if record and record.get("has_video") else []


def _part_mtimes(paths: list[str]) -> tuple[int, ...] | None:
    """Return each part's ``mtime_ns`` as a tuple, or None if any is missing."""
    out: list[int] = []
    for path in paths:
        try:
            out.append(Path(path).stat().st_mtime_ns)
        except OSError:
            return None
    return tuple(out)


def _participant_timeline(participant_id: str) -> list[tuple[str, int, int]] | None:
    """Cached source timeline for a 2+ part participant; None for single video.

    Keyed on every part's ``mtime_ns`` so replacing any part recomputes the
    cumulative offsets. Single-video participants return None (mapping is a
    no-op — no duration probe).
    """
    paths = _participant_video_paths(participant_id)
    if len(paths) < 2:
        return None
    mtimes = _part_mtimes(paths)
    if mtimes is None:
        return None
    with _participant_timeline_lock:
        cached = _participant_timeline_cache.get(participant_id)
        if cached is not None and cached[0] == mtimes:
            return cached[1]
    timeline = video.build_source_timeline(paths)
    with _participant_timeline_lock:
        _participant_timeline_cache[participant_id] = (mtimes, timeline)
    return timeline


def _mtime_or_zero(path: str) -> int:
    """Return a path's ``mtime_ns`` for cache keys, or 0 if it can't be stat'd."""
    try:
        return Path(path).stat().st_mtime_ns
    except OSError:
        return 0


def _map_participant_time(
    participant_id: str, global_ts: float
) -> tuple[str, float] | None:
    """Map a global timestamp to ``(sub_video_path, local_ts)``.

    Single-video participants map to ``(path, global_ts)`` unchanged (no probe,
    no stat). Multi-video participants resolve which sub-video owns *global_ts*
    and the local offset within it. Returns None when the participant has no
    video or (multi-video) the timestamp is out of range. Callers needing the
    file's mtime for a cache key obtain it via :func:`_mtime_or_zero`.
    """
    paths = _participant_video_paths(participant_id)
    if not paths:
        return None
    if len(paths) < 2:
        return (paths[0], global_ts)
    timeline = _participant_timeline(participant_id)
    if timeline is None:
        return None
    return utils.resolve_timeline_segment(timeline, global_ts)


def _participant_frame_extractor(
    participant_id: str,
) -> Callable[[float], "Any | None"]:
    """Return a ``frame_at(global_ts)`` closure mapping into the right sub-video.

    Used by task-media helpers that extract reference frames at global reference
    timestamps; single-video participants extract at the same time unchanged.
    """

    def _extract(global_ts: float) -> "Any | None":
        mapped = _map_participant_time(participant_id, global_ts)
        if mapped is None:
            return None
        return _decoded_video_frame(mapped[0], _mtime_or_zero(mapped[0]), mapped[1])

    return _extract


def _find_participant_video(participant_id: str) -> str | None:
    """Resolve the first source-video path for a participant (or None)."""
    paths = _participant_video_paths(participant_id)
    return paths[0] if paths else None


def _find_participant_video_with_mtime(participant_id: str) -> tuple[str, int] | None:
    """Resolve a participant's first source video path and its ``mtime_ns``.

    Returns ``(path, mtime_ns)`` or ``None`` if the participant has no video or
    the file no longer exists. For frame extraction at a specific timestamp use
    ``_map_participant_time`` instead (it maps multi-video global times into the
    owning sub-video).
    """
    path = _find_participant_video(participant_id)
    if path is None:
        return None
    try:
        return path, Path(path).stat().st_mtime_ns
    except OSError:
        return None


def _participant_video_version(participant_id: str) -> int | None:
    """Return a cache-bust version for a participant's source video(s), or None.

    For multi-video participants this combines every part's ``mtime_ns`` so the
    frontend's frame cache invalidates when any part is replaced.
    """
    paths = _participant_video_paths(participant_id)
    if not paths:
        return None
    mtimes = _part_mtimes(paths)
    if mtimes is None:
        return None
    return sum(mtimes)


def _participant_video_duration(participant_id: str) -> float | None:
    """Return the source video duration in seconds, or ``None`` if undeterminable.

    Reuses the ``api_video_info`` metadata cache when warm (the frontend probes
    it on participant load); otherwise probes once. Used to flag pins whose
    timestamp falls beyond a replaced/shortened video's duration.
    """
    # Multi-video participants: total timeline duration (pins use global times).
    timeline = _participant_timeline(participant_id)
    if timeline is not None:
        return float(timeline[-1][1] + timeline[-1][2])
    resolved = _find_participant_video_with_mtime(participant_id)
    if resolved is None:
        return None
    video_path, mtime_ns = resolved
    with _video_metadata_cache_lock:
        cached = _video_metadata_cache.get(participant_id)
    # Use the unrounded duration: the cached ``duration`` is rounded for display
    # and would mis-flag pins within ~0.5s of the true end.
    if cached is not None and cached[0] == mtime_ns:
        return cached[1].get("duration_seconds") or None
    props = video.probe_video_properties(video_path)
    if props is None:
        return None
    return props.get("duration", 0.0) or None


@screenspace_bp.route("/api/video/frame/<participant>/<timestamp>")
def api_video_frame(participant: str, timestamp: str) -> FlaskResponse:
    """Extract and return a single JPEG frame at the given timestamp.

    Optional query parameter ``w`` requests a scaled-down thumbnail
    (e.g. ``?w=200`` for a 200 px-wide JPEG).  Without ``w`` the frame
    is returned at full resolution.  Results are cached in-memory and the
    cache key includes the source file's ``mtime_ns`` so a re-encoded video
    yields fresh bytes. Frontend pairs the URL with ``?v=<mtime>`` so the
    browser HTTP cache invalidates on the same boundary, enabling the long
    ``immutable`` ``Cache-Control`` below.
    """
    try:
        ts = float(timestamp)
    except (ValueError, TypeError):
        return err("Invalid timestamp")

    # Map the global timestamp into the owning sub-video (multi-video) so the
    # frontend can keep requesting frames by global time; single-video unchanged.
    mapped = _map_participant_time(participant, ts)
    if mapped is None:
        return err_no_video(participant)
    video_path, local_ts = mapped
    mtime_ns = _mtime_or_zero(video_path)

    width = request.args.get("w", 0, type=int)
    cache_key = (video_path, mtime_ns, round(local_ts, 3), width)
    with _frame_cache_lock:
        cached = _frame_cache.get(cache_key)
        if cached is not None:
            # Refresh LRU recency.
            _frame_cache.move_to_end(cache_key)
    profiling.count(
        "screenspace.frame_cache.hit"
        if cached is not None
        else "screenspace.frame_cache.miss"
    )
    if cached is not None:
        return Response(
            cached,
            mimetype="image/jpeg",
            headers={"Cache-Control": "public, max-age=86400, immutable"},
        )

    if width > 0:
        jpeg_bytes = video.extract_thumbnail_bytes(video_path, local_ts, width=width)
    else:
        frame = video.extract_frame_at_timestamp(video_path, local_ts)
        if frame is None:
            return err("Could not read frame at timestamp")
        import cv2

        success, jpeg = cv2.imencode(".jpg", frame, [cv2.IMWRITE_JPEG_QUALITY, 85])
        if not success:
            return err("Could not extract frame")
        jpeg_bytes = jpeg.tobytes()

    if jpeg_bytes is None:
        return err("Could not extract frame")

    with _frame_cache_lock:
        _frame_cache[cache_key] = jpeg_bytes
        while len(_frame_cache) > _FRAME_CACHE_MAX:
            _frame_cache.popitem(last=False)
    return Response(
        jpeg_bytes,
        mimetype="image/jpeg",
        headers={"Cache-Control": "public, max-age=86400, immutable"},
    )


@screenspace_bp.route("/api/heatmap-sprite/<filename>")
def api_heatmap_sprite(filename: str) -> FlaskResponse:
    """Tiled PNG sprite sheet of a heatmap GIF's frames, for hover scrubbing.

    A GIF cannot be seeked in a browser, so the Results panel rests each animated
    heatmap thumb on a sprite cell and maps hover position to a frame. The sheet
    is derived from the GIF that already sits in the output directory rather than
    written beside it — that directory is the user's deliverable, and sprites are
    a derived, cacheable view everywhere else in clipgen too (studio's
    ``/api/sprite``, composer's timeline tiles).

    ``cols`` comes from the caller so the layout always matches the geometry the
    task descriptor already handed the frontend, even if the config default
    changed since the scan ran.
    """
    name = Path(filename).name  # never let a path escape the output directory
    if not name.startswith("heatmap_") or not name.endswith(".gif"):
        return err("Not a heatmap animation", 404)
    gif_path = Path(utils.get_effective_output_dir()) / name
    if not gif_path.is_file():
        return err("Heatmap animation not found", 404)

    cols = request.args.get("cols", config.SCREENSPACE_HEATMAP_SPRITE_COLS, type=int)
    cols = max(1, min(int(cols or 1), 64))

    sprite_bytes = _heatmap_sprite_cache.get_or_compute(
        (str(gif_path), _mtime_or_zero(str(gif_path)), cols),
        lambda: screenspace.build_gif_sprite_bytes(str(gif_path), cols),
    )
    if sprite_bytes is None:
        return err("Could not build heatmap sprite", 404)
    return Response(
        sprite_bytes,
        mimetype="image/png",
        headers={"Cache-Control": "public, max-age=86400"},
    )


@screenspace_bp.route("/api/preview/<participant>/<timestamp>", methods=["GET", "POST"])
def api_preview(participant: str, timestamp: str) -> FlaskResponse:
    """Render what the selected tool's CV pipeline sees at ``timestamp``.

    Returns a PNG composite image (grayscale crop, diff mask, edge map, flow
    vectors, pHash bit grid, etc.) tailored to the active tool.  Optional query
    params:

      tool=<name>          one of the screenspace tool types
      region=x,y,w,h       normalized 0–1 region coordinates (required for most tools)
      prev=<seconds>       prior timestamp for change/flow (defaults to ts-1s)
      ref=<seconds>        reference timestamp for similarity (region crop) or template
                           capture preview (same as ``reference_timestamp`` in tasks)
      noise=<int>          change tool's noise_threshold override
      h,s,v=<int>          color tool's target HSV override
      magnitude=<float>    flow tool's magnitude threshold override
      ocr_preprocess=<bool> text/numbers ROI enhancement override
      layer=<id>           if set, return that single overlay layer at native
                           region/frame resolution instead of the labeled
                           composite. See ``screenspace_preview.OVERLAY_LAYERS``
                           for valid (tool, layer) pairs.

    For **template** / **shape** with an **uploaded** PNG, send ``POST`` with a
    JSON body ``{"template_image_data": "<base64>"}`` (shape:
    ``shape_image_data``; same fields as task enqueue); query string still
    supplies ``tool``, ``region`` (optional), and ``_`` cache-bust.
    """
    import screenspace_preview

    try:
        ts = float(timestamp)
    except (ValueError, TypeError):
        return err("Invalid timestamp")

    video_path = _find_participant_video(participant)
    if video_path is None:
        return err_no_video(participant)

    tool = (request.args.get("tool") or "").strip() or "color"
    if tool not in _VALID_TASK_TYPES:
        return err(f"Unknown tool: {tool}")

    frame_at = _participant_frame_extractor(participant)
    frame = frame_at(ts)
    if frame is None:
        return err("Could not read frame")
    frame_h, frame_w = frame.shape[:2]

    region_coords: dict[str, Any] | None = None
    region_str = request.args.get("region", "").strip()
    if region_str:
        parts = region_str.split(",")
        if len(parts) == 4:
            try:
                rx, ry, rw, rh = (float(p) for p in parts)
            except ValueError:
                return err("Invalid region")
            region_coords = {
                "x": round(rx * frame_w),
                "y": round(ry * frame_h),
                "w": round(rw * frame_w),
                "h": round(rh * frame_h),
            }
            # Optional shaped-region contours: "u1,v1;u2,v2;..." bbox-relative
            # fractions, multiple contours joined with "|". Malformed values
            # are ignored (preview falls back to the plain rect) rather than
            # failing the whole preview.
            mask_str = request.args.get("mask", "").strip()
            if mask_str:
                try:
                    mask_points = [
                        [
                            [float(u), float(v)]
                            for u, v in (pair.split(",") for pair in part.split(";"))
                        ]
                        for part in mask_str.split("|")
                        if part
                    ]
                except ValueError:
                    mask_points = []
                mask_points = [c for c in mask_points if len(c) >= 3]
                if mask_points:
                    region_coords["mask_points"] = mask_points

    # Prev frame for tools that consume a temporal pair. Attention's motion
    # channel compares at its own (shorter) sampling interval by default.
    prev_frame = None
    if tool in ("change", "flow", "attention"):
        default_gap = (
            config.SCREENSPACE_ATTENTION_INTERVAL if tool == "attention" else 1.0
        )
        prev_raw = opt_number(request.args, "prev")
        prev_ts = prev_raw if prev_raw is not None else max(0.0, ts - default_gap)
        if prev_ts < ts:
            prev_frame = frame_at(prev_ts)

    # Build params dict for the preview (subset of task parameters)
    params: dict[str, Any] = {}
    for key in _PREVIEW_FLOAT_ARGS.get(tool, ()):
        value = opt_number(request.args, key)
        if value is not None:
            params[key] = value
    if tool == "change":
        value = opt_number(request.args, "noise")
        if value is not None and math.isfinite(value):
            params["noise_threshold"] = int(value)
    elif tool == "flow":
        value = opt_number(request.args, "magnitude")
        if value is not None:
            params["magnitude_threshold"] = value
    elif tool in ("text", "numbers"):
        raw = (request.args.get("ocr_preprocess") or "").strip().lower()
        if raw in ("1", "true", "yes", "on"):
            params["ocr_preprocess"] = True
    elif tool == "similarity":
        ref_ts = opt_number(request.args, "ref")
        if ref_ts is not None and region_coords is not None:
            ref_frame = frame_at(ref_ts)
            if ref_frame is not None:
                import screenspace as _ss

                params["reference_frame"] = _ss.extract_region(ref_frame, region_coords)

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
                return err("Could not decode uploaded image")
            params["template_image"] = bgr
            if mask is not None:
                params["template_mask"] = mask
        else:
            ref_ts_tpl = opt_number(request.args, "ref")
            tpl_rect = _preview_ref_rect(region_coords, frame_w, frame_h)
            if ref_ts_tpl is not None and tpl_rect is not None and tpl_rect.get("w"):
                ref_frame_tpl = frame_at(ref_ts_tpl)
                if ref_frame_tpl is not None:
                    params["template_image"] = _ss_tpl.extract_region(
                        ref_frame_tpl, tpl_rect
                    )

    elif tool == "shape":
        import screenspace as _ss_shp

        shape_b64: str | None = None
        if request.method == "POST":
            body = request.get_json(silent=True)
            if isinstance(body, dict):
                raw = body.get("shape_image_data")
                if isinstance(raw, str) and raw.strip():
                    shape_b64 = raw.strip()
        if shape_b64:
            try:
                bgr, mask = _template_bgr_and_mask_from_b64(shape_b64)
            except ValueError:
                return err("Could not decode uploaded image")
            params["shape_image"] = bgr
            if mask is not None:
                params["shape_mask"] = mask
        else:
            ref_ts_shp = opt_number(request.args, "ref")
            ref_rect = _preview_ref_rect(region_coords, frame_w, frame_h)
            if ref_ts_shp is not None and ref_rect is not None and ref_rect.get("w"):
                ref_frame_shp = frame_at(ref_ts_shp)
                if ref_frame_shp is not None:
                    params["shape_image"] = _ss_shp.extract_region(
                        ref_frame_shp, ref_rect
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
            return err(f"Layer '{layer}' not available for tool '{tool}'")
        layer_img = screenspace_preview.build_overlay_layer(
            frame, prev_frame, region_coords, tool, layer, params
        )
        if layer_img is None or getattr(layer_img, "size", 0) == 0:
            return err("Could not build overlay layer", 500)
        png_bytes = screenspace_preview.encode_png(layer_img, cap_width=False)
        if not png_bytes:
            return err("Could not encode overlay", 500)
        return Response(
            png_bytes,
            mimetype="image/png",
            headers={"Cache-Control": "no-cache"},
        )

    img = screenspace_preview.build_preview(
        frame, prev_frame, region_coords, tool, params
    )
    if img is None or getattr(img, "size", 0) == 0:
        return err("Could not build preview", 500)

    png_bytes = screenspace_preview.encode_png(img)
    if not png_bytes:
        return err("Could not encode preview", 500)
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
    return ok(layers=out)


@screenspace_bp.route("/api/video/info/<participant>")
def api_video_info(participant: str) -> FlaskResponse:
    """Return video metadata (duration, resolution, fps).

    Includes a ``version`` field carrying the source file's ``st_mtime_ns``.
    The frontend appends ``?v=<version>`` to frame and stream URLs so HTTP,
    backend, and blob caches all invalidate together when the file changes.
    """
    # Multi-video participant: report the total timeline duration + the per-part
    # breakdown so the frontend can switch the <video> source per part.
    timeline = _participant_timeline(participant)
    if timeline is not None:
        props = video.probe_video_properties(timeline[0][0]) or {}
        total = timeline[-1][1] + timeline[-1][2]
        info = {
            "participant": participant,
            "duration": total,
            "duration_seconds": float(total),
            "fps": props.get("fps", 0.0) or 30.0,
            "width": props.get("width") or None,
            "height": props.get("height") or None,
            "nb_frames": 0,
            "video_codec": props.get("video_codec") or "",
            "audio_tracks": props.get("audio_tracks") or [],
            "audio_track_count": props.get("audio_track_count") or 0,
            "version": _participant_video_version(participant),
            "parts": [
                {"filename": Path(p).name, "duration": d, "cumulativeStart": c}
                for p, d, c in timeline
            ],
        }
        return ok(info=info)

    resolved = _find_participant_video_with_mtime(participant)
    if resolved is None:
        return err_no_video(participant)
    video_path, mtime_ns = resolved

    with _video_metadata_cache_lock:
        cached = _video_metadata_cache.get(participant)
    if cached is not None and cached[0] == mtime_ns:
        return ok(info=cached[1])

    props = video.probe_video_properties(video_path)
    if props is None:
        return err("Could not probe video file", 500)

    vid_fps = props.get("fps", 0.0) or 30.0
    width = props.get("width", 0)
    height = props.get("height", 0)
    duration_seconds = props.get("duration", 0.0)
    duration = round(duration_seconds) if duration_seconds > 0 else 0

    info: dict[str, Any] = {
        "participant": participant,
        "duration": duration,
        # Unrounded duration retained for precise comparisons (e.g. pin
        # staleness); the rounded ``duration`` above stays the display value.
        "duration_seconds": duration_seconds,
        "fps": vid_fps,
        "width": width if width > 0 else None,
        "height": height if height > 0 else None,
        "nb_frames": props.get("nb_frames", 0) or 0,
        "video_codec": props.get("video_codec") or "",
        "audio_tracks": props.get("audio_tracks") or [],
        "audio_track_count": props.get("audio_track_count") or 0,
        "version": mtime_ns,
    }
    with _video_metadata_cache_lock:
        _video_metadata_cache[participant] = (mtime_ns, info)

    return ok(info=info)


@screenspace_bp.route("/api/video/stream/<participant>")
def api_video_stream(participant: str) -> FlaskResponse:
    """Stream the source video file for a participant (range-request aware).

    ``conditional=True`` lets Flask set ``Last-Modified``/``ETag`` from the
    file stat and answer ``If-Modified-Since`` with a cheap 304. We add
    ``Cache-Control: no-cache`` so the browser always revalidates, which
    pairs with the frontend's ``?v=<mtime>`` cache-bust to guarantee a fresh
    stream after a source-file replacement even when range requests are in
    flight.
    """
    paths = _participant_video_paths(participant)
    if not paths:
        return err_no_video(participant)
    # Multi-video participants: ?part=N selects the sub-video; the frontend swaps
    # the <video> source per part as it scrubs the global timeline. Defaults to
    # part 0 (and is the only file for single-video participants).
    part = request.args.get("part", type=int)
    video_path = (
        paths[part] if part is not None and 0 <= part < len(paths) else paths[0]
    )
    response = send_file(video_path, mimetype="video/mp4", conditional=True)
    response.headers["Cache-Control"] = "no-cache"
    return response


@screenspace_bp.route("/api/video/audio-track/<participant>/<int:idx>")
def api_video_audio_track(participant: str, idx: int) -> FlaskResponse:
    """Stream one demuxed audio track for the browser's per-track volume mixer."""
    paths = _participant_video_paths(participant)
    if not paths:
        return err_no_video(participant)
    part = request.args.get("part", type=int)
    video_path = (
        paths[part] if part is not None and 0 <= part < len(paths) else paths[0]
    )
    out = video.extract_audio_track(video_path, idx)
    if out is None:
        return err("Could not extract audio track", 500)
    response = send_file(str(out), mimetype="audio/mp4", conditional=True)
    response.headers["Cache-Control"] = "no-cache"
    return response


# ---- Region coordinate normalization ----


def _normalize_region(
    x: float,
    y: float,
    w: float,
    h: float,
    frame_w: int,
    frame_h: int,
    points: list[list[list[float]]] | None = None,
    shape: str | None = None,
) -> dict[str, Any]:
    """Convert pixel coordinates to normalized 0-1 fractions.

    For shaped regions, *points* is a list of polygon contours whose vertices
    are canvas-pixel ``[x, y]`` pairs; the bbox is recomputed from them (the
    caller's x/y/w/h are ignored) so bbox and points can never disagree, and
    the stored ``points`` contours are bbox-relative 0-1 fractions so
    move/resize only ever touch the bbox.
    """
    if points:
        xs = [p[0] for contour in points for p in contour]
        ys = [p[1] for contour in points for p in contour]
        x, y = min(xs), min(ys)
        w, h = max(xs) - x, max(ys) - y
    region: dict[str, Any] = {
        "x": x / frame_w,
        "y": y / frame_h,
        "w": w / frame_w,
        "h": h / frame_h,
        "source_width": frame_w,
        "source_height": frame_h,
    }
    if points:
        region["points"] = [
            [[round((px - x) / w, 4), round((py - y) / h, 4)] for px, py in contour]
            for contour in points
        ]
        region["shape"] = shape
    return region


# "combo" marks shapes produced by boolean region edits (shift-add /
# alt-subtract / merge) rather than a single drawing tool.
_REGION_SHAPES = ("lasso", "wand", "combo")
_REGION_MAX_POINTS = 400  # total vertices across all contours
_REGION_MAX_CONTOURS = 32
_REGION_MIN_BBOX_PX = 5
_REGION_MIN_AREA_PX = 64


def _validate_region_points(
    points: Any, shape: Any, canvas_w: float, canvas_h: float
) -> str | None:
    """Validate a shaped-region create payload; return an error string or None.

    Points arrive as a list of contours, each a list of canvas-pixel ``[x, y]``
    pairs. The bbox and shoelace-area guards reject shapes too small to
    rasterize meaningfully (the mask counterpart of the client's >5x5 rect
    minimum — rects keep their client-side check, so this is the only min-area
    gate in the system). The area guard sums per-contour areas: contours are
    disjoint by construction (client-side mask tracing), so the sum is the
    shape's true area.
    """
    if shape not in _REGION_SHAPES:
        return f"'shape' must be one of {', '.join(_REGION_SHAPES)}"
    contours_err = (
        f"'points' must be a list of 1-{_REGION_MAX_CONTOURS} contours, "
        f"each a list of 3+ [x, y] number pairs ({_REGION_MAX_POINTS} points total)"
    )
    if not isinstance(points, list) or not (1 <= len(points) <= _REGION_MAX_CONTOURS):
        return contours_err
    total = 0
    for contour in points:
        if not isinstance(contour, list) or len(contour) < 3:
            return contours_err
        total += len(contour)
        for p in contour:
            if (
                not isinstance(p, (list, tuple))
                or len(p) != 2
                or not all(isinstance(v, (int, float)) for v in p)
            ):
                return contours_err
    if total > _REGION_MAX_POINTS:
        return contours_err
    all_xs: list[float] = []
    all_ys: list[float] = []
    area = 0.0
    for contour in points:
        xs = [min(max(float(p[0]), 0.0), float(canvas_w)) for p in contour]
        ys = [min(max(float(p[1]), 0.0), float(canvas_h)) for p in contour]
        all_xs.extend(xs)
        all_ys.extend(ys)
        contour_area = 0.0
        for i in range(len(xs)):
            j = (i + 1) % len(xs)
            contour_area += xs[i] * ys[j] - xs[j] * ys[i]
        area += abs(contour_area) / 2.0
    if (
        max(all_xs) - min(all_xs) <= _REGION_MIN_BBOX_PX
        or max(all_ys) - min(all_ys) <= _REGION_MIN_BBOX_PX
    ):
        return "Region shape is too small"
    if area < _REGION_MIN_AREA_PX:
        return "Region shape is too small"
    return None


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
    """Resolve a task region request without flattening active/stashed duplicates.

    Thin Flask wrapper over the shared, pure ``screenspace.resolve_region_request`` so the
    server and the CLI re-run path can never drift on region semantics.
    """
    try:
        return screenspace.resolve_region_request(region_name, region_ref, _manifest)
    except ValueError as exc:
        return err(str(exc))


def _is_flask_error_response(value: Any) -> TypeGuard[FlaskResponse]:
    return isinstance(value, Response) or (
        isinstance(value, tuple) and len(value) == 2 and isinstance(value[1], int)
    )


# ---- Regions CRUD ----


@screenspace_bp.route("/api/regions")
def api_regions_list() -> FlaskResponse:
    """List all saved region definitions."""
    with _manifest_lock:
        regions = copy.deepcopy(_manifest.get("regions", {}))
    return ok(regions=regions)


@screenspace_bp.route("/api/regions", methods=["POST"])
def api_regions_create() -> FlaskResponse:
    """Create or update a named region."""
    data = request.get_json(silent=True)
    if not data:
        return err("JSON body required")

    name = data.get("name", "").strip()
    if not name:
        return err("Region name is required")

    for field in ("x", "y", "w", "h"):
        val = data.get(field)
        if val is None or not isinstance(val, (int, float)):
            return err(f"'{field}' must be a number")

    canvas_w = data.get("canvas_width")
    canvas_h = data.get("canvas_height")
    if (
        not isinstance(canvas_w, (int, float))
        or not isinstance(canvas_h, (int, float))
        or canvas_w <= 0
        or canvas_h <= 0
    ):
        return err("'canvas_width' and 'canvas_height' must be positive numbers")

    points = data.get("points")
    shape = data.get("shape")
    if points is not None or shape is not None:
        error = _validate_region_points(points, shape, canvas_w, canvas_h)
        if error:
            return err(error)
        points = [
            [
                [
                    min(max(float(p[0]), 0.0), float(canvas_w)),
                    min(max(float(p[1]), 0.0), float(canvas_h)),
                ]
                for p in contour
            ]
            for contour in points
        ]

    region = _normalize_region(
        data["x"],
        data["y"],
        data["w"],
        data["h"],
        int(canvas_w),
        int(canvas_h),
        points=points,
        shape=shape,
    )
    if "description" in data:
        region["description"] = str(data["description"])

    with _manifest_lock:
        _manifest.setdefault("regions", {})[name] = region
        _do_persist(drain_events=False)

    return ok(region=region)


@screenspace_bp.route("/api/regions/<name>", methods=["DELETE"])
def api_regions_delete(name: str) -> FlaskResponse:
    """Delete a region definition."""
    with _manifest_lock:
        regions = _manifest.get("regions", {})
        if name not in regions:
            return err(f"Region '{name}' not found", 404)
        del regions[name]
        _do_persist(drain_events=False)

    return ok()


@screenspace_bp.route("/api/regions", methods=["DELETE"])
def api_regions_delete_all() -> FlaskResponse:
    """Delete every active region. Stashes are left untouched."""
    with _manifest_lock:
        _manifest["regions"] = {}
        _do_persist(drain_events=False)

    return ok()


@screenspace_bp.route("/api/regions/reorder", methods=["PUT"])
def api_regions_reorder() -> FlaskResponse:
    """Reorder active regions to match the given name order."""
    data = request.get_json(silent=True)
    if not data or not isinstance(data.get("names"), list):
        return err("names list required")

    with _manifest_lock:
        regions = _manifest.get("regions", {})
        names = data["names"]
        if not all(isinstance(name, str) for name in names):
            return err("names must be strings")
        region_names = list(regions.keys())
        if len(names) != len(region_names) or set(names) != set(region_names):
            return err("names must match current regions exactly")
        _manifest["regions"] = {name: regions[name] for name in names}
        _do_persist(drain_events=False)

    return ok()


# ---- Stashes CRUD ----


@screenspace_bp.route("/api/stashes")
def api_stashes_list() -> FlaskResponse:
    """List all region stashes."""
    with _manifest_lock:
        stashes = copy.deepcopy(_manifest.get("stashes", []))
    return ok(stashes=stashes)


@screenspace_bp.route("/api/stashes", methods=["POST"])
def api_stashes_create() -> FlaskResponse:
    """Stash all current regions and clear the active set."""
    with _manifest_lock:
        regions = _manifest.get("regions", {})
        if not regions:
            return err("No regions to stash")

        stash = {
            "id": "stash_" + uuid.uuid4().hex[:8],
            "name": "Stashed Regions",
            "createdAt": datetime.now(UTC).isoformat(),
            "regions": copy.deepcopy(regions),
        }
        _manifest.setdefault("stashes", []).append(stash)
        _manifest["regions"] = {}
        _do_persist(drain_events=False)
    return ok(stash=stash)


@screenspace_bp.route("/api/stashes/<stash_id>", methods=["PUT"])
def api_stashes_update(stash_id: str) -> FlaskResponse:
    """Update a stash (rename)."""
    data = request.get_json(silent=True)
    if not data:
        return err("JSON body required")

    name = data.get("name", "").strip()
    with _manifest_lock:
        stashes = _manifest.get("stashes", [])
        stash = find_by_id(stashes, stash_id)
        if stash is None:
            return err("Stash not found", 404)
        if name:
            stash["name"] = name
        _do_persist(drain_events=False)
    return ok(stash=stash)


@screenspace_bp.route("/api/stashes/<stash_id>", methods=["DELETE"])
def api_stashes_delete(stash_id: str) -> FlaskResponse:
    """Dismiss a region stash."""
    with _manifest_lock:
        if remove_by_id(_manifest.get("stashes", []), stash_id) is None:
            return err("Stash not found", 404)
        _do_persist(drain_events=False)
    return ok()


@screenspace_bp.route("/api/stashes/<stash_id>/restore", methods=["POST"])
def api_stashes_restore(stash_id: str) -> FlaskResponse:
    """Restore a stash: replace active regions with stashed ones (stash is kept)."""
    with _manifest_lock:
        stash = find_by_id(_manifest.get("stashes", []), stash_id)
        if stash is None:
            return err("Stash not found", 404)
        _manifest["regions"] = copy.deepcopy(stash["regions"])
        _do_persist(drain_events=False)
    return ok(regions=_manifest["regions"])


@screenspace_bp.route("/api/stashes/<stash_id>/regions", methods=["POST"])
def api_stashes_add_region(stash_id: str) -> FlaskResponse:
    """Copy one active region into an existing stash (active set unchanged).

    If the stash already holds a region with that name, the active definition
    overwrites it (last-write-wins, matching api_regions_create's upsert).
    """
    data = request.get_json(silent=True)
    if not data:
        return err("JSON body required")
    name = data.get("name", "").strip()
    if not name:
        return err("Region name is required")

    with _manifest_lock:
        regions = _manifest.get("regions", {})
        if name not in regions:
            return err(f"Region '{name}' not found", 404)
        stash = find_by_id(_manifest.get("stashes", []), stash_id)
        if stash is None:
            return err("Stash not found", 404)
        stash.setdefault("regions", {})[name] = copy.deepcopy(regions[name])
        _do_persist(drain_events=False)

    return ok(stash=stash)


# ---- Tasks CRUD ----


@screenspace_bp.route("/api/tasks/stream")
def api_tasks_stream() -> FlaskResponse:
    """SSE endpoint for live task updates (replaces polling)."""
    return _sse_stream(_sse_task_payload)


@screenspace_bp.route("/api/tasks")
def api_tasks_list() -> FlaskResponse:
    """List all tasks with status and progress."""
    # Polling fallback for the SSE stream — slim, same as _sse_task_payload.
    tasks = _worker.get_all_tasks(include_results=False) if _worker else []
    clean = [_clean_task(t) for t in tasks]
    paused = _worker.is_paused if _worker else False
    alive = _worker.is_alive if _worker else False
    return ok(tasks=clean, paused=paused, worker_alive=alive)


@screenspace_bp.route("/api/tasks/<task_id>")
def api_tasks_get(task_id: str) -> FlaskResponse:
    """Get task detail including results."""
    if not _worker:
        return err("Worker not initialized", 500)
    task = _worker.get_task(task_id)
    if task is None:
        return err("Task not found", 404)
    return ok(task=_clean_task(task))


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
        return err(f"type must be one of: {', '.join(_VALID_TASK_TYPES)}")

    participant = data.get("participant", "").strip()
    if not participant:
        return err("participant is required")

    region_name = data.get("region", "").strip()
    region_ref = data.get("region_ref")
    raw_parameters = data.get("parameters")
    if raw_parameters is None:
        parameters: dict[str, Any] = {}
    elif isinstance(raw_parameters, dict):
        parameters = raw_parameters
    else:
        return err("parameters must be an object")

    # Template/shape tasks with an uploaded image scan the full frame; no
    # region needed.
    has_uploaded_template = (
        task_type == "template" and parameters.get("template_image_data")
    ) or (task_type == "shape" and parameters.get("shape_image_data"))

    # Multitool uses per-step regions; others need a global region (unless
    # template upload). Boundary always arrives with a forced full_frame
    # region_ref (set in api_tasks_create), so it satisfies this check.
    has_region_request = bool(region_name) or region_ref is not None
    if (
        not has_region_request
        and not has_uploaded_template
        and task_type != "multitool"
    ):
        return err("region is required")

    # Early validation for multitool steps
    if task_type == "multitool":
        mt_steps = parameters.get("steps")
        if not mt_steps or not isinstance(mt_steps, list) or len(mt_steps) < 2:
            return err("Multitool requires at least 2 steps")
        for i, step_raw in enumerate(mt_steps):
            if not isinstance(step_raw, dict):
                return err(f"Step {i}: must be an object")
            step_v = cast(dict[str, Any], step_raw)
            stype = step_v.get("type", "")
            if stype not in _VALID_STEP_TYPES:
                return err(f"Step {i}: invalid type '{stype}'")
            logic = step_v.get("logic")
            if logic is not None and logic not in ("AND", "NOT"):
                return err(f"Step {i}: logic must be 'AND' or 'NOT'")
            offset = step_v.get("offset")
            if offset is not None:
                if i == 0:
                    return err("Step 0: offset is not allowed on the first step")
                if (
                    not isinstance(offset, dict)
                    or offset.get("min") is None
                    or offset.get("max") is None
                ):
                    return err(f"Step {i}: offset requires numeric min and max")

    all_known_regions = _combined_region_lookup()
    requested_region: dict[str, Any] | None = None

    # Validate regions
    if task_type == "multitool":
        mt_steps_early: list[dict[str, Any]] = parameters.get("steps", [])
        for i, step in enumerate(mt_steps_early):
            step_region = (step.get("region") or "").strip()
            step_region_ref = step.get("region_ref")
            # A template step with an uploaded image scans the full frame and
            # needs no region (mirrors the top-level has_uploaded_template path).
            step_uploaded_template = step.get("type") == "template" and step.get(
                "template_image_data"
            )
            if not step_region and step_region_ref is None:
                if step_uploaded_template:
                    continue
                return err(f"Step {i}: region is required")
            if step_region_ref is not None:
                resolved = _resolve_region_request(step_region, step_region_ref)
                if _is_flask_error_response(resolved):
                    return resolved
            elif step_region not in all_known_regions:
                return err(f"Step {i}: region '{step_region}' not found")
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


def _coerce_tool_spec(spec: dict[str, Any], tool_type: str, context: str = "") -> None:
    """Coerce one tool spec's type-specific fields (task params or a multitool step).

    Mutates *spec* in place; raises ``ValueError`` on bad input. *context* prefixes
    error messages (e.g. ``"Step 0: "``). Shared by the task-level and per-step paths.
    """
    if (
        tool_type == "similarity"
        or tool_type == "template"
        and not spec.get("template_image_data")
        or tool_type == "shape"
        and not spec.get("shape_image_data")
    ):
        spec["reference_timestamp"] = _coerce_float(
            spec.get("reference_timestamp"),
            "reference_timestamp",
            required=True,
            context=context,
        )
    elif tool_type == "scene":
        spec["scene_references"] = _validate_scene_references(
            spec.get("scene_references"),
            context=context,
        )
    elif tool_type in ("text", "numbers"):
        _coerce_ocr_controls(spec, context=context)
    elif tool_type == "color":
        _coerce_color_controls(spec, context=context)
    if tool_type == "template":
        _coerce_template_controls(spec)
    elif tool_type == "shape":
        _coerce_shape_controls(spec, context=context)


def _coerce_task_params(
    task_type: str, parameters: dict[str, Any]
) -> dict[str, Any] | FlaskResponse:
    """Coerce and validate type-specific parameter values.

    Returns the updated parameters on success, or a Flask error response on failure.
    """
    try:
        if task_type == "multitool":
            for i, step in enumerate(parameters.get("steps", [])):
                step_context = f"Step {i}: "
                _coerce_tool_spec(step, step.get("type", ""), context=step_context)
                _coerce_offset(step, context=step_context)
        else:
            _coerce_tool_spec(parameters, task_type)
            if task_type in ("text", "numbers", "change", "flow"):
                _coerce_consecutive(parameters)
    except ValueError as exc:
        return err(str(exc))

    return parameters


def _extract_tool_media(
    spec: dict[str, Any],
    tool_type: str,
    frame_at: Callable[[float], "Any | None"],
    region_coords: dict[str, Any],
    context: str = "",
) -> None | FlaskResponse:
    """Extract the reference frame / template image into one spec.

    Mutates *spec* in place (task params or a multitool step). *frame_at* maps a
    GLOBAL reference timestamp into the owning sub-video. *context* prefixes error
    messages (e.g. ``"Step 0: "``). Returns a Flask error response on failure, else
    ``None``. Shared by the task-level and per-step multitool paths.
    """
    if tool_type == "similarity":
        ref_ts = cast(float, spec["reference_timestamp"])
        frame = frame_at(float(ref_ts))
        if frame is None:
            return err(f"{context}could not read reference frame")
        spec["reference_frame"] = screenspace.extract_region(frame, region_coords)

    elif tool_type == "template":
        upload_b64 = spec.pop("template_image_data", None)
        if upload_b64:
            try:
                bgr, mask = _template_bgr_and_mask_from_b64(upload_b64)
            except ValueError:
                return err(f"{context}could not decode uploaded image")
            spec["template_image"] = bgr
            if mask is not None:
                spec["template_mask"] = mask
        else:
            ref_ts = cast(float, spec["reference_timestamp"])
            frame = frame_at(float(ref_ts))
            if frame is None:
                return err(f"{context}could not read template frame")
            spec["template_image"] = screenspace.extract_region(frame, region_coords)
            # A shaped capture region doubles as the template's alpha mask.
            tpl_mask = screenspace.region_mask_for(
                region_coords, *spec["template_image"].shape[:2]
            )
            if tpl_mask is not None:
                spec["template_mask"] = tpl_mask

    elif tool_type == "shape":
        upload_b64 = spec.pop("shape_image_data", None)
        if upload_b64:
            try:
                bgr, mask = _template_bgr_and_mask_from_b64(upload_b64)
            except ValueError:
                return err(f"{context}could not decode uploaded image")
            spec["shape_image"] = bgr
            if mask is not None:
                spec["shape_mask"] = mask
        else:
            ref_ts = cast(float, spec["reference_timestamp"])
            frame = frame_at(float(ref_ts))
            if frame is None:
                return err(f"{context}could not read shape reference frame")
            spec["shape_image"] = screenspace.extract_region(frame, region_coords)
            # A shaped capture region doubles as the reference mask: inner
            # content (e.g. a button's label) can be lassoed out of the match.
            crop_mask = screenspace.region_mask_for(
                region_coords, *spec["shape_image"].shape[:2]
            )
            if crop_mask is not None:
                spec["shape_mask"] = crop_mask

    elif tool_type == "scene":
        scene_refs = cast(list[dict[str, Any]], spec["scene_references"])
        reference_scenes = []
        for ref in scene_refs:
            frame = frame_at(float(ref["timestamp"]))
            if frame is None:
                return err(f"{context}could not read frame for scene '{ref['name']}'")
            ref_region = screenspace.extract_region(frame, region_coords)
            scene_entry: dict = {"name": ref["name"], "frame": ref_region}
            if "threshold" in ref:
                scene_entry["threshold"] = float(ref["threshold"])
            reference_scenes.append(scene_entry)
        spec["reference_scenes"] = reference_scenes

    return None


def _extract_task_media(
    task_type: str,
    parameters: dict[str, Any],
    frame_at: Callable[[float], "Any | None"],
    region_coords: dict[str, Any],
) -> dict[str, Any] | FlaskResponse:
    """Extract reference frames / template images for non-multitool tasks.

    *frame_at* maps a GLOBAL reference timestamp into the owning sub-video and
    returns the frame (single-video participants extract unchanged). Returns the
    updated parameters on success, or a Flask error response on failure.
    """
    error = _extract_tool_media(parameters, task_type, frame_at, region_coords)
    if error is not None:
        return error
    return parameters


def _prepare_multitool_steps(
    parameters: dict[str, Any],
    all_known_regions: dict[str, Any],
    frame_at: Callable[[float], "Any | None"],
    region_coords: dict[str, Any],
    resolve_region_fn: Any,
) -> dict[str, Any] | FlaskResponse:
    """Resolve per-step regions and extract media for multitool tasks.

    *frame_at* maps a GLOBAL reference timestamp into the owning sub-video.
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

        error = _extract_tool_media(
            step, stype, frame_at, step_rc, context=f"Step {i}: "
        )
        if error is not None:
            return error

    return parameters


def _region_coords_resolver(
    props: dict[str, Any] | None, all_known_regions: dict[str, Any]
) -> Callable[..., dict[str, Any]]:
    """Return a closure that converts a named/inline region to pixel coords.

    When *region_data* is omitted the region is looked up by name in
    *all_known_regions*. Shared by the task-creation and calibration routes.
    """

    def _resolve(
        name: str, region_data: dict[str, Any] | None = None
    ) -> dict[str, Any]:
        rd = region_data if region_data is not None else all_known_regions[name]
        if props and props.get("width") and props.get("height"):
            return screenspace.denormalize_region(rd, props["width"], props["height"])
        coords = {k: rd[k] for k in ("x", "y", "w", "h") if k in rd}
        # Bbox-relative polygon points survive without denormalization.
        if rd.get("points"):
            coords["mask_points"] = rd["points"]
            if rd.get("shape"):
                coords["shape"] = rd["shape"]
        return coords

    return _resolve


def _prepare_task_media(
    task_type: str,
    participant: str,
    parameters: dict[str, Any],
    all_known_regions: dict[str, Any],
    region_coords: dict[str, Any],
    resolve_region_fn: Callable[..., dict[str, Any]],
) -> dict[str, Any] | FlaskResponse:
    """Coerce params, extract reference media, and resolve multitool steps.

    Shared preparation pipeline for the task-creation and calibration routes: each
    resolves ``region_coords`` its own way (their region-naming rules differ) then
    hands off here. Returns the enriched parameters, or a Flask error response.
    """
    coerced = _coerce_task_params(task_type, parameters)
    if isinstance(coerced, tuple):
        return coerced
    parameters = cast(dict[str, Any], coerced)

    frame_at = _participant_frame_extractor(participant)
    # Template/shape split the sample from the search scope: the reference is
    # cut from the *capture* region (``reference_region``, persisted so re-runs
    # keep extracting the same patch) while ``region_coords`` stays the run
    # window. Without it a Full-frame run would extract the whole frame as its
    # sample.
    extract_coords = region_coords
    if task_type in ("template", "shape"):
        ref_region = str(parameters.get("reference_region") or "").strip()
        if ref_region:
            parameters["reference_region"] = ref_region
            try:
                _ref_name, ref_norm = screenspace.resolve_region_request(
                    ref_region, None, _manifest
                )
            except ValueError as exc:
                return err(f"reference_region: {exc}")
            extract_coords = resolve_region_fn(ref_region, ref_norm)
        else:
            parameters.pop("reference_region", None)
    extracted = _extract_task_media(task_type, parameters, frame_at, extract_coords)
    if isinstance(extracted, tuple):
        return extracted
    parameters = cast(dict[str, Any], extracted)

    if task_type == "multitool":
        prepared = _prepare_multitool_steps(
            parameters, all_known_regions, frame_at, region_coords, resolve_region_fn
        )
        if isinstance(prepared, tuple):
            return prepared
        parameters = cast(dict[str, Any], prepared)

    return parameters


@screenspace_bp.route("/api/tasks", methods=["POST"])
def api_tasks_create() -> FlaskResponse:
    """Enqueue a new analysis task."""
    if not _worker:
        return err("Worker not initialized", 500)

    data = request.get_json(silent=True)
    if not data:
        return err("JSON body required")

    # Boundary and Attention are full-frame only by contract: ignore any
    # caller-supplied region so events and manifest metadata are never labeled
    # with a region the scan didn't use (both scanners always analyze the whole
    # frame). Forcing it here — before validation and before task["region_ref"]
    # is recorded — keeps the stored region_name/coords and region_ref
    # consistently full-frame.
    if (data.get("type") or "").strip() in ("boundary", "attention"):
        data["region"] = ""
        data["region_ref"] = {"source": "full_frame"}

    validated = _validate_task_request(data)
    if isinstance(validated, Response) or (
        isinstance(validated, tuple) and len(validated) == 2
    ):
        # cast() needed for older ty (<=0.0.33) which doesn't narrow `len == 2`;
        # newer ty flags this as redundant but only as a non-failing warning.
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

    video_paths = _participant_video_paths(participant)
    if not video_paths:
        return err_no_video(participant, 400)
    video_path = video_paths[0]

    # Default per-task source label; the per-part scan tags each event with the
    # specific sub-video it came from for multi-video participants.
    source_video = Path(video_paths[0]).name

    # Region coords come from the first part's resolution (parts share it).
    props = video.probe_video_properties(video_path)
    resolve_region = _region_coords_resolver(props, all_known_regions)

    if requested_region is not None:
        region_coords = resolve_region(region_name, requested_region)
    elif task_type == "multitool":
        first_step_region = parameters.get("steps", [{}])[0].get("region", "")
        if first_step_region and first_step_region in all_known_regions:
            region_name = first_step_region
            region_coords = resolve_region(first_step_region)
        else:
            region_name = "per_step"
            region_coords = {"x": 0, "y": 0, "w": 0, "h": 0}
    else:
        region_name = "full_frame"
        region_coords = {"x": 0, "y": 0, "w": 0, "h": 0}

    prepared = _prepare_task_media(
        task_type,
        participant,
        parameters,
        all_known_regions,
        region_coords,
        resolve_region,
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
        video_paths=video_paths,
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
    _schedule_persist()
    _notify_sse_clients("task_created")

    return ok(task=_clean_task(task))


@screenspace_bp.route("/api/tasks/<task_id>", methods=["DELETE"])
def api_tasks_cancel(task_id: str) -> FlaskResponse:
    """Cancel or dismiss a task.  ?dismiss=true fully removes the task."""
    if not _worker:
        return err("Worker not initialized", 500)
    if request.args.get("dismiss") == "true":
        if not _worker.remove_task(task_id):
            return err("Task not found", 404)
        _schedule_persist()
        _notify_sse_clients("task_dismissed")
        return ok()
    if _worker.cancel(task_id):
        _schedule_persist()
        _notify_sse_clients("task_cancelled")
        return ok()
    return err("Task not found or already finished")


@screenspace_bp.route("/api/tasks/reorder", methods=["PUT"])
def api_tasks_reorder() -> FlaskResponse:
    """Reorder queued tasks by priority."""
    if not _worker:
        return err("Worker not initialized", 500)
    data = request.get_json(silent=True)
    if not data or "task_ids" not in data:
        return err("task_ids list required")
    _worker.reorder(data["task_ids"])
    _schedule_persist()
    _notify_sse_clients("reorder")
    return ok()


@screenspace_bp.route("/api/tasks/pause", methods=["POST"])
def api_tasks_pause() -> FlaskResponse:
    """Pause the task queue."""
    if not _worker:
        return err("Worker not initialized", 500)
    _worker.pause()
    _schedule_persist()
    _notify_sse_clients("pause")
    return ok(paused=True)


@screenspace_bp.route("/api/tasks/resume", methods=["POST"])
def api_tasks_resume() -> FlaskResponse:
    """Resume the task queue."""
    if not _worker:
        return err("Worker not initialized", 500)
    _worker.resume()
    _schedule_persist()
    _notify_sse_clients("resume")
    return ok(paused=False)


@screenspace_bp.route("/api/tasks/<task_id>/results")
def api_tasks_results(task_id: str) -> FlaskResponse:
    """Get task results (timestamps, artifacts).

    ``?since=N`` returns only the result tail beyond index N (results are
    append-only during a scan), letting the frontend stream live results for a
    running task without re-fetching the whole list each tick. Omit for the full
    list (completion load). Always reports ``total`` for the next cursor.
    """
    if not _worker:
        return err("Worker not initialized", 500)
    try:
        since = int(request.args.get("since", 0))
    except (TypeError, ValueError):
        since = 0
    out = _worker.get_task_result_tail(task_id, since)
    if out is None:
        return err("Task not found", 404)
    results, total = out
    return ok(results=utils.sanitize_floats(results), total=total)


# ---- Events CRUD ----


def _apply_event_filters(
    events: list[dict[str, Any]], args: Any
) -> list[dict[str, Any]]:
    """Apply the standard excluded/participant/task_id query filters (pure)."""
    excluded_filter = args.get("excluded")
    if excluded_filter == "false":
        events = [e for e in events if not e.get("excluded")]
    elif excluded_filter == "true":
        events = [e for e in events if e.get("excluded")]
    participant = args.get("participant")
    if participant:
        events = [e for e in events if e.get("participant") == participant]
    task_id = args.get("task_id")
    if task_id:
        events = [e for e in events if e.get("task_id") == task_id]
    return events


def _events_payload(
    args: Any, client_version: int | None
) -> tuple[int, list[dict[str, Any]] | None]:
    """Version-aware events snapshot for the poll routes.

    Reads the events-version and deep-copies under a single lock acquisition (so
    a concurrent bump can't slip between the two). Returns ``(current_version,
    events)``; ``events`` is ``None`` — meaning "unchanged, skip the payload" —
    only when the caller passed a ``client_version`` equal to the current one.
    """
    with _manifest_lock:
        current = _events_version
        if client_version is not None and client_version == current:
            return current, None
        events = copy.deepcopy(_manifest.get("events", []))
    return current, utils.sanitize_floats(_apply_event_filters(events, args))


def _client_events_version(args: Any) -> int | None:
    """Parse the optional ``events_version`` poll cursor; ``None`` if absent/bad."""
    raw = args.get("events_version")
    if raw is None:
        return None
    try:
        return int(raw)
    except (TypeError, ValueError):
        return None


@screenspace_bp.route("/api/events")
def api_events_list() -> FlaskResponse:
    """List events with optional filtering.

    ``?events_version=N`` short-circuits: when N matches the current events
    version, returns ``events_unchanged`` with no payload. Omit for the full list.
    """
    version, events = _events_payload(
        request.args, _client_events_version(request.args)
    )
    if events is None:
        return ok(events_version=version, events_unchanged=True)
    return ok(events_version=version, events=events)


@screenspace_bp.route("/api/intake-poll")
def api_intake_poll() -> FlaskResponse:
    """Combined Studio-intake poll: task-status booleans + filtered events.

    Collapses the Studio intake client's separate /api/tasks and /api/events
    polls into a single request. Both reads are the same slim, in-memory ones
    those routes do (no results, no disk I/O). ``?events_version=N`` skips the
    events payload when nothing changed since the client's last tick."""
    tasks = [
        _clean_task(t)
        for t in (_worker.get_all_tasks(include_results=False) if _worker else [])
    ]
    running = any(t.get("status") == "running" for t in tasks)
    queued = any(t.get("status") == "queued" for t in tasks)
    alive = _worker.is_alive if _worker else False
    status = {"running": running, "worker_alive": alive, "queued": queued}
    version, events = _events_payload(
        request.args, _client_events_version(request.args)
    )
    if events is None:
        return ok(status=status, events_version=version, events_unchanged=True)
    return ok(status=status, events_version=version, events=events)


@screenspace_bp.route("/api/export/events")
def api_export_events() -> FlaskResponse:
    """Export Screenspace events as analysis-ready JSON or CSV.

    Both formats return a downloadable file body (not the ``ok()`` envelope):
    CSV rows, or the ``{exported_at, version, records}`` JSON envelope
    :func:`data_export.to_json` also writes for the bundle export.

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

    with _manifest_lock:
        manifest_snapshot = copy.deepcopy(_manifest)
    records = data_export.build_screenspace_events(
        manifest_snapshot,
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
        # Same {exported_at, version, records} envelope the bundle export writes,
        # so a saved screenspace_events.json is self-describing rather than an
        # API response that happened to be written to disk.
        body = data_export.to_json(utils.sanitize_floats(records))
        response = Response(body, mimetype="application/json")
        response.headers["Content-Disposition"] = (
            'attachment; filename="screenspace_events.json"'
        )
        return response
    return err(f"Unsupported format: {fmt}")


def _set_event_excluded(event_id: str, excluded: bool) -> FlaskResponse:
    """Set one event's excluded flag; 404 when the id is unknown."""
    with _manifest_lock:
        for e in _manifest.get("events", []):
            if e["id"] == event_id:
                e["excluded"] = excluded
                _bump_events_version()
                _do_persist(drain_events=False)
                return ok()
    return err("Event not found", 404)


def _bulk_set_excluded(excluded: bool) -> FlaskResponse:
    """Set the excluded flag on every event named in the request's id list."""
    data = request.get_json(silent=True) or {}
    ids = set(data.get("ids", []))
    if not ids:
        return err("ids list required")
    with _manifest_lock:
        count = 0
        for e in _manifest.get("events", []):
            if e["id"] in ids:
                e["excluded"] = excluded
                count += 1
        if count:
            _bump_events_version()
        _do_persist(drain_events=False)
    return ok(updated=count)


@screenspace_bp.route("/api/events/<event_id>/exclude", methods=["PUT"])
def api_event_exclude(event_id: str) -> FlaskResponse:
    return _set_event_excluded(event_id, True)


@screenspace_bp.route("/api/events/<event_id>/include", methods=["PUT"])
def api_event_include(event_id: str) -> FlaskResponse:
    return _set_event_excluded(event_id, False)


@screenspace_bp.route("/api/events/bulk-exclude", methods=["PUT"])
def api_events_bulk_exclude() -> FlaskResponse:
    return _bulk_set_excluded(True)


@screenspace_bp.route("/api/events/bulk-include", methods=["PUT"])
def api_events_bulk_include() -> FlaskResponse:
    return _bulk_set_excluded(False)


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


def _coerce_offset(step: dict[str, Any], *, context: str = "") -> None:
    """Validate a multitool step's offset window in place.

    A step's ``offset`` (idx > 0 only) declares a time window ``{min, max}`` in
    seconds, relative to the previous step's matched frame. Both bounds must be
    finite, ``min <= max``, and within ``±SCREENSPACE_MULTITOOL_MAX_OFFSET_SECONDS``
    (clamped). Absence of ``offset`` means same-frame matching (unchanged behavior).
    """
    offset = step.get("offset")
    if offset is None:
        return
    off_min = _coerce_float(
        offset.get("min"), "offset min", required=True, context=context
    )
    off_max = _coerce_float(
        offset.get("max"), "offset max", required=True, context=context
    )
    assert off_min is not None and off_max is not None  # required=True guarantees this
    if off_min > off_max:
        raise ValueError(f"{context}offset min must be <= max")
    bound = config.SCREENSPACE_MULTITOOL_MAX_OFFSET_SECONDS
    off_min = max(-bound, min(bound, off_min))
    off_max = max(-bound, min(bound, off_max))
    step["offset"] = {"min": off_min, "max": off_max}


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


def _coerce_shape_controls(params: dict[str, Any], *, context: str = "") -> None:
    """Validate shape-tool controls: the horizontal + optional vertical ladders."""
    for min_key, max_key, steps_key in (
        ("scale_min", "scale_max", "scale_steps"),
        ("scale_y_min", "scale_y_max", "scale_y_steps"),
    ):
        for key in (min_key, max_key):
            if params.get(key) is None:
                params.pop(key, None)
                continue
            val = _coerce_float(params[key], key, context=context)
            if val is None or val <= 0:
                raise ValueError(f"{context}{key} must be a positive number")
            params[key] = max(0.1, min(4.0, val))
        if (
            min_key in params
            and max_key in params
            and params[min_key] > params[max_key]
        ):
            raise ValueError(f"{context}{min_key} must not exceed {max_key}")
        if params.get(steps_key) is None:
            params.pop(steps_key, None)
            continue
        try:
            steps = int(params[steps_key])
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{context}{steps_key} must be an integer") from exc
        params[steps_key] = max(1, min(12, steps))


def _coerce_ocr_controls(params: dict[str, Any], *, context: str = "") -> None:
    """Validate optional OCR controls shared by Text and Numbers tools."""
    langs = params.get("languages")
    if langs is not None:
        # Closed set: each code maps onto a bundled recognition model
        # (screenspace_ocr._OCR_LANG_TO_MODEL). Unvalidated strings used to
        # reach the engine and die mid-scan; refuse at task creation instead.
        valid = tuple(screenspace._OCR_LANG_TO_MODEL)
        if (
            not isinstance(langs, list)
            or not langs
            or any(lang not in valid for lang in langs)
        ):
            raise ValueError(
                f"{context}languages must be a non-empty list drawn from {valid}"
            )
        # Known codes can still need two rec models (ja+ko). Refuse here so
        # the task never queues and fails mid-scan.
        try:
            screenspace._resolve_ocr_model([str(lang) for lang in langs])
        except ValueError as exc:
            raise ValueError(f"{context}{exc}") from exc
    if "ocr_confidence_threshold" not in params:
        return
    threshold = _coerce_float(
        params.get("ocr_confidence_threshold"),
        "ocr_confidence_threshold",
        required=True,
        context=context,
    )
    if threshold is None or threshold < 0 or threshold > 1:
        raise ValueError(f"{context}ocr_confidence_threshold must be between 0 and 1")
    params["ocr_confidence_threshold"] = threshold


def _coerce_consecutive(params: dict[str, Any], *, context: str = "") -> None:
    """Validate the optional require_consecutive control (Text/Numbers/Change/Flow).

    Clamps to [1, 10]; drops the key when it resolves to 1 (the default) so the
    manifest stays clean.
    """
    raw = params.get("require_consecutive")
    if raw is None:
        params.pop("require_consecutive", None)
        return
    try:
        count = int(raw)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{context}require_consecutive must be an integer") from exc
    count = max(1, min(10, count))
    if count == 1:
        params.pop("require_consecutive", None)
    else:
        params["require_consecutive"] = count


def _coerce_color_controls(params: dict[str, Any], *, context: str = "") -> None:
    """Validate the optional Color mode controls.

    ``color_mode`` is clamped to {"average", "presence"} (default "average",
    dropped to keep the manifest clean). ``min_coverage`` (presence-only minimum
    matching-pixel fraction) is clamped to [0, 1]; dropped when not in presence
    mode or when it resolves to 0 (the default).
    """
    mode = params.get("color_mode")
    if mode not in ("average", "presence"):
        params.pop("color_mode", None)
        params.pop("min_coverage", None)
        return
    if mode == "average":
        params.pop("color_mode", None)
        params.pop("min_coverage", None)
        return
    params["color_mode"] = "presence"
    raw = params.get("min_coverage")
    if raw is None:
        params.pop("min_coverage", None)
        return
    coverage = _coerce_float(raw, "min_coverage", required=True, context=context)
    if coverage is None or coverage < 0 or coverage > 1:
        raise ValueError(f"{context}min_coverage must be between 0 and 1")
    if coverage == 0:
        params.pop("min_coverage", None)
    else:
        params["min_coverage"] = coverage


def _validate_scene_references(
    scene_refs: Any, *, context: str = ""
) -> list[dict[str, Any]]:
    """Validate scene reference payloads and normalize numeric fields."""
    if not scene_refs or not isinstance(scene_refs, list):
        raise ValueError(f"{context}scene_references must be a non-empty list")
    validated_refs = []
    for i, ref_raw in enumerate(scene_refs):
        if not isinstance(ref_raw, dict):
            # ValueError, not TypeError: this validator's contract is that any bad
            # payload raises ValueError, which the calling route turns into a 400.
            raise ValueError(f"{context}scene_references[{i}] must be an object")  # noqa: TRY004
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


def _clean_task(task: dict[str, Any]) -> dict[str, Any]:
    """Remove internal fields from a task dict for API responses.

    Server-only ``change_grid`` per-frame data is already omitted upstream by
    ``ScreenspaceWorker.get_task``/``get_all_tasks`` (see ``_copy_task_for_read``).
    """
    cleaned = {k: v for k, v in task.items() if not k.startswith("_")}
    if "parameters" in cleaned:
        cleaned["parameters"] = screenspace.strip_task_param_binaries(
            cleaned["parameters"]
        )
    return utils.sanitize_floats(cleaned)


# ---- State initialization ----


def _backfill_missing_events(manifest: dict[str, Any]) -> None:
    """Heal manifests where completed tasks have results but no events.

    Why: events are generated only at task completion. Tasks completed before
    the events system existed (or whose events were lost) leave the frontend
    unable to render exclude toggles. Backfill so older results behave like
    new ones.
    """
    import screenspace

    events = manifest.setdefault("events", [])
    task_ids_with_events = {e.get("task_id") for e in events if e.get("task_id")}
    added = 0
    for task in manifest.get("tasks", []):
        if task.get("status") != screenspace.TASK_STATUS_COMPLETED:
            continue
        if task.get("id") in task_ids_with_events:
            continue
        result = task.get("result")
        if not isinstance(result, list) or not result:
            continue
        new_events = screenspace.generate_events_from_results(task, result)
        if new_events:
            events.extend(new_events)
            added += len(new_events)
    if added:
        screenspace.save_screenspace_manifest(
            manifest.get("regions", {}),
            manifest.get("tasks", []),
            events,
            stashes=manifest.get("stashes", []),
            per_participant=manifest.get("per_participant", {}),
            pins=manifest.get("pins") or {},
        )


def _init_screenspace_state(sheet_context: Any = None) -> None:
    """Initialize module-level state for Screenspace routes.

    Loads manifest, resolves participant video paths, and starts the
    background worker thread.

    Keeps a reference to *sheet_context* so :func:`_refresh_participants` can
    re-derive the participant union when the input directory changes.
    ``server._swap_worksheet`` re-inits this blueprint on every sheet swap, so
    the stored reference is replaced rather than going stale.
    """
    import screenspace

    global _manifest, _worker, _participant_source

    # Retire the previous session's worker before touching any state: its
    # callbacks resolve the module globals at call time, so left alive it
    # would persist the old study's task results into the *new* study's
    # manifest. Detach the callbacks and cancel first so the thread can only
    # finish inertly; the short join keeps a swap request from blocking on an
    # in-flight scan.
    if _worker is not None:
        _worker.on_task_complete = None
        _worker.on_progress_update = None
        _worker.cancel_all()
        _worker.stop(join_timeout=2.0)

    _manifest = screenspace.load_screenspace_manifest()
    _backfill_missing_events(_manifest)

    # mtime None forces the first _refresh_participants() call to build.
    _participant_source = {"sheet_context": sheet_context, "dir": "", "mtime": None}
    _refresh_participants()
    _participant_timeline_cache.clear()

    _worker = screenspace.ScreenspaceWorker()
    _worker.on_task_complete = _persist_manifest
    _worker.on_progress_update = lambda: _notify_sse_clients("progress")
    _worker.restore_tasks(_manifest.get("tasks", []))
    _worker.start()

    # Reclaim a stale empty manifest left by a prior abandoned session: the
    # guarded save removes the file when empty, idempotent rewrite otherwise.
    _persist_manifest()


def _do_persist(*, drain_events: bool = True) -> None:
    """Persist manifest to disk — caller must hold _manifest_lock."""
    import screenspace

    if _worker and drain_events:
        new_events = _worker.drain_new_events()
        if new_events:
            _manifest.setdefault("events", []).extend(new_events)
            _bump_events_version()
    tasks = _worker.get_all_tasks() if _worker else _manifest.get("tasks", [])
    _manifest["tasks"] = tasks
    screenspace.save_screenspace_manifest(
        _manifest.get("regions", {}),
        tasks,
        _manifest.get("events", []),
        stashes=_manifest.get("stashes", []),
        per_participant=_manifest.get("per_participant", {}),
        pins=_pin_manifest(),
    )


# Manifest-write debounce: rapid UI mutations (enqueue / cancel / reorder /
# pause / resume) coalesce into one disk write after a short quiet period.
# atexit fires the pending flush on normal exit / SIGINT / SIGTERM, but not
# on SIGKILL or hard power-loss — accepted because screenspace manifest is
# recreatable analysis state. The lambda looks up _do_persist at call time so
# tests monkeypatching it are seen.
(_schedule_persist, _flush_pending_persist, _cancel_pending_persist_timer) = (
    make_debounced_persist(lambda: _do_persist(drain_events=False), _manifest_lock)
)

atexit.register(_flush_pending_persist)


def _persist_manifest(*, drain_events: bool = True) -> None:
    """Persist the current manifest state through a single synchronized path.

    Synchronous callers (task completion, atexit flush) supersede any pending
    debounced write — cancel the timer here so we don't double-flush.
    """
    _cancel_pending_persist_timer()
    with _manifest_lock:
        _do_persist(drain_events=drain_events)
