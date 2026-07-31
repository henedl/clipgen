"""Transcripts Flask blueprint for clipgen.

Registered at /transcripts/ by start_combined_server(). Works with or without a
spreadsheet; auto-discovers participant videos from the input directory.
Module-level state: _manifest, _worker, _input_dir, _participants (initialized by
_init_transcripts_state()).

API endpoints (all under /transcripts/):
  GET  /media/<filename>                          - serve source video files
  GET  /api/participants                          - list discovered videos with transcription status
  GET  /api/transcript/<participant>              - full transcript segments (corrections applied)
  PUT  /api/transcript/<participant>/segment      - edit a segment's text, creates correction
  GET  /api/vtt/<participant>                     - serve transcript as WebVTT
  POST /api/embed-subtitle/<participant>          - mux participant transcript into a copy of their video
  POST /api/embed-all-subtitles                   - mux every participant's transcript into a subtitled copy of their video
  GET  /api/agent/<key>/<participant>            - a thinking agent's result (summary/citations/friction), status, or 404
  POST /api/agent/<key>/<participant>/regenerate - clear + re-trigger an agent (forces past its enabled config)
  POST /api/agent/<key>/<participant>/stop       - flag an in-flight agent run for discard
  GET  /api/agent/summary/<participant>/stream   - SSE token stream of the summary as it generates (summary-only)
  PUT  /api/agent/summary/<participant>          - save a user-edited summary (summary-only)
  GET  /api/corrections                           - list all study-local corrections
  POST /api/corrections                           - add a correction manually
  DELETE /api/corrections/<id>                    - remove a correction
  GET  /api/intake-poll                           - Studio-intake poll: running-state booleans + resolved marks
  GET  /api/marks                                 - list all marks with resolved segment data
  POST /api/marks                                 - create marks for segments
  PUT  /api/marks/<id>                            - update a mark's category or label
  DELETE /api/marks/<id>                          - remove a mark (or bulk delete with JSON body)
  GET  /api/search?q=<query>                      - keyword search across all participants
  POST /api/transcribe                            - enqueue participant(s) for transcription
  GET  /api/transcribe/status                     - poll transcription task status
  GET  /api/transcribe/<task_id>/segments         - running task's partial-segment tail (?since=N, append-only)
  DELETE /api/transcribe/<task_id>                 - cancel or dismiss a transcription task
  POST /api/transcribe/warmup                     - background-load Whisper when prewarm is enabled (confirms before downloading a non-cached model; force=true to proceed)
  GET  /api/transcribe/model-status               - whether the Whisper model is loaded or warming
  POST /api/models/ollama/pull                     - download (install) an Ollama model in the background
  GET  /api/models/ollama/pull-status              - poll progress of an in-flight Ollama model pull
"""

import atexit
import json
import os
import tempfile
import threading
import time
import uuid
from collections import defaultdict
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from flask import Blueprint, Response, jsonify, request, send_file, stream_with_context

import config
import files
import friction
import ollama_client
import thinking_agents
import transcripts
import utils
import video
from server_utils import err, make_debounced_persist, ok

FlaskResponse = Response | tuple[Response, int]

# ---- Module-level state (set once by _init_transcripts_state) ----

_manifest: dict[str, Any] = {}
_worker: transcripts.TranscriptWorker | None = None
_input_dir: str = ""
_participants: list[dict[str, Any]] = []
# What the current _participants list was built from: {"sheet_context", "dir",
# "mtime"}, or None before _init_transcripts_state has run (same "not configured
# yet" state _worker = None expresses). While None, _refresh_participants() is a
# no-op, so a directly-assigned _participants survives.
_participant_source: dict[str, Any] | None = None
_participants_lock = threading.Lock()
_manifest_lock = threading.Lock()
# Task ids whose result has already been merged into the in-memory manifest.
# Merging is idempotent per task so a later persist cannot re-apply a task's
# frozen segments over in-memory edits. Re-transcription mints a fresh task id,
# so its new segments still merge (and win) exactly once.
_merged_task_ids: set[str] = set()
# Participants whose transcript was freshly merged but whose completion side
# effects (clearing stale agent fields + starting the thinking-agent chain)
# have not run yet. The merge can be reached from either _on_task_complete or
# a debounced _do_persist — whichever wins _manifest_lock merges the task — so
# the reaction must drain this queue rather than key off which caller merged.
# Guarded by _manifest_lock.
_pending_chain_pids: list[str] = []
_transcript_model_warming = False
_transcript_model_warming_lock = threading.Lock()
# Thinking-agent orchestrator. Owns per-agent in-flight, cancel-event, and
# thread state. The `AgentOrchestrator` class is defined further down and the
# instance is bound at module load. Routes call methods on this; nothing else
# should reach into orchestrator internals.
_orchestrator: "AgentOrchestrator"

# After a Stop, schedule a delayed Ollama model unload so the model is evicted
# from memory if no follow-up run starts soon. Keyed by model name → Timer.
# A new run for the same model cancels the pending unload so we don't churn
# on rapid stop→run cycles.
_pending_model_unloads: dict[str, threading.Timer] = {}
_pending_model_unloads_lock = threading.Lock()

# Progress tracking for in-flight Ollama model pulls, keyed by model name.
# Each entry: {status, completed, total, done, succeeded, error}. The UI polls
# /api/models/ollama/pull-status after kicking off /api/models/ollama/pull so a
# new local model is only installed after explicit user confirmation.
_ollama_pull_status: dict[str, dict[str, Any]] = {}
_ollama_pull_lock = threading.Lock()


def _schedule_model_unload(model: str) -> None:
    """Schedule an Ollama model unload after ``config.OLLAMA_UNLOAD_DELAY_SECONDS``.

    Replaces any pending unload timer for the same model so the delay always
    measures from the most recent Stop.
    """
    delay = float(getattr(config, "OLLAMA_UNLOAD_DELAY_SECONDS", 15.0))
    if delay <= 0:
        ollama_client.unload_model(model)
        return

    def _unload() -> None:
        with _pending_model_unloads_lock:
            _pending_model_unloads.pop(model, None)
        ollama_client.unload_model(model)

    timer = threading.Timer(delay, _unload)
    timer.daemon = True
    with _pending_model_unloads_lock:
        existing = _pending_model_unloads.pop(model, None)
        if existing is not None:
            existing.cancel()
        _pending_model_unloads[model] = timer
    timer.start()


def _cancel_pending_unload(model: str) -> None:
    """Cancel any pending unload for *model* (because a new run is starting)."""
    with _pending_model_unloads_lock:
        timer = _pending_model_unloads.pop(model, None)
    if timer is not None:
        timer.cancel()


def _agent_model(agent_key: str) -> str | None:
    """Look up the Ollama model name configured for *agent_key*.

    A blank model knob means "inherit the summary model" (friction's default), so
    unload scheduling targets the model the agent actually loaded.
    """
    agent = thinking_agents.get_agent(agent_key)
    if agent is None:
        return None
    return thinking_agents.resolve_model(agent)


def _step_state_transcription(entry: dict[str, Any]) -> str:
    # Only the persisted result is known here; live running/queued/failed for
    # Whisper is merged in on the frontend from /api/transcribe/status.
    return "done" if entry.get("segments") else "idle"


def _step_state_agent(pid: str, entry: dict[str, Any], agent_key: str) -> str:
    field_name = next(
        (a["manifest_field"] for a in thinking_agents.AGENTS if a["key"] == agent_key),
        agent_key,
    )
    if _orchestrator.is_generating(pid, agent_key):
        return "running"
    if entry.get(field_name):
        return "done"
    return "idle"


def _mark_friction_stale(entry: dict[str, Any]) -> None:
    """Flag a participant's friction analysis stale after a transcript/summary edit.

    Friction's programmatic scores and LLM prompt both depend on segment text and
    the session summary, so any edit invalidates them. We flag rather than
    recompute — the user re-runs friction explicitly (no auto-rerun). Callers must
    hold ``_manifest_lock``.
    """
    fr = entry.get("friction")
    if isinstance(fr, dict):
        fr["stale"] = True


def _invalidate_dependents(entry: dict[str, Any], agent: thinking_agents.Agent) -> None:
    """Invalidate every agent whose result depends on *agent*'s field.

    Driven off ``depends_on`` + each dependent's ``on_upstream_change`` so this
    stays generic: ``"stale"`` keeps the result but flags it for a prompted
    re-run (friction), ``"clear"`` drops it (citations). Callers must hold
    ``_manifest_lock``.
    """
    for dep in thinking_agents.AGENTS:
        if agent["manifest_field"] not in dep["depends_on"]:
            continue
        if dep.get("on_upstream_change") == "stale":
            _mark_friction_stale(entry)  # only the friction shape has a stale flag
        else:
            entry.pop(dep["manifest_field"], None)


def _transcribe_prewarm_setting() -> str:
    """Return a validated TRANSCRIBE_PREWARM value for API clients."""
    v = getattr(config, "TRANSCRIBE_PREWARM", "queue_open")
    if v in ("off", "queue_open", "page_load"):
        return v
    return "queue_open"


# ---- Blueprint ----

transcripts_bp = Blueprint("transcripts", __name__)

utils.register_static_routes(
    transcripts_bp,
    "transcripts.html",
    media_dir_getter=lambda: _input_dir,
    media_error="Input directory not configured",
    icons=True,
)


# ---- Participants ----


def _refresh_participants() -> None:
    """Rebuild ``_participants`` when the input directory changed since the last build.

    Keyed on the input dir's ``st_mtime_ns`` (which advances on add/remove/rename),
    mirroring ``utils.discover_participant_videos``' own memo — the steady-state
    cost is one ``stat()``. This is what lets a video dropped into ``-i`` mid-session
    show up without a server restart.

    No-op until :func:`_init_transcripts_state` has configured the source. The
    rebuild rebinds ``_participants`` (atomic under the GIL), so a concurrent
    reader sees either the old list or the new one, never a torn one.
    """
    global _participants

    source = _participant_source
    if source is None:
        return
    input_dir = str(Path(utils.get_effective_input_dir()))
    try:
        mtime: int | None = Path(input_dir).stat().st_mtime_ns
    except OSError:
        mtime = None
    if source["dir"] == input_dir and source["mtime"] == mtime:
        return
    with _participants_lock:
        # A racing request may have rebuilt, or a sheet swap may have replaced the
        # source entirely, while we waited on the lock.
        if _participant_source is not source:
            return
        if source["dir"] == input_dir and source["mtime"] == mtime:
            return
        _participants = files.resolve_participant_videos(source["sheet_context"])
        source["dir"] = input_dir
        source["mtime"] = mtime


@transcripts_bp.route("/api/participants")
def api_participants() -> FlaskResponse:
    """List discovered source videos with transcription status."""
    import viewer

    _refresh_participants()
    result = []
    with _manifest_lock:
        src = _manifest.get("source_transcripts", {})
        for p in _participants:
            pid = p["id"]
            entry = src.get(pid, {})
            has_transcript = bool(entry.get("segments"))
            video_paths = p["video_paths"]
            first_path = video_paths[0]
            # Combine every part's mtime so the frontend cache-bust (?v=) on any
            # part URL invalidates when a non-first part is replaced too.
            video_version: int | None = None
            if p["has_video"]:
                try:
                    video_version = sum(
                        Path(vp).stat().st_mtime_ns for vp in video_paths
                    )
                except OSError:
                    video_version = None
            info: dict[str, Any] = {
                "id": pid,
                "video_path": first_path,
                "video_paths": video_paths,
                "has_video": p["has_video"],
                "in_sheet": p.get("in_sheet", False),
                "has_transcript": has_transcript,
                "segment_count": len(entry.get("segments", [])),
                "video_filename": Path(first_path).name,
                "video_filenames": [Path(vp).name for vp in video_paths],
                "video_version": video_version,
                "agents": {
                    "transcription": _step_state_transcription(entry),
                    "summary": _step_state_agent(pid, entry, "summary"),
                    "citations": _step_state_agent(pid, entry, "citations"),
                    "friction": _step_state_agent(pid, entry, "friction"),
                },
            }
            # Multi-video: expose the timeline so the frontend can switch the
            # <video> source per part and seek the local offset. Omitted for a
            # single video (no probe) → the frontend keeps its one-file path.
            if p["has_video"] and len(video_paths) >= 2:
                timeline = video.timeline_or_none(video_paths)
                if timeline is not None:
                    info["timeline"] = [
                        {
                            "filename": Path(path).name,
                            "duration": dur,
                            "cumulativeStart": cum,
                        }
                        for path, dur, cum in timeline
                    ]
            if has_transcript:
                info["language"] = entry.get("language", "")
                info["model"] = entry.get("model", "")
                info["transcribed_at"] = entry.get("transcribed_at", "")
                # What the last run actually transcribed — the picker shows it
                # back so an auto-detect deviation is explainable after the fact.
                info["audio_index"] = entry.get("audio_index", 0)
                info["audio_track_label"] = entry.get("audio_track_label", "")
                info["has_summary"] = bool(entry.get("summary"))
            result.append(info)

    # Check for stale artifacts (transcript outdated relative to source)
    artifacts = viewer.load_manifest_artifacts()
    artifacts_by_participant: defaultdict[str, list[dict[str, Any]]] = defaultdict(list)
    for art in artifacts:
        artifacts_by_participant[art.get("participant", "")].append(art)
    for info in result:
        if not info.get("has_transcript"):
            info["has_stale_artifacts"] = False
            continue
        pid = info["id"]
        current_ta = info.get("transcribed_at", "")
        if not current_ta:
            info["has_stale_artifacts"] = False
            continue
        has_stale = False
        for art in artifacts_by_participant.get(pid, []):
            if not art.get("transcript"):
                continue
            art_tv = art.get("transcript_version", "")
            if not art_tv or art_tv < current_ta:
                has_stale = True
                break
        info["has_stale_artifacts"] = has_stale

    # ``has_sheet`` gates the off-sheet badge: with no sheet loaded every entry
    # is ``in_sheet: False`` and marking them all would be noise.
    return ok(
        participants=result,
        has_sheet=bool(_participant_source and _participant_source["sheet_context"]),
        transcribe_prewarm=_transcribe_prewarm_setting(),
    )


# ---- Transcript data ----


# ---- Corrected-segments cache ----
#
# apply_corrections() is pure given (segments, corrections) but was re-run on
# every read — api_transcript per request, api_search across *all* participants
# per query. Corrections change rarely and segments are static post-transcription,
# so memoize the corrected list per participant, invalidated by a version counter
# bumped whenever corrections or a participant's segments change. Own lock so both
# the lock-holding caller (api_search) and the lock-free one (api_transcript) are
# safe; lock order is always _manifest_lock -> _corrected_cache_lock.
_corrected_cache: dict[str, tuple[int, list[Any]]] = {}
_corrected_cache_lock = threading.Lock()
_corrections_version = 0


def _bump_corrections_version() -> None:
    """Invalidate the corrected-segments cache (corrections or segments changed)."""
    global _corrections_version
    with _corrected_cache_lock:
        _corrections_version += 1
        _corrected_cache.clear()


def _corrected_segments(
    participant: str,
    raw_segments: list[Any],
    corrections: list[Any],
) -> list[Any]:
    """apply_corrections() for *participant*, memoized by corrections version."""
    with _corrected_cache_lock:
        version = _corrections_version
        cached = _corrected_cache.get(participant)
        if cached is not None and cached[0] == version:
            return cached[1]
    corrected = transcripts.apply_corrections(raw_segments, corrections)
    with _corrected_cache_lock:
        # Skip the store if a concurrent mutation bumped the version mid-compute.
        if _corrections_version == version:
            _corrected_cache[participant] = (version, corrected)
    return corrected


@transcripts_bp.route("/api/transcript/<participant>")
def api_transcript(participant: str) -> FlaskResponse:
    """Return full transcript segments for a participant (corrections applied)."""
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        if not entry or not entry.get("segments"):
            return err("No transcript for participant", 404)
        # Snapshot under the lock so a concurrent edit/transcribe/mark can't
        # mutate segments, corrections, or marks mid-iteration (mirrors api_vtt).
        raw_segments = list(entry["segments"])
        corrections = list(_manifest.get("corrections", []))
        marks_snapshot = list(_manifest.get("marks", []))
        language = entry.get("language", "")
        model = entry.get("model", "")
        transcribed_at = entry.get("transcribed_at", "")

    # Apply corrections to get corrected text (memoized per participant)
    corrected_segments = _corrected_segments(participant, raw_segments, corrections)

    # Build marks-by-segment-id lookup
    marks_by_seg: dict[str, list[dict[str, Any]]] = {}
    for mark in marks_snapshot:
        sid = mark.get("segment_id", "")
        marks_by_seg.setdefault(sid, []).append(mark)

    # Build response segments with corrected flag and marks
    segments = []
    for raw, corrected in zip(raw_segments, corrected_segments):
        seg_id = raw.get("id", "")
        seg: dict[str, Any] = {
            "id": seg_id,
            "start": corrected["start"],
            "end": corrected["end"],
            "text": corrected["text"],
            "corrected": raw["text"] != corrected["text"],
            "marks": marks_by_seg.get(seg_id, []),
        }
        segments.append(seg)

    return ok(
        participant=participant,
        segments=segments,
        language=language,
        model=model,
        transcribed_at=transcribed_at,
    )


@transcripts_bp.route("/api/transcript/<participant>/segment", methods=["PUT"])
def api_edit_segment(participant: str) -> FlaskResponse:
    """Edit a segment's text. Creates a correction entry automatically."""
    data = request.get_json(silent=True)
    if not data:
        return err("Missing JSON body")

    segment_id = data.get("segment_id", "")
    new_text = data.get("text", "").strip()
    if not segment_id or not new_text:
        return err("segment_id and text required")

    with _manifest_lock:
        src = _manifest.get("source_transcripts", {})
        entry = src.get(participant)
        if not entry:
            return err("No transcript for participant", 404)

        # Find the raw segment by ID
        raw_seg = None
        for seg in entry.get("segments", []):
            if seg.get("id") == segment_id:
                raw_seg = seg
                break
        if raw_seg is None:
            return err("Segment not found", 404)

        original_text = raw_seg["text"]
        if original_text == new_text:
            return ok(correction=None)

        # Create correction
        correction = {
            "id": f"c_{uuid.uuid4().hex[:8]}",
            "from": original_text,
            "to": new_text,
            "created": datetime.now(UTC).isoformat(),
        }
        _manifest.setdefault("corrections", []).append(correction)
        _bump_corrections_version()  # new correction invalidates corrected cache
        _mark_friction_stale(entry)  # edited segment text invalidates friction scores

    _schedule_persist()
    return ok(correction=correction)


# ---- WebVTT ----


@transcripts_bp.route("/api/vtt/<participant>")
def api_vtt(participant: str) -> FlaskResponse:
    """Serve transcript as WebVTT for <track> subtitle support."""
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        if not entry or not entry.get("segments"):
            return Response("WEBVTT\n", content_type="text/vtt")
        # Snapshot under the lock so a concurrent edit/transcribe can't mutate
        # corrections or segments mid-iteration.
        segments_snapshot = list(entry["segments"])
        corrections_snapshot = list(_manifest.get("corrections", []))
        language = entry.get("language", "")
        source_file = entry.get("source_file", "")
        model = entry.get("model", "")

    corrected = _corrected_segments(
        participant, segments_snapshot, corrections_snapshot
    )
    result = transcripts.TranscriptResult(
        segments=corrected,
        language=language,
        source_file=source_file,
        model=model,
    )
    vtt_text = transcripts._format_vtt(result)
    return Response(vtt_text, content_type="text/vtt")


# ---- Embed subtitles into video ----


def _video_paths_for_participant(participant: str) -> list[str]:
    """Return the ordered source video path(s) for *participant*, or [] if unknown."""
    _refresh_participants()
    for p in _participants:
        if p["id"] == participant:
            return list(p["video_paths"])
    return []


@transcripts_bp.route("/api/audio-info/<participant>")
def api_audio_info(participant: str) -> FlaskResponse:
    """Return the participant's audio-track layout (count + per-track labels).

    Lazy per-participant probe (cached in ``video`` by file mtime) so the
    ``/api/participants`` list endpoint stays free of an ffprobe per participant.
    Multi-part participants share the first part's audio setup.

    ``auto_index`` is the track transcription would auto-pick, computed by the
    same helper the worker uses — the transcribe picker labels its Auto option
    from this rather than re-deriving the heuristic in JS.
    """
    video_paths = _video_paths_for_participant(participant)
    if not video_paths:
        return err(f"No video for participant {participant}", 404)
    props = video.probe_video_properties(video_paths[0])
    if props is None:
        return err("Could not probe video file", 500)
    tracks = props.get("audio_tracks") or []
    return ok(
        audio_tracks=tracks,
        audio_track_count=props.get("audio_track_count") or 0,
        auto_index=video.pick_speech_audio_track(tracks),
    )


@transcripts_bp.route("/api/audio-track/<participant>/<int:idx>")
def api_audio_track(participant: str, idx: int) -> FlaskResponse:
    """Stream one demuxed audio track for the browser's per-track volume mixer."""
    video_paths = _video_paths_for_participant(participant)
    if not video_paths:
        return err(f"No video for participant {participant}", 404)
    out = video.extract_audio_track(video_paths[0], idx)
    if out is None:
        return err("Could not extract audio track", 500)
    response = send_file(str(out), mimetype="audio/mp4", conditional=True)
    response.headers["Cache-Control"] = "no-cache"
    return response


def _embed_subtitle_for_participant(
    participant: str, output_dir: Path
) -> dict[str, Any]:
    """Mux *participant*'s transcript into a copy of their source video.

    Returns a dict describing the outcome:
        {"participant": pid, "ok": True, "output_path": str, "output_filename": str}
        {"participant": pid, "ok": False, "error": str}
    Snapshots manifest state under the lock so concurrent edits cannot mutate
    segments or corrections mid-format.
    """
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        if not entry or not entry.get("segments"):
            return {
                "participant": participant,
                "ok": False,
                "error": "No transcript for participant",
            }
        segments_snapshot = list(entry["segments"])
        corrections_snapshot = list(_manifest.get("corrections", []))
        language = entry.get("language", "")
        source_file = entry.get("source_file", "")
        model = entry.get("model", "")

    video_paths = _video_paths_for_participant(participant)
    if not video_paths or not Path(video_paths[0]).is_file():
        return {
            "participant": participant,
            "ok": False,
            "error": "Source video not found",
        }
    if len(video_paths) > 1:
        # The global transcript spans several source files; muxing it back into a
        # single file would require concatenating the parts first. Not supported.
        return {
            "participant": participant,
            "ok": False,
            "error": "Subtitle embedding isn't supported for multi-video participants.",
        }
    video_path = video_paths[0]

    corrected = _corrected_segments(
        participant, segments_snapshot, corrections_snapshot
    )
    result = transcripts.TranscriptResult(
        segments=corrected,
        language=language,
        source_file=source_file,
        model=model,
    )
    srt_text = transcripts._format_srt(result)
    if not srt_text:
        return {
            "participant": participant,
            "ok": False,
            "error": "Transcript produced no SRT output",
        }

    src = Path(video_path)
    output_dir.mkdir(parents=True, exist_ok=True)
    desired_name = f"{src.stem}-subtitled{src.suffix}"
    output_path = files.get_unique_filename(
        str(output_dir / desired_name), file_format=src.suffix
    )

    tmp_path = ""
    try:
        # delete=False + the explicit unlink below: ffmpeg reopens the sidecar by
        # path once we've closed it, which Windows won't allow while the
        # NamedTemporaryFile handle is still open.
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".srt", delete=False, encoding="utf-8"
        ) as tmp:
            tmp_path = tmp.name
            tmp.write(srt_text)
        ok = video.mux_subtitles(
            str(video_path),
            tmp_path,
            output_path,
            track_language=language or "und",
        )
    finally:
        if tmp_path:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    if not ok:
        return {
            "participant": participant,
            "ok": False,
            "error": "ffmpeg failed to mux subtitles",
        }
    return {
        "participant": participant,
        "ok": True,
        "output_path": output_path,
        "output_filename": Path(output_path).name,
    }


@transcripts_bp.route("/api/embed-subtitle/<participant>", methods=["POST"])
def api_embed_subtitle(participant: str) -> FlaskResponse:
    """Mux *participant*'s transcript into a copy of their source video.

    Writes ``<video-stem>-subtitled<ext>`` into the effective output dir
    (uniquified if a previous run already wrote that name).
    """
    output_dir = Path(utils.get_effective_output_dir())
    outcome = _embed_subtitle_for_participant(participant, output_dir)
    if not outcome["ok"]:
        status = 404 if outcome["error"] == "No transcript for participant" else 500
        return err(outcome["error"], status)
    return ok(
        output_path=outcome["output_path"],
        output_filename=outcome["output_filename"],
        output_dir=str(output_dir),
    )


@transcripts_bp.route("/api/embed-all-subtitles", methods=["POST"])
def api_embed_all_subtitles() -> FlaskResponse:
    """Mux every participant with a transcript into a subtitled copy."""
    output_dir = Path(utils.get_effective_output_dir())
    with _manifest_lock:
        src = _manifest.get("source_transcripts", {})
        targets = [pid for pid, entry in src.items() if entry.get("segments")]
    if not targets:
        return err("No transcripts available to embed.", 404)
    results = [_embed_subtitle_for_participant(pid, output_dir) for pid in targets]
    return ok(
        results=results,
        output_dir=str(output_dir),
    )


# ---- AI thinking agents (summary / citations / friction) ----
#
# Three generic routes cover every agent in thinking_agents.AGENTS, keyed by
# <agent_key>, so appending an Agent needs no new endpoints here. Summary keeps
# two extra routes below that are genuinely unique to it (the SSE token stream
# and the user-edit PUT).


@transcripts_bp.route("/api/agent/<agent_key>/<participant>")
def api_agent_get(agent_key: str, participant: str) -> FlaskResponse:
    """Return an agent's result, its generation status, or 404.

    Success carries the result under the agent's ``manifest_field`` key, so the
    summary/citations/friction responses keep their historic shape. For every
    agent that depends on this one, a ``<dep_field>_generating`` flag (and
    ``<dep_field>_started_at`` when true) is added generically — that is how the
    summary poll still surfaces citation-generation status to the UI without
    special-casing it.
    """
    agent = thinking_agents.get_agent(agent_key)
    if agent is None:
        return jsonify({"ok": False}), 404
    field = agent["manifest_field"]
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        result = entry.get(field) if entry else None
    if result:
        resp: dict[str, Any] = {"ok": True, field: result}
        for dep in thinking_agents.AGENTS:
            if field in dep["depends_on"]:
                dep_field = dep["manifest_field"]
                generating = _orchestrator.is_generating(participant, dep["key"])
                resp[f"{dep_field}_generating"] = generating
                if generating:
                    resp[f"{dep_field}_started_at"] = _orchestrator.started_at(
                        participant, dep["key"]
                    )
        return jsonify(resp)
    if _orchestrator.is_generating(participant, agent_key):
        return jsonify(
            {
                "ok": False,
                "generating": True,
                "started_at": _orchestrator.started_at(participant, agent_key),
                "partial": _orchestrator.partial_text(participant, agent_key),
            }
        )
    # Friction's per-segment scores + session stats come from a pure, deterministic
    # scorer (friction.py) that needs no summary and no LLM. Surface them before the
    # summary-gated friction agent runs so the heatmap/timeline/stats are usable
    # immediately; the LLM-refined "moments" stay empty until the agent writes the
    # manifest field, at which point the stored branch above wins. The `deterministic`
    # flag lets the client show programmatic-only copy and keeps the friction poll
    # from mistaking this display fallback for a completed run.
    if agent_key == "friction" and entry and entry.get("segments"):
        segments = entry["segments"]
        scored = friction.score_segments(segments)
        stats = friction.compute_stats(
            scored, thinking_agents._segments_duration(segments)
        )
        return jsonify(
            {
                "ok": True,
                "friction": {
                    "segments": scored,
                    "moments": [],
                    "stats": stats,
                    "model": None,
                    "llm_ok": None,
                    "stale": False,
                    "deterministic": True,
                },
            }
        )
    return jsonify({"ok": False}), 404


@transcripts_bp.route(
    "/api/agent/<agent_key>/<participant>/regenerate", methods=["POST"]
)
def api_agent_regenerate(agent_key: str, participant: str) -> FlaskResponse:
    """Clear an agent's result and re-trigger it (bypassing its enabled config).

    Manual trigger: runs even when the agent's enabled config is False so the
    frontend's per-participant controls can force a run. Requires a transcript
    and every dependency's result. Regenerating an agent also invalidates its
    dependents (``_invalidate_dependents``): e.g. regenerating summary clears
    citations and flags friction stale. ``run_agent`` auto-advances the chain
    on completion, so downstream enabled agents still run.
    """
    agent = thinking_agents.get_agent(agent_key)
    if agent is None:
        return jsonify({"ok": False}), 404
    if _orchestrator.is_generating(participant, agent_key):
        return ok(generating=True)
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        if (
            not entry
            or not entry.get("segments")
            or not _agent_dependencies_met(agent, entry)
        ):
            return err("No transcript or unmet dependency", 404)
        entry.pop(agent["manifest_field"], None)
        _invalidate_dependents(entry, agent)
    _persist_manifest()
    _orchestrator.run_agent(agent_key, participant, force=True)
    return ok(generating=True)


@transcripts_bp.route("/api/agent/<agent_key>/<participant>/stop", methods=["POST"])
def api_agent_stop(agent_key: str, participant: str) -> FlaskResponse:
    """Abort an in-flight agent run.

    Sets the cancel event so the streaming Ollama call closes its response
    promptly, freeing the model for another run. The UI flips to idle
    immediately. After a short delay, the model is unloaded from memory if no
    new run has started in the meantime.
    """
    if thinking_agents.get_agent(agent_key) is None:
        return jsonify({"ok": False}), 404
    if _orchestrator.stop(agent_key, participant):
        model = _agent_model(agent_key)
        if model:
            _schedule_model_unload(model)
    return ok(running=False)


# ---- Summary-only routes (SSE token stream + user edit) ----


# Server-side cadence for the summary token stream. The agent runs in the
# orchestrator's daemon thread and appends tokens to the shared partial buffer;
# this SSE generator samples that buffer in-process and pushes deltas to the
# browser, so the client sees near-real-time word-by-word text over one
# connection instead of hammering the GET poll.
_SUMMARY_STREAM_TICK = 0.1  # seconds between buffer samples
_SUMMARY_STREAM_START_GRACE = 2.0  # seconds to wait for the run to claim its slot


@transcripts_bp.route("/api/agent/summary/<participant>/stream")
def api_summary_stream(participant: str) -> FlaskResponse:
    """Stream summary tokens to the browser via SSE as the model produces them.

    Emits ``data: {"partial": "<text-so-far>"}`` each time the buffer grows and a
    final ``data: {"done": true}`` when the run finishes (or was never running).
    The client closes the stream on ``done`` — EventSource would otherwise treat
    the server-side close as an error and reconnect. On any transport failure the
    client falls back to the GET poll, so this is a pure enhancement.
    """

    def _events():  # type: ignore[no-untyped-def]
        sent = 0
        deadline = time.monotonic() + _SUMMARY_STREAM_START_GRACE
        while True:
            generating = _orchestrator.is_generating(participant, "summary")
            text = _orchestrator.partial_text(participant, "summary")
            if len(text) > sent:
                sent = len(text)
                yield f"data: {json.dumps({'partial': text})}\n\n"
            if not generating:
                # Guard the open race: the client opens the stream right after
                # regenerate claims the slot, but tolerate a brief window where
                # is_generating hasn't flipped true yet before declaring done.
                if sent == 0 and time.monotonic() < deadline:
                    time.sleep(_SUMMARY_STREAM_TICK)
                    continue
                yield f"data: {json.dumps({'done': True})}\n\n"
                return
            time.sleep(_SUMMARY_STREAM_TICK)

    return Response(
        stream_with_context(_events()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@transcripts_bp.route("/api/agent/summary/<participant>", methods=["PUT"])
def api_summary_save(participant: str) -> FlaskResponse:
    """Save a user-edited summary for a participant."""
    data = request.get_json(silent=True)
    if not data or not data.get("summary", "").strip():
        return err("Summary text is required")
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        if not entry:
            return err("Participant not found", 404)
        entry["summary"] = data["summary"].strip()
        entry.pop("citations", None)  # invalidate citations on edit
        _mark_friction_stale(entry)  # summary text feeds the friction prompt
    _schedule_persist()
    return ok()


# ---- Corrections ----


@transcripts_bp.route("/api/corrections")
def api_corrections_list() -> FlaskResponse:
    """List all study-local corrections."""
    with _manifest_lock:
        corrections = list(_manifest.get("corrections", []))
    return ok(corrections=corrections)


@transcripts_bp.route("/api/corrections", methods=["POST"])
def api_corrections_add() -> FlaskResponse:
    """Add a manual correction."""
    data = request.get_json(silent=True)
    if not data:
        return err("Missing JSON body")

    from_text = data.get("from", "").strip()
    to_text = data.get("to", "").strip()
    if not from_text or not to_text:
        return err("'from' and 'to' required")

    removed_id = None
    with _manifest_lock:
        corrections = _manifest.setdefault("corrections", [])

        # Check for an existing correction whose `to` chains into this `from`.
        # e.g. existing "teh"→"the" + new "the"→"they" → update to "teh"→"they".
        # If the update would make from == to, delete the correction instead.
        chained = None
        for c in corrections:
            if c.get("to", "").lower() == from_text.lower():
                chained = c
                break

        if chained:
            if chained["from"].lower() == to_text.lower():
                corrections.remove(chained)
                removed_id = chained["id"]
                correction = None
            else:
                chained["to"] = to_text
                correction = chained
        else:
            correction = {
                "id": f"c_{uuid.uuid4().hex[:8]}",
                "from": from_text,
                "to": to_text,
                "created": datetime.now(UTC).isoformat(),
            }
            corrections.append(correction)

        _bump_corrections_version()  # add/update/remove invalidates corrected cache
    # Schedule the debounced write OUTSIDE _manifest_lock, like the other edit
    # routes, so we never nest _manifest_lock -> the debounce timer lock.
    _schedule_persist()
    if removed_id is not None:
        return ok(correction=None, removed=removed_id)
    return ok(correction=correction)


@transcripts_bp.route("/api/corrections/<correction_id>", methods=["DELETE"])
def api_corrections_delete(correction_id: str) -> FlaskResponse:
    """Remove a correction by ID."""
    with _manifest_lock:
        corrections = _manifest.get("corrections", [])
        before = len(corrections)
        _manifest["corrections"] = [
            c for c in corrections if c.get("id") != correction_id
        ]
        removed = before - len(_manifest["corrections"])
        if removed:
            _bump_corrections_version()  # deletion invalidates corrected cache

    if removed == 0:
        return err("Correction not found", 404)

    _schedule_persist()
    return ok()


# ---- Marks ----


def _resolve_mark(
    mark: dict[str, Any],
    partial_lookup: dict[str, list] | None = None,
) -> dict[str, Any]:
    """Enrich a mark with its segment's data (participant, start, end, text, valid).

    *partial_lookup* maps participant IDs to ``partial_segments`` lists from
    running transcription tasks, allowing marks made during streaming to
    resolve before the transcript is persisted.
    """
    seg_id = mark.get("segment_id", "")
    # segment IDs are "{participant}:{index}"
    parts = seg_id.split(":", 1)
    if len(parts) != 2:
        return {
            **mark,
            "valid": False,
            "participant": "",
            "start": 0,
            "end": 0,
            "text": "",
        }
    pid, idx_str = parts
    try:
        idx = int(idx_str)
    except ValueError:
        return {
            **mark,
            "valid": False,
            "participant": pid,
            "start": 0,
            "end": 0,
            "text": "",
        }

    # Try persisted transcript first
    src = _manifest.get("source_transcripts", {})
    entry = src.get(pid, {})
    segments = entry.get("segments", [])
    if 0 <= idx < len(segments):
        corrections = _manifest.get("corrections", [])
        # Correct the whole participant list once (memoized by corrections
        # version) instead of recompiling the regex set per mark.
        corrected = _corrected_segments(pid, segments, corrections)
        seg = corrected[idx]
        return {
            **mark,
            "valid": True,
            "participant": pid,
            "start": seg["start"],
            "end": seg["end"],
            "text": seg["text"],
        }

    # Fall back to partial segments from a running transcription task
    if partial_lookup:
        partial_segs = partial_lookup.get(pid, [])
        if 0 <= idx < len(partial_segs):
            seg = partial_segs[idx]
            return {
                **mark,
                "valid": True,
                "participant": pid,
                "start": seg["start"],
                "end": seg["end"],
                "text": seg["text"],
            }

    return {
        **mark,
        "valid": False,
        "participant": pid,
        "start": 0,
        "end": 0,
        "text": "",
    }


def _build_partial_lookup() -> dict[str, list]:
    """Build a participant→partial_segments map from running transcription tasks."""
    if not _worker:
        return {}
    lookup: dict[str, list] = {}
    for task in _worker.get_all_tasks():
        if task["status"] == transcripts.TASK_STATUS_RUNNING:
            segs = task.get("partial_segments", [])
            if segs:
                lookup[task["participant"]] = segs
    return lookup


def marks_for_participant(pid: str) -> list[dict[str, Any]]:
    """Resolved, valid marks for one participant, sorted by start time.

    Reuses the in-memory manifest (no disk read) so the Screenspace blueprint
    can surface transcript marks without re-loading state. Mirrors the
    resolution in ``api_marks_list``; filters to marks whose resolved
    participant equals *pid* and drops any that no longer resolve.
    """
    # Build partial lookup outside _manifest_lock (get_all_tasks acquires worker lock)
    partial_lookup = _build_partial_lookup()
    with _manifest_lock:
        raw_marks = list(_manifest.get("marks", []))
        resolved = [_resolve_mark(m, partial_lookup) for m in raw_marks]
    out = [m for m in resolved if m.get("valid") and m.get("participant") == pid]
    out.sort(key=lambda m: m.get("start", 0))
    return out


@transcripts_bp.route("/api/marks")
def api_marks_list() -> FlaskResponse:
    """List all marks, enriched with resolved segment data."""
    # Build partial lookup outside _manifest_lock (get_all_tasks acquires worker lock)
    partial_lookup = _build_partial_lookup()
    with _manifest_lock:
        raw_marks = list(_manifest.get("marks", []))
        resolved = [_resolve_mark(m, partial_lookup) for m in raw_marks]
    return ok(
        marks=resolved,
        categories=config.MARK_CATEGORIES,
    )


@transcripts_bp.route("/api/intake-poll")
def api_intake_poll() -> FlaskResponse:
    """Combined Studio-intake poll: running-state booleans + resolved marks.

    Collapses the Studio intake client's four transcript polls (transcribe
    status, model-status, participants, marks) into one request. ``agents_running``
    is read straight from the orchestrator's in-memory in-flight tracking, so —
    unlike /api/participants — this never stat()s or ffprobe()s a video file."""
    tasks_running = False
    if _worker:
        for t in _worker.get_all_tasks(include_partials=False):
            if t["status"] == transcripts.TASK_STATUS_RUNNING:
                tasks_running = True
                break
    with _transcript_model_warming_lock:
        model_warming = _transcript_model_warming
    agents_running = any(
        _orchestrator.is_generating(p["id"], "summary")
        or _orchestrator.is_generating(p["id"], "citations")
        for p in _participants
    )
    # Same resolve path as /api/marks (partial lookup outside _manifest_lock).
    partial_lookup = _build_partial_lookup()
    with _manifest_lock:
        raw_marks = list(_manifest.get("marks", []))
        resolved = [_resolve_mark(m, partial_lookup) for m in raw_marks]
    return ok(
        status={
            "tasks_running": tasks_running,
            "model_warming": model_warming,
            "agents_running": agents_running,
        },
        marks={"marks": resolved, "categories": config.MARK_CATEGORIES},
    )


@transcripts_bp.route("/api/marks", methods=["POST"])
def api_marks_add() -> FlaskResponse:
    """Create marks for one or more segments."""
    data = request.get_json(silent=True)
    if not data:
        return err("Missing JSON body")

    segment_ids = data.get("segment_ids", [])
    if not segment_ids:
        return err("segment_ids required")

    category = data.get("category") or None
    label = data.get("label") or None
    severity = data.get("severity") or None
    now = datetime.now(UTC).isoformat()

    created = []
    with _manifest_lock:
        marks = _manifest.setdefault("marks", [])
        existing_by_seg = {m["segment_id"]: m for m in marks}

        for sid in segment_ids:
            if sid in existing_by_seg:
                # Update existing mark
                m = existing_by_seg[sid]
                if category is not None:
                    m["category"] = category
                if label is not None:
                    m["label"] = label
                if severity is not None:
                    m["severity"] = severity
                created.append(m)
            else:
                m = {
                    "id": f"m_{uuid.uuid4().hex[:8]}",
                    "segment_id": sid,
                    "category": category,
                    "label": label,
                    "severity": severity,
                    "created": now,
                }
                marks.append(m)
                existing_by_seg[sid] = m
                created.append(m)

    _schedule_persist()
    return ok(marks=created)


@transcripts_bp.route("/api/marks/<mark_id>", methods=["PUT"])
def api_marks_update(mark_id: str) -> FlaskResponse:
    """Update a mark's category or label."""
    data = request.get_json(silent=True)
    if not data:
        return err("Missing JSON body")

    with _manifest_lock:
        marks = _manifest.get("marks", [])
        target = None
        for m in marks:
            if m.get("id") == mark_id:
                target = m
                break
        if not target:
            return err("Mark not found", 404)

        if "category" in data:
            target["category"] = data["category"] or None
        if "label" in data:
            target["label"] = data["label"] or None
        if "severity" in data:
            target["severity"] = data["severity"] or None

    _schedule_persist()
    return ok(mark=target)


@transcripts_bp.route("/api/marks/<mark_id>", methods=["DELETE"])
def api_marks_delete(mark_id: str) -> FlaskResponse:
    """Remove a mark by ID, or bulk-delete with JSON body {ids: [...]}."""
    # Bulk delete: DELETE /api/marks with {ids: [...]} — mark_id may be a placeholder
    data = request.get_json(silent=True)
    ids_to_remove: list[str] = []

    if data and data.get("ids"):
        ids_to_remove = data["ids"]
    else:
        ids_to_remove = [mark_id]

    with _manifest_lock:
        marks = _manifest.get("marks", [])
        remove_set = set(ids_to_remove)
        before = len(marks)
        _manifest["marks"] = [m for m in marks if m.get("id") not in remove_set]
        removed = before - len(_manifest["marks"])

    if removed == 0:
        return err("Mark not found", 404)

    _schedule_persist()
    return ok(removed=removed)


# ---- Search ----


@transcripts_bp.route("/api/search")
def api_search() -> FlaskResponse:
    """Keyword search across all transcribed participants."""
    query = request.args.get("q", "").strip()
    if not query:
        return ok(
            query="",
            total_count=0,
            results=[],
            counts_by_participant={},
        )

    query_lower = query.lower()
    results: list[dict[str, Any]] = []
    counts: dict[str, int] = {}

    with _manifest_lock:
        src = _manifest.get("source_transcripts", {})
        corrections = _manifest.get("corrections", [])

        for pid, entry in src.items():
            raw_segments = entry.get("segments", [])
            if not raw_segments:
                continue
            corrected = _corrected_segments(pid, raw_segments, corrections)
            participant_count = 0
            for raw, seg in zip(raw_segments, corrected):
                text_lower = seg["text"].lower()
                n = text_lower.count(query_lower)
                if n > 0:
                    participant_count += n
                    results.append(
                        {
                            "participant": pid,
                            "segment_id": raw.get("id", ""),
                            "start": seg["start"],
                            "end": seg["end"],
                            "text": seg["text"],
                            "count": n,
                        }
                    )
            if participant_count > 0:
                counts[pid] = participant_count

    total = sum(counts.values())
    return ok(
        query=query,
        total_count=total,
        results=results,
        counts_by_participant=counts,
    )


# ---- Transcription queue ----


def _whisper_model_size_mb(model: str) -> int | None:
    """Download size (MB) for a Whisper model name, or None if unknown."""
    return next(
        (m["size_mb"] for m in transcripts.WHISPER_MODELS if m["name"] == model),
        None,
    )


@transcripts_bp.route("/api/transcribe/warmup", methods=["POST"])
def api_transcribe_warmup() -> FlaskResponse:
    """Background-load the Whisper model when automatic prewarm is not ``off``.

    Never downloads a model silently: if the configured model isn't cached yet,
    skips with ``reason: model_not_cached`` (plus ``model``/``size_mb``) so the
    frontend can confirm the download. Re-post with ``{"force": true}`` to
    proceed after the user agrees.
    """
    if _transcribe_prewarm_setting() == "off":
        return ok(
            skipped=True,
            reason="prewarm_disabled",
        )

    if transcripts.is_transcription_model_loaded():
        return ok(already_loaded=True)

    # Never download a model silently during prewarm. When the configured model
    # isn't cached yet, skip and report it so the frontend can confirm the
    # download with the user; a re-post with {"force": true} then proceeds.
    data = request.get_json(silent=True) or {}
    force = bool(data.get("force"))
    model = config.TRANSCRIBE_MODEL
    if not force and not transcripts.is_whisper_model_cached(model):
        return ok(
            skipped=True,
            reason="model_not_cached",
            model=model,
            size_mb=_whisper_model_size_mb(model),
        )

    global _transcript_model_warming

    with _transcript_model_warming_lock:
        if _transcript_model_warming:
            return ok(already_warming=True)
        if transcripts.is_transcription_model_loaded():
            return ok(already_loaded=True)
        _transcript_model_warming = True

    def _run_warmup() -> None:
        global _transcript_model_warming
        try:
            transcripts.warmup_transcription_model()
        finally:
            with _transcript_model_warming_lock:
                _transcript_model_warming = False

    threading.Thread(
        target=_run_warmup, daemon=True, name="transcript-model-warmup"
    ).start()
    return ok(started=True)


@transcripts_bp.route("/api/transcribe/model-status")
def api_transcribe_model_status() -> FlaskResponse:
    """Report whether faster-whisper is loaded or a warm-up is in progress."""
    with _transcript_model_warming_lock:
        warming = _transcript_model_warming
    return ok(
        loaded=transcripts.is_transcription_model_loaded(),
        warming=warming,
        model=config.TRANSCRIBE_MODEL,
        prewarm=_transcribe_prewarm_setting(),
    )


@transcripts_bp.route("/api/models/ollama/pull", methods=["POST"])
def api_ollama_pull() -> FlaskResponse:
    """Download (install) an Ollama model in the background, tracking progress.

    The frontend calls this only after the user confirms the install, then
    polls /api/models/ollama/pull-status for progress.
    """
    data = request.get_json(silent=True) or {}
    model = (data.get("model") or "").strip()
    if not model:
        return err("Missing model")

    with _ollama_pull_lock:
        existing = _ollama_pull_status.get(model)
        if existing is not None and not existing.get("done"):
            return ok(already_pulling=True)
        _ollama_pull_status[model] = {
            "status": "starting",
            "completed": 0,
            "total": 0,
            "done": False,
            "succeeded": False,
            "error": None,
        }

    def _on_progress(chunk: dict[str, Any]) -> None:
        with _ollama_pull_lock:
            st = _ollama_pull_status.get(model)
            if st is None:
                return
            status = chunk.get("status")
            if status:
                st["status"] = status
            total = chunk.get("total")
            completed = chunk.get("completed")
            if isinstance(total, (int, float)):
                st["total"] = int(total)
            if isinstance(completed, (int, float)):
                st["completed"] = int(completed)

    def _run_pull() -> None:
        succeeded = False
        try:
            succeeded = ollama_client.pull_model(model, on_progress=_on_progress)
        finally:
            with _ollama_pull_lock:
                st = _ollama_pull_status.get(model)
                if st is not None:
                    st["done"] = True
                    st["succeeded"] = succeeded
                    if succeeded:
                        st["status"] = "success"
                    elif not st.get("error"):
                        st["error"] = "Pull failed"

    threading.Thread(target=_run_pull, daemon=True, name=f"ollama-pull-{model}").start()
    return ok(started=True)


@transcripts_bp.route("/api/models/ollama/pull-status")
def api_ollama_pull_status() -> FlaskResponse:
    """Report progress for a model pull started via /api/models/ollama/pull."""
    model = (request.args.get("model") or "").strip()
    if not model:
        return err("Missing model")
    with _ollama_pull_lock:
        st = _ollama_pull_status.get(model)
        snapshot = dict(st) if st is not None else None
    if snapshot is None:
        return ok(found=False)
    snapshot["ok"] = True
    snapshot["found"] = True
    return jsonify(snapshot)


@transcripts_bp.route("/api/transcribe", methods=["POST"])
def api_transcribe() -> FlaskResponse:
    """Enqueue participant(s) for background transcription."""
    data = request.get_json(silent=True)
    if not data:
        return err("Missing JSON body")

    participant_ids = data.get("participants", [])
    force = data.get("force", False)
    overrides = data.get("overrides") or {}
    allow_download = bool(data.get("allow_download"))

    if not participant_ids:
        return err("No participants specified")

    # Build a lookup of available participants (refresh first so a video dropped
    # into the input dir since page load can be enqueued rather than 404'd).
    _refresh_participants()
    available = {p["id"]: p for p in _participants}
    enqueued = []

    with _manifest_lock:
        src = _manifest.get("source_transcripts", {})

        # Resolve the participants that would actually be enqueued, with the
        # effective Whisper model for each (per-participant override → default).
        eligible: list[
            tuple[dict[str, Any], str | None, str | None, int | None, str]
        ] = []
        for pid in participant_ids:
            p = available.get(pid)
            if not p or not p.get("has_video"):
                continue

            # Skip already-transcribed unless force
            if not force and pid in src and src[pid].get("segments"):
                continue

            o = overrides.get(pid) or {}
            model_override = o.get("model") or None
            language_override = o.get("language") or None
            # Explicitly, not `or None`: track 0 is a valid selection and would
            # be swallowed as "no override" by a falsy test.
            raw_index = o.get("audio_index")
            audio_override: int | None = None
            if raw_index is not None and str(raw_index).strip() != "":
                try:
                    audio_override = int(raw_index)
                except (TypeError, ValueError):
                    return err(f"Invalid audio_index for {pid}")
                if audio_override < 0:
                    return err(f"Invalid audio_index for {pid}")
            # No upper bound here — that needs an ffprobe per participant, and
            # the worker already fails the task with the real track count.
            effective_model = model_override or config.TRANSCRIBE_MODEL
            eligible.append(
                (p, model_override, language_override, audio_override, effective_model)
            )

        # Authoritative download gate: never let a worker silently pull an
        # uncached faster-whisper model. The browser confirmation is advisory;
        # this enforces it for direct API calls and the /api/models fallback.
        if not allow_download:
            uncached: list[str] = []
            for _p, _mo, _lo, _ao, effective_model in eligible:
                if (
                    effective_model not in uncached
                    and not transcripts.is_whisper_model_cached(effective_model)
                ):
                    uncached.append(effective_model)
            if uncached:
                return jsonify(
                    {
                        "ok": False,
                        "reason": "model_not_cached",
                        "uncached": [
                            {"model": m, "size_mb": _whisper_model_size_mb(m)}
                            for m in uncached
                        ],
                    }
                )

        for p, model_override, language_override, audio_override, _em in eligible:
            task = transcripts.create_transcript_task(
                p["id"],
                p["video_paths"],
                model=model_override,
                language=language_override,
                audio_index=audio_override,
            )
            if _worker:
                _worker.enqueue(task)
            enqueued.append(
                {
                    "id": task["id"],
                    "participant": p["id"],
                    "status": task["status"],
                }
            )

    return ok(tasks=enqueued)


@transcripts_bp.route("/api/transcribe/status")
def api_transcribe_status() -> FlaskResponse:
    """Poll transcription task status."""
    tasks = []
    if _worker:
        # include_partials=False keeps the 3 s poll from deep-copying every
        # running task's growing segment tail; clients pull new segments via
        # /api/transcribe/<task_id>/segments?since=N using partial_count.
        for t in _worker.get_all_tasks(include_partials=False):
            task_info = {
                "id": t["id"],
                "participant": t["participant"],
                "status": t["status"],
                "progress": t["progress"],
                "error": t.get("error"),
                "created_at": t.get("created_at"),
                "completed_at": t.get("completed_at"),
            }
            if t["status"] == transcripts.TASK_STATUS_RUNNING:
                task_info["partial_count"] = t.get("partial_count", 0)
            tasks.append(task_info)
    return ok(
        tasks=tasks,
        worker_alive=_worker.is_alive if _worker else False,
    )


@transcripts_bp.route("/api/transcribe/<task_id>/segments")
def api_transcribe_segments(task_id: str) -> FlaskResponse:
    """Return a running task's partial-segment tail from ``?since=N`` (append-only)."""
    if not _worker:
        return err("Worker not initialized", 500)
    try:
        since = int(request.args.get("since", 0))
    except (TypeError, ValueError):
        since = 0
    segments, total = _worker.get_partial_segments(task_id, since)
    return ok(segments=segments, total=total)


@transcripts_bp.route("/api/transcribe/<task_id>", methods=["DELETE"])
def api_transcribe_cancel(task_id: str) -> FlaskResponse:
    """Cancel or dismiss a transcription task. ?dismiss=true fully removes it."""
    if not _worker:
        return err("Worker not initialized", 500)

    if request.args.get("dismiss") == "true":
        if not _worker.remove_task(task_id):
            return err("Task not found", 404)
        _persist_manifest()
        return ok()

    if _worker.cancel(task_id):
        _persist_manifest()
        return ok()

    return err("Task not found or already finished")


# ---- Manifest persistence ----


def _persist_manifest() -> None:
    """Persist the current manifest state through a single synchronized path.

    Synchronous callers (task completion, atexit flush) supersede any pending
    debounced write — cancel the timer here so we don't double-flush.
    """
    _cancel_pending_persist_timer()
    with _manifest_lock:
        _do_persist()


def _merge_completed_results_locked() -> list[str]:
    """Merge completed task results into the in-memory manifest.

    Caller must hold _manifest_lock. Merges (not replaces) so that extra keys
    like "summary" added by other threads are preserved. Kept separate from the
    disk write so callers can register the thinking-agent chain off the freshly
    merged segments *before* the (slower) persist.

    Each task is merged exactly once (tracked in ``_merged_task_ids``). Returns
    the participant ids whose segments were freshly merged on this call, in
    completion order, so the completion handler can refresh their agent outputs.
    """
    if not _worker:
        return []
    merged_pids: list[str] = []
    for task in _worker.get_all_tasks():
        if (
            task["status"] == transcripts.TASK_STATUS_COMPLETED
            and task.get("result")
            and task["id"] not in _merged_task_ids
        ):
            pid = task["participant"]
            src = _manifest.setdefault("source_transcripts", {})
            existing = src.get(pid, {})
            existing.update(task["result"])
            src[pid] = existing
            _merged_task_ids.add(task["id"])
            merged_pids.append(pid)
    _manifest["tasks"] = _worker.get_all_tasks()
    # Queue the completion side effects for _on_task_complete no matter which
    # caller performed the merge (a debounced _do_persist can win the race).
    _pending_chain_pids.extend(merged_pids)
    if merged_pids:
        # Each task merges exactly once, so a non-empty merged_pids means a
        # participant's segments were just (re)placed — invalidate the
        # corrected-segments cache so reads recompute against the new text.
        _bump_corrections_version()
    return merged_pids


def _do_persist() -> None:
    """Persist manifest to disk - caller must hold _manifest_lock."""
    _merge_completed_results_locked()
    transcripts.save_transcripts_manifest(
        _manifest.get("source_transcripts", {}),
        _manifest.get("corrections", []),
        marks=_manifest.get("marks"),
    )


# Manifest-write debounce: rapid UI edits (adding marks/corrections, segment and
# summary edits) coalesce into one disk write after a short quiet period instead
# of blocking each request on a full save_transcripts_manifest() of every
# participant's segments. In-session reads are unaffected — they read the
# in-memory _manifest under _manifest_lock, never disk. atexit fires the pending
# flush on normal exit / SIGINT / SIGTERM, but not on SIGKILL or hard power-loss
# — accepted because the transcripts manifest is recreatable (re-transcribe) and
# a sub-2s window of edits is cheap to redo. The lambda looks up _do_persist at
# call time so tests monkeypatching it are seen.
(_schedule_persist, _flush_pending_persist, _cancel_pending_persist_timer) = (
    make_debounced_persist(lambda: _do_persist(), _manifest_lock)
)

atexit.register(_flush_pending_persist)


# ---- Thinking-agent orchestration ----


def _on_task_complete() -> None:
    """Merge completed results into memory, start the thinking-agent chain for
    newly completed participants, then persist to disk.

    Ordering matters: run_chain registers the summary agent in _in_flight (which
    drives the UI's "generating" state) *before* the slow disk write, shrinking
    the window where a task reads as completed but no agent reads as running.
    """
    with _manifest_lock:
        # Merge first so next_eligible() sees the freshly completed segments.
        # A debounced _do_persist may already have merged this task; either
        # way the freshly merged participants (first transcription and
        # re-transcription alike — a new task id) sit in _pending_chain_pids,
        # so drain that instead of trusting this call's own merge return.
        _merge_completed_results_locked()
        merged_pids = list(dict.fromkeys(_pending_chain_pids))
        _pending_chain_pids.clear()
        src = _manifest.get("source_transcripts", {})
        for pid in merged_pids:
            entry = src.get(pid)
            if not entry:
                continue
            # A fresh transcript invalidates any prior AI outputs; clear every
            # agent's field so the chain regenerates them against the new
            # segments instead of leaving stale results from the old transcript.
            for agent in thinking_agents.AGENTS:
                entry.pop(agent["manifest_field"], None)

    # run_chain -> next_eligible/run_agent re-acquire _manifest_lock, so this
    # must run OUTSIDE the block above (the lock is non-reentrant).
    for pid in merged_pids:
        _orchestrator.run_chain(pid)

    _persist_manifest()


def _agent_enabled(agent: thinking_agents.Agent) -> bool:
    """Return True if *agent* is enabled in the current config."""
    return bool(getattr(config, agent["enabled_config_key"], False))


def _agent_dependencies_met(
    agent: thinking_agents.Agent, entry: dict[str, Any]
) -> bool:
    """Return True if every dependency agent's manifest field is populated
    on *entry*."""
    for dep_key in agent["depends_on"]:
        dep = thinking_agents.get_agent(dep_key)
        if dep is None:
            return False
        if not entry.get(dep["manifest_field"]):
            return False
    return True


class AgentOrchestrator:
    """Owns per-agent in-flight, cancel-event, and thread state.

    Encapsulates the dicts and orchestration helpers that previously lived as
    module-level globals. Methods reach into module-level ``_manifest``,
    ``_persist_manifest``, and ``_cancel_pending_unload`` directly — fine
    because this class is module-internal and those names resolve at call
    time.

    The shared ``_manifest_lock`` is passed in at construction so atomicity
    between manifest reads and in-flight slot claims is preserved verbatim
    against the pre-refactor behavior.

    Per-run cancel events: when an agent starts for a participant, a fresh
    ``threading.Event`` is registered; the Ollama transport closes its
    streaming HTTP response on ``event.set()``, which unblocks the read loop
    and frees the model promptly. ``stop`` looks up the event and sets it.
    The ``cancel_event.is_set()`` re-check in ``run_agent`` gates the
    result-and-chain-advance step as defense-in-depth for finish-vs-cancel
    races.
    """

    def __init__(self, lock: threading.Lock) -> None:
        self._lock = lock
        self._in_flight: dict[str, set[str]] = {
            a["key"]: set() for a in thinking_agents.AGENTS
        }
        self._cancel_events: dict[str, dict[str, threading.Event]] = {
            a["key"]: {} for a in thinking_agents.AGENTS
        }
        # Wall-clock epoch (seconds) stamped when each run claims its in-flight
        # slot, so a UI reattach after page navigation can show accurate elapsed
        # time instead of restarting from zero. Maintained in lockstep with
        # _in_flight / _cancel_events (claimed, released, and ownership-guarded
        # at the same points).
        self._started_at: dict[str, dict[str, float]] = {
            a["key"]: {} for a in thinking_agents.AGENTS
        }
        self._threads: dict[str, set[threading.Thread]] = {
            a["key"]: set() for a in thinking_agents.AGENTS
        }
        # Accumulated streamed tokens per in-flight run, so the poll endpoint can
        # surface partial text (only the summary agent fills this — structured
        # agents don't stream). Guarded by its own lock, kept off _manifest_lock
        # so per-token appends never contend with manifest reads / poll handlers.
        self._partial: dict[str, dict[str, list[str]]] = {
            a["key"]: {} for a in thinking_agents.AGENTS
        }
        self._partial_lock = threading.Lock()

    def partial_text(self, participant: str, agent_key: str) -> str:
        """Return the tokens streamed so far for an in-flight run (``""`` if none)."""
        with self._partial_lock:
            return "".join(self._partial.get(agent_key, {}).get(participant, []))

    def is_generating(self, participant: str, agent_key: str) -> bool:
        """Return True if *agent_key* is currently running for *participant*."""
        return participant in self._in_flight.get(agent_key, set())

    def started_at(self, participant: str, agent_key: str) -> float | None:
        """Epoch seconds when *agent_key* started running for *participant*.

        ``None`` when no run is in flight. Lets the frontend seed its elapsed
        clock from the true start so navigating away and back doesn't reset it.
        """
        return self._started_at.get(agent_key, {}).get(participant)

    def stop(self, agent_key: str, participant: str) -> bool:
        """Abort an in-flight run. Returns True iff a run was actually stopped.

        Releases the in-flight slot immediately so subsequent ``is_generating``
        polls (e.g. from the UI) flip to idle without waiting for the daemon
        thread's ``finally`` block to fire.
        """
        with self._lock:
            if participant not in self._in_flight.get(agent_key, set()):
                return False
            event = self._cancel_events.get(agent_key, {}).get(participant)
            self._in_flight[agent_key].discard(participant)
            self._started_at.get(agent_key, {}).pop(participant, None)
        with self._partial_lock:
            self._partial.get(agent_key, {}).pop(participant, None)
        if event is not None:
            event.set()
        return True

    def next_eligible(
        self,
        participant: str,
        force: bool = False,
        skip: set[str] | None = None,
    ) -> thinking_agents.Agent | None:
        """Find the first agent that should run next for *participant*.

        An agent is eligible when:
          - it is enabled (unless *force* is True — manual triggers bypass this),
          - its result is not already on the entry,
          - all of its dependencies are satisfied,
          - it is not already running for this participant,
          - its key is not in *skip*.

        *skip* exists for the chain advance after a run that stored nothing: the
        agent's field is still empty, so it would otherwise be picked again
        immediately and loop.
        """
        with self._lock:
            entry = _manifest.get("source_transcripts", {}).get(participant)
            if not entry or not entry.get("segments"):
                return None
            for agent in thinking_agents.AGENTS:
                if skip and agent["key"] in skip:
                    continue
                if not force and not _agent_enabled(agent):
                    continue
                if entry.get(agent["manifest_field"]):
                    continue
                if participant in self._in_flight.get(agent["key"], set()):
                    continue
                if not _agent_dependencies_met(agent, entry):
                    continue
                return agent
        return None

    def run_agent(
        self,
        agent_key: str,
        participant: str,
        force: bool = False,
        skip: set[str] | None = None,
    ) -> None:
        """Spawn a daemon thread to run a single agent for *participant*.

        On success, the agent's result is written to the manifest and the chain
        advances to the next eligible agent. Guards against double-spawning via
        the in-flight set. Pass *force* to bypass the config enabled check
        (used by manual triggers from the UI).

        *skip* carries the keys of agents that already had their turn in this
        chain and stored nothing; it is forwarded to the next ``run_chain`` so
        the exclusion accumulates. Without that, two agents that both fail to
        commit take turns re-qualifying each other forever — each one's field
        stays empty and it leaves the in-flight set when its thread ends.
        """
        agent = thinking_agents.get_agent(agent_key)
        if agent is None:
            return
        if not force and not _agent_enabled(agent):
            return
        chain_skip = set(skip or ())

        cancel_event = threading.Event()
        # Atomically check-and-claim the in-flight slot so two near-simultaneous
        # chain triggers (e.g. user click + auto-chain from a completed
        # dependency) can't both spawn a thread for the same participant.
        with self._lock:
            if participant in self._in_flight[agent_key]:
                return
            self._in_flight[agent_key].add(participant)
            self._cancel_events[agent_key][participant] = cancel_event
            self._started_at[agent_key][participant] = datetime.now(UTC).timestamp()
        with self._partial_lock:
            self._partial[agent_key][participant] = []

        # If a Stop just scheduled an unload for this model, cancel it — the
        # next request would only force a reload.
        model = _agent_model(agent_key)
        if model:
            _cancel_pending_unload(model)

        def _run() -> None:
            try:
                with self._lock:
                    entry = _manifest.get("source_transcripts", {}).get(participant)
                    if not entry or not entry.get("segments"):
                        return
                    # Snapshot the entry so the agent does not hold the lock
                    # during the (potentially slow) model call.
                    snapshot = dict(entry)

                def _sink(tok: str) -> None:
                    with self._partial_lock:
                        buf = self._partial.get(agent_key, {}).get(participant)
                        if buf is not None:
                            buf.append(tok)

                result = agent["run"](snapshot, cancel_event, _sink)
                # Defense in depth: if the model finished in the same tick as a
                # Stop click, drop the result and skip the chain advance.
                if cancel_event.is_set():
                    return
                committed = False
                if result is not None:
                    with self._lock:
                        # Re-check the cancel event inside the lock so a Stop
                        # that arrives between the snapshot read above and this
                        # commit cannot race past us and leave a stale result
                        # behind.
                        if cancel_event.is_set():
                            return
                        entry = _manifest.get("source_transcripts", {}).get(participant)
                        if entry is not None:
                            entry[agent["manifest_field"]] = result
                            committed = True
                if committed:
                    _persist_manifest()
                    # Chain into the next eligible agent (e.g. summary →
                    # citations). The auto-advance is always force=False: a manual
                    # trigger forces only the one agent the user asked for (it ran
                    # above with its own *force*); downstream agents should still
                    # respect their enabled config. Otherwise a single-agent
                    # regenerate (e.g. Citations) would cross-trigger a disabled
                    # sibling (e.g. Friction), since they share depends_on=["summary"].
                    self.run_chain(participant, force=False, skip=chain_skip)
                else:
                    # The run finished but stored nothing (the model call failed,
                    # or its inputs were empty). Its field stays empty so it is
                    # still eligible on the next chain entry — but this chain must
                    # advance *past* it, or a sibling that does not depend on it
                    # (friction needs only the summary) would never start. Skip
                    # this agent for that lookup: without it next_eligible picks
                    # the same empty field straight back and the chain spins.
                    self.run_chain(
                        participant, force=False, skip=chain_skip | {agent_key}
                    )
            except Exception as exc:
                utils.warning_print(
                    f"{agent_key} generation failed for {participant}: {exc}"
                )
                # A raising agent stores nothing either, so the chain has to move
                # on for the same reason as the else-branch above. Skip it after a
                # cancel, though: Stop is not a failure to route around, and the
                # early returns above deliberately leave the chain where it is.
                if not cancel_event.is_set():
                    self.run_chain(
                        participant, force=False, skip=chain_skip | {agent_key}
                    )
            finally:
                with self._lock:
                    # Only clean up if the slot is still ours. A
                    # Stop-then-Regenerate cycle can claim the slot for a
                    # successor run between when stop() sets our cancel_event
                    # and when we reach this finally — in which case
                    # _cancel_events[...][participant] now holds the
                    # successor's event, and clobbering would orphan it
                    # (uncancellable, invisible to is_generating).
                    slot = self._cancel_events.get(agent_key, {}).get(participant)
                    if slot is cancel_event:
                        self._in_flight[agent_key].discard(participant)
                        self._cancel_events[agent_key].pop(participant, None)
                        self._started_at[agent_key].pop(participant, None)
                        with self._partial_lock:
                            self._partial.get(agent_key, {}).pop(participant, None)
                    self._threads[agent_key].discard(t)

        t = threading.Thread(
            target=_run,
            daemon=True,
            name=f"{agent['thread_name_prefix']}-{participant}",
        )
        with self._lock:
            self._threads[agent_key].add(t)
        t.start()

    def run_chain(
        self, participant: str, force: bool = False, skip: set[str] | None = None
    ) -> None:
        """Advance the thinking-agent chain for *participant* by one step.

        Starts the first eligible agent (if any). When that agent completes,
        ``run_agent`` re-enters this method to start the next step. Pass
        *force* to bypass config enable checks for manual/UI-triggered runs, and
        *skip* to exclude agent keys from consideration (see ``next_eligible``).
        """
        agent = self.next_eligible(participant, force=force, skip=skip)
        if agent is not None:
            self.run_agent(agent["key"], participant, force=force, skip=skip)


_orchestrator = AgentOrchestrator(_manifest_lock)


# ---- State initialization ----


def _init_transcripts_state(sheet_context: Any = None) -> None:
    """Initialize module-level state for Transcript routes.

    Loads manifest, resolves participant video paths, and starts the
    background worker thread.

    Keeps a reference to *sheet_context* so :func:`_refresh_participants` can
    re-derive the participant union when the input directory changes.
    ``server._swap_worksheet`` re-inits this blueprint on every sheet swap, so
    the stored reference is replaced rather than going stale.
    """
    global _manifest, _worker, _input_dir, _participant_source

    _input_dir = str(utils.get_effective_input_dir())
    _manifest = transcripts.load_transcripts_manifest()
    _merged_task_ids.clear()
    _pending_chain_pids.clear()

    # mtime None forces the first _refresh_participants() call to build.
    _participant_source = {"sheet_context": sheet_context, "dir": "", "mtime": None}
    _refresh_participants()

    _worker = transcripts.TranscriptWorker()
    _worker.on_task_complete = _on_task_complete
    _worker.restore_tasks(_manifest.get("tasks", []))
    # Restored completed tasks already have their segments saved in the manifest
    # (possibly with later edits); seed the merged set so a startup persist does
    # not re-apply their frozen results over those segments.
    for task in _worker.get_all_tasks():
        if task["status"] == transcripts.TASK_STATUS_COMPLETED and task.get("result"):
            _merged_task_ids.add(task["id"])
    _worker.start()

    # Reclaim a stale empty manifest left by a prior abandoned session: the
    # guarded save removes the file when empty, idempotent rewrite otherwise.
    _persist_manifest()
