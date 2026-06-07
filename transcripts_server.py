# -*- coding: utf-8 -*-
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
  GET  /api/summary/<participant>                - AI-generated transcript summary
  POST /api/summary/<participant>/regenerate    - re-trigger AI summary generation
  POST /api/summary/<participant>/stop          - flag in-flight summary for discard
  PUT  /api/summary/<participant>               - save user-edited summary
  GET  /api/citations/<participant>             - citation refs for summary sentences
  POST /api/citations/<participant>/regenerate  - re-trigger citation generation
  POST /api/citations/<participant>/stop        - flag in-flight citations for discard
  GET  /api/corrections                           - list all study-local corrections
  POST /api/corrections                           - add a correction manually
  DELETE /api/corrections/<id>                    - remove a correction
  GET  /api/marks                                 - list all marks with resolved segment data
  POST /api/marks                                 - create marks for segments
  PUT  /api/marks/<id>                            - update a mark's category or label
  DELETE /api/marks/<id>                          - remove a mark (or bulk delete with JSON body)
  GET  /api/search?q=<query>                      - keyword search across all participants
  POST /api/transcribe                            - enqueue participant(s) for transcription
  GET  /api/transcribe/status                     - poll transcription task status
  DELETE /api/transcribe/<task_id>                 - cancel or dismiss a transcription task
  POST /api/transcribe/warmup                     - background-load Whisper when prewarm is enabled
  GET  /api/transcribe/model-status               - whether the Whisper model is loaded or warming
"""

import os
import tempfile
import threading
import uuid
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from flask import Blueprint, Response, jsonify, request

import config
import files
import ollama_client
import thinking_agents
import transcripts
import utils
import video

FlaskResponse = Response | tuple[Response, int]

# ---- Module-level state (set once by _init_transcripts_state) ----

_manifest: dict[str, Any] = {}
_worker: transcripts.TranscriptWorker | None = None
_input_dir: str = ""
_participants: list[dict[str, Any]] = []
_manifest_lock = threading.Lock()
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
    model = getattr(config, agent["model_config_key"], None)
    if not model:
        model = getattr(config, "OLLAMA_SUMMARY_MODEL", None)
    return model


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


@transcripts_bp.route("/api/participants")
def api_participants() -> FlaskResponse:
    """List discovered source videos with transcription status."""
    import viewer

    result = []
    with _manifest_lock:
        src = _manifest.get("source_transcripts", {})
        for p in _participants:
            pid = p["id"]
            entry = src.get(pid, {})
            has_transcript = bool(entry.get("segments"))
            video_version: int | None = None
            if p["has_video"]:
                try:
                    video_version = Path(p["video_path"]).stat().st_mtime_ns
                except OSError:
                    video_version = None
            info: dict[str, Any] = {
                "id": pid,
                "video_path": p["video_path"],
                "has_video": p["has_video"],
                "has_transcript": has_transcript,
                "segment_count": len(entry.get("segments", [])),
                "video_filename": Path(p["video_path"]).name,
                "video_version": video_version,
                "agents": {
                    "transcription": _step_state_transcription(entry),
                    "summary": _step_state_agent(pid, entry, "summary"),
                    "citations": _step_state_agent(pid, entry, "citations"),
                    "friction": _step_state_agent(pid, entry, "friction"),
                },
            }
            if has_transcript:
                info["language"] = entry.get("language", "")
                info["model"] = entry.get("model", "")
                info["transcribed_at"] = entry.get("transcribed_at", "")
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

    return jsonify(
        {
            "ok": True,
            "participants": result,
            "transcribe_prewarm": _transcribe_prewarm_setting(),
        }
    )


# ---- Transcript data ----


@transcripts_bp.route("/api/transcript/<participant>")
def api_transcript(participant: str) -> FlaskResponse:
    """Return full transcript segments for a participant (corrections applied)."""
    with _manifest_lock:
        src = _manifest.get("source_transcripts", {})
        entry = src.get(participant)
    if not entry or not entry.get("segments"):
        return jsonify({"ok": False, "error": "No transcript for participant"}), 404

    raw_segments = entry["segments"]
    corrections = _manifest.get("corrections", [])

    # Apply corrections to get corrected text
    corrected_segments = transcripts.apply_corrections(raw_segments, corrections)

    # Build marks-by-segment-id lookup
    marks_by_seg: dict[str, list[dict[str, Any]]] = {}
    for mark in _manifest.get("marks", []):
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

    return jsonify(
        {
            "ok": True,
            "participant": participant,
            "segments": segments,
            "language": entry.get("language", ""),
            "model": entry.get("model", ""),
            "transcribed_at": entry.get("transcribed_at", ""),
        }
    )


@transcripts_bp.route("/api/transcript/<participant>/segment", methods=["PUT"])
def api_edit_segment(participant: str) -> FlaskResponse:
    """Edit a segment's text. Creates a correction entry automatically."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "Missing JSON body"}), 400

    segment_id = data.get("segment_id", "")
    new_text = data.get("text", "").strip()
    if not segment_id or not new_text:
        return jsonify({"ok": False, "error": "segment_id and text required"}), 400

    with _manifest_lock:
        src = _manifest.get("source_transcripts", {})
        entry = src.get(participant)
        if not entry:
            return jsonify({"ok": False, "error": "No transcript for participant"}), 404

        # Find the raw segment by ID
        raw_seg = None
        for seg in entry.get("segments", []):
            if seg.get("id") == segment_id:
                raw_seg = seg
                break
        if raw_seg is None:
            return jsonify({"ok": False, "error": "Segment not found"}), 404

        original_text = raw_seg["text"]
        if original_text == new_text:
            return jsonify({"ok": True, "correction": None})

        # Create correction
        correction = {
            "id": f"c_{uuid.uuid4().hex[:8]}",
            "from": original_text,
            "to": new_text,
            "created": datetime.now(timezone.utc).isoformat(),
        }
        _manifest.setdefault("corrections", []).append(correction)
        _mark_friction_stale(entry)  # edited segment text invalidates friction scores

    _persist_manifest()
    return jsonify({"ok": True, "correction": correction})


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

    corrected = transcripts.apply_corrections(segments_snapshot, corrections_snapshot)
    result = transcripts.TranscriptResult(
        segments=corrected,
        language=language,
        source_file=source_file,
        model=model,
    )
    vtt_text = transcripts._format_vtt(result)
    return Response(vtt_text, content_type="text/vtt")


# ---- Embed subtitles into video ----


def _video_path_for_participant(participant: str) -> str | None:
    """Return the source video path for *participant*, or None if unknown."""
    for p in _participants:
        if p["id"] == participant:
            return p["video_path"]
    return None


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

    video_path = _video_path_for_participant(participant)
    if not video_path or not Path(video_path).is_file():
        return {
            "participant": participant,
            "ok": False,
            "error": "Source video not found",
        }

    corrected = transcripts.apply_corrections(segments_snapshot, corrections_snapshot)
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

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".srt", delete=False, encoding="utf-8"
    )
    try:
        tmp.write(srt_text)
        tmp.close()
        ok = video.mux_subtitles(
            str(video_path),
            tmp.name,
            output_path,
            track_language=language or "und",
        )
    finally:
        try:
            os.unlink(tmp.name)
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
        return jsonify({"ok": False, "error": outcome["error"]}), status
    return jsonify(
        {
            "ok": True,
            "output_path": outcome["output_path"],
            "output_filename": outcome["output_filename"],
            "output_dir": str(output_dir),
        }
    )


@transcripts_bp.route("/api/embed-all-subtitles", methods=["POST"])
def api_embed_all_subtitles() -> FlaskResponse:
    """Mux every participant with a transcript into a subtitled copy."""
    output_dir = Path(utils.get_effective_output_dir())
    with _manifest_lock:
        src = _manifest.get("source_transcripts", {})
        targets = [pid for pid, entry in src.items() if entry.get("segments")]
    if not targets:
        return jsonify(
            {"ok": False, "error": "No transcripts available to embed."}
        ), 404
    results = [_embed_subtitle_for_participant(pid, output_dir) for pid in targets]
    return jsonify(
        {
            "ok": True,
            "results": results,
            "output_dir": str(output_dir),
        }
    )


# ---- AI Summary ----


@transcripts_bp.route("/api/summary/<participant>")
def api_summary(participant: str) -> FlaskResponse:
    """Return AI-generated summary, generation status, or 404."""
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        if not entry or not entry.get("summary"):
            if _orchestrator.is_generating(participant, "summary"):
                return jsonify({"ok": False, "generating": True})
            return jsonify({"ok": False}), 404
        summary = entry["summary"]
        citations = entry.get("citations")
    resp: dict[str, Any] = {"ok": True, "summary": summary}
    if citations:
        resp["citations"] = citations
    resp["citations_generating"] = _orchestrator.is_generating(participant, "citations")
    return jsonify(resp)


@transcripts_bp.route("/api/summary/<participant>/regenerate", methods=["POST"])
def api_summary_regenerate(participant: str) -> FlaskResponse:
    """Clear existing summary and re-trigger AI generation.

    Manual trigger: runs even when config.OLLAMA_SUMMARY_ENABLED is False
    so the frontend's per-participant controls can force a run.
    """
    if _orchestrator.is_generating(participant, "summary"):
        return jsonify({"ok": True, "generating": True})
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        if not entry or not entry.get("segments"):
            return jsonify({"ok": False, "error": "No transcript found"}), 404
        entry["summary"] = ""
        entry.pop("citations", None)
    _persist_manifest()
    _orchestrator.run_chain(participant, force=True)
    return jsonify({"ok": True, "generating": True})


@transcripts_bp.route("/api/summary/<participant>/stop", methods=["POST"])
def api_summary_stop(participant: str) -> FlaskResponse:
    """Abort an in-flight summary run.

    Sets the cancel event so the streaming Ollama call closes its response
    promptly, freeing the model for another run. The UI flips to idle
    immediately. After a short delay, the model is unloaded from memory if
    no new run has started in the meantime.
    """
    if _orchestrator.stop("summary", participant):
        model = _agent_model("summary")
        if model:
            _schedule_model_unload(model)
    return jsonify({"ok": True, "running": False})


@transcripts_bp.route("/api/summary/<participant>", methods=["PUT"])
def api_summary_save(participant: str) -> FlaskResponse:
    """Save a user-edited summary for a participant."""
    data = request.get_json(silent=True)
    if not data or not data.get("summary", "").strip():
        return jsonify({"ok": False, "error": "Summary text is required"}), 400
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        if not entry:
            return jsonify({"ok": False, "error": "Participant not found"}), 404
        entry["summary"] = data["summary"].strip()
        entry.pop("citations", None)  # invalidate citations on edit
        _mark_friction_stale(entry)  # summary text feeds the friction prompt
    _persist_manifest()
    return jsonify({"ok": True})


# ---- Citations ----


@transcripts_bp.route("/api/citations/<participant>")
def api_citations(participant: str) -> FlaskResponse:
    """Return citation refs for a participant's summary, or generation status."""
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        if not entry:
            return jsonify({"ok": False}), 404
        citations = entry.get("citations")
    if citations:
        return jsonify({"ok": True, "citations": citations})
    if _orchestrator.is_generating(participant, "citations"):
        return jsonify({"ok": False, "generating": True})
    return jsonify({"ok": False}), 404


@transcripts_bp.route("/api/citations/<participant>/regenerate", methods=["POST"])
def api_citations_regenerate(participant: str) -> FlaskResponse:
    """Clear existing citations and re-trigger citation generation (Pass 2).

    Manual trigger: runs even when config.OLLAMA_SUMMARY_ENABLED is False.
    """
    if _orchestrator.is_generating(participant, "citations"):
        return jsonify({"ok": True, "generating": True})
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        if not entry or not entry.get("summary") or not entry.get("segments"):
            return jsonify(
                {"ok": False, "error": "No summary or transcript found"}
            ), 404
        entry.pop("citations", None)
    _persist_manifest()
    _orchestrator.run_agent("citations", participant, force=True)
    return jsonify({"ok": True, "generating": True})


@transcripts_bp.route("/api/citations/<participant>/stop", methods=["POST"])
def api_citations_stop(participant: str) -> FlaskResponse:
    """Abort an in-flight citations run.

    Sets the cancel event so the streaming Ollama call closes its response
    promptly, freeing the model for another run. The UI flips to idle
    immediately. After a short delay, the model is unloaded from memory if
    no new run has started in the meantime.
    """
    if _orchestrator.stop("citations", participant):
        model = _agent_model("citations")
        if model:
            _schedule_model_unload(model)
    return jsonify({"ok": True, "running": False})


# ---- Friction ----


@transcripts_bp.route("/api/friction/<participant>")
def api_friction(participant: str) -> FlaskResponse:
    """Return friction analysis for a participant, or generation status.

    The returned object carries its own ``stale`` flag (set after segment/summary
    edits) so the UI can prompt for a re-run without a separate request.
    """
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        if not entry:
            return jsonify({"ok": False}), 404
        friction_data = entry.get("friction")
    if friction_data:
        return jsonify({"ok": True, "friction": friction_data})
    if _orchestrator.is_generating(participant, "friction"):
        return jsonify({"ok": False, "generating": True})
    return jsonify({"ok": False}), 404


@transcripts_bp.route("/api/friction/<participant>/regenerate", methods=["POST"])
def api_friction_regenerate(participant: str) -> FlaskResponse:
    """Clear friction and re-trigger analysis (Pass 3).

    Manual trigger: runs even when config.OLLAMA_FRICTION_ENABLED is False so the
    frontend's per-participant controls can force a run. Requires a summary.
    """
    if _orchestrator.is_generating(participant, "friction"):
        return jsonify({"ok": True, "generating": True})
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
        if not entry or not entry.get("summary") or not entry.get("segments"):
            return jsonify(
                {"ok": False, "error": "No summary or transcript found"}
            ), 404
        entry.pop("friction", None)
    _persist_manifest()
    _orchestrator.run_agent("friction", participant, force=True)
    return jsonify({"ok": True, "generating": True})


@transcripts_bp.route("/api/friction/<participant>/stop", methods=["POST"])
def api_friction_stop(participant: str) -> FlaskResponse:
    """Abort an in-flight friction run.

    Sets the cancel event so the streaming Ollama call closes its response
    promptly, freeing the model for another run. The UI flips to idle
    immediately. After a short delay, the model is unloaded from memory if no new
    run has started in the meantime.
    """
    if _orchestrator.stop("friction", participant):
        model = _agent_model("friction")
        if model:
            _schedule_model_unload(model)
    return jsonify({"ok": True, "running": False})


# ---- Corrections ----


@transcripts_bp.route("/api/corrections")
def api_corrections_list() -> FlaskResponse:
    """List all study-local corrections."""
    with _manifest_lock:
        corrections = list(_manifest.get("corrections", []))
    return jsonify({"ok": True, "corrections": corrections})


@transcripts_bp.route("/api/corrections", methods=["POST"])
def api_corrections_add() -> FlaskResponse:
    """Add a manual correction."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "Missing JSON body"}), 400

    from_text = data.get("from", "").strip()
    to_text = data.get("to", "").strip()
    if not from_text or not to_text:
        return jsonify({"ok": False, "error": "'from' and 'to' required"}), 400

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
                _persist_manifest()
                return jsonify(
                    {"ok": True, "correction": None, "removed": chained["id"]}
                )
            chained["to"] = to_text
            correction = chained
        else:
            correction = {
                "id": f"c_{uuid.uuid4().hex[:8]}",
                "from": from_text,
                "to": to_text,
                "created": datetime.now(timezone.utc).isoformat(),
            }
            corrections.append(correction)

    _persist_manifest()
    return jsonify({"ok": True, "correction": correction})


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

    if removed == 0:
        return jsonify({"ok": False, "error": "Correction not found"}), 404

    _persist_manifest()
    return jsonify({"ok": True})


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
        raw_seg = segments[idx]
        corrections = _manifest.get("corrections", [])
        corrected = transcripts.apply_corrections([raw_seg], corrections)
        seg = corrected[0]
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


@transcripts_bp.route("/api/marks")
def api_marks_list() -> FlaskResponse:
    """List all marks, enriched with resolved segment data."""
    # Build partial lookup outside _manifest_lock (get_all_tasks acquires worker lock)
    partial_lookup = _build_partial_lookup()
    with _manifest_lock:
        raw_marks = list(_manifest.get("marks", []))
        resolved = [_resolve_mark(m, partial_lookup) for m in raw_marks]
    return jsonify(
        {
            "ok": True,
            "marks": resolved,
            "categories": config.MARK_CATEGORIES,
        }
    )


@transcripts_bp.route("/api/marks", methods=["POST"])
def api_marks_add() -> FlaskResponse:
    """Create marks for one or more segments."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "Missing JSON body"}), 400

    segment_ids = data.get("segment_ids", [])
    if not segment_ids:
        return jsonify({"ok": False, "error": "segment_ids required"}), 400

    category = data.get("category") or None
    label = data.get("label") or None
    now = datetime.now(timezone.utc).isoformat()

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
                created.append(m)
            else:
                m = {
                    "id": f"m_{uuid.uuid4().hex[:8]}",
                    "segment_id": sid,
                    "category": category,
                    "label": label,
                    "created": now,
                }
                marks.append(m)
                existing_by_seg[sid] = m
                created.append(m)

    _persist_manifest()
    return jsonify({"ok": True, "marks": created})


@transcripts_bp.route("/api/marks/<mark_id>", methods=["PUT"])
def api_marks_update(mark_id: str) -> FlaskResponse:
    """Update a mark's category or label."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "Missing JSON body"}), 400

    with _manifest_lock:
        marks = _manifest.get("marks", [])
        target = None
        for m in marks:
            if m.get("id") == mark_id:
                target = m
                break
        if not target:
            return jsonify({"ok": False, "error": "Mark not found"}), 404

        if "category" in data:
            target["category"] = data["category"] or None
        if "label" in data:
            target["label"] = data["label"] or None

    _persist_manifest()
    return jsonify({"ok": True, "mark": target})


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
        return jsonify({"ok": False, "error": "Mark not found"}), 404

    _persist_manifest()
    return jsonify({"ok": True, "removed": removed})


# ---- Search ----


@transcripts_bp.route("/api/search")
def api_search() -> FlaskResponse:
    """Keyword search across all transcribed participants."""
    query = request.args.get("q", "").strip()
    if not query:
        return jsonify(
            {
                "ok": True,
                "query": "",
                "total_count": 0,
                "results": [],
                "counts_by_participant": {},
            }
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
            corrected = transcripts.apply_corrections(raw_segments, corrections)
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
    return jsonify(
        {
            "ok": True,
            "query": query,
            "total_count": total,
            "results": results,
            "counts_by_participant": counts,
        }
    )


# ---- Transcription queue ----


@transcripts_bp.route("/api/transcribe/warmup", methods=["POST"])
def api_transcribe_warmup() -> FlaskResponse:
    """Background-load the Whisper model when automatic prewarm is not ``off``."""
    if _transcribe_prewarm_setting() == "off":
        return jsonify(
            {
                "ok": True,
                "skipped": True,
                "reason": "prewarm_disabled",
            }
        )

    if transcripts.is_transcription_model_loaded():
        return jsonify({"ok": True, "already_loaded": True})

    global _transcript_model_warming  # noqa: PLW0603

    with _transcript_model_warming_lock:
        if _transcript_model_warming:
            return jsonify({"ok": True, "already_warming": True})
        if transcripts.is_transcription_model_loaded():
            return jsonify({"ok": True, "already_loaded": True})
        _transcript_model_warming = True

    def _run_warmup() -> None:
        global _transcript_model_warming  # noqa: PLW0603
        try:
            transcripts.warmup_transcription_model()
        finally:
            with _transcript_model_warming_lock:
                _transcript_model_warming = False

    threading.Thread(
        target=_run_warmup, daemon=True, name="transcript-model-warmup"
    ).start()
    return jsonify({"ok": True, "started": True})


@transcripts_bp.route("/api/transcribe/model-status")
def api_transcribe_model_status() -> FlaskResponse:
    """Report whether faster-whisper is loaded or a warm-up is in progress."""
    with _transcript_model_warming_lock:
        warming = _transcript_model_warming
    return jsonify(
        {
            "ok": True,
            "loaded": transcripts.is_transcription_model_loaded(),
            "warming": warming,
            "model": config.TRANSCRIBE_MODEL,
            "prewarm": _transcribe_prewarm_setting(),
        }
    )


@transcripts_bp.route("/api/transcribe", methods=["POST"])
def api_transcribe() -> FlaskResponse:
    """Enqueue participant(s) for background transcription."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "Missing JSON body"}), 400

    participant_ids = data.get("participants", [])
    force = data.get("force", False)
    overrides = data.get("overrides") or {}

    if not participant_ids:
        return jsonify({"ok": False, "error": "No participants specified"}), 400

    # Build a lookup of available participants
    available = {p["id"]: p for p in _participants}
    enqueued = []

    with _manifest_lock:
        src = _manifest.get("source_transcripts", {})
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
            task = transcripts.create_transcript_task(
                pid,
                p["video_path"],
                model=model_override,
                language=language_override,
            )
            if _worker:
                _worker.enqueue(task)
            enqueued.append(
                {
                    "id": task["id"],
                    "participant": pid,
                    "status": task["status"],
                }
            )

    return jsonify({"ok": True, "tasks": enqueued})


@transcripts_bp.route("/api/transcribe/status")
def api_transcribe_status() -> FlaskResponse:
    """Poll transcription task status."""
    tasks = []
    if _worker:
        for t in _worker.get_all_tasks():
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
                task_info["partial_segments"] = t.get("partial_segments", [])
            tasks.append(task_info)
    return jsonify(
        {
            "ok": True,
            "tasks": tasks,
            "worker_alive": _worker.is_alive if _worker else False,
        }
    )


@transcripts_bp.route("/api/transcribe/<task_id>", methods=["DELETE"])
def api_transcribe_cancel(task_id: str) -> FlaskResponse:
    """Cancel or dismiss a transcription task. ?dismiss=true fully removes it."""
    if not _worker:
        return jsonify({"ok": False, "error": "Worker not initialized"}), 500

    if request.args.get("dismiss") == "true":
        if not _worker.remove_task(task_id):
            return jsonify({"ok": False, "error": "Task not found"}), 404
        _persist_manifest()
        return jsonify({"ok": True})

    if _worker.cancel(task_id):
        _persist_manifest()
        return jsonify({"ok": True})

    return jsonify({"ok": False, "error": "Task not found or already finished"}), 400


# ---- Manifest persistence ----


def _persist_manifest() -> None:
    """Persist the current manifest state through a single synchronized path."""
    with _manifest_lock:
        _do_persist()


def _do_persist() -> None:
    """Persist manifest to disk - caller must hold _manifest_lock."""
    # Collect completed task results into source_transcripts (merge, not replace,
    # so that extra keys like "summary" added by other threads are preserved)
    if _worker:
        for task in _worker.get_all_tasks():
            if task["status"] == transcripts.TASK_STATUS_COMPLETED and task.get(
                "result"
            ):
                pid = task["participant"]
                src = _manifest.setdefault("source_transcripts", {})
                existing = src.get(pid, {})
                existing.update(task["result"])
                src[pid] = existing
        _manifest["tasks"] = _worker.get_all_tasks()

    transcripts.save_transcripts_manifest(
        _manifest.get("source_transcripts", {}),
        _manifest.get("corrections", []),
        marks=_manifest.get("marks"),
    )


# ---- Thinking-agent orchestration ----


def _on_task_complete() -> None:
    """Persist manifest, then trigger the thinking-agent chain for newly
    completed participants (those without a summary yet)."""
    newly_completed: list[str] = []
    if _worker:
        for task in _worker.get_all_tasks():
            if task["status"] == transcripts.TASK_STATUS_COMPLETED and task.get(
                "result"
            ):
                pid = task["participant"]
                with _manifest_lock:
                    existing = _manifest.get("source_transcripts", {}).get(pid, {})
                    if not existing.get("summary"):
                        newly_completed.append(pid)

    _persist_manifest()

    for pid in newly_completed:
        _orchestrator.run_chain(pid)


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
        self._threads: dict[str, set[threading.Thread]] = {
            a["key"]: set() for a in thinking_agents.AGENTS
        }

    def is_generating(self, participant: str, agent_key: str) -> bool:
        """Return True if *agent_key* is currently running for *participant*."""
        return participant in self._in_flight.get(agent_key, set())

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
        if event is not None:
            event.set()
        return True

    def next_eligible(
        self, participant: str, force: bool = False
    ) -> thinking_agents.Agent | None:
        """Find the first agent that should run next for *participant*.

        An agent is eligible when:
          - it is enabled (unless *force* is True — manual triggers bypass this),
          - its result is not already on the entry,
          - all of its dependencies are satisfied,
          - it is not already running for this participant.
        """
        with self._lock:
            entry = _manifest.get("source_transcripts", {}).get(participant)
            if not entry or not entry.get("segments"):
                return None
            for agent in thinking_agents.AGENTS:
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

    def run_agent(self, agent_key: str, participant: str, force: bool = False) -> None:
        """Spawn a daemon thread to run a single agent for *participant*.

        On success, the agent's result is written to the manifest and the chain
        advances to the next eligible agent. Guards against double-spawning via
        the in-flight set. Pass *force* to bypass the config enabled check
        (used by manual triggers from the UI).
        """
        agent = thinking_agents.get_agent(agent_key)
        if agent is None:
            return
        if not force and not _agent_enabled(agent):
            return

        cancel_event = threading.Event()
        # Atomically check-and-claim the in-flight slot so two near-simultaneous
        # chain triggers (e.g. user click + auto-chain from a completed
        # dependency) can't both spawn a thread for the same participant.
        with self._lock:
            if participant in self._in_flight[agent_key]:
                return
            self._in_flight[agent_key].add(participant)
            self._cancel_events[agent_key][participant] = cancel_event

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
                result = agent["run"](snapshot, cancel_event)
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
                    # citations). Propagate *force* so a manual run cascades
                    # through disabled downstream agents too.
                    self.run_chain(participant, force=force)
            except Exception as exc:
                utils.warning_print(
                    f"{agent_key} generation failed for {participant}: {exc}"
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
                    self._threads[agent_key].discard(t)

        t = threading.Thread(
            target=_run,
            daemon=True,
            name=f"{agent['thread_name_prefix']}-{participant}",
        )
        with self._lock:
            self._threads[agent_key].add(t)
        t.start()

    def run_chain(self, participant: str, force: bool = False) -> None:
        """Advance the thinking-agent chain for *participant* by one step.

        Starts the first eligible agent (if any). When that agent completes,
        ``run_agent`` re-enters this method to start the next step. Pass
        *force* to bypass config enable checks for manual/UI-triggered runs.
        """
        agent = self.next_eligible(participant, force=force)
        if agent is not None:
            self.run_agent(agent["key"], participant, force=force)


_orchestrator = AgentOrchestrator(_manifest_lock)


# ---- State initialization ----


def _init_transcripts_state(
    sheet_context: Any = None,
    participant_list: list[str] | None = None,
) -> None:
    """Initialize module-level state for Transcript routes.

    Loads manifest, resolves participant video paths, and starts the
    background worker thread.
    """
    global _manifest, _worker, _input_dir, _participants  # noqa: PLW0603

    _input_dir = str(utils.get_effective_input_dir())
    _manifest = transcripts.load_transcripts_manifest()

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
        _participants = utils.discover_participant_videos(study_name)

    _worker = transcripts.TranscriptWorker()
    _worker.on_task_complete = _on_task_complete
    _worker.restore_tasks(_manifest.get("tasks", []))
    _worker.start()
