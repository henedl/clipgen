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
  GET  /api/corrections                           - list all study-local corrections
  POST /api/corrections                           - add a correction manually
  DELETE /api/corrections/<id>                    - remove a correction
  GET  /api/search?q=<query>                      - keyword search across all participants
  POST /api/transcribe                            - enqueue participant(s) for transcription
  GET  /api/transcribe/status                     - poll transcription task status
"""

from __future__ import annotations

import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from flask import Blueprint, Response, jsonify, request, send_from_directory

import files
import transcripts
import utils

FlaskResponse = Union[Response, Tuple[Response, int]]

# ---- Module-level state (set once by _init_transcripts_state) ----

_manifest: Dict[str, Any] = {}
_worker: Optional[transcripts.TranscriptWorker] = None
_input_dir: str = ""
_participants: List[Dict[str, Any]] = []
_manifest_lock = threading.Lock()

_assets_dir = utils.get_bundled_assets_root() / "assets" / "web"

# ---- Blueprint ----

transcripts_bp = Blueprint("transcripts", __name__)


# ---- Static file serving ----


@transcripts_bp.route("/")
def serve_index() -> FlaskResponse:
    return send_from_directory(_assets_dir, "transcripts.html")


@transcripts_bp.route("/<path:filename>")
def serve_static(filename: str) -> FlaskResponse:
    return send_from_directory(_assets_dir, filename)


@transcripts_bp.route("/icons/<path:filename>")
def serve_icons(filename: str) -> FlaskResponse:
    icons_dir = utils.get_bundled_assets_root() / "assets" / "icons"
    return send_from_directory(icons_dir, filename)


@transcripts_bp.route("/media/<path:filename>")
def serve_media(filename: str) -> FlaskResponse:
    if not _input_dir:
        return jsonify({"ok": False, "error": "Input directory not configured"}), 500
    return send_from_directory(_input_dir, filename)


# ---- Participants ----


@transcripts_bp.route("/api/participants")
def api_participants() -> FlaskResponse:
    """List discovered source videos with transcription status."""
    result = []
    with _manifest_lock:
        src = _manifest.get("source_transcripts", {})
        for p in _participants:
            pid = p["id"]
            entry = src.get(pid, {})
            has_transcript = bool(entry.get("segments"))
            info: Dict[str, Any] = {
                "id": pid,
                "video_path": p["video_path"],
                "has_video": p["has_video"],
                "has_transcript": has_transcript,
                "segment_count": len(entry.get("segments", [])),
                "video_filename": Path(p["video_path"]).name,
            }
            if has_transcript:
                info["language"] = entry.get("language", "")
                info["model"] = entry.get("model", "")
                info["transcribed_at"] = entry.get("transcribed_at", "")
            result.append(info)
    return jsonify({"ok": True, "participants": result})


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

    # Build response segments with corrected flag
    segments = []
    for raw, corrected in zip(raw_segments, corrected_segments):
        seg: Dict[str, Any] = {
            "id": raw.get("id", ""),
            "start": corrected["start"],
            "end": corrected["end"],
            "text": corrected["text"],
            "corrected": raw["text"] != corrected["text"],
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

    _persist_manifest()
    return jsonify({"ok": True, "correction": correction})


# ---- WebVTT ----


@transcripts_bp.route("/api/vtt/<participant>")
def api_vtt(participant: str) -> FlaskResponse:
    """Serve transcript as WebVTT for <track> subtitle support."""
    with _manifest_lock:
        src = _manifest.get("source_transcripts", {})
        entry = src.get(participant)
    if not entry or not entry.get("segments"):
        return Response("WEBVTT\n", content_type="text/vtt")

    corrections = _manifest.get("corrections", [])
    corrected = transcripts.apply_corrections(entry["segments"], corrections)

    result = transcripts.TranscriptResult(
        segments=corrected,
        language=entry.get("language", ""),
        source_file=entry.get("source_file", ""),
        model=entry.get("model", ""),
    )
    vtt_text = transcripts._format_vtt(result)
    return Response(vtt_text, content_type="text/vtt")


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
                return jsonify({"ok": True, "correction": None, "removed": chained["id"]})
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
    results: List[Dict[str, Any]] = []
    counts: Dict[str, int] = {}

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


@transcripts_bp.route("/api/transcribe", methods=["POST"])
def api_transcribe() -> FlaskResponse:
    """Enqueue participant(s) for background transcription."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"ok": False, "error": "Missing JSON body"}), 400

    participant_ids = data.get("participants", [])
    force = data.get("force", False)

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

            task = transcripts.create_transcript_task(pid, p["video_path"])
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
            tasks.append(
                {
                    "id": t["id"],
                    "participant": t["participant"],
                    "status": t["status"],
                    "progress": t["progress"],
                    "error": t.get("error"),
                    "created_at": t.get("created_at"),
                    "completed_at": t.get("completed_at"),
                }
            )
    return jsonify(
        {
            "ok": True,
            "tasks": tasks,
            "worker_alive": _worker.is_alive if _worker else False,
        }
    )


# ---- Manifest persistence ----


def _persist_manifest() -> None:
    """Persist the current manifest state through a single synchronized path."""
    with _manifest_lock:
        _do_persist()


def _do_persist() -> None:
    """Persist manifest to disk - caller must hold _manifest_lock."""
    # Collect completed task results into source_transcripts
    if _worker:
        for task in _worker.get_all_tasks():
            if task["status"] == transcripts.TASK_STATUS_COMPLETED and task.get(
                "result"
            ):
                pid = task["participant"]
                _manifest.setdefault("source_transcripts", {})[pid] = task["result"]
        _manifest["tasks"] = _worker.get_all_tasks()

    transcripts.save_transcripts_manifest(
        _manifest.get("source_transcripts", {}),
        _manifest.get("corrections", []),
    )


# ---- State initialization ----


def _init_transcripts_state(
    sheet_context: Any = None,
    participant_list: Optional[List[str]] = None,
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
    _worker.on_task_complete = _persist_manifest
    _worker.restore_tasks(_manifest.get("tasks", []))
    _worker.start()
