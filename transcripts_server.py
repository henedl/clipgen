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
  GET  /api/summary/<participant>                - AI-generated transcript summary
  POST /api/summary/<participant>/regenerate    - re-trigger AI summary generation
  PUT  /api/summary/<participant>               - save user-edited summary
  GET  /api/citations/<participant>             - citation refs for summary sentences
  POST /api/citations/<participant>/regenerate  - re-trigger citation generation
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
"""

import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from flask import Blueprint, Response, jsonify, request

import config
import files
import ollama_client
import transcripts
import utils

FlaskResponse = Response | tuple[Response, int]

# ---- Module-level state (set once by _init_transcripts_state) ----

_manifest: dict[str, Any] = {}
_worker: transcripts.TranscriptWorker | None = None
_input_dir: str = ""
_participants: list[dict[str, Any]] = []
_manifest_lock = threading.Lock()
_summary_threads: set[threading.Thread] = set()
_generating_summaries: set[str] = (
    set()
)  # participant IDs with in-flight summary generation
_citation_threads: set[threading.Thread] = set()
_generating_citations: set[str] = (
    set()
)  # participant IDs with in-flight citation generation

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
            info: dict[str, Any] = {
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
                info["has_summary"] = bool(entry.get("summary"))
            result.append(info)

    # Check for stale artifacts (transcript outdated relative to source)
    artifacts = viewer.load_manifest_artifacts()
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
        for art in artifacts:
            if art.get("participant") != pid:
                continue
            if not art.get("transcript"):
                continue
            art_tv = art.get("transcript_version", "")
            if not art_tv or art_tv < current_ta:
                has_stale = True
                break
        info["has_stale_artifacts"] = has_stale

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


# ---- AI Summary ----


@transcripts_bp.route("/api/summary/<participant>")
def api_summary(participant: str) -> FlaskResponse:
    """Return AI-generated summary, generation status, or 404."""
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
    if not entry or not entry.get("summary"):
        if participant in _generating_summaries:
            return jsonify({"ok": False, "generating": True})
        return jsonify({"ok": False}), 404
    resp: dict[str, Any] = {"ok": True, "summary": entry["summary"]}
    citations = entry.get("citations")
    if citations:
        resp["citations"] = citations
    resp["citations_generating"] = participant in _generating_citations
    return jsonify(resp)


@transcripts_bp.route("/api/summary/<participant>/regenerate", methods=["POST"])
def api_summary_regenerate(participant: str) -> FlaskResponse:
    """Clear existing summary and re-trigger AI generation."""
    if not config.OLLAMA_SUMMARY_ENABLED:
        return jsonify({"ok": False, "error": "Summary generation is disabled"}), 400
    if participant in _generating_summaries:
        return jsonify({"ok": True, "generating": True})
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
    if not entry or not entry.get("segments"):
        return jsonify({"ok": False, "error": "No transcript found"}), 404
    with _manifest_lock:
        entry["summary"] = ""
        entry.pop("citations", None)
    _persist_manifest()
    _trigger_summary_generation(participant)
    return jsonify({"ok": True, "generating": True})


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
    citations = entry.get("citations") if entry else None
    if citations:
        return jsonify({"ok": True, "citations": citations})
    if participant in _generating_citations:
        return jsonify({"ok": False, "generating": True})
    return jsonify({"ok": False}), 404


@transcripts_bp.route("/api/citations/<participant>/regenerate", methods=["POST"])
def api_citations_regenerate(participant: str) -> FlaskResponse:
    """Clear existing citations and re-trigger citation generation (Pass 2)."""
    if not config.OLLAMA_SUMMARY_ENABLED:
        return jsonify({"ok": False, "error": "Summary generation is disabled"}), 400
    if participant in _generating_citations:
        return jsonify({"ok": True, "generating": True})
    with _manifest_lock:
        entry = _manifest.get("source_transcripts", {}).get(participant)
    if not entry or not entry.get("summary") or not entry.get("segments"):
        return jsonify({"ok": False, "error": "No summary or transcript found"}), 404
    with _manifest_lock:
        entry.pop("citations", None)
    _persist_manifest()
    _trigger_citation_generation(participant)
    return jsonify({"ok": True, "generating": True})


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
            "categories": transcripts.MARK_CATEGORIES,
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


# ---- AI Summary generation ----


def _on_task_complete() -> None:
    """Persist manifest, then trigger summary generation for newly completed participants."""
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

    if not config.OLLAMA_SUMMARY_ENABLED:
        return
    for pid in newly_completed:
        _trigger_summary_generation(pid)


def _trigger_summary_generation(participant: str) -> None:
    """Spawn daemon thread to generate AI summary for a participant."""

    def _run() -> None:
        try:
            with _manifest_lock:
                entry = _manifest.get("source_transcripts", {}).get(participant)
                if not entry or not entry.get("segments"):
                    return
                segments = list(entry["segments"])
            summary = ollama_client.summarize_transcript(segments)
            if summary:
                with _manifest_lock:
                    entry = _manifest.get("source_transcripts", {}).get(participant)
                    if entry:
                        entry["summary"] = summary
                _persist_manifest()
                # Chain into Pass 2: citation linking
                _trigger_citation_generation(participant)
        except Exception as exc:
            utils.warning_print(f"Summary generation failed for {participant}: {exc}")
        finally:
            _generating_summaries.discard(participant)
            _summary_threads.discard(t)

    _generating_summaries.add(participant)
    t = threading.Thread(target=_run, daemon=True, name=f"summary-{participant}")
    _summary_threads.add(t)
    t.start()


# ---- Citation generation (Pass 2) ----


def _trigger_citation_generation(participant: str) -> None:
    """Spawn daemon thread to find citation refs for a participant's summary."""

    def _run() -> None:
        try:
            with _manifest_lock:
                entry = _manifest.get("source_transcripts", {}).get(participant)
                if not entry or not entry.get("summary") or not entry.get("segments"):
                    return
                summary = entry["summary"]
                segments = list(entry["segments"])
            citations = ollama_client.find_citations(summary, segments)
            if citations is not None:
                with _manifest_lock:
                    entry = _manifest.get("source_transcripts", {}).get(participant)
                    if entry:
                        entry["citations"] = citations
                _persist_manifest()
        except Exception as exc:
            utils.warning_print(f"Citation generation failed for {participant}: {exc}")
        finally:
            _generating_citations.discard(participant)
            _citation_threads.discard(t)

    _generating_citations.add(participant)
    t = threading.Thread(target=_run, daemon=True, name=f"citations-{participant}")
    _citation_threads.add(t)
    t.start()


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
