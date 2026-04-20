# -*- coding: utf-8 -*-
"""Combined Flask server for clipgen Studio, Insights, and Screenspace.

Entry point: start_combined_server(worksheet, port, default_page) registers
Studio, Insights, and Screenspace blueprints on one app at config.SERVER_PORT (8089).
Module-level state: _worksheet, _sheet_context, _generated_artifacts, _generated_reels
(initialized by _init_studio_state()).

Studio API endpoints (all under /studio/):
  GET  /api/sheet              – spreadsheet grid data (rows, participants, timestamps)
  POST /api/sheet/refresh      – re-fetch spreadsheet data from source (Google/Excel)
  GET  /api/thumbnail/<p>/<t>  – JPEG thumbnail frame from participant video
  POST /api/generate           – generate clip/screen/gif artifacts for specified cells
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
  GET  /api/settings           – read current config settings
  PUT  /api/settings           – update config settings
  GET  /api/status             – reports which interfaces are active (studio/insights/screenspace)
"""

import concurrent.futures
import hashlib
import json
import os
import sys
import threading
import webbrowser
from contextlib import contextmanager
from pathlib import Path
from collections.abc import Iterator
from typing import Any

from flask import (
    Blueprint,
    Flask,
    Response,
    jsonify,
    redirect,
    request,
)

import config
import files
import spreadsheet
import pipeline
import utils
import video
import viewer

FlaskResponse = Response | tuple[Response, int]

# ---- Module-level state (set once by _init_studio_state) ----

_worksheet: Any = None
_sheet_context: spreadsheet.SheetContext | None = None
_generated_artifacts: list[dict[str, Any]] = []
_generated_reels: list[dict[str, Any]] = []
_thumbnail_cache: dict[tuple, bytes] = {}
_reel_cancel_event = threading.Event()

# Snapshot config defaults before any settings file is loaded.
_settings_defaults: dict[str, Any] = {
    name: getattr(config, name) for name in getattr(config, "STUDIO_SETTINGS", {})
}


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


def _resolve_source_video(participant: str) -> Path | None:
    """Return the resolved path to a participant's source video, or None."""
    if _sheet_context is None:
        return None
    ctx = _sheet_context
    participants = spreadsheet.get_participant_list(
        ctx.header_row, ctx.id_cell, ctx.num_participants
    )
    if participant not in participants:
        return None
    p_idx = participants.index(participant)
    col_idx = ctx.id_cell.col + p_idx

    override = None
    if ctx.filename_row_idx is not None:
        row_data = ctx.sheet_data[ctx.filename_row_idx]
        if col_idx < len(row_data) and row_data[col_idx].strip():
            override = row_data[col_idx].strip()

    filename = files.get_source_video_filename(ctx.study_name, participant, override)
    return utils.resolve_input_path(filename)


# ---- API endpoints ----


@studio_bp.route("/api/thumbnail/<participant>/<start_seconds>")
def api_thumbnail(participant: str, start_seconds: str) -> FlaskResponse:
    if _sheet_context is None:
        return jsonify({"ok": False, "error": "No sheet loaded"}), 500

    try:
        start_sec = max(0, int(start_seconds))
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "Invalid timestamp"}), 400
    video_path = _resolve_source_video(participant)
    if video_path is None or not video_path.is_file():
        return jsonify({"ok": False, "error": "Source video not found"}), 404

    cache_key = (str(video_path), start_sec)
    cached = _thumbnail_cache.get(cache_key)
    if cached is not None:
        return Response(
            cached,
            mimetype="image/jpeg",
            headers={"Cache-Control": "public, max-age=86400"},
        )

    jpeg_bytes = video.extract_thumbnail_bytes(
        str(video_path), start_sec, width=config.STUDIO_THUMBNAIL_WIDTH
    )
    if jpeg_bytes is None:
        return jsonify({"ok": False, "error": "Thumbnail extraction failed"}), 404

    _thumbnail_cache[cache_key] = jpeg_bytes
    return Response(
        jpeg_bytes,
        mimetype="image/jpeg",
        headers={"Cache-Control": "public, max-age=86400"},
    )


@studio_bp.route("/api/sheet")
def api_sheet() -> FlaskResponse:
    if _sheet_context is None:
        return jsonify({"ok": False, "error": "No sheet loaded"}), 500

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
        for p_idx, pid in enumerate(participants):
            col_idx = ctx.id_cell.col + p_idx
            value = row_data[col_idx] if col_idx < len(row_data) else ""
            has_text = bool(value.strip())
            valid = False
            if has_text:
                cleaned, _, _ = utils.parse_cell_annotations(value)
                parsed = utils.parse_timestamps(cleaned)
                valid = bool(parsed)
            cells[pid] = {"value": value.strip(), "valid": valid, "hasText": has_text}

        rows.append(
            {
                "rowNum": row_idx + 1,
                "observation": observation,
                "category": category,
                "severity": severity,
                "cells": cells,
            }
        )

    return jsonify(
        {
            "ok": True,
            "study": ctx.study_name,
            "version": config.VERSIONNUM,
            "highlightsDuration": config.HIGHLIGHTS_REEL_DURATION_SECONDS,
            "titlecardsEnabled": config.TITLECARDS_ENABLED,
            "titlecardDuration": config.TITLECARD_DURATION_SECONDS,
            "cellExpandHover": config.STUDIO_CELL_EXPAND_HOVER,
            "defaultDuration": config.DEFAULT_DURATION_SECONDS,
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
        return jsonify({"ok": False, "error": "No sheet loaded"}), 500

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


@studio_bp.route("/api/sheet/refresh", methods=["POST"])
def api_sheet_refresh() -> FlaskResponse:
    global _sheet_context
    if _worksheet is None:
        return jsonify({"ok": False, "error": "No worksheet available"}), 500
    new_context = spreadsheet.build_sheet_context(_worksheet)
    if new_context is None:
        return jsonify({"ok": False, "error": "Failed to refresh sheet data"}), 500
    _sheet_context = new_context
    return jsonify({"ok": True})


def _save_manifest_quiet() -> None:
    """Save manifest after generate/reel; swallow errors so the caller still succeeds."""
    try:
        artifacts = _generated_artifacts
        reels = _generated_reels
        if not artifacts and not reels:
            return
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
    except Exception as e:
        utils.warning_print(f"Failed to save manifest: {e}")


def _resolve_intake_video_path(participant: str, source: str = "") -> str | None:
    """Resolve a video path for an intake participant.

    Tries the source-specific participant list first, then falls back to the
    other.  Both lists are populated from the same source videos so the
    fallback is a safety net.
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
                return p["video_path"]
    return None


def _generate_intake_clips(
    items: list[dict[str, Any]],
    output_format: str = "clip",
    study: str = "",
) -> list[dict[str, Any]]:
    """Generate clips from intake items (Screenspace or Transcript).

    Each returned dict has an extra ``"_ok"`` key (True on success, False if
    the video was missing or ffmpeg failed) plus an ``"_error"`` string so
    callers can report per-item results without duplicating the loop.
    """
    results: list[dict[str, Any]] = []

    for item in items:
        participant = item.get("participant", "")
        start = float(item.get("start", 0))
        end = float(item.get("end", 0))
        event_type = item.get("event_type", "")
        event_ids = item.get("event_ids", [])
        source = item.get("source", "screenspace")
        mark_ids = item.get("mark_ids", [])

        video_path = _resolve_intake_video_path(participant, source)

        if not video_path:
            results.append({"_ok": False, "_error": f"No video for {participant}"})
            continue

        span_hash = hashlib.md5(f"{participant}_{start}_{end}".encode()).hexdigest()[:8]
        safe_event_type = utils.sanitize_filename(event_type) if event_type else ""
        desc_part = f"{safe_event_type} " if safe_event_type else ""
        out_name = (
            f"{study} {participant} {desc_part}intake {span_hash}{config.FILEFORMAT}"
            if study
            else f"intake_{span_hash}{config.FILEFORMAT}"
        )
        out_path = files.get_unique_filename(out_name)

        start_str = utils.seconds_to_timestamp(int(round(start)))
        end_str = utils.seconds_to_timestamp(int(round(end)))

        success = video.run_ffmpeg(
            video_path, out_path, start_str, end_str, config.REENCODING
        )

        if success:
            default_desc = (
                "Transcript intake" if source == "transcript" else "Screenspace intake"
            )
            artifact: dict[str, Any] = {
                "id": f"intake_{span_hash}_s0",
                "type": output_format,
                "file": Path(out_path).name,
                "start": start,
                "end": end,
                "thumbnail": "",
                "study": study,
                "participant": participant,
                "category": "",
                "severity": "",
                "description": event_type or default_desc,
                "cellRow": None,
                "cellCol": None,
                "cellA1": "",
                "annotations": [],
                "source": source,
                "event_ids": event_ids,
                "mark_ids": mark_ids,
                "intake_label": event_type,
                "sourceVideo": Path(video_path).name,
                "_ok": True,
                "_error": "",
            }
            if source == "transcript":
                import transcripts_server

                with transcripts_server._manifest_lock:
                    src_entry = transcripts_server._manifest.get(
                        "source_transcripts", {}
                    ).get(participant, {})
                artifact["transcript_version"] = src_entry.get("transcribed_at", "")
            results.append(artifact)
        else:
            results.append({"_ok": False, "_error": "ffmpeg failed"})

    return results


def _load_stashes() -> list[dict[str, Any]]:
    data = utils.load_json_manifest(config.STASHES_MANIFEST_FILENAME, default=[])
    if not isinstance(data, list):
        return []
    return data


def _save_stashes(stashes: list[dict[str, Any]]) -> Path | None:
    return utils.save_json_manifest(config.STASHES_MANIFEST_FILENAME, stashes)


def _load_artifact_stashes() -> list[dict[str, Any]]:
    data = utils.load_json_manifest(
        config.ARTIFACT_STASHES_MANIFEST_FILENAME, default=[]
    )
    if not isinstance(data, list):
        return []
    return data


def _save_artifact_stashes(stashes: list[dict[str, Any]]) -> Path | None:
    return utils.save_json_manifest(config.ARTIFACT_STASHES_MANIFEST_FILENAME, stashes)


def _load_studio_settings() -> dict[str, Any]:
    """Load studio_settings.json and apply non-default values to config module."""
    data = utils.load_json_manifest(config.STUDIO_SETTINGS_FILENAME, default={})
    if not isinstance(data, dict):
        return {}

    applied: dict[str, Any] = {}
    for name, value in data.items():
        if name not in config.STUDIO_SETTINGS:
            continue
        default = _settings_defaults.get(name)
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
        settings_path = (
            Path(utils.get_effective_output_dir()) / config.STUDIO_SETTINGS_FILENAME
        )
        if settings_path.is_file():
            try:
                settings_path.unlink()
            except OSError:
                pass
        return None
    return utils.save_json_manifest(config.STUDIO_SETTINGS_FILENAME, to_save)


def _find_existing_artifacts(
    cell_row: int, cell_col: int, artifact_type: str
) -> list[dict[str, Any]]:
    """Return cached artifact records for a cell+type whose files still exist on disk."""
    matches = [
        a
        for a in _generated_artifacts
        if a.get("cellRow") == cell_row
        and a.get("cellCol") == cell_col
        and a.get("type") == artifact_type
    ]
    return [a for a in matches if Path(utils.resolve_output_path(a["file"])).is_file()]


@studio_bp.route("/api/generate", methods=["POST"])
def api_generate() -> FlaskResponse:
    if _worksheet is None:
        return jsonify({"ok": False, "error": "No worksheet loaded"}), 500

    data = request.get_json(silent=True) or {}
    cell_strings = data.get("cells", [])
    output_format = data.get("format", "clip")
    tc_enabled = data.get("titlecards_enabled")
    tc_duration = data.get("titlecard_duration")

    if not cell_strings:
        return jsonify({"ok": False, "error": "No cells specified"}), 400

    if output_format not in ("clip", "screen", "gif"):
        return jsonify({"ok": False, "error": f"Invalid format: {output_format}"}), 400

    try:
        cell_input = ", ".join(cell_strings)
        cell_specs = spreadsheet.parse_cell_specifications(cell_input)
        if not cell_specs:
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
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

    def stream() -> Any:
        overrides: dict[str, Any] = {}
        if tc_enabled is not None:
            overrides["TITLECARDS_ENABLED"] = bool(tc_enabled)
        if tc_duration is not None:
            try:
                val = int(tc_duration)
                if val > 0:
                    overrides["TITLECARD_DURATION_SECONDS"] = val
            except (ValueError, TypeError):
                pass
        with _override_config(**overrides):
            clip_cells: set[str] = set()

            # Pass 1: yield already-existing artifacts, collect clips that need generation
            to_generate: list[tuple[Any, str]] = []
            for clip in clips:
                cell_str = clip["participant"] + "." + str(clip["cell"].row)
                clip_cells.add(cell_str)

                existing = _find_existing_artifacts(
                    clip["cell"].row, clip["cell"].col, output_format
                )
                if existing:
                    yield (
                        json.dumps(
                            {
                                "cell": cell_str,
                                "ok": True,
                                "generated": len(existing),
                                "artifacts": existing,
                                "skipped": True,
                            }
                        )
                        + "\n"
                    )
                else:
                    to_generate.append((clip, cell_str))

            # Pass 2: generate in parallel and yield as each completes
            if to_generate:
                workers = min(4, os.cpu_count() or 1)
                if workers >= 2 and len(to_generate) >= 2:
                    with concurrent.futures.ThreadPoolExecutor(
                        max_workers=workers
                    ) as pool:
                        future_to_cell: dict[
                            concurrent.futures.Future, tuple[Any, str]
                        ] = {
                            pool.submit(
                                pipeline.process_clips,
                                [clip],
                                output_format=output_format,
                            ): (clip, cell_str)
                            for clip, cell_str in to_generate
                        }
                        for future in concurrent.futures.as_completed(future_to_cell):
                            clip, cell_str = future_to_cell[future]
                            try:
                                generated, artifacts = future.result()
                                _generated_artifacts.extend(artifacts)
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
                        try:
                            generated, artifacts = pipeline.process_clips(
                                [clip], output_format=output_format
                            )
                            _generated_artifacts.extend(artifacts)
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

            for cs in cell_strings:
                if cs not in clip_cells:
                    yield (
                        json.dumps({"cell": cs, "ok": False, "error": "No clip found"})
                        + "\n"
                    )
            _save_manifest_quiet()

    return Response(
        stream(), mimetype="application/x-ndjson", headers={"X-Accel-Buffering": "no"}
    )


@studio_bp.route("/api/highlights-preview", methods=["POST"])
def api_highlights_preview() -> FlaskResponse:
    if _worksheet is None:
        return jsonify({"ok": False, "error": "No worksheet loaded"}), 500

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
        return jsonify({"ok": False, "error": "No worksheet loaded"}), 500

    data = request.get_json(silent=True) or {}
    cell_strings = data.get("cells", [])
    highlights_duration = data.get("highlights_duration")
    tc_enabled = data.get("titlecards_enabled")
    tc_duration = data.get("titlecard_duration")

    if not cell_strings:
        return jsonify({"ok": False, "error": "No cells specified"}), 400

    overrides: dict[str, Any] = {}
    if highlights_duration is not None:
        try:
            val = int(highlights_duration)
            if val > 0:
                overrides["HIGHLIGHTS_REEL_DURATION_SECONDS"] = val
        except (ValueError, TypeError):
            pass
    if tc_enabled is not None:
        overrides["TITLECARDS_ENABLED"] = bool(tc_enabled)
    if tc_duration is not None:
        try:
            val = int(tc_duration)
            if val > 0:
                overrides["TITLECARD_DURATION_SECONDS"] = val
        except (ValueError, TypeError):
            pass

    with _override_config(**overrides):
        try:
            reel_input = ", ".join(cell_strings)

            # generate_list needs HIGHLIGHTS override only during this call
            clips = spreadsheet.generate_list(
                _worksheet,
                "reel",
                ctx=_sheet_context,
                reel_input=reel_input,
                skip_prompts=True,
            )

            if not clips:
                return jsonify(
                    {"ok": False, "error": "No clips found for the specified cells"}
                ), 400

            # Check if an identical reel already exists
            components: list[dict[str, Any]] = []
            for clip in clips:
                files.prepare_clip(clip)
                cell = clip.get("cell")
                for start_str, end_str in clip.get("times", []):
                    components.append(
                        {
                            "cellRow": getattr(cell, "row", None),
                            "cellCol": getattr(cell, "col", None),
                            "start": utils.timestamp_to_seconds(start_str),
                            "end": utils.timestamp_to_seconds(end_str),
                        }
                    )
            if components:
                expected_id = pipeline.compute_reel_id(components)
                for reel in _generated_reels:
                    if (
                        reel.get("id") == expected_id
                        and Path(utils.resolve_output_path(reel["file"])).is_file()
                    ):
                        return jsonify(
                            {
                                "ok": True,
                                "generated": 1,
                                "reels": [reel],
                                "skipped": True,
                            }
                        )

            _reel_cancel_event.clear()
            cancel_flag = _reel_cancel_event.is_set
            generated, reel_records = pipeline.process_reel(
                clips, cancel_flag=cancel_flag
            )
            if _reel_cancel_event.is_set():
                return jsonify(
                    {
                        "ok": False,
                        "error": "Reel generation cancelled",
                        "cancelled": True,
                    }
                )
            _generated_reels.extend(reel_records)
            _save_manifest_quiet()
            return jsonify({"ok": True, "generated": generated, "reels": reel_records})

        except Exception as e:
            return jsonify({"ok": False, "error": str(e)}), 500


@studio_bp.route("/api/viewer", methods=["POST"])
def api_viewer() -> FlaskResponse:
    artifacts = _generated_artifacts
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


@studio_bp.route("/api/timeline-viewer", methods=["POST"])
def api_timeline_viewer() -> FlaskResponse:
    if _worksheet is None:
        return jsonify({"ok": False, "error": "No worksheet loaded"}), 500

    try:
        req = request.get_json(silent=True) or {}
        include_intake = req.get("include_intake", False)
        intake_items = req.get("intake_items", [])

        clips_list = spreadsheet.generate_list(
            _worksheet, "batch", ctx=_sheet_context, skip_prompts=True
        )
        if not clips_list:
            return jsonify({"ok": False, "error": "No clips found in sheet"}), 400

        generated, artifacts = pipeline.process_clips(clips_list, output_format="clip")
        if not artifacts:
            return jsonify({"ok": False, "error": "No artifacts were generated"}), 400

        _generated_artifacts.extend(artifacts)

        # Generate intake clips if requested
        intake_artifacts: list[dict[str, Any]] = []
        if include_intake and intake_items:
            raw = _generate_intake_clips(intake_items)
            for r in raw:
                if r.pop("_ok", False):
                    r.pop("_error", None)
                    intake_artifacts.append(r)
            artifacts = artifacts + intake_artifacts
            _generated_artifacts.extend(intake_artifacts)

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


@studio_bp.route("/api/gallery", methods=["POST"])
def api_gallery() -> FlaskResponse:
    if _sheet_context is None:
        return jsonify({"ok": False, "error": "No sheet loaded"}), 500

    data = request.get_json(silent=True) or {}
    participant = data.get("participant", "")
    output_format = data.get("format", "screen")
    interval = data.get("interval", config.GALLERY_INTERVAL_SECONDS)
    bundle = bool(data.get("bundle", config.GALLERY_BUNDLE_ENABLED))

    if not participant:
        return jsonify({"ok": False, "error": "No participant specified"}), 400

    if output_format not in ("screen", "gif"):
        return jsonify({"ok": False, "error": f"Invalid format: {output_format}"}), 400

    try:
        interval = int(interval)
        if interval < 1:
            interval = config.GALLERY_INTERVAL_SECONDS
    except (ValueError, TypeError):
        interval = config.GALLERY_INTERVAL_SECONDS

    video_path = _resolve_source_video(participant)
    if video_path is None or not video_path.is_file():
        return jsonify(
            {"ok": False, "error": f"Source video not found for {participant}"}
        ), 404

    try:
        artifacts = video.generate_interval_captures(
            str(video_path),
            interval_seconds=interval,
            output_format=output_format,
            gif_duration_seconds=config.GALLERY_GIF_DURATION_SECONDS,
        )
        if not artifacts:
            return jsonify({"ok": False, "error": "No captures generated"}), 500

        duration = video.get_file_duration(str(video_path)) or 0
        gallery_data = viewer.finalize_gallery_data(
            artifacts,
            source_video=video_path.name,
            video_duration=duration,
            output_format=output_format,
            interval=interval,
            bundle=bundle,
        )
        gallery_path = viewer.generate_gallery_viewer(gallery_data)
        if gallery_path:
            return jsonify({"ok": True, "file": str(gallery_path)})
        return jsonify({"ok": False, "error": "Failed to generate gallery viewer"}), 500

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


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
        artifacts = viewer.load_manifest_artifacts()
        reels = viewer.load_manifest_reels()
        return jsonify({"ok": True, "artifacts": artifacts, "reels": reels})

    artifacts = _generated_artifacts
    reels = _generated_reels
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
                "step": meta.get("step"),
                "provider": meta.get("provider"),
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
            default = _settings_defaults.get(name)
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

    applied = {}
    for name, value in settings_data.items():
        if name not in config.STUDIO_SETTINGS:
            continue
        default = _settings_defaults.get(name)
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

    _save_studio_settings(applied)
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


# ---- Screenspace Intake ----


@studio_bp.route("/api/generate-intake", methods=["POST"])
def api_generate_intake() -> FlaskResponse:
    """Generate clips from intake spans (Screenspace or Transcript)."""
    data = request.get_json(silent=True) or {}
    items = data.get("items", [])
    if not items:
        return jsonify({"ok": False, "error": "No intake items specified"}), 400

    output_format = data.get("format", "clip")
    study = _sheet_context.study_name if _sheet_context else ""

    raw = _generate_intake_clips(items, output_format=output_format, study=study)
    results: list[dict[str, Any]] = []
    for r in raw:
        ok = r.pop("_ok", False)
        error = r.pop("_error", "")
        if ok:
            _generated_artifacts.append(r)
            results.append({"ok": True, "artifact": r})
        else:
            results.append({"ok": False, "error": error})

    _save_manifest_quiet()
    return jsonify({"ok": True, "results": results})


@studio_bp.route("/api/reel-direct", methods=["POST"])
def api_reel_direct() -> FlaskResponse:
    """Build a reel from direct timestamp segments (for intake / mixed queues)."""
    import tempfile

    data = request.get_json(silent=True) or {}
    segments = data.get("segments", [])
    if not segments:
        return jsonify({"ok": False, "error": "No segments specified"}), 400

    tc_enabled = data.get("titlecards_enabled")
    tc_duration = data.get("titlecard_duration")
    overrides: dict[str, Any] = {}
    if tc_enabled is not None:
        overrides["TITLECARDS_ENABLED"] = bool(tc_enabled)
    if tc_duration is not None:
        try:
            val = int(tc_duration)
            if val > 0:
                overrides["TITLECARD_DURATION_SECONDS"] = val
        except (ValueError, TypeError):
            pass

    output_dir = Path(utils.get_effective_output_dir())
    clip_paths: list[str] = []
    temp_clips: list[str] = []

    _reel_cancel_event.clear()

    with _override_config(**overrides):
        try:
            for seg in segments:
                if _reel_cancel_event.is_set():
                    break

                participant = seg.get("participant", "")
                start = float(seg.get("start", 0))
                end = float(seg.get("end", 0))
                if end <= start:
                    continue

                source = seg.get("source", "screenspace")
                video_path = _resolve_intake_video_path(participant, source)

                if not video_path:
                    continue

                start_str = utils.seconds_to_timestamp(int(round(start)))
                end_str = utils.seconds_to_timestamp(int(round(end)))

                fd, tmp_path = tempfile.mkstemp(
                    suffix=config.FILEFORMAT, dir=str(output_dir)
                )
                os.close(fd)
                temp_clips.append(tmp_path)

                ok = video.run_ffmpeg(
                    video_path, tmp_path, start_str, end_str, config.REENCODING
                )
                if ok:
                    clip_paths.append(tmp_path)

            if _reel_cancel_event.is_set():
                return jsonify(
                    {
                        "ok": False,
                        "error": "Reel generation cancelled",
                        "cancelled": True,
                    }
                )

            if not clip_paths:
                return (
                    jsonify({"ok": False, "error": "No clips could be generated"}),
                    400,
                )

            reel_study = _sheet_context.study_name if _sheet_context else ""
            reel_base = f"{reel_study} intake reel" if reel_study else "intake_reel"
            reel_name = files.get_unique_filename(f"{reel_base}{config.FILEFORMAT}")
            ok = video.concatenate_clips(clip_paths, reel_name, reencode_on_fail=True)

            if ok:
                reel_record: dict[str, Any] = {
                    "id": f"reel_intake_{hashlib.md5(reel_name.encode()).hexdigest()[:8]}",
                    "file": Path(reel_name).name,
                    "source": "intake",
                    "description": f"Intake reel ({len(clip_paths)} segments)",
                }
                _generated_reels.append(reel_record)
                _save_manifest_quiet()
                return jsonify({"ok": True, "generated": 1, "reels": [reel_record]})
            else:
                return (
                    jsonify({"ok": False, "error": "Reel concatenation failed"}),
                    500,
                )
        finally:
            for tmp in temp_clips:
                try:
                    Path(tmp).unlink(missing_ok=True)
                except OSError:
                    pass


@studio_bp.route("/api/reel/cancel", methods=["POST"])
def api_reel_cancel() -> FlaskResponse:
    """Signal cancellation for the in-progress reel build."""
    _reel_cancel_event.set()
    return jsonify({"ok": True})


# ---- State initialization ----


def _init_studio_state(worksheet: Any) -> None:
    """Initialize module-level state for Studio routes."""
    global \
        _worksheet, \
        _sheet_context, \
        _generated_artifacts, \
        _generated_reels, \
        _thumbnail_cache

    _load_studio_settings()
    _worksheet = worksheet
    _sheet_context = spreadsheet.build_sheet_context(worksheet)
    _generated_artifacts, _generated_reels = viewer._load_manifest_both()
    _thumbnail_cache = {}

    if _sheet_context is None:
        utils.error_print("Could not load spreadsheet data for Studio.")
        sys.exit(1)


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


def start_combined_server(
    worksheet: Any = None,
    port: int | None = None,
    default_page: str = "studio",
) -> None:
    """Start a combined Studio + Insights + Screenspace server on one port.

    When worksheet is provided, both Studio and Insights are available.
    When worksheet is None, only Insights is registered.
    Screenspace is always registered (auto-discovers videos when no
    spreadsheet is provided).
    """
    import insights_server
    import screenspace_server
    import transcripts_server

    combined = Flask(__name__, static_folder=None)

    # Always register Insights (only needs manifest files on disk)
    insights_server._init_insights_state()
    combined.register_blueprint(insights_server.insights_bp, url_prefix="/insights")

    # Register Studio only if a worksheet is available
    has_studio = worksheet is not None
    if has_studio:
        _init_studio_state(worksheet)
        combined.register_blueprint(studio_bp, url_prefix="/studio")

    # Always register Screenspace (auto-discovers videos from input dir)
    screenspace_server._init_screenspace_state(
        sheet_context=_sheet_context if has_studio else None,
        participant_list=_resolve_participants() if has_studio else None,
    )
    combined.register_blueprint(
        screenspace_server.screenspace_bp, url_prefix="/screenspace"
    )

    # Always register Transcripts (auto-discovers videos from input dir)
    transcripts_server._init_transcripts_state(
        sheet_context=_sheet_context if has_studio else None,
        participant_list=_resolve_participants() if has_studio else None,
    )
    combined.register_blueprint(
        transcripts_server.transcripts_bp, url_prefix="/transcripts"
    )

    @combined.after_request
    def _set_cache_headers(response):
        # Skip if a route already set Cache-Control (e.g. thumbnails, sprites)
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
        return jsonify(
            {
                "studio": has_studio,
                "insights": True,
                "screenspace": True,
                "transcripts": True,
            }
        )

    # ---- Shared settings (available from any page) ----

    # Load persisted settings unconditionally so Transcripts/Screenspace can
    # read and write model preferences even when Studio is not active.
    if not has_studio:
        _load_studio_settings()

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

    # ---- Model discovery ----

    @combined.route("/api/models")
    def api_models() -> Response:
        import ollama_client
        import transcripts

        whisper_models = [
            {
                "name": m["name"],
                "size_mb": m["size_mb"],
                "description": m["description"],
                "selected": m["name"] == config.TRANSCRIBE_MODEL,
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

        return jsonify(
            {
                "ok": True,
                "whisper": {"models": whisper_models},
                "ollama": {
                    "available": ollama_available,
                    "models": ollama_models,
                    "base_url": config.OLLAMA_BASE_URL,
                },
            }
        )

    port = port or config.SERVER_PORT
    url = f"http://127.0.0.1:{port}/{default_page}/"

    utils.info_print(f"clipgen server running at http://127.0.0.1:{port}")
    webbrowser.open(url)

    combined.run(host="127.0.0.1", port=port, debug=False, threaded=True)
