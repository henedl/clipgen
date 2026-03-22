# -*- coding: utf-8 -*-
"""Web server for clipgen Studio and Insights Builder.

Serves the Studio and Insights Builder front-ends on a single port via
start_combined_server(), and exposes REST endpoints for sheet data
access, artifact generation, reel building, and viewer creation.
"""

import sys
import webbrowser
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from flask import (
    Blueprint,
    Flask,
    Response,
    jsonify,
    redirect,
    request,
    send_from_directory,
)

import clipgen
import config
import spreadsheet
import utils
import viewer

FlaskResponse = Union[Response, Tuple[Response, int]]

# ---- Module-level state (set once by _init_studio_state) ----

_worksheet: Any = None
_sheet_context: Optional[spreadsheet.SheetContext] = None
_generated_artifacts: List[Dict[str, Any]] = []
_generated_reels: List[Dict[str, Any]] = []

_assets_dir = Path(__file__).resolve().parent / "assets" / "web"

# ---- Blueprint ----

studio_bp = Blueprint("studio", __name__)


# ---- Static file serving ----


@studio_bp.route("/")
def serve_index() -> FlaskResponse:
    return send_from_directory(_assets_dir, "studio.html")


@studio_bp.route("/<path:filename>")
def serve_static(filename: str) -> FlaskResponse:
    return send_from_directory(_assets_dir, filename)


# ---- API endpoints ----


@studio_bp.route("/api/sheet")
def api_sheet() -> FlaskResponse:
    if _sheet_context is None:
        return jsonify({"ok": False, "error": "No sheet loaded"}), 500

    ctx = _sheet_context
    participants = spreadsheet.get_participant_list(
        ctx.header_row, ctx.id_cell, ctx.num_participants
    )

    rows: List[Dict[str, Any]] = []
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

        cells: Dict[str, Dict[str, Any]] = {}
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
            "participants": participants,
            "rows": rows,
        }
    )


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
            is_excel=clipgen._is_excel_worksheet(_worksheet) if _worksheet else False,
            mode="studio",
        )
    except Exception:
        pass


@studio_bp.route("/api/generate", methods=["POST"])
def api_generate() -> FlaskResponse:
    if _worksheet is None:
        return jsonify({"ok": False, "error": "No worksheet loaded"}), 500

    data = request.get_json(silent=True) or {}
    cell_strings = data.get("cells", [])
    output_format = data.get("format", "clip")

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
            _worksheet, "cell", cell_specs=cell_specs, skip_prompts=True
        )
        if not clips:
            return jsonify(
                {"ok": False, "error": "No clips found for the specified cells"}
            ), 400

        generated, artifacts = clipgen.process_clips(clips, output_format=output_format)
        _generated_artifacts.extend(artifacts)
        _save_manifest_quiet()

        return jsonify({"ok": True, "generated": generated, "artifacts": artifacts})

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@studio_bp.route("/api/reel", methods=["POST"])
def api_reel() -> FlaskResponse:
    if _worksheet is None:
        return jsonify({"ok": False, "error": "No worksheet loaded"}), 500

    data = request.get_json(silent=True) or {}
    cell_strings = data.get("cells", [])
    highlights_duration = data.get("highlights_duration")

    if not cell_strings:
        return jsonify({"ok": False, "error": "No cells specified"}), 400

    try:
        reel_input = ", ".join(cell_strings)

        original_duration = config.HIGHLIGHTS_REEL_DURATION_SECONDS
        if highlights_duration is not None:
            try:
                val = int(highlights_duration)
                if val > 0:
                    config.HIGHLIGHTS_REEL_DURATION_SECONDS = val
            except (ValueError, TypeError):
                pass

        try:
            clips = spreadsheet.generate_list(
                _worksheet, "reel", reel_input=reel_input, skip_prompts=True
            )
        finally:
            config.HIGHLIGHTS_REEL_DURATION_SECONDS = original_duration

        if not clips:
            return jsonify(
                {"ok": False, "error": "No clips found for the specified cells"}
            ), 400

        generated, reel_records = clipgen.process_reel(clips)
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
        data = viewer.finalize_timeline_data(
            artifacts, study=study, participant=participant, mode="studio"
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
        clips_list = spreadsheet.generate_list(_worksheet, "batch", skip_prompts=True)
        if not clips_list:
            return jsonify({"ok": False, "error": "No clips found in sheet"}), 400

        generated, artifacts = clipgen.process_clips(clips_list, output_format="clip")
        if not artifacts:
            return jsonify({"ok": False, "error": "No artifacts were generated"}), 400

        _generated_artifacts.extend(artifacts)

        study = artifacts[0].get("study", "")
        data = viewer.finalize_timeline_data(
            artifacts,
            study=study,
            worksheet_title=getattr(_worksheet, "title", ""),
            is_excel=clipgen._is_excel_worksheet(_worksheet),
            mode="timeline-viewer",
            output_format="clip",
        )
        viewer_path = viewer.generate_timeline_viewer(
            data,
            template_name="timeline-viewer.html",
            output_basename="timeline_viewer.html",
        )
        if viewer_path:
            return jsonify(
                {
                    "ok": True,
                    "file": str(viewer_path),
                    "generated": generated,
                }
            )
        return jsonify(
            {"ok": False, "error": "Failed to generate timeline viewer"}
        ), 500

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@studio_bp.route("/api/manifest", methods=["POST"])
def api_manifest() -> FlaskResponse:
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
            is_excel=clipgen._is_excel_worksheet(_worksheet) if _worksheet else False,
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
        artifacts = viewer.load_manifest_artifacts()
        reels = viewer.load_manifest_reels()
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

        regenerated = clipgen.regenerate_from_manifest(artifacts, reels=reels)
        return jsonify({"ok": True, "regenerated": regenerated, "total": total})

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ---- State initialization ----


def _init_studio_state(worksheet: Any) -> None:
    """Initialize module-level state for Studio routes."""
    global _worksheet, _sheet_context, _generated_artifacts, _generated_reels

    _worksheet = worksheet
    _sheet_context = spreadsheet.build_sheet_context(worksheet)
    _generated_artifacts = []
    _generated_reels = []

    if _sheet_context is None:
        utils.error_print("Could not load spreadsheet data for Studio.")
        sys.exit(1)


# ---- Entry point ----


def start_combined_server(
    worksheet: Any = None,
    port: Optional[int] = None,
    default_page: str = "studio",
) -> None:
    """Start a combined Studio + Insights Builder server on one port.

    When worksheet is provided, both Studio and Insights are available.
    When worksheet is None, only Insights is registered.
    """
    import insights_server

    combined = Flask(__name__, static_folder=None)

    # Always register Insights (only needs manifest files on disk)
    insights_server._init_insights_state()
    combined.register_blueprint(insights_server.insights_bp, url_prefix="/insights")

    # Register Studio only if a worksheet is available
    has_studio = worksheet is not None
    if has_studio:
        _init_studio_state(worksheet)
        combined.register_blueprint(studio_bp, url_prefix="/studio")

    @combined.route("/")
    def root() -> Response:
        return redirect(f"/{default_page}/")

    @combined.route("/api/status")
    def status() -> Response:
        return jsonify({"studio": has_studio, "insights": True})

    port = port or config.SERVER_PORT
    url = f"http://127.0.0.1:{port}/{default_page}/"

    utils.info_print(f"clipgen server running at http://127.0.0.1:{port}")
    webbrowser.open(url)

    combined.run(host="127.0.0.1", port=port, debug=False)
