# -*- coding: utf-8 -*-
"""Insights Builder Flask blueprint for clipgen.

Registered at /insights/ by start_combined_server(). When --insights is used without
a spreadsheet, only this blueprint is active.

API endpoints (all under /insights/):
  GET  /media/<filename>             – serve artifact media files from output directory
  GET  /api/artifacts                – artifacts from manifest enriched with sprite data
  GET  /api/insights                 – list all insights
  POST /api/insights                 – create new insight
  GET  /api/insights/<id>            – read single insight
  PUT  /api/insights/<id>            – update insight fields
  DELETE /api/insights/<id>          – delete insight
  GET  /api/sprites/<filename>       – serve or generate a sprite sheet for a clip
  POST /api/generate-viewer          – export standalone insights_viewer.html
"""

from __future__ import annotations

import tempfile
from math import ceil, sqrt
from pathlib import Path
from typing import Any, Dict, List, Tuple, Union

from flask import Blueprint, Response, jsonify, request, send_from_directory

import config
import insights
import utils
import video
import viewer

FlaskResponse = Union[Response, Tuple[Response, int]]

# ---- Module-level state (set once by _init_insights_state) ----

_artifacts: List[Dict[str, Any]] = []
_insights_data: Dict[str, Any] = {}
_output_dir: Union[str, Path] = ""
_sprite_cache: Dict[str, bytes] = {}

_assets_dir = utils.get_bundled_assets_root() / "assets" / "web"

# ---- Blueprint ----

insights_bp = Blueprint("insights", __name__)


# ---- Static file serving ----


@insights_bp.route("/")
def serve_index() -> FlaskResponse:
    return send_from_directory(_assets_dir, "insights-builder.html")


@insights_bp.route("/<path:filename>")
def serve_static(filename: str) -> FlaskResponse:
    return send_from_directory(_assets_dir, filename)


@insights_bp.route("/media/<path:filename>")
def serve_media(filename: str) -> FlaskResponse:
    return send_from_directory(_output_dir, filename)


# ---- API endpoints ----


@insights_bp.route("/api/artifacts")
def api_artifacts() -> FlaskResponse:
    # Re-read manifest on every request so artifacts generated in Studio
    # appear immediately when the user navigates to Insights.
    fresh_artifacts = viewer.load_manifest_artifacts()
    for art in fresh_artifacts:
        if art.get("type") != "clip":
            continue
        start = art.get("start")
        end = art.get("end")
        if start is None or end is None:
            continue
        duration = (end or 0) - (start or 0)
        if duration > 0:
            art["spriteData"] = _compute_sprite_metadata(duration)
    return jsonify({"ok": True, "artifacts": fresh_artifacts})


@insights_bp.route("/api/insights")
def api_insights_list() -> FlaskResponse:
    return jsonify({"ok": True, "insights": _insights_data.get("insights", [])})


@insights_bp.route("/api/insights/<insight_id>")
def api_insights_get(insight_id: str) -> FlaskResponse:
    for ins in _insights_data.get("insights", []):
        if ins["id"] == insight_id:
            return jsonify({"ok": True, "insight": ins})
    return jsonify({"ok": False, "error": "Insight not found"}), 404


@insights_bp.route("/api/insights", methods=["POST"])
def api_insights_create() -> FlaskResponse:
    data = request.get_json(silent=True) or {}
    new_insight = insights.create_insight(
        title=data.get("title", ""),
        severity=data.get("severity", ""),
        status=data.get("status", "draft"),
    )
    _insights_data.setdefault("insights", []).append(new_insight)
    _save_insights()
    return jsonify({"ok": True, "insight": new_insight})


@insights_bp.route("/api/insights/<insight_id>", methods=["PUT"])
def api_insights_update(insight_id: str) -> FlaskResponse:
    data = request.get_json(silent=True) or {}
    updated = insights.update_insight(
        _insights_data.get("insights", []), insight_id, data
    )
    if updated is None:
        return jsonify({"ok": False, "error": "Insight not found"}), 404
    _save_insights()
    return jsonify({"ok": True, "insight": updated})


@insights_bp.route("/api/insights/<insight_id>", methods=["DELETE"])
def api_insights_delete(insight_id: str) -> FlaskResponse:
    removed = insights.delete_insight(_insights_data.get("insights", []), insight_id)
    if not removed:
        return jsonify({"ok": False, "error": "Insight not found"}), 404
    _save_insights()
    return jsonify({"ok": True})


@insights_bp.route("/api/sprites/<path:filename>")
def api_sprite(filename: str) -> FlaskResponse:
    cached = _sprite_cache.get(filename)
    if cached:
        return Response(
            cached,
            mimetype="image/png",
            headers={"Cache-Control": "public, max-age=86400"},
        )

    clip_path = Path(_output_dir) / filename
    if not clip_path.is_file():
        return jsonify({"ok": False, "error": "Clip not found"}), 404

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmp_path = tmp.name
    tmp.close()
    try:
        sprite_data = video.generate_sprite_sheet(str(clip_path), tmp_path)
        if not sprite_data or not Path(tmp_path).is_file():
            return jsonify({"ok": False, "error": "Sprite generation failed"}), 500
        png_bytes = Path(tmp_path).read_bytes()
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    _sprite_cache[filename] = png_bytes
    return Response(
        png_bytes,
        mimetype="image/png",
        headers={"Cache-Control": "public, max-age=86400"},
    )


@insights_bp.route("/api/generate-viewer", methods=["POST"])
def api_generate_viewer() -> FlaskResponse:
    ins_list = _insights_data.get("insights", [])
    if not ins_list:
        return jsonify({"ok": False, "error": "No insights to export"}), 400

    try:
        # Check for existing timeline viewer
        timeline_viewer_file = ""
        output_path = Path(_output_dir)
        for html_file in sorted(output_path.glob("clips_viewer*.html")):
            timeline_viewer_file = html_file.name
            break

        study = _insights_data.get("meta", {}).get("study", "")
        if not study and _artifacts:
            study = _artifacts[0].get("study", "")

        fresh_artifacts = viewer.load_manifest_artifacts()
        data = viewer.finalize_insights_viewer_data(
            ins_list,
            fresh_artifacts,
            study=study,
            timeline_viewer_file=timeline_viewer_file,
        )
        viewer_path = viewer.generate_insights_viewer(data)
        if viewer_path:
            return jsonify({"ok": True, "file": str(viewer_path)})
        return jsonify({"ok": False, "error": "Failed to generate viewer"}), 500
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ---- Internal helpers ----


def _save_insights() -> None:
    meta = _insights_data.get("meta", {})
    insights.save_insights_manifest(meta, _insights_data.get("insights", []))


def _compute_sprite_metadata(duration_seconds: float) -> Dict[str, Any]:
    """Compute sprite sheet layout metadata from a clip duration (no I/O)."""
    frame_count = config.SPRITE_SHEET_FRAME_COUNT
    thumb_width = config.SPRITE_SHEET_THUMB_WIDTH
    min_interval = config.SPRITE_SHEET_MIN_INTERVAL
    interval = max(min_interval, int(duration_seconds) // frame_count)
    actual_frames = min(frame_count, max(1, int(duration_seconds) // interval))
    cols = ceil(sqrt(actual_frames))
    rows = ceil(actual_frames / cols)
    return {
        "cols": cols,
        "rows": rows,
        "frameCount": actual_frames,
        "frameWidth": thumb_width,
        "frameHeight": round(thumb_width * 9 / 16),
        "interval": interval,
    }


# ---- State initialization ----


def _init_insights_state() -> None:
    """Initialize module-level state for Insights routes.

    Loads artifacts from clipgen_manifest.json and insights from
    insights_manifest.json. Sprite sheets are generated lazily on request.
    """
    global _artifacts, _insights_data, _output_dir

    _output_dir = utils.get_effective_output_dir()
    _artifacts = viewer.load_manifest_artifacts()
    _insights_data = insights.load_insights_manifest()

    if not _artifacts:
        utils.warning_print(
            "No artifacts found in manifest.",
            [
                "Generate artifacts first, then run the insights builder.",
                f"Expected manifest: {Path(_output_dir) / config.MANIFEST_FILENAME}",
            ],
        )

    # Set study in insights meta from artifacts if not already set
    if not _insights_data.get("meta", {}).get("study") and _artifacts:
        _insights_data.setdefault("meta", {})["study"] = _artifacts[0].get("study", "")
