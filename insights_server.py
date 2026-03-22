# -*- coding: utf-8 -*-
"""Insights Builder web server for clipgen.

Serves the Insights Builder front-end and exposes REST endpoints for
insight CRUD, artifact browsing, sprite sheet generation, and viewer export.
"""

from __future__ import annotations

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
_output_dir: str = ""

_assets_dir = Path(__file__).resolve().parent / "assets" / "web"

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
    # Enrich with sprite data from the cached list
    sprite_map = {a.get("file"): a for a in _artifacts if a.get("thumbnail")}
    missing_sprites = []
    for art in fresh_artifacts:
        cached = sprite_map.get(art.get("file"))
        if cached:
            art["thumbnail"] = cached["thumbnail"]
            art["spriteData"] = cached.get("spriteData", {})
        elif art.get("type") == "clip":
            missing_sprites.append(art)

    # Generate sprites for newly discovered clip artifacts (e.g. from Studio)
    if missing_sprites:
        _artifacts.extend(missing_sprites)
        generated = _generate_missing_sprites()
        if generated:
            sprite_map = {a.get("file"): a for a in _artifacts if a.get("thumbnail")}
            for art in fresh_artifacts:
                if not art.get("spriteData"):
                    cached = sprite_map.get(art.get("file"))
                    if cached:
                        art["thumbnail"] = cached["thumbnail"]
                        art["spriteData"] = cached.get("spriteData", {})

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


@insights_bp.route("/api/sprites/generate", methods=["POST"])
def api_sprites_generate() -> FlaskResponse:
    generated = _generate_missing_sprites()
    return jsonify({"ok": True, "generated": generated})


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

        data = viewer.finalize_insights_viewer_data(
            ins_list,
            _artifacts,
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


def _generate_missing_sprites() -> int:
    """Generate sprite sheets for clip artifacts that lack them. Returns count generated."""
    generated = 0
    output_path = Path(_output_dir)
    for artifact in _artifacts:
        if artifact.get("type") != "clip":
            continue
        filename = artifact.get("file", "")
        if not filename:
            continue
        stem = Path(filename).stem
        sprite_name = f"{stem}_sprite.png"
        sprite_path = output_path / sprite_name

        if sprite_path.is_file():
            artifact["thumbnail"] = sprite_name
            if "spriteData" not in artifact or not artifact["spriteData"]:
                # Sprite exists but no metadata — fill in defaults
                artifact["spriteData"] = {
                    "cols": 5,
                    "rows": 4,
                    "frameCount": config.SPRITE_SHEET_FRAME_COUNT,
                    "frameWidth": config.SPRITE_SHEET_THUMB_WIDTH,
                    "frameHeight": round(config.SPRITE_SHEET_THUMB_WIDTH * 9 / 16),
                    "interval": 1,
                }
            continue

        source_video = artifact.get("sourceVideo", "")
        if not source_video or not Path(source_video).is_file():
            # Try finding clip file in output dir
            clip_path = output_path / filename
            if clip_path.is_file():
                source_video = str(clip_path)
            else:
                continue

        sprite_data = video.generate_sprite_sheet(source_video, str(sprite_path))
        if sprite_data:
            artifact["thumbnail"] = sprite_name
            artifact["spriteData"] = sprite_data
            generated += 1

    return generated


# ---- State initialization ----


def _init_insights_state() -> None:
    """Initialize module-level state for Insights routes.

    Loads artifacts from clipgen_manifest.json and insights from
    insights_manifest.json. Generates missing sprite sheets.
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

    # Generate missing sprite sheets (blocking, before browser opens)
    sprite_count = _generate_missing_sprites()
    if sprite_count:
        utils.info_print(
            f"Generated {sprite_count} sprite sheet{'s' if sprite_count != 1 else ''}."
        )


