# -*- coding: utf-8 -*-
"""Timeline and gallery viewer generation, manifest persistence, and insights viewer export.

Timeline viewer (--viewer / interactive 'viewer'):
  Injects window.CLIPGEN_DATA into viewer.html, replacing <!-- CLIPGEN_DATA_HERE -->.
  Data shape: { meta: {study, participant, generatedAt, mode, sourceSpreadsheet,
    sourceFileType, filmstripEnabled}, artifacts: [{id, type, file, start, end,
    study, participant, category, description, cellRow, cellCol, cellA1, annotations,
    sourceVideo}], timeline: {duration, startOffset} }
  Key functions: build_artifact_records_for_clip(), finalize_timeline_data(),
    generate_timeline_viewer().

Gallery viewer (--gallery):
  Same inlining pattern using gallery.html.
  Data shape: { meta: {sourceVideo, generatedAt, mode, format, interval, videoDuration},
    artifacts: [{file, timestamp, timestamp_formatted, type, duration}] }
  Key functions: finalize_gallery_data(), generate_gallery_viewer().
  Gallery artifacts are NOT written to the manifest by default.

Insights viewer (generate_insights_viewer()):
  Produces a standalone insights_viewer.html. Shows only 'final' insights when any exist;
  falls back to all insights. Only artifacts referenced by visible insights are included.
  Key functions: finalize_insights_viewer_data(), generate_insights_viewer().

Artifact manifest (save_manifest / load_manifest_*):
  Merges new artifacts/reels into clipgen_manifest.json, deduplicating by id (newer wins).
  Consumed by Insights Builder, --regenerate, and standalone --viewer.
"""

import json
import math
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import gspread

import config
import files
import utils

# Mutable lists of records collected during an interactive session.
INTERACTIVE_ARTIFACTS: List[Dict[str, Any]] = []
INTERACTIVE_REELS: List[Dict[str, Any]] = []


def _is_valid_artifact(a: Dict[str, Any]) -> bool:
    """Return True if artifact has minimum required fields for viewer rendering."""
    if not a.get("id"):
        return False
    if not a.get("file"):
        return False
    if a.get("start") is None and a.get("end") is None:
        return False
    return True


def build_artifact_records_for_clip(
    clip: utils.ClipRecord,
    base_video: str,
    segment_details: list,
    output_format: str,
) -> List[Dict[str, Any]]:
    """Build artifact metadata records from a processed clip's successful outputs.

    Args:
        clip: Prepared clip dict with 'times', 'category', 'study', etc.
        base_video: Source video filename
        segment_details: List of (output_path, start_time_str, end_time_str) tuples
        output_format: 'clip', 'screen', or 'gif'

    Returns:
        List of artifact dicts ready for JSON serialization
    """
    artifacts: List[Dict[str, Any]] = []
    artifact_type = (
        output_format if output_format in ("clip", "screen", "gif") else "clip"
    )

    cell = clip.get("cell")
    cell_row = getattr(cell, "row", None)
    cell_col = getattr(cell, "col", None)
    try:
        cell_a1 = (
            gspread.utils.rowcol_to_a1(cell_row, cell_col)
            if cell_row and cell_col
            else ""
        )
    except Exception:
        cell_a1 = ""

    annotations = list(clip.get("cell_annotations", []))

    for seg_idx, (out_path, start_str, end_str) in enumerate(segment_details):
        start_sec = utils.timestamp_to_seconds(start_str)
        end_sec = utils.timestamp_to_seconds(end_str)

        artifacts.append(
            {
                "id": f"a{cell_row or 0}c{cell_col or 0}s{seg_idx}",
                "type": artifact_type,
                "file": Path(out_path).name,
                "start": start_sec,
                "end": end_sec,
                "thumbnail": "",
                "study": clip.get("study", ""),
                "participant": clip.get("participant", ""),
                "category": clip.get("category", ""),
                "severity": clip.get("severity", ""),
                "description": clip.get("desc", ""),
                "cellRow": cell_row,
                "cellCol": cell_col,
                "cellA1": cell_a1,
                "annotations": annotations,
                "sourceVideo": base_video,
            }
        )
    return artifacts


def finalize_timeline_data(
    artifacts: List[Dict[str, Any]],
    *,
    reels: Optional[List[Dict[str, Any]]] = None,
    study: str = "",
    participant: str = "",
    worksheet_title: str = "",
    is_excel: bool = False,
    mode: str = "",
    output_format: str = "clip",
    screenspace_events: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """Construct the full window.CLIPGEN_DATA structure for the timeline viewer."""
    valid_artifacts = [a for a in artifacts if _is_valid_artifact(a)]
    dropped = len(artifacts) - len(valid_artifacts)
    if dropped:
        dropped_ids = [
            a.get("id", "<no-id>") for a in artifacts if not _is_valid_artifact(a)
        ]
        utils.warning_print(
            f"Skipped {dropped} artifact(s) with missing required fields "
            f"(ids: {', '.join(str(i) for i in dropped_ids[:5])})."
        )
    artifacts = valid_artifacts

    max_time = 0.0
    for a in artifacts:
        end = a.get("end") or a.get("start") or 0
        if end and end > max_time:
            max_time = float(end)

    duration = max_time * 1.05 if max_time > 0 else 0.0

    data: Dict[str, Any] = {
        "meta": {
            "study": study,
            "participant": participant,
            "generatedAt": datetime.now(timezone.utc).isoformat(),
            "mode": mode,
            "sourceSpreadsheet": worksheet_title,
            "sourceFileType": "excel" if is_excel else "google",
            "filmstripEnabled": config.FILMSTRIP_ENABLED,
        },
        "artifacts": artifacts,
        "timeline": {
            "duration": duration,
            "startOffset": 0.0,
        },
    }
    if reels:
        data["reels"] = reels
    if screenspace_events:
        data["meta"]["screenspaceEnabled"] = True
        data["screenspaceEvents"] = screenspace_events
    return data


def _sanitize_event_metadata(obj: Any) -> Any:
    """Replace non-finite floats (inf, nan) with None for JSON safety."""
    if isinstance(obj, float):
        return obj if math.isfinite(obj) else None
    if isinstance(obj, dict):
        return {k: _sanitize_event_metadata(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize_event_metadata(v) for v in obj]
    return obj


def load_screenspace_events_for_viewer() -> List[Dict[str, Any]]:
    """Load non-excluded events from screenspace manifest for viewer export."""
    import screenspace

    manifest = screenspace.load_screenspace_manifest()
    return [
        {
            "id": e["id"],
            "type": e["detector"],
            "eventType": e["event_type"],
            "participant": e["participant"],
            "timeIn": e["time_in"],
            "timeOut": e["time_out"],
            "confidence": e["confidence"],
            "region": e.get("region", ""),
            "metadata": _sanitize_event_metadata(e.get("metadata", {})),
        }
        for e in manifest.get("events", [])
        if not e.get("excluded")
    ]


_CLIPGEN_DATA_PLACEHOLDER = "<!-- CLIPGEN_DATA_HERE -->"


def _generate_viewer_html(
    data: Dict[str, Any],
    *,
    template_name: str,
    js_name: str,
    css_name: str,
    output_basename: str,
    viewer_label: str,
) -> Optional[Path]:
    """Build a self-contained HTML viewer by inlining JS/CSS and injecting data.

    Shared implementation for timeline and gallery viewer generation.
    Returns the path to the generated HTML, or None on failure.
    """
    assets_base = utils.get_bundled_assets_root()
    assets_dir = assets_base / "assets" / "web"
    template_path = assets_dir / template_name
    js_path = assets_dir / js_name
    css_path = assets_dir / css_name

    for required in (template_path, js_path, css_path):
        if not required.is_file():
            utils.warning_print(
                f"{viewer_label} asset not found: {required}",
                [f"{viewer_label} HTML will not be generated."],
            )
            return None

    try:
        template_html = template_path.read_text(encoding="utf-8")
    except OSError as e:
        utils.warning_print(f"Could not read {viewer_label.lower()} template: {e}")
        return None

    try:
        css_text = css_path.read_text(encoding="utf-8")
        js_text = js_path.read_text(encoding="utf-8")
    except OSError as e:
        utils.warning_print(f"Could not read {viewer_label.lower()} assets: {e}")
        return None

    # Prepend design tokens so standalone viewers have the full token set
    tokens_path = assets_dir / "tokens.css"
    if tokens_path.is_file():
        try:
            css_text = tokens_path.read_text(encoding="utf-8") + "\n" + css_text
        except OSError:
            pass

    # Prepend shared utilities so standalone viewers have them
    utils_js_path = assets_dir / "utils.js"
    if utils_js_path.is_file():
        try:
            js_text = utils_js_path.read_text(encoding="utf-8") + "\n" + js_text
        except OSError:
            pass

    # Inline CSS
    css_link_tag = f'<link rel="stylesheet" href="{css_name}">'
    inline_css_block = f"<style>\n{css_text}\n</style>"
    if css_link_tag in template_html:
        template_html = template_html.replace(css_link_tag, inline_css_block)
    elif "</head>" in template_html:
        template_html = template_html.replace("</head>", f"{inline_css_block}\n</head>")

    # Remove utils.js script tag (content already prepended to main JS)
    utils_js_tag = '<script src="utils.js" defer></script>\n  '
    template_html = template_html.replace(utils_js_tag, "")

    # Inline JS
    js_script_tag = f'<script src="{js_name}" defer></script>'
    inline_js_block = f"<script defer>\n{js_text}\n</script>"
    if js_script_tag in template_html:
        template_html = template_html.replace(js_script_tag, inline_js_block)
    elif "</body>" in template_html:
        template_html = template_html.replace("</body>", f"{inline_js_block}\n</body>")

    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    script_block = f"<script>window.CLIPGEN_DATA={data_json};</script>"

    if _CLIPGEN_DATA_PLACEHOLDER in template_html:
        output_html = template_html.replace(_CLIPGEN_DATA_PLACEHOLDER, script_block)
    else:
        output_html = template_html.replace("</body>", f"{script_block}\n</body>")

    out_name = files.get_unique_filename(output_basename, file_format=".html")
    out_path = Path(out_name)

    try:
        out_path.write_text(output_html, encoding="utf-8")
    except OSError as e:
        utils.warning_print(f"Could not write {viewer_label.lower()} HTML: {e}")
        return None

    return out_path


def generate_timeline_viewer(
    data: Dict[str, Any],
    *,
    output_basename: str = "clips_viewer.html",
    template_name: str = "viewer.html",
) -> Optional[Path]:
    """Create a per-run timeline viewer HTML file with inlined JS/CSS."""
    return _generate_viewer_html(
        data,
        template_name=template_name,
        js_name="viewer.js",
        css_name="viewer.css",
        output_basename=output_basename,
        viewer_label="Timeline viewer",
    )


def finalize_gallery_data(
    artifacts: List[Dict[str, Any]],
    *,
    source_video: str = "",
    video_duration: int = 0,
    output_format: str = "screen",
    interval: int = 10,
) -> Dict[str, Any]:
    """Construct the window.CLIPGEN_DATA structure for the gallery viewer."""
    return {
        "meta": {
            "sourceVideo": source_video,
            "generatedAt": datetime.now(timezone.utc).isoformat(),
            "mode": "gallery",
            "format": output_format,
            "interval": interval,
            "videoDuration": video_duration,
        },
        "artifacts": artifacts,
    }


def generate_gallery_viewer(
    data: Dict[str, Any],
    *,
    output_basename: str = "gallery_viewer.html",
) -> Optional[Path]:
    """Create a gallery viewer HTML file with inlined JS/CSS."""
    return _generate_viewer_html(
        data,
        template_name="gallery.html",
        js_name="gallery.js",
        css_name="gallery.css",
        output_basename=output_basename,
        viewer_label="Gallery viewer",
    )


def finalize_insights_viewer_data(
    insights_list: List[Dict[str, Any]],
    artifacts: List[Dict[str, Any]],
    *,
    study: str = "",
    timeline_viewer_file: str = "",
) -> Dict[str, Any]:
    """Construct the window.CLIPGEN_DATA structure for the insights viewer."""
    # Show only "final" insights if any exist, otherwise show all
    final_insights = [i for i in insights_list if i.get("status") == "final"]
    visible = final_insights if final_insights else list(insights_list)

    # Collect all referenced artifact IDs
    referenced_ids: set = set()
    for ins in visible:
        for bucket in ("causes", "behaviors", "impacts"):
            referenced_ids.update(ins.get(bucket, {}).get("artifacts", []))

    referenced_artifacts = [a for a in artifacts if a.get("id") in referenced_ids]

    return {
        "meta": {
            "study": study,
            "generatedAt": datetime.now(timezone.utc).isoformat(),
            "version": config.VERSIONNUM,
            "timelineViewerFile": timeline_viewer_file,
        },
        "insights": visible,
        "artifacts": referenced_artifacts,
    }


def generate_insights_viewer(
    data: Dict[str, Any],
    *,
    output_basename: str = "insights_viewer.html",
) -> Optional[Path]:
    """Create an insights viewer HTML file with inlined JS/CSS."""
    return _generate_viewer_html(
        data,
        template_name="insights-viewer.html",
        js_name="insights-viewer.js",
        css_name="insights-viewer.css",
        output_basename=output_basename,
        viewer_label="Insights viewer",
    )


def _load_manifest_both() -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Load artifact and reel records from the manifest in a single read.

    Returns (artifacts, reels). Both default to [] on missing/corrupt file.
    """
    data = utils.load_json_manifest(config.MANIFEST_FILENAME)
    if not isinstance(data, dict):
        return ([], [])
    raw = data.get("artifacts", [])
    valid = [a for a in raw if _is_valid_artifact(a)]
    if len(valid) < len(raw):
        utils.warning_print(
            f"Manifest contained {len(raw) - len(valid)} artifact(s) with "
            "missing fields; skipped."
        )
    return (valid, data.get("reels", []))


def load_manifest_artifacts() -> List[Dict[str, Any]]:
    """Load artifact records from the manifest file, or return [] if unavailable."""
    artifacts, _ = _load_manifest_both()
    return artifacts


def load_manifest_reels() -> List[Dict[str, Any]]:
    """Load reel records from the manifest file, or return [] if unavailable."""
    _, reels = _load_manifest_both()
    return reels


def save_manifest(
    new_artifacts: List[Dict[str, Any]],
    *,
    new_reels: Optional[List[Dict[str, Any]]] = None,
    study: str = "",
    participant: str = "",
    worksheet_title: str = "",
    is_excel: bool = False,
    mode: str = "",
    output_format: str = "clip",
) -> Optional[Path]:
    """Merge new artifacts and reels into the manifest file and write it back.

    Deduplicates by ``id``; newer entries win.
    Returns the manifest path on success, or None on failure.
    """
    existing, existing_reels = _load_manifest_both()
    merged = {a["id"]: a for a in existing}
    for a in new_artifacts:
        merged[a["id"]] = a
    all_artifacts = list(merged.values())

    reel_merged = {r["id"]: r for r in existing_reels}
    for r in new_reels or []:
        reel_merged[r["id"]] = r
    all_reels = list(reel_merged.values())

    data = finalize_timeline_data(
        all_artifacts,
        reels=all_reels or None,
        study=study,
        participant=participant,
        worksheet_title=worksheet_title,
        is_excel=is_excel,
        mode=mode,
        output_format=output_format,
    )

    return utils.save_json_manifest(
        config.MANIFEST_FILENAME, data, warn_label="manifest"
    )
