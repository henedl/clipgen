# -*- coding: utf-8 -*-
"""Timeline viewer generation for clipgen.

Builds artifact metadata records from processed clips and generates
a self-contained HTML timeline viewer with inlined CSS/JS.
"""

import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

import gspread

import config
import files
import utils

# Mutable list of artifact records collected during an interactive session.
INTERACTIVE_ARTIFACTS: List[Dict[str, Any]] = []


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
                "id": f"a{cell_row}c{cell_col}s{seg_idx}",
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
    study: str = "",
    participant: str = "",
    worksheet_title: str = "",
    is_excel: bool = False,
    mode: str = "",
    output_format: str = "clip",
) -> Dict[str, Any]:
    """Construct the full window.CLIPGEN_DATA structure for the timeline viewer."""
    max_time = 0.0
    for a in artifacts:
        end = a.get("end") or a.get("start") or 0
        if end and end > max_time:
            max_time = float(end)

    duration = max_time * 1.05 if max_time > 0 else 0.0

    return {
        "meta": {
            "study": study,
            "participant": participant,
            "generatedAt": datetime.now(timezone.utc).isoformat(),
            "mode": mode,
            "sourceSpreadsheet": worksheet_title,
            "sourceFileType": "excel" if is_excel else "google",
        },
        "artifacts": artifacts,
        "timeline": {
            "duration": duration,
            "startOffset": 0.0,
        },
    }


_CLIPGEN_DATA_PLACEHOLDER = "<!-- CLIPGEN_DATA_HERE -->"


def generate_timeline_viewer(
    data: Dict[str, Any],
    *,
    output_basename: str = "clips_viewer.html",
    template_name: str = "viewer.html",
) -> Optional[Path]:
    """Create a per-run viewer HTML file with inlined JS/CSS.

    Reads the static template from the assets/web directory,
    injects the serialized data as window.CLIPGEN_DATA, writes the result
    into the effective output directory.

    Returns the path to the generated HTML, or None on failure.
    """
    if getattr(sys, "frozen", False):
        assets_base = Path(sys.executable).resolve().parent
    else:
        assets_base = Path(__file__).resolve().parent
    assets_dir = assets_base / "assets" / "web"
    template_path = assets_dir / template_name
    js_path = assets_dir / "viewer.js"
    css_path = assets_dir / "viewer.css"

    for required in (template_path, js_path, css_path):
        if not required.is_file():
            utils.warning_print(
                f"Timeline viewer asset not found: {required}",
                ["Viewer HTML will not be generated."],
            )
            return None

    try:
        template_html = template_path.read_text(encoding="utf-8")
    except OSError as e:
        utils.warning_print(f"Could not read viewer template: {e}")
        return None

    try:
        css_text = css_path.read_text(encoding="utf-8")
        js_text = js_path.read_text(encoding="utf-8")
    except OSError as e:
        utils.warning_print(f"Could not read viewer assets: {e}")
        return None

    # Inline CSS
    css_link_tag = '<link rel="stylesheet" href="viewer.css">'
    inline_css_block = f"<style>\n{css_text}\n</style>"
    if css_link_tag in template_html:
        template_html = template_html.replace(css_link_tag, inline_css_block)
    elif "</head>" in template_html:
        template_html = template_html.replace("</head>", f"{inline_css_block}\n</head>")

    # Inline JS
    js_script_tag = '<script src="viewer.js" defer></script>'
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

    # Let files.get_unique_filename resolve against the effective output directory
    out_name = files.get_unique_filename(output_basename, file_format=".html")
    out_path = Path(out_name)

    try:
        out_path.write_text(output_html, encoding="utf-8")
    except OSError as e:
        utils.warning_print(f"Could not write viewer HTML: {e}")
        return None

    return out_path


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
    """Create a gallery viewer HTML file with inlined JS/CSS.

    Reads the gallery template from assets/web, injects serialized data
    as window.CLIPGEN_DATA, and writes the result to the output directory.

    Returns the path to the generated HTML, or None on failure.
    """
    if getattr(sys, "frozen", False):
        assets_base = Path(sys.executable).resolve().parent
    else:
        assets_base = Path(__file__).resolve().parent
    assets_dir = assets_base / "assets" / "web"
    template_path = assets_dir / "gallery.html"
    js_path = assets_dir / "gallery.js"
    css_path = assets_dir / "gallery.css"

    for required in (template_path, js_path, css_path):
        if not required.is_file():
            utils.warning_print(
                f"Gallery viewer asset not found: {required}",
                ["Gallery HTML will not be generated."],
            )
            return None

    try:
        template_html = template_path.read_text(encoding="utf-8")
    except OSError as e:
        utils.warning_print(f"Could not read gallery template: {e}")
        return None

    try:
        css_text = css_path.read_text(encoding="utf-8")
        js_text = js_path.read_text(encoding="utf-8")
    except OSError as e:
        utils.warning_print(f"Could not read gallery assets: {e}")
        return None

    css_link_tag = '<link rel="stylesheet" href="gallery.css">'
    inline_css_block = f"<style>\n{css_text}\n</style>"
    if css_link_tag in template_html:
        template_html = template_html.replace(css_link_tag, inline_css_block)
    elif "</head>" in template_html:
        template_html = template_html.replace("</head>", f"{inline_css_block}\n</head>")

    js_script_tag = '<script src="gallery.js" defer></script>'
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
        utils.warning_print(f"Could not write gallery HTML: {e}")
        return None

    return out_path


def load_manifest_artifacts() -> List[Dict[str, Any]]:
    """Load artifact records from the manifest file, or return [] if unavailable."""
    manifest_path = Path(utils.get_effective_output_dir()) / config.MANIFEST_FILENAME
    if not manifest_path.is_file():
        return []
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
        return data.get("artifacts", [])
    except (OSError, json.JSONDecodeError, AttributeError):
        return []


def save_manifest(
    new_artifacts: List[Dict[str, Any]],
    *,
    study: str = "",
    participant: str = "",
    worksheet_title: str = "",
    is_excel: bool = False,
    mode: str = "",
    output_format: str = "clip",
) -> Optional[Path]:
    """Merge new artifacts into the manifest file and write it back.

    Deduplicates by artifact ``id``; newer entries win.
    Returns the manifest path on success, or None on failure.
    """
    existing = load_manifest_artifacts()
    merged = {a["id"]: a for a in existing}
    for a in new_artifacts:
        merged[a["id"]] = a
    all_artifacts = list(merged.values())

    data = finalize_timeline_data(
        all_artifacts,
        study=study,
        participant=participant,
        worksheet_title=worksheet_title,
        is_excel=is_excel,
        mode=mode,
        output_format=output_format,
    )

    manifest_path = Path(utils.get_effective_output_dir()) / config.MANIFEST_FILENAME
    try:
        manifest_path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return manifest_path
    except OSError as e:
        utils.warning_print(f"Could not write manifest: {e}")
        return None
