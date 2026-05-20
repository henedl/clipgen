# -*- coding: utf-8 -*-
"""Timeline and gallery viewer generation, manifest persistence.

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

Artifact manifest (save_manifest / load_manifest_*):
  Merges new artifacts/reels into clipgen_manifest.json, deduplicating by id (newer wins).
  Consumed by --regenerate and standalone --viewer.
"""

import base64
import json
import math
import re
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


import config
import files
import utils

# Mutable lists of records collected during an interactive session.
INTERACTIVE_ARTIFACTS: list[dict[str, Any]] = []
INTERACTIVE_REELS: list[dict[str, Any]] = []


def build_artifact_records_for_clip(
    clip: utils.ClipRecord,
    base_video: str,
    segment_details: list[tuple[str, int]],
    output_format: str,
) -> list[dict[str, Any]]:
    """Build artifact metadata records from a processed clip's successful outputs.

    Args:
        clip: Prepared clip dict with 'times', 'category', 'study', etc.
        base_video: Source video filename
        segment_details: List of (output_path, time_index) pairs.
            ``time_index`` is the index into ``clip['times']`` for this segment.
        output_format: 'clip', 'screen', or 'gif'

    Returns:
        List of artifact dicts ready for JSON serialization
    """
    artifact_type = (
        output_format if output_format in ("clip", "screen", "gif") else "clip"
    )
    times = clip.get("times", [])
    return [
        utils.build_artifact_record(
            clip,
            base_video,
            out_path,
            times[time_idx][0],
            times[time_idx][1],
            artifact_type=artifact_type,
            seg_idx=seg_idx,
        )
        for seg_idx, (out_path, time_idx) in enumerate(segment_details)
    ]


def finalize_timeline_data(
    artifacts: list[dict[str, Any]],
    *,
    reels: list[dict[str, Any]] | None = None,
    study: str = "",
    participant: str = "",
    worksheet_title: str = "",
    is_excel: bool = False,
    mode: str = "",
    output_format: str = "clip",
    screenspace_events: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Construct the full window.CLIPGEN_DATA structure for the timeline viewer."""
    max_time = 0.0
    for a in artifacts:
        end = a.get("end") or a.get("start") or 0
        if end and end > max_time:
            max_time = float(end)

    duration = max_time * 1.05 if max_time > 0 else 0.0

    data: dict[str, Any] = {
        "meta": {
            "study": study,
            "participant": participant,
            "generatedAt": datetime.now(timezone.utc).isoformat(),
            "mode": mode,
            "sourceSpreadsheet": worksheet_title,
            "sourceFileType": "excel" if is_excel else "google",
            "filmstripEnabled": config.FILMSTRIP_ENABLED,
        },
        "config": utils.get_frontend_config(),
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


def load_screenspace_events_for_viewer() -> list[dict[str, Any]]:
    """Load non-excluded events from screenspace manifest for viewer export."""
    import screenspace

    manifest = screenspace.load_screenspace_manifest()
    return [
        {
            "id": e.get("id", ""),
            "type": e.get("detector", ""),
            "eventType": e.get("event_type", ""),
            "participant": e.get("participant", ""),
            "timeIn": e.get("time_in", 0.0),
            "timeOut": e.get("time_out", 0.0),
            "confidence": _sanitize_event_metadata(e.get("confidence", 0.0)),
            "region": e.get("region", ""),
            "metadata": _sanitize_event_metadata(e.get("metadata", {})),
        }
        for e in manifest.get("events", [])
        if not e.get("excluded")
    ]


_CLIPGEN_DATA_PLACEHOLDER = "<!-- CLIPGEN_DATA_HERE -->"


def _generate_viewer_html(
    data: dict[str, Any],
    *,
    template_name: str,
    js_name: str,
    css_name: str,
    output_basename: str,
    viewer_label: str,
) -> Path | None:
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

    # Strip dev-only tags (e.g. dev-token-tweak.js) so they never ship in exports.
    template_html = re.sub(
        r"<script\b[^>]*\bdata-dev-only\b[^>]*>\s*</script>\s*",
        "",
        template_html,
        flags=re.IGNORECASE,
    )
    template_html = re.sub(
        r"<link\b[^>]*\bdata-dev-only\b[^>]*/?>\s*",
        "",
        template_html,
        flags=re.IGNORECASE,
    )

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
    data: dict[str, Any],
    *,
    output_basename: str = "clips_viewer.html",
    template_name: str = "viewer.html",
) -> Path | None:
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
    artifacts: list[dict[str, Any]],
    *,
    source_video: str = "",
    video_duration: int = 0,
    output_format: str = "screen",
    interval: int = 10,
    bundle: bool = False,
) -> dict[str, Any]:
    """Construct the window.CLIPGEN_DATA structure for the gallery viewer."""
    if bundle:
        ext_mime_map = {
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".gif": "image/gif",
            ".webp": "image/webp",
            ".webm": "video/webm",
        }
        output_dir = Path(utils.get_effective_output_dir())
        for a in artifacts:
            file_path = output_dir / a["file"]
            if not file_path.is_file():
                utils.warning_print(
                    f"Bundle: file not found, skipping embed: {a['file']}"
                )
                continue
            mime = ext_mime_map.get(file_path.suffix.lower(), "image/png")
            try:
                raw = file_path.read_bytes()
                encoded = base64.b64encode(raw).decode("ascii")
                a["data"] = f"data:{mime};base64,{encoded}"
            except OSError as e:
                utils.warning_print(f"Bundle: could not read {a['file']}: {e}")

    return {
        "meta": {
            "sourceVideo": source_video,
            "generatedAt": datetime.now(timezone.utc).isoformat(),
            "mode": "gallery",
            "format": output_format,
            "interval": interval,
            "videoDuration": video_duration,
            "bundled": bundle,
        },
        "config": utils.get_frontend_config(),
        "artifacts": artifacts,
    }


def generate_gallery_viewer(
    data: dict[str, Any],
    *,
    output_basename: str = "gallery_viewer.html",
) -> Path | None:
    """Create a gallery viewer HTML file with inlined JS/CSS."""
    return _generate_viewer_html(
        data,
        template_name="gallery.html",
        js_name="gallery.js",
        css_name="gallery.css",
        output_basename=output_basename,
        viewer_label="Gallery viewer",
    )


# Module-level cache for the parsed manifest, keyed on the file's path and
# mtime_ns. Studio/Transcripts all hit `load_manifest_artifacts()`
# repeatedly on every request; re-reading and re-parsing the JSON each time is
# pure overhead. The cache is invalidated automatically whenever the file is
# rewritten (save_manifest bumps mtime) so no explicit bust is required in the
# normal happy path — _reset_manifest_cache() exists only for tests that reuse
# the same output directory across mutations without touching mtime.
_MANIFEST_CACHE_LOCK = threading.Lock()
_manifest_cache: dict[str, Any] = {
    "path": None,
    "mtime_ns": None,
    "artifacts": [],
    "reels": [],
}


def _reset_manifest_cache() -> None:
    """Drop the in-memory manifest cache. Intended for test fixtures."""
    with _MANIFEST_CACHE_LOCK:
        _manifest_cache["path"] = None
        _manifest_cache["mtime_ns"] = None
        _manifest_cache["artifacts"] = []
        _manifest_cache["reels"] = []


def _load_manifest_both() -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Load artifact and reel records from the manifest in a single read.

    Returns (artifacts, reels). Both default to [] on missing/corrupt file.

    Memoizes the parsed result keyed on the manifest's path + mtime_ns so
    repeated calls from the same process share a single read/parse until the
    file is rewritten. Returns shallow copies so callers that mutate the
    returned lists do not corrupt the cached state.
    """
    path = Path(utils.get_effective_output_dir()) / config.MANIFEST_FILENAME
    path_str = str(path)
    try:
        mtime_ns: int | None = path.stat().st_mtime_ns if path.is_file() else None
    except OSError:
        mtime_ns = None

    with _MANIFEST_CACHE_LOCK:
        if (
            mtime_ns is not None
            and _manifest_cache["path"] == path_str
            and _manifest_cache["mtime_ns"] == mtime_ns
        ):
            return (
                list(_manifest_cache["artifacts"]),
                list(_manifest_cache["reels"]),
            )

    data = utils.load_json_manifest(
        config.MANIFEST_FILENAME, default={"artifacts": [], "reels": []}
    )
    artifacts = data.get("artifacts", [])
    reels = data.get("reels", [])

    with _MANIFEST_CACHE_LOCK:
        _manifest_cache["path"] = path_str
        _manifest_cache["mtime_ns"] = mtime_ns
        _manifest_cache["artifacts"] = artifacts
        _manifest_cache["reels"] = reels

    return (list(artifacts), list(reels))


def load_manifest_artifacts() -> list[dict[str, Any]]:
    """Load artifact records from the manifest file, or return [] if unavailable."""
    artifacts, _ = _load_manifest_both()
    return artifacts


def load_manifest_reels() -> list[dict[str, Any]]:
    """Load reel records from the manifest file, or return [] if unavailable."""
    _, reels = _load_manifest_both()
    return reels


def save_manifest(
    new_artifacts: list[dict[str, Any]],
    *,
    new_reels: list[dict[str, Any]] | None = None,
    study: str = "",
    participant: str = "",
    worksheet_title: str = "",
    is_excel: bool = False,
    mode: str = "",
    output_format: str = "clip",
) -> Path | None:
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

    result = utils.save_json_manifest(
        config.MANIFEST_FILENAME, data, warn_label="manifest"
    )
    # Invalidate the cache so the next load picks up what we just wrote,
    # even if the filesystem's mtime resolution elides the change.
    if result is not None:
        _reset_manifest_cache()
    return result
