"""Timeline and gallery viewer generation, manifest persistence.

Both viewers inject ``window.CLIPGEN_DATA`` into their HTML template, replacing
``<!-- CLIPGEN_DATA_HERE -->``.

Timeline (--viewer / interactive 'viewer'), from viewer.html::

    { meta: {study, participant, generatedAt, clipgenVersion, repoUrl, mode,
      sourceSpreadsheet, sourceFileType, filmstripEnabled},
      artifacts: [{id, type, file, start, end, study, participant, category,
      description, cellRow, cellCol, cellA1, annotations, sourceVideo}],
      timeline: {duration, startOffset} }

Artifact ``type`` is clip / screen / gif / reel (timeline events) or the
non-timeline "attachment" types timelapse / heatmap / export, which carry
start/end 0 — the viewer JS branches on type and gives those their own panel.

Gallery (--gallery), from gallery.html::

    { meta: {sourceVideo, generatedAt, clipgenVersion, repoUrl, mode, format,
      interval, videoDuration},
      artifacts: [{file, timestamp, timestamp_formatted, type, duration}] }

Gallery artifacts are NOT written to the manifest by default.

The manifest helpers merge new artifacts/reels into the manifest's clips section,
deduplicating by id (newer wins), for --regenerate and standalone --viewer.
"""

import base64
import functools
import json
import re
import threading
from datetime import UTC, datetime
from pathlib import Path
from typing import Any


import config
import files
import utils

# Mutable lists of records collected during an interactive session.
INTERACTIVE_ARTIFACTS: list[dict[str, Any]] = []
INTERACTIVE_REELS: list[dict[str, Any]] = []


def _export_config() -> dict[str, Any]:
    """Frontend config for exported viewers.

    Exported HTML always runs the default keymap (a standalone file has no
    settings UI to inspect or fix a surprising binding), so the user's
    hotkey overrides are stripped from the embedded payload.
    """
    cfg = utils.get_frontend_config()
    cfg.pop("hotkeyOverrides", None)
    # Exported HTML never profiles: the flag mirrors a live --profile launch,
    # and a standalone file has no report sink to read the marks from.
    cfg.pop("profiling", None)
    return cfg


def build_artifact_records_for_clip(
    clip: utils.ClipRecord,
    base_video: str,
    segment_details: list[tuple[str, int]],
    output_format: str,
    *,
    titlecards: bool = False,
    titlecard_duration: int = 0,
    titlecard_image: str = "",
    endcard_image: str = "",
) -> list[dict[str, Any]]:
    """Build artifact metadata records from a processed clip's successful outputs.

    Args:
        clip: Prepared clip dict with 'times', 'category', 'study', etc.
        base_video: Source video filename
        segment_details: List of (output_path, time_index) pairs.
            ``time_index`` is the index into ``clip['times']`` for this segment.
        output_format: 'clip', 'screen', or 'gif'
        titlecards: Whether titlecards were applied (recorded on clip artifacts
            only; lets the Studio skip logic and manifest regeneration detect a
            mismatch with the requested titlecard setting).
        titlecard_duration: Titlecard duration in seconds (clip artifacts only).
        titlecard_image: Selected titlecard background id (clip artifacts only).
        endcard_image: Selected endcard background id (clip artifacts only).

    Returns:
        List of artifact dicts ready for JSON serialization
    """
    artifact_type = (
        output_format if output_format in ("clip", "screen", "gif") else "clip"
    )
    times = clip.get("times", [])
    records = [
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
    if artifact_type == "clip":
        for record in records:
            record["titlecards"] = titlecards
            record["titlecardDuration"] = titlecard_duration
            record["titlecardImage"] = titlecard_image
            record["endcardImage"] = endcard_image
    return records


def finalize_timeline_data(
    artifacts: list[dict[str, Any]],
    *,
    reels: list[dict[str, Any]] | None = None,
    study: str = "",
    participant: str = "",
    worksheet_title: str = "",
    is_excel: bool = False,
    mode: str = "",
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
            "generatedAt": datetime.now(UTC).isoformat(),
            "clipgenVersion": utils.get_version(),
            "repoUrl": config.REPO_URL,
            "mode": mode,
            "sourceSpreadsheet": worksheet_title,
            "sourceFileType": "excel" if is_excel else "google",
            "filmstripEnabled": config.FILMSTRIP_ENABLED,
        },
        "config": _export_config(),
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


# Module-level mtime cache for the screenspace-events-for-viewer transform.
# Viewer exports call load_screenspace_events_for_viewer() repeatedly; re-parsing
# the screenspace section each time is pure overhead. Keyed on the manifest's
# (output dir, mtime_ns), so any rewrite invalidates it. One entry suffices.
_SS_EVENTS_CACHE_LOCK = threading.Lock()
_ss_events_cache: dict[str, Any] = {
    "path": None,
    "mtime_ns": None,
    "events": [],
}


def _reset_screenspace_events_cache() -> None:
    """Drop the in-memory screenspace-events cache. Intended for test fixtures."""
    with _SS_EVENTS_CACHE_LOCK:
        _ss_events_cache["path"] = None
        _ss_events_cache["mtime_ns"] = None
        _ss_events_cache["events"] = []


def load_screenspace_events_for_viewer() -> list[dict[str, Any]]:
    """Load non-excluded events from screenspace manifest for viewer export.

    Memoizes the transformed event list keyed on the manifest's path + mtime_ns
    so repeated viewer exports share a single read/parse until the file changes.
    """
    import screenspace

    path_str = str(utils.get_effective_output_dir())
    mtime_ns: int | None = utils.manifest_mtime() or None

    with _SS_EVENTS_CACHE_LOCK:
        if (
            mtime_ns is not None
            and _ss_events_cache["path"] == path_str
            and _ss_events_cache["mtime_ns"] == mtime_ns
        ):
            return list(_ss_events_cache["events"])

    manifest = screenspace.load_screenspace_manifest()
    events = [
        {
            "id": e.get("id", ""),
            "type": e.get("detector", ""),
            "eventType": e.get("event_type", ""),
            "participant": e.get("participant", ""),
            "timeIn": utils.sanitize_floats(e.get("time_in", 0.0)),
            "timeOut": utils.sanitize_floats(e.get("time_out", 0.0)),
            "confidence": utils.sanitize_floats(e.get("confidence", 0.0)),
            "region": e.get("region", ""),
            "navigational": bool(e.get("navigational", False)),
            "metadata": utils.sanitize_floats(e.get("metadata", {})),
        }
        for e in manifest.get("events", [])
        if not e.get("excluded")
    ]

    with _SS_EVENTS_CACHE_LOCK:
        _ss_events_cache["path"] = path_str
        _ss_events_cache["mtime_ns"] = mtime_ns
        _ss_events_cache["events"] = events

    return list(events)


_CLIPGEN_DATA_PLACEHOLDER = "<!-- CLIPGEN_DATA_HERE -->"


@functools.cache
def _read_bundled_asset(path_str: str) -> str:
    """Read a bundled asset file as text, cached for the process lifetime.

    Bundled CSS/JS/HTML templates never change at runtime, so repeated viewer
    exports share a single read instead of re-hitting disk each time.
    """
    return Path(path_str).read_text(encoding="utf-8")


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
        template_html = _read_bundled_asset(str(template_path))
    except OSError as e:
        utils.warning_print(f"Could not read {viewer_label.lower()} template: {e}")
        return None

    try:
        css_text = _read_bundled_asset(str(css_path))
        js_text = _read_bundled_asset(str(js_path))
    except OSError as e:
        utils.warning_print(f"Could not read {viewer_label.lower()} assets: {e}")
        return None

    # Prepend design tokens so standalone viewers have the full token set
    tokens_path = assets_dir / "tokens.css"
    if tokens_path.is_file():
        try:
            css_text = _read_bundled_asset(str(tokens_path)) + "\n" + css_text
        except OSError:
            pass

    # Inline the shared hotkey registry (cheatsheet + dispatcher). Exported
    # viewers always run the default keymap: the embedded config deliberately
    # omits hotkeyOverrides (see _export_config).
    hk_css_tag = '<link rel="stylesheet" href="hotkeys.css">'
    hk_js_tag = '<script src="hotkeys.js" defer></script>'
    if hk_css_tag in template_html:
        hk_css_path = assets_dir / "hotkeys.css"
        if hk_css_path.is_file():
            try:
                css_text = css_text + "\n" + _read_bundled_asset(str(hk_css_path))
            except OSError:
                pass
        template_html = template_html.replace(hk_css_tag, "")
    if hk_js_tag in template_html:
        hk_js_path = assets_dir / "hotkeys.js"
        if hk_js_path.is_file():
            try:
                # utils.js is prepended below, ahead of this, so the final
                # order stays utils -> hotkeys -> page modules.
                js_text = _read_bundled_asset(str(hk_js_path)) + "\n" + js_text
            except OSError:
                pass
        template_html = template_html.replace(hk_js_tag, "")

    # Inline the shared motion engine (ClipgenMotion) so exported viewers animate
    # the same way the live pages do — an export has no asset routes, so without
    # this the toast's guarded fade silently no-ops and it snaps instead. JS-only:
    # motion.js ships no stylesheet. Position among these blocks is immaterial —
    # utils.js is prepended last (so it always ends up first) and every consumer
    # reads window.ClipgenMotion lazily inside a function, never at load time.
    mo_js_tag = '<script src="motion.js" defer></script>'
    if mo_js_tag in template_html:
        mo_js_path = assets_dir / "motion.js"
        if mo_js_path.is_file():
            try:
                js_text = _read_bundled_asset(str(mo_js_path)) + "\n" + js_text
            except OSError:
                pass
        template_html = template_html.replace(mo_js_tag, "")

    # Inline the card-scrubber module into viewers that reference it (timeline,
    # not gallery). Its CSS/JS join the shared bundles so the export stays
    # self-contained; the external tags are stripped below.
    cs_css_tag = '<link rel="stylesheet" href="card-scrubber.css">'
    cs_js_tag = '<script src="card-scrubber.js" defer></script>'
    if cs_css_tag in template_html:
        cs_css_path = assets_dir / "card-scrubber.css"
        if cs_css_path.is_file():
            try:
                css_text = css_text + "\n" + _read_bundled_asset(str(cs_css_path))
            except OSError:
                pass
        template_html = template_html.replace(cs_css_tag, "")
    if cs_js_tag in template_html:
        cs_js_path = assets_dir / "card-scrubber.js"
        if cs_js_path.is_file():
            try:
                js_text = _read_bundled_asset(str(cs_js_path)) + "\n" + js_text
            except OSError:
                pass
        template_html = template_html.replace(cs_js_tag, "")

    # Prepend shared utilities so standalone viewers have them
    utils_js_path = assets_dir / "utils.js"
    if utils_js_path.is_file():
        try:
            js_text = _read_bundled_asset(str(utils_js_path)) + "\n" + js_text
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
    # Escape chars that could break out of the <script> tag or a JS string literal
    # (</script>, <!--, <script, and the U+2028/U+2029 line separators). All stay valid
    # JSON via \uXXXX, so JSON.parse decodes the payload back to the original unchanged.
    data_json = (
        data_json.replace("<", "\\u003c")
        .replace(">", "\\u003e")
        .replace("&", "\\u0026")
        .replace(" ", "\\u2028")
        .replace(" ", "\\u2029")
    )
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
            "generatedAt": datetime.now(UTC).isoformat(),
            "clipgenVersion": utils.get_version(),
            "repoUrl": config.REPO_URL,
            "mode": "gallery",
            "format": output_format,
            "interval": interval,
            "videoDuration": video_duration,
            "bundled": bundle,
        },
        "config": _export_config(),
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


# Serializes the full load-merge-write cycle in save_manifest() so concurrent
# Studio completions cannot last-writer-wins a partial merge.
_MANIFEST_WRITE_LOCK = threading.Lock()


def load_manifest_both() -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Load artifact and reel records from the manifest's ``clips`` section.

    Returns (artifacts, reels). Both default to [] on missing/corrupt file.
    """
    data = utils.load_manifest_section("clips", default={})
    if not isinstance(data, dict):
        return [], []
    return data.get("artifacts", []), data.get("reels", [])


def load_manifest_artifacts() -> list[dict[str, Any]]:
    """Load artifact records from the manifest file, or return [] if unavailable."""
    artifacts, _ = load_manifest_both()
    return artifacts


def save_manifest(
    new_artifacts: list[dict[str, Any]],
    *,
    new_reels: list[dict[str, Any]] | None = None,
    study: str = "",
    participant: str = "",
    worksheet_title: str = "",
    is_excel: bool = False,
    mode: str = "",
) -> Path | None:
    """Merge new artifacts and reels into the manifest file and write it back.

    Deduplicates by ``id``; newer entries win.
    Returns the manifest path on success, or None on failure.
    """
    # Hold the write lock across the whole load-merge-write cycle: a concurrent
    # writer must see this writer's persisted result before computing its merge,
    # otherwise both read the same old manifest and the second write drops the
    # first writer's records.
    with _MANIFEST_WRITE_LOCK:
        existing, existing_reels = load_manifest_both()
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
        )

        return utils.save_manifest_section("clips", data)
