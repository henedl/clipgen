"""Verify that JS fallback constants in utils.js match their config.py defaults.

The live values are repopulated at runtime from the server, but utils.js carries
hardcoded defaults so exported viewers and the brief boot window before the
first /api/marks fetch still render with the same colors.
"""

import json
import re
from pathlib import Path

import cli
import config
import screenspace_tools
import utils
import workflows_catalog

from _frontend_source import WEB

JS_PATH = WEB / "utils.js"


def _js_source() -> str:
    return JS_PATH.read_text(encoding="utf-8")


def _parse_js_object_literal(name: str) -> dict:
    """Extract a top-level JS object literal `var <name> = { ... };` as a dict.

    Handles unquoted keys, trailing commas, and whole-line `//` comments (the
    fallback block documents which Python constant each key mirrors); values
    must be JSON-compatible. Only line-leading comments are stripped, so a
    value containing `//` is never mangled.
    """
    match = re.search(
        r"var\s+" + re.escape(name) + r"\s*=\s*(\{.+?\n\});",
        _js_source(),
        re.DOTALL,
    )
    assert match, f"{name} not found in utils.js"
    raw = match.group(1)
    raw = re.sub(r"^\s*//.*$", "", raw, flags=re.MULTILINE)
    raw = re.sub(r"(\b\w+)\s*:", r'"\1":', raw)
    raw = re.sub(r",\s*([}\]])", r"\1", raw)
    return json.loads(raw)


def _parse_js_mark_categories() -> dict:
    """Extract the MARK_CATEGORIES fallback from assets/web/utils.js."""
    match = re.search(
        r"var\s+MARK_CATEGORIES\s*=\s*\{(.+?)\};",
        _js_source(),
        re.DOTALL,
    )
    assert match, "MARK_CATEGORIES not found in utils.js"
    raw = "{" + match.group(1) + "}"
    # Convert JS object to valid JSON: unquoted keys → quoted keys
    raw = re.sub(r"(\w+)\s*:", r'"\1":', raw)
    # Remove trailing commas
    raw = re.sub(r",\s*([}\]])", r"\1", raw)
    return json.loads(raw)


def test_mark_categories_match_python():
    js_cats = _parse_js_mark_categories()
    py_cats = config.MARK_CATEGORIES
    assert set(js_cats.keys()) == set(py_cats.keys()), (
        f"Key mismatch: JS={sorted(js_cats)} vs Python={sorted(py_cats)}"
    )
    for key in py_cats:
        assert js_cats[key]["label"] == py_cats[key]["label"], (
            f"MARK_CATEGORIES[{key!r}].label differs"
        )
        assert js_cats[key]["color"] == py_cats[key]["color"], (
            f"MARK_CATEGORIES[{key!r}].color differs"
        )


def test_clipgen_config_defaults_match_python():
    """utils.js CLIPGEN_CONFIG fallback must match utils.get_frontend_config().

    The fallback is what a re-opened older exported viewer (lacking the
    `config` field in window.CLIPGEN_DATA) sees, so it must mirror the
    canonical Python config exactly.
    """
    js_config = _parse_js_object_literal("CLIPGEN_CONFIG")
    py_config = utils.get_frontend_config()

    assert js_config["defaultDuration"] == py_config["defaultDuration"]
    assert js_config["annotationKeyphrases"] == py_config["annotationKeyphrases"]
    assert js_config["annotations"] == py_config["annotations"]
    assert js_config["ignoredTimestampTokens"] == py_config["ignoredTimestampTokens"]
    assert (
        js_config["screenspaceOcrMinConfidence"]
        == py_config["screenspaceOcrMinConfidence"]
    )
    assert (
        js_config["screenspaceOcrFuzzyThreshold"]
        == py_config["screenspaceOcrFuzzyThreshold"]
    )
    assert (
        js_config["screenspaceMultitoolMaxOffset"]
        == py_config["screenspaceMultitoolMaxOffset"]
    )
    assert (
        js_config["screenspaceMaskFallbackTools"]
        == py_config["screenspaceMaskFallbackTools"]
    )

    assert len(js_config["severity"]) == len(py_config["severity"])
    for js_sev, py_sev in zip(js_config["severity"], py_config["severity"]):
        assert js_sev["label"] == py_sev["label"]
        assert js_sev["rank"] == py_sev["rank"]
        assert js_sev["cssClass"] == py_sev["cssClass"]

    assert js_config["frictionColorToken"] == py_config["frictionColorToken"]
    assert js_config["frictionMomentLimit"] == py_config["frictionMomentLimit"]
    assert js_config["frictionCategories"] == py_config["frictionCategories"]
    assert js_config["convergenceSources"] == py_config["convergenceSources"]
    assert js_config["cardScrubberSpriteCols"] == py_config["cardScrubberSpriteCols"]
    assert js_config["cardScrubberSpriteRows"] == py_config["cardScrubberSpriteRows"]
    assert js_config["clipFormat"] == py_config["clipFormat"]
    assert js_config["screenshotFormat"] == py_config["screenshotFormat"]
    assert js_config["gifFormat"] == py_config["gifFormat"]
    assert js_config["composerAnnotationColor"] == py_config["composerAnnotationColor"]
    assert (
        js_config["composerAnnotationColorSecondary"]
        == py_config["composerAnnotationColorSecondary"]
    )
    assert (
        js_config["composerAnnotationStrokeWidth"]
        == py_config["composerAnnotationStrokeWidth"]
    )
    assert (
        js_config["composerAnnotationStrokeStyle"]
        == py_config["composerAnnotationStrokeStyle"]
    )
    assert (
        js_config["composerAnnotationFontSize"]
        == py_config["composerAnnotationFontSize"]
    )
    assert (
        js_config["composerAnnotationSpanSeconds"]
        == py_config["composerAnnotationSpanSeconds"]
    )
    assert (
        js_config["composerScrubMaxAudioSeconds"]
        == py_config["composerScrubMaxAudioSeconds"]
    )
    assert js_config["composerDoubleClickCuts"] == py_config["composerDoubleClickCuts"]
    assert js_config["crossReferences"] == py_config["crossReferences"]
    assert js_config["mediaContainerWarning"] == py_config["mediaContainerWarning"]
    # The Embed Subtitles dialog filters its target list against these, so JS
    # drifting from video.SUBTITLE_CODEC_BY_CONTAINER means promising output
    # ffmpeg will refuse to write (or hiding one it would have written).
    assert js_config["subtitleContainers"] == py_config["subtitleContainers"]
    # Profiling defaults off on both sides; live launches overlay --profile's
    # True via clipgenApplyConfig, and exports strip the key entirely.
    assert js_config["profiling"] is False
    assert py_config["profiling"] == config.PROFILING


def test_clipgen_apply_config_covers_frontend_config():
    """clipgenApplyConfig must have a branch for every get_frontend_config key.

    test_clipgen_config_defaults_match_python compares default *values*, so a
    key the server ships but the applier silently drops stays green while the
    live frontend runs the JS defaults forever (the six composerAnnotation*
    keys shipped un-applied this way). Coverage is asserted as payload.<key>
    access inside the function body.
    """
    match = re.search(
        r"var\s+clipgenApplyConfig\s*=\s*function\s*\(payload\)\s*\{(.*?)\n\};",
        _js_source(),
        re.DOTALL,
    )
    assert match, "clipgenApplyConfig not found in utils.js"
    handled = set(re.findall(r"payload\.(\w+)", match.group(1)))
    missing = set(utils.get_frontend_config().keys()) - handled
    assert not missing, (
        f"clipgenApplyConfig drops config keys the server ships: "
        f"{sorted(missing)}. Add a type-guarded branch for each in utils.js."
    )


def test_get_frontend_config_shape():
    """Contract: every consumer of get_frontend_config relies on these keys."""
    cfg = utils.get_frontend_config()
    assert set(cfg.keys()) == {
        "defaultDuration",
        "severity",
        "annotationKeyphrases",
        "annotations",
        "ignoredTimestampTokens",
        "screenspaceOcrMinConfidence",
        "screenspaceOcrFuzzyThreshold",
        "screenspaceMultitoolMaxOffset",
        "screenspaceMaskFallbackTools",
        "frictionCategories",
        "frictionColorToken",
        "frictionMomentLimit",
        "convergenceSources",
        "cardScrubberSpriteCols",
        "cardScrubberSpriteRows",
        "clipFormat",
        "screenshotFormat",
        "gifFormat",
        "composerAnnotationColor",
        "composerAnnotationColorSecondary",
        "composerAnnotationStrokeWidth",
        "composerAnnotationStrokeStyle",
        "composerAnnotationFontSize",
        "composerAnnotationSpanSeconds",
        "composerScrubMaxAudioSeconds",
        "composerDoubleClickCuts",
        "crossReferences",
        "mediaContainerWarning",
        "subtitleContainers",
        "hotkeyOverrides",
        "profiling",
    }
    assert isinstance(cfg["defaultDuration"], int)
    assert cfg["defaultDuration"] == config.DEFAULT_DURATION_SECONDS
    assert isinstance(cfg["severity"], list) and cfg["severity"]
    for entry in cfg["severity"]:
        assert set(entry.keys()) == {"label", "rank", "cssClass"}
        assert entry["cssClass"].startswith("sev-")
    assert sorted(cfg["annotationKeyphrases"]) == sorted(
        utils.get_known_annotation_map().keys()
    )
    assert isinstance(cfg["annotations"], list)
    annotation_map = utils.get_known_annotation_map()
    assert {(a["token"], a["id"]) for a in cfg["annotations"]} == set(
        annotation_map.items()
    )
    for entry in cfg["annotations"]:
        assert set(entry.keys()) == {"id", "token"}
    assert sorted(cfg["ignoredTimestampTokens"]) == sorted(
        utils.get_ignored_timestamp_tokens()
    )
    assert isinstance(cfg["screenspaceOcrMinConfidence"], float)
    assert cfg["screenspaceOcrMinConfidence"] == config.SCREENSPACE_OCR_MIN_CONFIDENCE
    assert isinstance(cfg["screenspaceOcrFuzzyThreshold"], float)
    assert cfg["screenspaceOcrFuzzyThreshold"] == config.SCREENSPACE_OCR_FUZZY_THRESHOLD
    assert isinstance(cfg["screenspaceMultitoolMaxOffset"], float)
    assert (
        cfg["screenspaceMultitoolMaxOffset"]
        == config.SCREENSPACE_MULTITOOL_MAX_OFFSET_SECONDS
    )
    assert cfg["screenspaceMaskFallbackTools"] == list(
        config.SCREENSPACE_MASK_FALLBACK_TOOLS
    )
    assert cfg["frictionColorToken"] == "--color-friction"
    assert cfg["frictionMomentLimit"] == config.FRICTION_MOMENT_LIMIT
    assert [c["key"] for c in cfg["frictionCategories"]] == list(
        config.FRICTION_CATEGORIES.keys()
    )
    for entry in cfg["frictionCategories"]:
        assert set(entry.keys()) == {"key", "label"}
    assert cfg["convergenceSources"] == list(config.CONVERGENCE_SOURCES)
    assert cfg["cardScrubberSpriteCols"] == config.STUDIO_SCRUBBER_SPRITE_COLS
    assert cfg["cardScrubberSpriteRows"] == config.STUDIO_SCRUBBER_SPRITE_ROWS
    assert cfg["clipFormat"] == config.FILEFORMAT
    assert cfg["screenshotFormat"] == config.SCREENSHOT_FORMAT
    assert cfg["gifFormat"] == config.GIF_FORMAT


def test_severity_css_class_mapping():
    """severity_css_class returns the CSS classes tokens.css defines."""
    expected = {
        "Critical": "sev-critical",
        "High": "sev-high",
        "Medium": "sev-medium",
        "Low": "sev-low",
        "N/A": "sev-na",
        "Positive": "sev-positive",
        "Very Positive": "sev-very-positive",
    }
    for label, css in expected.items():
        assert utils.severity_css_class(label) == css
    assert utils.severity_css_class("") == ""
    assert utils.severity_css_class("Bogus") == "sev-unknown"


def test_studio_api_payload_includes_config():
    """server.api_sheet must pass utils.get_frontend_config() to the JS layer."""
    src = (Path(__file__).resolve().parent.parent / "source" / "server.py").read_text(
        "utf-8"
    )
    assert '"config": utils.get_frontend_config()' in src, (
        "server.py /api/sheet response must embed `config: utils.get_frontend_config()` "
        "so JS reads canonical values instead of hardcoding them"
    )


def test_exported_viewer_payloads_include_config():
    """finalize_timeline_data / gallery viewer embed `config`.

    Exports go through _export_config(), which strips hotkeyOverrides so a
    standalone HTML file always runs the default keymap.
    """
    src = (Path(__file__).resolve().parent.parent / "source" / "viewer.py").read_text(
        "utf-8"
    )
    occurrences = src.count('"config": _export_config()')
    assert occurrences >= 2, (
        f"Expected at least 2 occurrences of config payload in viewer.py "
        f"(timeline/gallery), found {occurrences}"
    )
    import viewer

    assert "hotkeyOverrides" not in viewer._export_config()
    assert "profiling" not in viewer._export_config()


def _parse_js_string_array(name: str) -> list[str]:
    """Extract a top-level `var <name> = [ ... ];` of string literals."""
    match = re.search(
        r"var\s+" + re.escape(name) + r"\s*=\s*\[(.+?)\];",
        _js_source(),
        re.DOTALL,
    )
    assert match, f"{name} not found in utils.js"
    raw = match.group(1)
    return re.findall(r'"([^"]+)"', raw)


def test_detector_palette_stays_aligned():
    """Detector colours have one canonical home: `--color-task-*` in tokens.css.

    Three places carry detector palette knowledge today:
      1. `--color-task-{type}` tokens in `assets/web/tokens.css` (canonical)
      2. `_DETECTOR_TYPES` list in `assets/web/utils.js` (the JS catalogue
         of detectors that participates in the colour system)
      3. `_DETECTOR_FALLBACK` map in `assets/web/utils.js` (export-only
         safety net for HTML files that ship without tokens.css)

    Adding a detector — or renaming one — must touch all three. This test
    catches the most likely drift: keys diverging between the JS catalogue
    and the CSS palette. It does not compare hex equality (that would
    require an oklch-aware parser); the manual rule, captured in the
    comment block above `_DETECTOR_FALLBACK` in utils.js, is to mirror the
    dark-theme `--color-task-*` values into `_DETECTOR_FALLBACK`.
    """
    detector_types = _parse_js_string_array("_DETECTOR_TYPES")
    fallback = _parse_js_object_literal("_DETECTOR_FALLBACK")

    assert set(detector_types) == set(fallback.keys()), (
        f"_DETECTOR_TYPES and _DETECTOR_FALLBACK keys diverged: "
        f"types={sorted(detector_types)} vs fallback={sorted(fallback)}. "
        f"Update both lists in utils.js together."
    )

    tokens_css = (
        Path(__file__).resolve().parent.parent / "assets" / "web" / "tokens.css"
    ).read_text("utf-8")
    css_detector_keys = set(re.findall(r"--color-task-([\w-]+)\s*:", tokens_css))
    # tokens.css declares dark + light variants of every key, so the regex picks
    # them up twice — `set()` dedupes naturally.
    assert set(detector_types) <= css_detector_keys, (
        f"Detectors missing a `--color-task-*` token in tokens.css: "
        f"{sorted(set(detector_types) - css_detector_keys)}. "
        f"Add the token (dark + light blocks) to assets/web/tokens.css."
    )


# CSS mask classes that are not engine tools (workflow toolbar chrome).
_SS_TASK_ICON_NON_TOOLS = frozenset({"info", "play"})
# CLI --ss-task has no create path for these (rerun-only / UI-only).
_CLI_TASK_EXCLUDES = frozenset({"boundary", "multitool"})
# Own NODE_TYPES entries, not generated ss_* detect nodes.
_WORKFLOWS_SEPARATE_NODES = frozenset({"multitool", "timelapse"})
# In TOOLS but not yet a workflows node. A 14th tool must join specs, a
# separate NODE_TYPES entry, or this set — otherwise the equality below fails.
_WORKFLOWS_UNWIRED: frozenset[str] = frozenset()
# Timeline viewer has no per-event icon for timelapse (single output file).
_VIEWER_ICON_SKIP = frozenset({"timelapse"})


def _js_object_body(source: str, name: str) -> str:
    """Return the inside of `var <name> = { ... }` (brace-matched)."""
    match = re.search(r"var\s+" + re.escape(name) + r"\s*=\s*\{", source)
    assert match, f"{name} not found"
    start = match.end() - 1
    depth = 0
    for i, ch in enumerate(source[start:]):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return source[start + 1 : start + i]
    raise AssertionError(f"{name} object not closed")


def _flat_js_object_keys(source: str, name: str) -> set[str]:
    """Keys of a flat `var <name> = { a: ..., b: ... }` object."""
    return set(re.findall(r"(\w+)\s*:", _js_object_body(source, name)))


def test_detector_registries_stay_aligned():
    """Engine TOOLS is the source of truth; parallel catalogues must match it.

    test_detector_palette_stays_aligned already locks JS types ↔ fallback ↔
    tokens.css colours. This catches Python↔Python and JS↔JS name drift the
    palette test never sees (CODE-REVIEW.md "Parallel registries").
    """
    engine = set(screenspace_tools.TOOLS)
    js_types = set(_parse_js_string_array("_DETECTOR_TYPES"))
    assert engine == js_types, (
        f"screenspace_tools.TOOLS and utils.js _DETECTOR_TYPES diverged: "
        f"engine={sorted(engine)} vs js={sorted(js_types)}. "
        f"Add the tool to both (and the other registries this test lists)."
    )

    assert set(cli._SS_VALID_TASK_TYPES) == engine - _CLI_TASK_EXCLUDES, (
        f"cli._SS_VALID_TASK_TYPES drifted from TOOLS minus {_CLI_TASK_EXCLUDES}. "
        f"CLI={sorted(cli._SS_VALID_TASK_TYPES)} vs expected="
        f"{sorted(engine - _CLI_TASK_EXCLUDES)}."
    )

    specs = set(workflows_catalog._SS_DETECTOR_SPECS)
    assert specs == engine - _WORKFLOWS_SEPARATE_NODES - _WORKFLOWS_UNWIRED, (
        f"workflows_catalog._SS_DETECTOR_SPECS drifted from TOOLS. "
        f"specs={sorted(specs)} vs expected="
        f"{sorted(engine - _WORKFLOWS_SEPARATE_NODES - _WORKFLOWS_UNWIRED)}. "
        f"Separate nodes={sorted(_WORKFLOWS_SEPARATE_NODES)}; "
        f"unwired={sorted(_WORKFLOWS_UNWIRED)}."
    )
    assert _WORKFLOWS_SEPARATE_NODES <= set(workflows_catalog.NODE_TYPES), (
        f"Workflows NODE_TYPES missing {_WORKFLOWS_SEPARATE_NODES - set(workflows_catalog.NODE_TYPES)}. "
        f"multitool/timelapse live as their own nodes, not ss_* specs."
    )

    ss_js = (WEB / "screenspace.js").read_text(encoding="utf-8")
    icon_types = _flat_js_object_keys(ss_js, "SS_TASK_ICON_TYPES")
    icon_names = _flat_js_object_keys(ss_js, "TOOL_ICON_NAMES")
    assert icon_types == engine, (
        f"screenspace.js SS_TASK_ICON_TYPES drifted from TOOLS: "
        f"{sorted(icon_types)} vs {sorted(engine)}."
    )
    assert icon_names == engine, (
        f"screenspace.js TOOL_ICON_NAMES drifted from TOOLS: "
        f"{sorted(icon_names)} vs {sorted(engine)}."
    )

    ss_css = (WEB / "screenspace.css").read_text(encoding="utf-8")
    css_icons = set(re.findall(r"\.ss-task-icon--([\w-]+)", ss_css))
    assert css_icons - _SS_TASK_ICON_NON_TOOLS == engine, (
        f"screenspace.css .ss-task-icon--* drifted from TOOLS "
        f"(ignoring {_SS_TASK_ICON_NON_TOOLS}): "
        f"{sorted(css_icons - _SS_TASK_ICON_NON_TOOLS)} vs {sorted(engine)}."
    )

    viewer_js = (WEB / "viewer.js").read_text(encoding="utf-8")
    viewer_icons = set(
        re.findall(
            r"^\s+(\w+):\s*\{\s*viewBox:",
            _js_object_body(viewer_js, "SS_DETECTOR_ICON_PATHS"),
            re.MULTILINE,
        )
    )
    assert viewer_icons == engine - _VIEWER_ICON_SKIP, (
        f"viewer.js SS_DETECTOR_ICON_PATHS drifted from TOOLS minus "
        f"{_VIEWER_ICON_SKIP}: {sorted(viewer_icons)} vs "
        f"{sorted(engine - _VIEWER_ICON_SKIP)}. Timelapse skips this map "
        f"(single output file); every other tool needs an inline path."
    )
