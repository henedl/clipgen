"""Verify that JS fallback constants in utils.js match their config.py defaults.

The live values are repopulated at runtime from the server, but utils.js carries
hardcoded defaults so exported viewers and the brief boot window before the
first /api/marks fetch still render with the same colors.
"""

import json
import re
from pathlib import Path

import config
import utils

JS_PATH = Path(__file__).resolve().parent.parent / "assets" / "web" / "utils.js"


def _js_source() -> str:
    return JS_PATH.read_text(encoding="utf-8")


def _parse_js_object_literal(name: str) -> dict:
    """Extract a top-level JS object literal `var <name> = { ... };` as a dict.

    Handles unquoted keys and trailing commas; values must be JSON-compatible.
    """
    match = re.search(
        r"var\s+" + re.escape(name) + r"\s*=\s*(\{.+?\n\});",
        _js_source(),
        re.DOTALL,
    )
    assert match, f"{name} not found in utils.js"
    raw = match.group(1)
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

    assert len(js_config["severity"]) == len(py_config["severity"])
    for js_sev, py_sev in zip(js_config["severity"], py_config["severity"]):
        assert js_sev["label"] == py_sev["label"]
        assert js_sev["rank"] == py_sev["rank"]
        assert js_sev["cssClass"] == py_sev["cssClass"]


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
    src = (Path(__file__).resolve().parent.parent / "server.py").read_text("utf-8")
    assert '"config": utils.get_frontend_config()' in src, (
        "server.py /api/sheet response must embed `config: utils.get_frontend_config()` "
        "so JS reads canonical values instead of hardcoding them"
    )


def test_exported_viewer_payloads_include_config():
    """finalize_timeline_data / gallery viewer embed `config`."""
    src = (Path(__file__).resolve().parent.parent / "viewer.py").read_text("utf-8")
    occurrences = src.count('"config": utils.get_frontend_config()')
    assert occurrences >= 2, (
        f"Expected at least 2 occurrences of config payload in viewer.py "
        f"(timeline/gallery), found {occurrences}"
    )


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
