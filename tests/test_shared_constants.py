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
    assert js_config["ignoredTimestampTokens"] == py_config["ignoredTimestampTokens"]

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
        "ignoredTimestampTokens",
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
    assert sorted(cfg["ignoredTimestampTokens"]) == sorted(
        utils.get_ignored_timestamp_tokens()
    )


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


def test_insights_api_payload_includes_config():
    src = (Path(__file__).resolve().parent.parent / "insights_server.py").read_text(
        "utf-8"
    )
    assert '"config": utils.get_frontend_config()' in src, (
        "insights_server.py /api/artifacts response must embed `config`"
    )


def test_exported_viewer_payloads_include_config():
    """finalize_timeline_data / gallery / insights viewer all embed `config`."""
    src = (Path(__file__).resolve().parent.parent / "viewer.py").read_text("utf-8")
    occurrences = src.count('"config": utils.get_frontend_config()')
    assert occurrences >= 3, (
        f"Expected at least 3 occurrences of config payload in viewer.py "
        f"(timeline/gallery/insights viewer), found {occurrences}"
    )
