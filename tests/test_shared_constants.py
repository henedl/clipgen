"""Verify that JS constants in utils.js stay in sync with their Python sources."""

import json
import re
from pathlib import Path

import transcripts


def _parse_js_mark_categories() -> dict:
    """Extract MARK_CATEGORIES from assets/web/utils.js as a Python dict."""
    js_path = Path(__file__).resolve().parent.parent / "assets" / "web" / "utils.js"
    text = js_path.read_text(encoding="utf-8")
    # Grab the object literal assigned to MARK_CATEGORIES
    match = re.search(
        r"var\s+MARK_CATEGORIES\s*=\s*\{(.+?)\};",
        text,
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
    py_cats = transcripts.MARK_CATEGORIES
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
