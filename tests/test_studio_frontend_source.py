"""Static regression checks for Studio sheet selection in assets/web/studio.js."""

from pathlib import Path

STUDIO_JS = Path(__file__).resolve().parent.parent / "assets" / "web" / "studio.js"


def _studio_js() -> str:
    return STUDIO_JS.read_text(encoding="utf-8")


def test_studio_selection_requires_valid_timestamp_cells():
    src = _studio_js()
    assert "function isSelectableTimestampCell(td)" in src
    assert 'td.classList.contains("valid-ts")' in src
    assert "if (!isSelectableTimestampCell(td)) continue;" in src
    assert "if (!isSelectableTimestampCell(td)) return;" in src


def test_studio_parse_clip_timestamps_no_zero_fallback():
    src = _studio_js()
    start = src.index("function parseClipTimestamps")
    end = src.index("  // Cross-referencing:", start)
    body = src[start:end]
    assert "segments.push({ startSeconds: 0" not in body
    assert "return parseClipSegmentsForCell" in body
