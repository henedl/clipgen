"""Static regression checks for Studio frontend sources."""

from pathlib import Path

STUDIO_JS = Path(__file__).resolve().parent.parent / "assets" / "web" / "studio.js"
STUDIO_CSS = Path(__file__).resolve().parent.parent / "assets" / "web" / "studio.css"


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


def test_studio_symmetric_generation_job_state():
    """Artifact and reel jobs use independent locks, not one global generating flag."""
    src = _studio_js()
    assert "artifactGenerating: false" in src
    assert "reelGenerating: false" in src
    assert "overlayJobRunning: false" in src
    assert "function isArtifactQueueLocked()" in src
    assert "function isReelQueueLocked()" in src
    assert "function isAnyStudioJobRunning()" in src
    assert "function setArtifactGenerating(active)" in src
    assert "function setReelGenerating(active)" in src
    assert "function updateArtifactActions()" in src
    assert "function updateReelActions()" in src
    assert "setGeneratingLock" not in src
    assert "state.generating" not in src


def test_studio_css_does_not_freeze_drop_targets_during_generation():
    """Wheel scroll on reel/artifact strips must not be blocked by pointer-events: none."""
    css = STUDIO_CSS.read_text(encoding="utf-8")
    assert ".studio-generating .drop-target" not in css
    assert "studio-generating" not in css
