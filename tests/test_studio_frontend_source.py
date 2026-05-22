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


def test_studio_generate_uses_abort_controllers_for_both_branches():
    """Generate sheet + intake fetches must be cancellable via AbortController
    so the network connections are torn down promptly on cancel."""
    src = _studio_js()
    assert "new AbortController()" in src
    assert "sheetAbort = new AbortController()" in src
    assert "intakeAbort = new AbortController()" in src
    assert "signal: sheetAbort.signal" in src
    assert "signal: intakeAbort.signal" in src


def test_on_cancel_generate_aborts_and_posts_intake_cancel():
    """Cancel button must abort live fetches AND post both server cancel
    endpoints so in-flight ffmpeg subprocesses get terminated."""
    src = _studio_js()
    start = src.index("function onCancelGenerate()")
    end = src.index("\n  }", start)
    body = src[start:end]
    assert "aborts[i].abort()" in body
    assert 'apiPost("api/generate/cancel")' in body
    assert 'apiPost("api/generate-intake/cancel")' in body


# ---- Task 3: studio streaming, selectors, finishBranch ----


def test_studio_defines_shared_ndjson_reader():
    """One shared NDJSON helper guards response.body and is reused by the
    sheet/intake/reel readers (the three sites used to duplicate the loop)."""
    src = _studio_js()
    assert "function readNDJSONStream(response, onLine)" in src
    # response.body guard
    assert "if (!response.body" in src
    # The duplicated raw .getReader() blocks should be gone (only the helper
    # itself calls it). Allow the helper to be the sole getReader site.
    assert src.count("response.body.getReader()") == 1
    # Used by all three streaming endpoints.
    assert "readNDJSONStream(response, handleLine).then(finishBranch)" in src
    assert "readNDJSONStream(response, handleIntakeLine).then(finishBranch)" in src
    assert "readNDJSONStream(response, handleLine).then(finish)" in src


def test_sheet_branch_catch_marks_failures():
    """A fetch failure on the sheet branch must mark captured sheet cards as
    failed (not leave them visually queued) and tally totalFail."""
    src = _studio_js()
    assert "var sheetCardEls = [];" in src
    assert "setCardResult(sheetCardEls[j], false)" in src
    assert "totalFail += sheetItems.length;" in src


def test_finish_branch_handles_zero_zero_case():
    """Stream that ends with no successes and no failures (not cancelled)
    should surface an error, not silently report '0 artifacts'."""
    src = _studio_js()
    assert "No artifacts were generated" in src


def test_clear_card_status_selects_card_gen_badge():
    """clearCardStatus must select .card-gen-badge (the class actually set
    by createResultBadge), not the stale .card-result-badge name."""
    src = _studio_js()
    assert ".card-result-badge" not in src
    start = src.index("function clearCardStatus")
    end = src.index("\n  }", start)
    body = src[start:end]
    assert ".card-gen-badge" in body


def test_cancel_cleanup_uses_queue_card_queued_selector():
    """Cancel cleanup must query .queue-card-queued (the actual class set
    by setCardQueued), not the broken compound .queue-card.queued."""
    src = _studio_js()
    assert ".queue-card.queued" not in src
    assert 'querySelectorAll(".queue-card-queued")' in src


def test_reel_409_treated_as_json_error():
    """409 from /api/reel-direct (reel busy) is a JSON error, not NDJSON —
    the previous status !== 409 exemption was a bug."""
    src = _studio_js()
    assert "response.status !== 409" not in src
