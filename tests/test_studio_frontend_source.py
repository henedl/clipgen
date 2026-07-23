"""Static regression checks for Studio frontend sources."""

from pathlib import Path

_WEB = Path(__file__).resolve().parent.parent / "assets" / "web"
STUDIO_CSS = _WEB / "studio.css"
STUDIO_HTML = _WEB / "studio.html"


def _studio_js() -> str:
    # Studio is a hub (studio.js) + feature satellites (studio-*.js). Concatenate
    # all of them so assertions stay valid wherever a function lives. Satellites
    # sort before the hub ("-" < "."), so each src.index(a)..src.index(b) slice
    # below still resolves within the single file that owns both anchors.
    return "".join(
        p.read_text(encoding="utf-8") for p in sorted(_WEB.glob("studio*.js"))
    )


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


def test_studio_uses_shared_ndjson_reader():
    """The shared NDJSON helper lives in utils.js (also used by Composer) and is
    reused by the sheet/intake/reel readers (the three sites used to duplicate
    the loop); Studio must not grow a local copy back."""
    utils_src = (_WEB / "utils.js").read_text(encoding="utf-8")
    assert "var readNDJSONStream = function (response, onLine)" in utils_src
    # response.body guard
    assert "if (!response.body" in utils_src
    src = _studio_js()
    # The duplicated raw .getReader() blocks should be gone (only the utils.js
    # helper calls it).
    assert "response.body.getReader()" not in src
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


def test_generate_abort_treated_as_cancel_not_failure():
    """Aborting the generate fetch (Cancel button) must set cancelled and must
    not mark every card failed or increment totalFail in the catch path."""
    src = _studio_js()
    assert "state.generateCancelledByUser = true" in src
    assert "function isGenerateFetchAborted(err)" in src
    assert 'err.name === "AbortError"' in src
    start = src.index("function onGenerate()")
    end = src.index("function onCancelReel()", start)
    body = src[start:end]
    assert "if (isGenerateFetchAborted(err))" in body
    assert "cancelled = true" in body
    # Abort path clears queued cards; real failures still use setCardResult(..., false).
    abort_blocks = body.split("if (isGenerateFetchAborted(err))")
    assert len(abort_blocks) >= 3
    for block in abort_blocks[1:]:
        abort_section = block.split("finishBranch();")[0]
        assert "setCardResult" not in abort_section
        assert "totalFail +=" not in abort_section


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


def test_load_manifest_state_hydrates_reels_without_artifacts():
    """Reel-only manifests must still populate generatedReels and renderLog."""
    src = _studio_js()
    start = src.index("function loadManifestState()")
    end = src.index("function applyJobStatus(", start)
    body = src[start:end]
    assert "var reels = data.reels || [];" in body
    assert "artifacts.length === 0 && reels.length === 0" in body
    assert "state.generatedReels.push(stampLog(reel))" in body
    assert "renderLog();" in body
    assert "if (artifacts.length === 0) return;" not in body


def test_job_status_poll_includes_intake():
    """Re-attach polling must keep running while intake generation is active, and
    sheet + intake (one Generate action) share a single combined progress state
    rather than two branches that clobber each other's button/elapsed/idle."""
    src = _studio_js()
    assert "var intake = status.intake || {};" in src
    assert "data.intake && data.intake.in_progress" in src
    assert "var genActive = !!gen.in_progress || !!intake.in_progress;" in src


def test_reel_409_treated_as_json_error():
    """409 from /api/reel-direct (reel busy) is a JSON error, not NDJSON —
    the previous status !== 409 exemption was a bug."""
    src = _studio_js()
    assert "response.status !== 409" not in src


def test_add_to_queue_handles_intake_sources():
    """Intake items in the artifact queue must reach the reel via addToQueue(),
    not expandCellToSegments() which requires spreadsheet row/timestamp shape."""
    src = _studio_js()
    start = src.index("function addToQueue(")
    end = src.index("\n  // Collect selectable timestamp cell infos", start)
    body = src[start:end]
    intake_idx = body.index("if (isIntakeSource(info.source))")
    expand_idx = body.index("expandCellToSegments")
    assert intake_idx >= 0
    assert intake_idx < expand_idx
    assert "if (info.row) updateSingleCellClass" in body


def test_intake_drop_targets_route_through_add_to_queue():
    """Intake drag/drop must use addToQueue(), not duplicate push+render blocks."""
    src = _studio_js()
    start = src.index("function initDropTargets()")
    end = src.index("\n  function setupDropTarget(", start)
    body = src[start:end]
    assert "state.artifactQueue.push(info)" not in body
    assert "state.reelQueue.push(info)" not in body
    assert body.count("if (isIntakeSource(info.source))") == 2
    assert body.count("addToQueue(state.artifactQueue, info, renderArtifactQueue)") >= 1
    assert body.count("addToQueue(state.reelQueue, info, renderReelQueue)") >= 2


def test_studio_card_scrubber_wiring():
    """Opt-in card scrubber: state flag, attach hook, settings re-read, assets."""
    src = _studio_js()
    assert "cardScrubberEnabled: false" in src
    assert "function attachQueueScrubbers(listEl)" in src
    assert "window.clipgenCardScrubber.attach(thumb" in src
    assert "window.clipgenCardScrubber.detachStale()" in src
    assert "CLIPGEN_CONFIG.cardScrubberSpriteCols" in src
    assert '_findSetting("STUDIO_CARD_SCRUBBER")' in src

    html = STUDIO_HTML.read_text(encoding="utf-8")
    assert '<link rel="stylesheet" href="card-scrubber.css">' in html
    assert '<script src="card-scrubber.js" defer></script>' in html


def test_studio_card_scrubber_gates_on_thumbnail_and_prefetches():
    """Scrubber only wires cards with a real thumbnail frame (invalid-timestamp
    cards 404), and warms sprite sheets in a throttled background queue."""
    src = _studio_js()
    # Gate: activate only when the thumbnail <img> actually loaded a frame.
    assert "function wireCardScrubber(thumb, cols, rows, frameCount)" in src
    assert "img.complete && img.naturalWidth > 0" in src
    assert "if (img.naturalWidth > 0) activate();" in src
    # Eager, throttled prefetch of sprite sheets.
    assert "function enqueueSpritePrefetch(thumb)" in src
    assert "function processSpritePrefetch()" in src
    assert "function loadCardSprite(thumb, done)" in src
    assert "SPRITE_PREFETCH_CONCURRENCY" in src
