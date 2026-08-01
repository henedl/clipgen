"""Guard: static element IDs the Transcripts page script toggles must exist.

A missing element makes ``qs("#id")`` return ``null`` and the first
``.classList`` call throw, silently aborting an entire render/handler flow. That
is exactly what happened when ``#frictionGenerating`` was referenced by
transcripts.js but never added to transcripts.html: the Friction tab showed its
status header (rendered first) but never un-hid its content, and the Re-run
button threw before sending its request.

This is a cheap structural check — it does not exercise behaviour (the project
has no JS DOM harness), but it catches HTML/JS drift for the elements that the
render flow assumes exist.
"""

import re

from _frontend_source import WEB, concat_js, read

_ICONS = WEB.parent / "icons"
_CSS = read("transcripts.css")
_HTML = read("transcripts.html")
# The page script is a hub (transcripts.js) plus feature satellites
# (transcripts-{corrections,search,video,pills,agents}.js); the friction/summary
# element IDs live in the agents satellite, so read all of them together.
_JS = concat_js("transcripts")

# Static element IDs the page script references via qs("#...") that must be
# present in transcripts.html. Dynamically-created nodes are intentionally
# excluded; these are the ones the tabbed analysis panel / friction flow toggle.
REQUIRED_IDS = [
    "summarySection",
    "tabBtnSummary",
    "tabBtnFriction",
    "summaryTab",
    "frictionTab",
    "summaryActions",
    "summaryBody",
    "summaryContent",
    "summaryEmpty",
    "summaryRunCta",
    "frictionStatus",
    "frictionRerun",
    "frictionCancel",
    "frictionEmpty",
    "frictionEmptyHint",
    "frictionGenerating",
    "frictionContent",
    "frictionModeMount",
    "frictionCounter",
    "frictionMarkAll",
    "frictionHistogram",
    "frictionBounds",
    "frictionEvidence",
    "frictionJumpStrip",
    "frictionJumpPrev",
    "frictionJumpNext",
    "frictionStaleDot",
    "clipMarksModal",
    "clipMarksScope",
    "clipMarksGap",
    "clipMarksPad",
    "clipMarksSummary",
    "clipMarksProgress",
    "clipMarksBarFill",
    "clipMarksProgressText",
    "clipMarksCancel",
    "clipMarksConfirm",
]


def test_required_ids_present_in_html():
    missing = [i for i in REQUIRED_IDS if f'id="{i}"' not in _HTML]
    assert not missing, (
        "transcripts.js toggles these elements but they are absent from "
        f"transcripts.html: {missing}"
    )


def test_required_ids_referenced_in_js():
    # Keeps the guard honest: every ID we require must actually be used by the JS.
    unused = [i for i in REQUIRED_IDS if i not in _JS]
    assert not unused, f"REQUIRED_IDS not referenced in transcripts.js: {unused}"


def test_boot_placeholders_are_skeletons_not_empty_states():
    """Shipping "No participants" / "No transcript available" before the fetch
    resolves states something false for as long as the load takes."""
    assert "pill-row-empty" not in _HTML, (
        "#participantPills must ship skeletons, not the post-fetch empty state"
    )
    assert _HTML.count('class="skeleton pill-skeleton"') >= 3
    assert ".pill-skeleton" in _CSS, ".pill-skeleton is used in HTML but never styled"
    assert 'id="transcriptEmpty" class="transcript-empty hidden"' in _HTML, (
        "#transcriptEmpty must ship with the .hidden CLASS — the JS toggles the "
        "class, so the bare `hidden` attribute would never be removed"
    )


def test_segment_rebuild_keeps_the_reader_in_place():
    """The rebuild wipes #segmentList, but scroll lives on #trMain — restoring
    the wrong element is a silent no-op, and restoring without marking the write
    as programmatic pauses playhead auto-follow for three seconds."""
    start = _JS.index("function renderSegments(")
    body = _JS[start : _JS.index("\n  // Which participant", start)]
    assert 'qs("#trMain")' in body, "scroll lives on #trMain, not #segmentList"
    assert "_renderedSegmentsPid" in body, (
        "restore must be gated on the participant being unchanged"
    )
    assert body.index("ignoreNextScroll()") < body.index("scrollHost.scrollTop ="), (
        "mark the scroll as programmatic before writing it, or the auto-follow "
        "pause treats the restore as a reader scroll"
    )
    assert body.index("applyFrictionDecorations()") < body.index(
        "ignoreNextScroll()"
    ), (
        "friction Isolate hides rows and the moment callouts add height, so the "
        "decoration pass must run before the scroll restore — otherwise the "
        "browser clamps the restored offset against a stale scrollHeight"
    )


def test_friction_decorations_are_the_only_friction_writer_on_the_segment_list():
    """renderSegments builds one big HTML string; applyFrictionDecorations then
    toggles classes on the result. If the string ALSO emitted friction markup the
    two paths would drift, and a threshold drag (which only runs the decoration
    pass) would disagree with the next full rebuild."""
    start = _JS.index("function renderSegments(")
    body = _JS[start : _JS.index("\n  // Which participant", start)]
    assert "segment-friction" not in body and "seg-friction-alpha" not in body, (
        "friction markup must not be emitted by renderSegments' HTML string; "
        "applyFrictionDecorations owns every friction class and inline var"
    )


def test_isolate_mode_hides_rows_without_breaking_the_index_cache():
    """state.cachedSegmentRows is indexed positionally against state.segments, so
    Isolate must hide rows rather than remove them — and the hide rule needs the
    compound selector, since .segment-row's own display:flex is equal specificity
    and would otherwise win purely on source order."""
    assert ".segment-row.segment-hidden {" in _CSS, (
        "a bare .segment-hidden rule ties Isolate mode to CSS source order"
    )
    assert 'classList.toggle("segment-hidden"' in _JS, (
        "Isolate must toggle a class; removing rows misaligns cachedSegmentRows"
    )


def test_scroll_to_segment_bails_on_an_isolated_row():
    """A display:none row reports an all-zero getBoundingClientRect(), so the
    scroll math lands ~188px above the reader. With PiP forcing auto-follow, that
    fires on every playhead transition — a continuous upward yank."""
    start = _JS.index("function scrollToSegment(")
    body = _JS[start : _JS.index("\n  }", start)]
    assert 'classList.contains("segment-hidden")' in body, (
        "scrollToSegment must return before any layout read on a hidden row"
    )
    assert body.index("segment-hidden") < body.index("getBoundingClientRect"), (
        "the guard has to precede the rect read it exists to prevent"
    )


def test_friction_consumers_all_read_a_shared_derived_map():
    """The threshold/category filter must drive the pane, the segment tints AND
    the timeline band. Before the redesign the band and the tints used raw
    score > 0, so the transcript disagreed with the tab that filtered it.

    Two derived maps now, for two different questions: frictionMatchBySegId is
    the keyword score (it sets the tint alpha), frictionBandBySegId is the union
    of both evidence sources (what counts as flagged at all). Every consumer
    reads one of them; none re-derive from raw scores."""
    for name in ("transcripts.js", "transcripts-video.js", "transcripts-agents.js"):
        src = read(name)
        assert "frictionMatchBySegId" in src or "frictionBandBySegId" in src, (
            f"{name} must read a derived map, not raw friction scores"
        )
    assert "frictionHeatmapEnabled" not in _JS, (
        "the boolean heatmap flag was replaced by the three-way state.frictionMode"
    )


def test_a_collapsed_score_band_can_be_dragged_back_open():
    """The histogram has two handles clamped against each other. Picking the
    nearest handle is not enough: once min === max every press ties, ties resolve
    to min, and min can never exceed max — the band is shut for good. A press
    outside the band must grab the bound on that side so it widens toward the
    press."""
    start = _JS.index("function _initFrictionHistogramDrag(")
    body = _JS[start : _JS.index("\n  // ----", start)]
    assert 'grabbed = "max"' in body and 'grabbed = "min"' in body, (
        "the outside-the-band case must pick a side explicitly"
    )
    assert body.index("v > state.frictionMax") < body.index("Math.abs("), (
        "the outside-the-band check has to run BEFORE the nearest-handle tie-break"
    )


def test_histogram_drag_does_not_depend_on_pointer_capture():
    """setPointerCapture is a nicety some engines refuse — and the desktop bundle
    hosts this app in WKWebView. Gating the move/up handlers on it would make the
    drag silently dead there, so they gate on the grabbed handle and self-heal
    when a pointerup is missed."""
    start = _JS.index("function _initFrictionHistogramDrag(")
    body = _JS[start : _JS.index("\n  // ----", start)]
    assert "host.hasPointerCapture(" not in body, (
        "gate the drag on `grabbed`, not on whether capture was granted"
    )
    assert "e.buttons === 0" in body, (
        "a move with no button held means the pointerup was missed; end the drag"
    )


def test_every_friction_hover_explains_itself_from_one_builder():
    """A score on its own says nothing. All three friction hover surfaces — the
    histogram bins, the hot segment rows and the timeline density band — have to
    answer 'why is this flagged' with the same words, so they share one builder
    rather than each formatting scores and quotes their own way."""
    assert _JS.count("function _frictionWhyLine(") == 1
    assert _JS.count("function _frictionQuote(") == 1
    for caller in ("_frictionBinTooltip", "_showFrictionTooltip"):
        assert caller in _JS
    video = read("transcripts-video.js")
    assert "TS._showFrictionTooltip" in video, (
        "the timeline band must reuse the agents satellite's friction tooltip, "
        "not grow a second one that drifts from it"
    )
    assert "hitTestFrictionBand" in video


def test_transcripts_does_not_pull_in_the_studio_primitives():
    """primitives.{js,css} are Studio/Overview-only; the friction chips are
    deliberately page-local so this page keeps its current asset set."""
    assert "primitives.js" not in _HTML and "primitives.css" not in _HTML


def test_sheet_xref_leg_stops_polling_an_empty_studio():
    """/studio/api/sheet answers ok with no rows when no spreadsheet is open, so
    the 30s poll can never learn anything — but it must re-arm on tab focus, or a
    sheet opened from another tab never reaches this page."""
    start = _JS.index("function loadCrossRefData(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert "if (_sheetXrefIdle) return;" in body
    assert "data.sheet_loaded === false" in body
    arm = _JS.index("function startXrefPolling(")
    assert "_sheetXrefIdle = false" in _JS[arm : _JS.index("\n  }", arm)], (
        "startXrefPolling also runs on tab focus — re-arm the sheet leg there"
    )


def test_a_failed_boot_fetch_clears_the_placeholders():
    """Without this the pill row shimmers forever and the transcript pane stays
    blank, which reads as "still loading" rather than "server unreachable"."""
    start = _JS.index("function loadParticipants(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert body.count("_clearBootPlaceholders()") == 2, (
        "both the !data.ok branch and the .catch must fall back to the real "
        "empty states"
    )


def test_every_thinking_agent_run_button_carries_the_local_ai_badge():
    """The badge marks the control that *starts* an Ollama run, not the pane that
    shows its output — so it belongs on the run buttons and not on the tabs (which
    could never mark Citations, having no tab of its own)."""
    tabbar = _HTML[_HTML.index('class="panel-tabbar"') : _HTML.index('id="summaryTab"')]
    assert "ai-agent-badge" not in tabbar, (
        "the badge moved to the run buttons; a tab must not carry it again"
    )
    for btn_id in ("summaryRunCta", "frictionRerun", "frictionStaleRerun"):
        start = _HTML.index(f'id="{btn_id}"')
        button = _HTML[start : _HTML.index("</button>", start)]
        assert "ai-agent-badge" in button, f"#{btn_id} starts an agent but has no badge"
        assert "data-tooltip=" in button, f"#{btn_id}'s badge must explain itself"
        # The badge would not survive a textContent relabel of the button.
        assert "agent-run-label" in button, (
            f"#{btn_id}'s label must be its own span, beside the badge"
        )


def test_icon_only_agent_buttons_get_the_badge_via_the_tooltip_sidecar():
    """The 24px action-row buttons have no room for a second glyph, so the badge
    rides in on the tooltip's icon sidecar. The accessible name still says it in
    words — the sidecar is decorative and screen readers never see it."""
    for btn_id in ("summaryRegenerate", "citationsRegenerate"):
        start = _HTML.index(f'id="{btn_id}"')
        button = _HTML[start : _HTML.index("</button>", start)]
        assert 'data-tooltip-icon="octicon/dependabot-16"' in button, (
            f"#{btn_id} starts an agent but its tooltip carries no badge"
        )
        assert "local AI" in button, (
            f"#{btn_id}'s aria-label must name the local AI; the sidecar is decorative"
        )


def test_pill_dropdown_badges_the_thinking_agents_but_not_transcription():
    """Transcription is Whisper, not an Ollama thinking agent — badging its row
    would make the marker meaningless. All four rows share one builder, so the
    flag is the only thing keeping them apart."""
    pills = read("transcripts-pills.js")
    section = pills[pills.index("function buildPillAgentsSection(") :]
    section = section[: section.index("\n  function buildAgentRow(")]
    for agent in ("summary", "citations", "friction"):
        row = section[section.index(f'agent: "{agent}"') :]
        assert "aiBadge: true" in row[: row.index("onStart")], (
            f"the {agent} pill row is Ollama-backed but passes no aiBadge"
        )
    transcription = section[section.index('agent: "transcription"') :]
    assert "aiBadge" not in transcription[: transcription.index("onStart")], (
        "Transcription is Whisper, not a thinking agent — it must stay unbadged"
    )
    assert 'className = "ai-agent-badge"' in pills, (
        "buildAgentRow must actually build the badge node"
    )


def test_local_ai_badge_is_styled_and_its_icon_exists():
    """The badge is a pure CSS mask with no JS behind it, so an unstyled class or
    a mistyped icon path renders a zero-size invisible span with no error."""
    assert ".ai-agent-badge {" in _CSS, (
        ".ai-agent-badge is used in HTML but never styled"
    )
    # .btn is not a flex container, so without an explicit display the inline
    # span collapses to zero size inside every run button.
    assert ".btn > .ai-agent-badge {" in _CSS, (
        "the badge needs its in-button display rule or it renders invisible"
    )
    masks = re.findall(r"mask-image:\s*url\((?:'|\")?([^'\")]+)", _CSS)
    referenced = {m for m in masks if m.startswith("icons/")}
    assert "icons/octicon/dependabot-16.svg" in referenced
    missing = [m for m in referenced if not (_ICONS / m[len("icons/") :]).is_file()]
    assert not missing, f"transcripts.css masks nonexistent icons: {missing}"


def test_audio_track_override_survives_track_zero():
    """`<select>.value` for track 0 is the string "0" (truthy), so the falsy gate
    in startTranscribe is correct — but only by accident of that stringiness. A
    refactor to a numeric override would silently drop "transcribe track 1"."""
    start = _JS.index("function startTranscribe(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert "ov.audioTrack" in body, (
        "startTranscribe must forward the pill's audio-track override"
    )
    assert "audio_index" in body, "the POST key the server parses is audio_index"


def test_audio_track_row_is_labelled_and_styled():
    """The row is built in JS, so an unstyled wrapper class collapses the hint
    onto the select's line with no error anywhere."""
    assert '"Audio track"' in _JS, "the pill dropdown must label the track picker"
    for cls in (".pill-options-group", ".pill-options-hint"):
        assert cls + " {" in _CSS, f"{cls} is created in JS but never styled"


def test_pill_nav_cursor_survives_a_pane_rebuild_by_identity():
    """The options pane is rebuilt wholesale by the poll and can gain the
    audio-track row between two builds (its layout is fetched asynchronously, and
    a rebuild can race that fetch). The cursor therefore has to be carried across
    by control identity: index arithmetic cannot work, because replaceChild drops
    the old pane before anything could count what the index was measured against.
    Getting this wrong silently moves the highlight onto a different control."""
    start = _JS.index("function _refreshPillOptionsContent(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert "_pillNavCursorId()" in body and "_pillNavRestoreCursor(" in body, (
        "the pane rebuild must capture the cursor before the swap and restore it "
        "after, or a pane that gained a row repaints the cursor one control off"
    )
    # Every control pillNavControls() can land on needs a stable identity.
    assert _JS.count('setAttribute("data-nav-id"') >= 4, (
        "model/language/audio-track selects and the agent buttons must all carry "
        "data-nav-id for the cursor to be restorable"
    )
    assert "pillOptionsCursor +=" not in _JS, (
        "cursor index arithmetic is the bug this replaced — restore by nav id"
    )


# ---- Clip Marked Lines (Quick action -> Studio intake generation) ----


def test_clip_marks_loads_the_shared_clusterer_before_the_hub():
    """intake-cluster.js is not part of the transcripts-*.js glob, so a missing
    or late <script> is a runtime ReferenceError the concat-based scans above
    cannot see — the hub reads window.ClipgenIntakeCluster on the first open."""
    assert "intake-cluster.js" in _HTML, (
        "transcripts.html must load intake-cluster.js for the clip-marks action"
    )
    assert _HTML.index("intake-cluster.js") < _HTML.index('src="transcripts.js"'), (
        "intake-cluster.js must load before the page hub"
    )


def test_clip_marks_previews_and_cuts_from_one_clusterer():
    """The summary tells the user how many clips they will get. If the preview
    and the payload clustered differently, that number would be a lie — so both
    go through the shared helper, and neither hand-rolls a merge loop."""
    assert _JS.count("ClipgenIntakeCluster.clusterTranscriptMarks(") == 1, (
        "clustering belongs in one place (_clipMarksClusters); the summary and "
        "the payload both call it"
    )
    start = _JS.index("function submitClipMarks(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert "_clipMarksClusters()" in body


def test_clip_marks_sends_the_text_and_label_studio_drops():
    """server.py's _process_intake_item titles a transcript clip from label ->
    truncated text -> category. Studio's queue path omits both, so its clips are
    all named after the category; sending them is the whole reason these clips
    read as anything useful."""
    start = _JS.index("function submitClipMarks(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert '"../studio/api/generate-intake"' in body
    assert 'source: "transcript"' in body
    for key in ("text:", "label:", "mark_ids:"):
        assert key in body, f"the intake payload must carry {key}"


def test_clip_marks_pads_the_span_without_going_negative():
    """A mark's segment boundaries sit tight against the speech, so the cut is
    padded — but a mark near t=0 would otherwise ask ffmpeg for a negative
    start."""
    start = _JS.index("function submitClipMarks(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert "Math.max(0, c.start - pad)" in body, (
        "the padded start must be floored at zero"
    )


def test_clip_marks_ignores_the_trailing_cancelled_line():
    """/api/generate-intake closes a cancelled stream with {"cancelled": true},
    which has no index. Counting it as a completed item over-reports progress by
    one (the bug live in overview-reports.js)."""
    start = _JS.index("function submitClipMarks(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert 'typeof data.index !== "number"' in body, (
        "the NDJSON handler must bail on lines with no index"
    )


def test_clip_marks_modal_classes_are_all_styled():
    """Standing toggle-completeness check: an unstyled modal class ships a
    dialog that renders as unlaid-out text with no error anywhere."""
    for cls in re.findall(r'class="([^"]*clip-marks[^"]*)"', _HTML):
        for name in cls.split():
            if name.startswith("clip-marks"):
                assert "." + name in _CSS, f".{name} is used in HTML but never styled"


def test_clip_marks_run_outlives_its_dialog():
    """Escape/backdrop only close the modal; the batch keeps streaming and the
    quick action must reopen onto live progress. Storing the run inside the open
    handler (or clearing it on close) would strand a running job with no way to
    stop it."""
    assert "var _clipMarksRun = null;" in _JS
    close_start = _JS.index("function closeClipMarksModal(")
    close_body = _JS[close_start : _JS.index("\n  function ", close_start + 1)]
    assert "_clipMarksRun" not in close_body, (
        "closing the dialog must not touch the in-flight run"
    )
    open_start = _JS.index("function openClipMarksModal(")
    open_body = _JS[open_start : _JS.index("\n  function ", open_start + 1)]
    assert "if (_clipMarksRun) return;" in open_body, (
        "reopening mid-run must show progress, not re-fetch and reset the pickers"
    )


def test_friction_refetch_keeps_the_programmatic_scores():
    """The deterministic scorer's output does not depend on the LLM, but
    loadFriction used to blank it before every GET. Combined with the server
    answering `generating` for the whole agent run, a tab refocus / re-select /
    reload mid-run emptied the histogram, chips, tinting and timeline band until
    the agent finished. Blank only on a real participant change."""
    start = _JS.index("function loadFriction(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert "if (state.frictionPid !== pid) {" in body, (
        "the wipe must be gated on the participant actually changing"
    )
    assert body.index("frictionPid !== pid") < body.index(
        "state.frictionData = null"
    ), "the gate has to precede the wipe it exists to prevent"
    # Mid-run the server ships the scores alongside the generating flag.
    gen = body[body.index("data.generating") :]
    assert "if (data.friction) _setFrictionData(data.friction);" in gen, (
        "the generating branch must adopt the deterministic scores the server "
        "sends with it, or it renders the empty 'Analyzing…' box"
    )
    assert gen.index("_setFrictionData") < gen.index(
        "state.frictionGenerating = true"
    ), "_setFrictionData clears the generating flag, so it has to run first"


# ---- Friction evidence table (programmatic vs agentic) ----


def test_friction_counts_the_two_evidence_sources_apart():
    """The keyword scorer labels segments; the agent labels its own moments, and
    never reconciles them against the line it quotes. The old single chip row
    counted only the first, so a category could read 0 while its moment sat in
    the jump strip."""
    start = _JS.index("function _frictionEvidenceCounts(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert "frictionData.segments" in body and "frictionData.moments" in body, (
        "both sources must be counted"
    )
    assert "return { prog: prog, ai: ai };" in body, "counted apart, not summed"


def test_friction_cell_is_inert_only_on_total_absence():
    """The original bug: a chip with a 0 count was unclickable, so the score band
    could lock the user out of the very control that widens it — and a category
    with AI evidence but no keyword hit could never be toggled at all. Inertness
    must key on whether that kind of evidence exists AT ALL, not on the banded
    count the cell happens to display."""
    start = _JS.index("function renderFrictionEvidence(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert "var everyAny = totals[source] > 0;" in body
    assert 'cell.classList.toggle("is-empty", !everyAny);' in body, (
        "is-empty must come from the unbanded totals, never from the banded count"
    )
    rows = _JS.index("function _frictionEvidenceRows(")
    rows_body = _JS[rows : _JS.index("\n  function ", rows + 1)]
    assert "_frictionScoreInBand" not in rows_body, (
        "the row set must be band-independent, or rows appear and vanish under "
        "the pointer mid-drag"
    )


def test_friction_filters_are_per_source():
    """One shared filter dict meant hiding a category on the keyword side
    silently hid the agent's moments in that category too — including moments in
    categories whose chip was inert, which no control could then bring back."""
    assert "state.frictionMomentFilter" in _JS
    start = _JS.index("function _frictionMomentMatches(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert "state.frictionMomentFilter" in body, "moments read their own filter"
    assert "frictionCategoryFilter" not in body, (
        "the keyword filter must not gate moments"
    )


def test_unvalidated_model_categories_are_bucketed_not_dropped():
    """thinking_agents only lowercases and underscores the model's category, so
    it can emit anything. An unrecognized string must still land in a row the
    user can toggle, rather than falling through every filter."""
    start = _JS.index("function _frictionMomentCategory(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert "FRICTION_OTHER" in body and "_frictionCatKeys()" in body
    # The strip and callout still show what the model actually said.
    label = _JS.index("function _frictionCatLabel(")
    label_body = _JS[label : _JS.index("\n  function ", label + 1)]
    assert 'if (key === "other") return "Other";' in label_body
    assert "toUpperCase()" in label_body, (
        "an invented category should render its own wording, not be blanked"
    )


def test_mark_all_matching_covers_the_ai_only_segments():
    """frictionMatchBySegId holds keyword matches only, so segments the agent
    alone flagged (score 0, no regex category) were silently skipped — the very
    lines the jump strip exists to surface."""
    start = _JS.index("function _frictionMarkAll(")
    body = _JS[start : _JS.index("\n  function initFriction(", start)]
    assert "state.frictionVisibleMoments" in body, (
        "AI-cited segments must be marked too"
    )
    assert "claimed[segId]" in body, (
        "a segment both sources flag must be marked once, not twice"
    )


def test_friction_evidence_classes_are_all_styled():
    """An unstyled row class renders the table as a run-on line of numbers."""
    for cls in re.findall(r'"(friction-ev-[a-z-]+)"', _JS):
        assert "." + cls in _CSS, f".{cls} is created in JS but never styled"
    for state_cls in (
        ".friction-ev-cell.is-on",
        ".friction-ev-cell.is-muted",
        ".friction-ev-cell.is-empty",
    ):
        assert state_cls in _CSS, f"{state_cls} is toggled in JS but never styled"


def test_the_density_band_covers_both_evidence_sources():
    """An AI-cited line scores 0 with the keyword scorer, so a band reading the
    keyword map alone drew nothing for the very moments the jump strip is built
    around. The union is derived once, in the same producer as everything else,
    so the canvas can never disagree with the pane about what is flagged."""
    start = _JS.index("function _recomputeFrictionMatches(")
    body = _JS[start : _JS.index("\n  // The single entry point", start)]
    assert "state.frictionBandBySegId = band;" in body, (
        "the union belongs in the one derived-state producer"
    )
    assert "visible[i].moment.score" in body, (
        "an AI-only line has no keyword score; the band needs the moment's"
    )
    video = read("transcripts-video.js")
    assert "state.frictionBandBySegId" in video
    draw = video.index("function _drawFrictionBand(")
    draw_body = video[draw : video.index("\n  // The selected participant", draw)]
    assert "frictionMatchBySegId" not in draw_body, (
        "the band must draw from the union, not the keyword-only map"
    )
    hit = video.index("function hitTestFrictionBand(")
    hit_body = video[hit : video.index("\n  function ", hit + 1)]
    assert "frictionBandBySegId" in hit_body, (
        "hover has to key on the same map the band draws, or an AI-only stripe "
        "hovers as if it were not there"
    )


def test_flagged_by_either_source_is_derived_once():
    """'keyword match OR cited' was spelled out at three call sites — the band
    skipped the clause entirely, which is how it lost the AI moments."""
    assert (
        "frictionMatchBySegId[seg.id] !== undefined || state.frictionCitedBySegId"
        not in _JS
    ), "read the derived union map instead of re-deriving the OR"
    start = _JS.index("function _decorateSegmentList(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert "var flagged = !!seg && band[seg.id] !== undefined;" in body
    assert 'row.classList.toggle("segment-hidden", isolate && !flagged);' in body
    # ...but the tint alpha stays the keyword score, and the rail stays citation.
    assert 'row.style.setProperty("--seg-friction-alpha", score);' in body
    assert 'row.classList.toggle("segment-cited", on && isCited);' in body


def test_the_band_tooltip_names_which_score_it_is_showing():
    """The band now includes AI-cited lines, which score 0 with the keyword
    scorer — a bare 'score 0.00' on one of those stripes reads as a bug. The two
    numbers are on different scales, so each is labelled."""
    start = _JS.index("function _showFrictionTooltip(")
    body = _JS[start : _JS.index("\n  function ", start + 1)]
    assert "_visibleMomentsCiting(seg)" in body
    assert '"keyword "' in body and '"AI "' in body, (
        "label each score with its source rather than printing a bare number"
    )
    assert ".friction-cat-badge--ai" in _CSS, (
        "agent-labelled categories need a treatment distinct from the scorer's"
    )


def test_the_wip_badge_is_shared_not_forked():
    """The Friction tab is the second consumer of Overview's WIP pill, so the
    rule moved to tokens.css. A page-local copy is the drift this project keeps
    re-fixing: tokens.css loads first, so an equal-specificity fork silently
    wins on source order in whichever page defines it."""
    assert ".cg-wip-badge {" in read("tokens.css")
    assert 'class="cg-wip-badge"' in _HTML, "the Friction tab carries the badge"
    for page in ("transcripts.css", "overview.css"):
        assert ".cg-wip-badge {" not in read(page), (
            f"{page} must not redefine the shared badge"
        )
    assert "ov-wip-badge" not in read("overview.html"), (
        "Overview's copy was migrated onto the shared class"
    )
