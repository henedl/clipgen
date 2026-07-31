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
    "frictionChips",
    "frictionJumpStrip",
    "frictionJumpPrev",
    "frictionJumpNext",
    "frictionStaleDot",
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


def test_friction_consumers_all_read_the_shared_match_map():
    """The threshold/category filter must drive the pane, the segment tints AND
    the timeline band. Before the redesign the band and the tints used raw
    score > 0, so the transcript disagreed with the tab that filtered it."""
    for name in ("transcripts.js", "transcripts-video.js", "transcripts-agents.js"):
        assert "frictionMatchBySegId" in read(name), (
            f"{name} must read the derived match map, not raw friction scores"
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


def test_every_thinking_agent_tab_carries_the_local_ai_badge():
    """Summary and Friction are the two tabs rendered by Ollama thinking agents,
    so both must show the badge that says so. A tab that grows an agent later
    (or an agent-backed tab added to another page) is the thing this catches."""
    tabbar = _HTML[_HTML.index('class="panel-tabbar"') : _HTML.index('id="summaryTab"')]
    assert tabbar.count('class="ai-agent-badge"') == 2, (
        "both #tabBtnSummary and #tabBtnFriction must carry .ai-agent-badge"
    )
    for tab_id in ("tabBtnSummary", "tabBtnFriction"):
        button = tabbar[
            tabbar.index(tab_id) : tabbar.index("</button>", tabbar.index(tab_id))
        ]
        assert "ai-agent-badge" in button, f"#{tab_id} is agent-backed but has no badge"
        assert "data-tooltip=" in button, f"#{tab_id}'s badge must explain itself"


def test_local_ai_badge_is_styled_and_its_icon_exists():
    """The badge is a pure CSS mask with no JS behind it, so an unstyled class or
    a mistyped icon path renders a zero-size invisible span with no error."""
    assert ".ai-agent-badge {" in _CSS, (
        ".ai-agent-badge is used in HTML but never styled"
    )
    masks = re.findall(r"mask-image:\s*url\((?:'|\")?([^'\")]+)", _CSS)
    referenced = {m for m in masks if m.startswith("icons/")}
    assert "icons/octicon/dependabot-16.svg" in referenced
    missing = [m for m in referenced if not (_ICONS / m[len("icons/") :]).is_file()]
    assert not missing, f"transcripts.css masks nonexistent icons: {missing}"
