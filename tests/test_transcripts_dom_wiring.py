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
    "frictionStats",
    "frictionThreshold",
    "frictionThresholdVal",
    "frictionCategoryToggles",
    "frictionMarkAll",
    "frictionMoments",
    "frictionStaleDot",
    "frictionHeatmapBtn",
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
