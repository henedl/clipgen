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

from _frontend_source import concat_js, read

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
