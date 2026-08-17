"""Source-level wiring assertions for the Start overlay's right column.

The two Refresh buttons and their failure handling are invisible to the Python
tests: the markup, the ``?refresh=true`` query the Google one sends (the only
user-facing escape hatch from ``server._cached_spreadsheet_meta``'s 5-minute
TTL), the fact that both buttons live *outside* the status node that several
code paths rewrite, and the recovery each loader owes its panel when a fetch
fails mid-load.

The column's three top-level tabs (Open / About / Recent updates) are the same
kind of blind spot: which panel a block of markup landed in is pure DOM order,
and the classes JS toggles need matching CSS to do anything at all.
"""

from __future__ import annotations

import re

from _frontend_source import assert_es5, read, strip_comments


def _button(html: str, role: str) -> str:
    match = re.search(rf"<button[^>]*data-role=\"{role}\"[^>]*>", html)
    assert match, f"no <button data-role={role!r}> in start-overlay.html"
    return match.group(0)


def test_right_column_is_tabbed():
    html = read("start-overlay.html")
    assert re.findall(r'data-start-tab="(\w+)"', html) == ["open", "about", "updates"]
    assert re.findall(r'data-start-panel="(\w+)"', html) == ["open", "about", "updates"]
    # Only the default tab's panel is visible on mount.
    for panel, hidden in (("open", False), ("about", True), ("updates", True)):
        tag = re.search(rf'<section[^>]*data-start-panel="{panel}"[^>]*>', html)
        assert tag, f"no panel for the {panel} tab"
        assert ("hidden" in tag.group(0)) is hidden


def test_confirm_footer_lives_in_the_open_panel():
    """The action button is per-tab: it may not float over About/Updates."""
    html = read("start-overlay.html")
    footer = html.index('class="start-footer')
    assert html.index('data-start-panel="open"') < footer
    assert footer < html.index('data-start-panel="about"')


def test_about_tab_carries_the_tool_tiles_then_the_about_rows():
    html = read("start-overlay.html")
    about = html.index('data-start-panel="about"')
    updates = html.index('data-start-panel="updates"')
    tiles = html.index('class="tool-tiles"')
    rows = html.index('data-role="about-grid"')
    # Attribution closes the tab, below clipgen's own license.
    attribution = html.index('data-role="attribution-list"')
    assert about < tiles < rows < attribution < updates
    assert html.index('data-role="changelog-list"') > updates


def test_attribution_link_opens_externally():
    """desktop.py's shim keys on target="_blank" to route clicks through
    open_external; without it a native-window user navigates the app to GitHub
    with no back button."""
    html = read("start-overlay.html")
    link = html[html.index('class="about__link attribution__more"') :][:400]
    assert 'target="_blank"' in link
    assert 'rel="noopener noreferrer"' in link
    assert "THIRD-PARTY-LICENSES" in link


def test_attribution_classes_built_by_js_have_css():
    """renderAttribution() builds these; without CSS the list renders unstyled."""
    css = read("start-overlay.css")
    for rule in (
        ".attribution__intro",
        ".attribution__group",
        ".attribution__row",
        ".attribution__row--nested",
        ".attribution__name",
        ".attribution__version",
        ".attribution__license",
        ".attribution__more",
    ):
        assert rule in css, f"{rule} built by start-overlay.js but absent from CSS"


def test_tab_classes_toggled_by_js_have_css():
    """.is-active / .is-entering / [hidden] do nothing without these rules."""
    css = read("start-overlay.css")
    for rule in (
        ".start-tabs",
        ".start-tab.is-active",
        ".start-tab__badge",
        ".start-tabpanel[hidden]",
        ".start-tabpanel__scroll",
        ".start-tabpanel.is-entering",
    ):
        assert rule in css, f"{rule} toggled/used in the overlay but absent from CSS"
    # The panel scroller owns scrolling now; a nested 180px window would clip
    # the changelog and about rows inside it.
    body = css[css.index(".start-overlay .changelog {") :][:400]
    assert "max-height" not in body


def test_overlay_always_opens_on_the_open_tab():
    src = strip_comments(read("start-overlay.js"))
    body = src[src.index("function open()") :]
    body = body[: body.index("\n  function ")]
    assert 'setStartTab("open")' in body, (
        "a launcher left on About would open with no way to confirm"
    )


def _hotkey_register_block() -> str:
    src = strip_comments(read("start-overlay.js"))
    start = src.index("ClipgenHotkeys.register([")
    return src[start : src.index("]);", start)]


def _hotkey_entry(block: str, hid: str) -> str:
    marker = f'id: "{hid}"'
    i = block.index(marker)
    nxt = block.find("{ id:", i + 1)
    return block[i : nxt if nxt != -1 else len(block)]


def test_form_hotkeys_require_the_open_tab():
    """G/E/I/Cmd+Enter must not fire while About or Recent updates is showing."""
    block = _hotkey_register_block()
    for hid in (
        "start.tabGoogle",
        "start.tabExcel",
        "start.tabMindnode",
        "start.tabNone",
        "start.browseInput",
        "start.browseOutput",
        "start.confirm",
    ):
        assert "isOpenForm" in _hotkey_entry(block, hid), hid
    for hid in ("start.tabOpen", "start.tabAbout", "start.tabUpdates"):
        entry = _hotkey_entry(block, hid)
        assert "when: isOpen," in entry, hid
        assert "isOpenForm" not in entry, hid


def test_both_panels_ship_a_refresh_button():
    html = read("start-overlay.html")
    for role in ("google-refresh", "excel-refresh"):
        button = _button(html, role)
        assert 'class="sheet-panel__refresh"' in button
        assert 'data-icon="arrow-path"' in html
    # Google ships hidden (the unauthenticated panel shows a Connect CTA
    # instead); Excel has no auth gate, so it is visible from the start.
    assert "hidden" in _button(html, "google-refresh")
    assert "hidden" not in _button(html, "excel-refresh")


def test_refresh_buttons_are_not_inside_the_status_node():
    """loadGoogleSheets / renderGoogleConnectCTA rewrite the status node."""
    html = read("start-overlay.html")
    for status, refresh in (
        ("google-status", "google-refresh"),
        ("excel-status", "excel-refresh"),
    ):
        opens = html.index(f'data-role="{status}"')
        # The status <div> is self-contained: it closes before the button opens.
        assert html.index("</div>", opens) < html.index(f'data-role="{refresh}"')


def test_refresh_always_clears_its_spinner():
    js = read("start-overlay.js")
    assert_es5(js, "start-overlay.js")
    src = strip_comments(js)
    assert '"?refresh=true"' in src
    assert "loadGoogleSheets(true)" in src
    # The teardown is a final .then() after .catch(), so a failed refresh can't
    # leave the button spinning and disabled forever.
    helper = src[src.index("function runPanelRefresh") :][:500]
    assert helper.index(".catch(") < helper.index('classList.remove("is-spinning")')


def test_failed_listings_leave_the_panels_usable():
    """A rejected fetch must not strand a panel in its loading state."""
    src = strip_comments(read("start-overlay.js"))
    google = src[
        src.index("function loadGoogleSheets") : src.index(
            "function keepPreviousGoogleList"
        )
    ]
    assert ".catch(" in google, "loadGoogleSheets hides the picker with no recovery"
    # Both the transport failure and the server-reported auth_error land on the
    # same recovery, which re-reveals the last list rather than leaving the
    # picker hidden with sheets still in state.
    assert google.count("keepPreviousGoogleList(") == 2

    excel = src[
        src.index("function loadExcelFiles") : src.index("function renderExcelList")
    ]
    assert ".catch(" in excel, "loadExcelFiles leaves the status on 'Scanning…'"


def test_refresh_button_has_its_css():
    css = read("start-overlay.css")
    for rule in (".sheet-panel__status-row", ".sheet-panel__refresh", "is-spinning"):
        assert rule in css, f"{rule} toggled/used in the overlay but absent from CSS"


def test_session_prefill_follows_active_source_not_mindnode_loaded():
    """A mind map and a spreadsheet coexist by design — _open_mindnode never
    clears the sheet and api_spreadsheets_open never clears the map. Branching
    on mindnode_loaded therefore meant that once a map had been opened it
    hijacked the prefill forever: the recents list highlighted the spreadsheet
    opened afterwards (currentSessionKey keys on active_source) while the panel
    switched to the Mind map tab, and confirming re-opened the map."""
    src = strip_comments(read("start-overlay.js"))
    body = src[src.index("function applyCurrentSessionPrefill") :]
    body = body[: body.index("\n  function ")]
    assert "active_source" in body, (
        "the mindnode branch must agree with currentSessionKey's active_source"
    )
    assert body.index("active_source") < body.index("mindnode_loaded"), (
        "active_source has to gate the mindnode branch, not follow it"
    )


def test_no_spreadsheet_actually_closes_the_open_source():
    """The 'No spreadsheet' tab only recorded a session. Nothing in the UI ever
    posted to /api/spreadsheets/close, so a source opened earlier stayed loaded
    for the rest of the process — a mind map especially, since opening a sheet
    does not clear one."""
    src = strip_comments(read("start-overlay.js"))
    assert "/api/spreadsheets/close" in src, (
        "the overlay must be able to close what it opened"
    )
    body = src[src.index("var skipSpreadsheet") :][:1600]
    assert "/api/spreadsheets/close" in body, (
        "the close belongs on the skipSpreadsheet path"
    )
    # One call per source: the route drops the map for {type: "mindnode"} and
    # the worksheet otherwise, and both can be open at once.
    assert "mindnode_loaded" in body and "sheet_loaded" in body, (
        "both sources have to be closed, not just whichever is active"
    )


def test_close_failures_do_not_record_a_session_or_reload():
    """/api/spreadsheets/close answers 409 while a generation is running. A
    catch-only handler treats that as success: it records a no-spreadsheet
    session and reloads with the source still loaded. The open path right
    below has always checked r.ok; this one must too."""
    src = strip_comments(read("start-overlay.js"))
    body = src[src.index("var skipSpreadsheet") :][:2200]
    assert "r.ok" in body, "the close response status must be checked"
    assert "markSheetError(" in body, (
        "a refused close has to surface, not fall through to the reload"
    )
    # The reload and the session record both sit behind the ok branch.
    ok_gate = body.index("if (!res.ok)")
    assert ok_gate < body.index("recordSession("), (
        "recordSession must not run when the source is still loaded"
    )
    assert ok_gate < body.index("window.location.reload()"), (
        "the reload must not run when the close was refused"
    )
