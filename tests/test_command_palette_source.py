"""Static regression checks for the global command palette.

command-palette.js is a shared opt-in module (card-scrubber.js pattern): one
IIFE exposing window.ClipgenCommandPalette, loaded on all six hub pages after
topnav.js and before each page hub. These checks pin the contracts the
satellite-wiring test can't see (IIFE structure, script order, icon
references, CSS toggle completeness, the summon-chord shape).
"""

import re

from _frontend_source import WEB as _WEB
from _frontend_source import assert_es5, strip_comments

_ICONS = _WEB.parent / "icons"

PALETTE_JS = _WEB / "command-palette.js"
PALETTE_CSS = _WEB / "command-palette.css"

HUB_PAGES = {
    "studio.html": "studio.js",
    "screenspace.html": "screenspace.js",
    "transcripts.html": "transcripts.js",
    "workflows.html": "workflows.js",
    "composer.html": "composer.js",
    "overview.html": "overview.js",
}


def test_es5_discipline():
    """House style: no arrows / async-await in the palette module."""
    assert_es5(PALETTE_JS.read_text(encoding="utf-8"), "command-palette.js")


def test_iife_with_single_namespace():
    """Must stay an IIFE (keeps it out of the wiring test's ambient-global
    scan) exposing exactly the ClipgenCommandPalette namespace."""
    src = strip_comments(PALETTE_JS.read_text(encoding="utf-8")).strip()
    assert src.startswith("(function"), "command-palette.js must be an IIFE"
    assert "window.ClipgenCommandPalette =" in src


def test_wired_on_all_hub_pages():
    """Every hub page links the CSS and loads the JS after topnav.js and
    before the page hub (the palette must exist when hubs register)."""
    for page, hub_js in HUB_PAGES.items():
        html = (_WEB / page).read_text(encoding="utf-8")
        assert '<link rel="stylesheet" href="command-palette.css">' in html, (
            f"{page} does not link command-palette.css"
        )
        scripts = re.findall(r'<script src="([^"]+)"', html)
        assert "command-palette.js" in scripts, f"{page} does not load the palette"
        assert (
            scripts.index("topnav.js")
            < scripts.index("command-palette.js")
            < scripts.index(hub_js)
        ), f"{page} loads command-palette.js out of order"


def test_referenced_icons_exist():
    """Every icon stem referenced by a command (in the module and the six hub
    registration blocks) must resolve to a real Heroicon file; the palette
    must never inline SVG."""
    sources = [PALETTE_JS] + [_WEB / js for js in HUB_PAGES.values()]
    for path in sources:
        src = path.read_text(encoding="utf-8")
        for name in set(re.findall(r'icon: "([a-z0-9-]+)"', src)):
            assert (_ICONS / f"{name}.svg").exists(), (
                f"{path.name} references missing icon {name}.svg"
            )
    assert "<svg" not in PALETTE_JS.read_text(encoding="utf-8")


def test_nav_icons_cover_all_surfaces():
    """The built-in nav provider carries an icon for each of the six
    surfaces, and each maps to a real icon file."""
    src = PALETTE_JS.read_text(encoding="utf-8")
    block = src[src.index("var NAV_ICONS") :]
    block = block[: block.index("};")]
    icons = dict(re.findall(r'(\w+): "([a-z0-9-]+)"', block))
    assert sorted(icons) == [
        "composer",
        "overview",
        "screenspace",
        "studio",
        "transcripts",
        "workflows",
    ]
    for name in icons.values():
        assert (_ICONS / f"{name}.svg").exists(), f"missing nav icon {name}.svg"


def test_css_toggle_completeness_and_tokens():
    """The overlay's .hidden state must be self-contained (never rely on a
    page stylesheet), sit at --z-overlay, and use no raw z-index."""
    css = PALETTE_CSS.read_text(encoding="utf-8")
    assert re.search(
        r"^\.cmdp-overlay\.hidden \{\n  display: none !important;", css, re.MULTILINE
    )
    assert "z-index: var(--z-overlay);" in css
    assert not re.search(r"z-index: \d", css), "raw z-index in command-palette.css"


def test_summon_chord_via_hotkey_registry():
    """The chord is a hotkeys.js catalog action ("global.palette", default
    Mod+Shift+P / Mod+K — Firefox reserves Ctrl+Shift+P), registered with
    allowInInput so it fires while typing (Spotlight behavior). The registry
    suppresses combos while a blocking modal is open, so toggling closed is
    handled by the palette input's own keydown via resolvedCombos (which
    respects Settings → Hotkeys rebinds). No raw document keydown listener —
    that contract is also enforced by test_hotkeys_frontend_source.py."""
    hotkeys_src = (_WEB / "hotkeys.js").read_text(encoding="utf-8")
    entry = re.search(r'\{ id: "global\.palette".*\}', hotkeys_src)
    assert entry, "hotkeys.js catalog lacks the global.palette action"
    assert '"Mod+Shift+P"' in entry.group(0)
    assert '"Mod+K"' in entry.group(0)
    src = PALETTE_JS.read_text(encoding="utf-8")
    assert re.search(
        r'\{ id: "global\.palette", handler: toggle, allowInInput: true \}', src
    ), "palette does not register the chord through ClipgenHotkeys"
    assert 'resolvedCombos("global.palette")' in src  # toggle-while-open
    assert not re.search(r'document\.addEventListener\(\s*"keydown"', src), (
        "command-palette.js must not attach document-level keydown listeners"
    )


def test_command_ids_stay_out_of_hotkey_namespace():
    """Palette command ids use ":" separators; dotted "a.b" ids are reserved
    for hotkey catalog actions (test_hotkeys_frontend_source.py scans every
    id: \"a.b\" literal and requires it in HOTKEY_CATALOG). global.palette is
    the one deliberate exception — it IS a catalog action."""
    src = PALETTE_JS.read_text(encoding="utf-8")
    dotted = re.findall(r'id: "([a-z][a-zA-Z0-9]*(?:\.[a-zA-Z0-9]+)+)"', src)
    assert dotted == ["global.palette"], (
        f"unexpected dotted command ids in command-palette.js: {dotted}"
    )


def test_modal_guard():
    """open() refuses to steal an existing overlay's trap — both the settings
    modal (body.modal-open, no openBlockingModal) and the openBlockingModal
    holders (Studio gallery/status/confirm, Transcripts install dialog,
    hotkeys cheatsheet) that never set the class."""
    src = PALETTE_JS.read_text(encoding="utf-8")
    assert 'classList.contains("modal-open")' in src
    assert "isBlockingModalOpen()" in src
    utils_src = (_WEB / "utils.js").read_text(encoding="utf-8")
    assert "var isBlockingModalOpen = function () {" in utils_src


def test_cross_page_deep_links():
    """Cross-page tab commands emit /PAGE/#tab=KEY hashes; every NAV_TABS
    destination must actually consume them (clipgenHashTab), and every
    participant destination must consume #Pxx (clipgenHashParticipant)."""
    src = PALETTE_JS.read_text(encoding="utf-8")
    assert "var NAV_TABS" in src
    assert '"#tab=" + key' in src
    utils_src = (_WEB / "utils.js").read_text(encoding="utf-8")
    assert "var clipgenHashTab = function () {" in utils_src
    tabs_block = src[src.index("var NAV_TABS") : src.index("var PARTICIPANT_PAGES")]
    for dest in re.findall(r"^    (\w[\w-]*): \[", tabs_block, re.MULTILINE):
        page_src = (_WEB / f"{dest}.js").read_text(encoding="utf-8")
        assert "clipgenHashTab()" in page_src, f"{dest}.js ignores #tab= deep links"
    pages_block = src[src.index("var PARTICIPANT_PAGES") : src.index("var providers")]
    dests = re.findall(r'id: "(\w+)"', pages_block)
    assert sorted(dests) == ["composer", "screenspace", "transcripts"]
    for dest in dests:
        page_src = (_WEB / f"{dest}.js").read_text(encoding="utf-8")
        assert "clipgenHashParticipant()" in page_src, (
            f"{dest}.js ignores #Pxx deep links"
        )


def test_every_hub_feeds_participants():
    """Cross-page participant jumps exist on every page: each hub must hand
    its participant list to the palette via setParticipants."""
    src = PALETTE_JS.read_text(encoding="utf-8")
    assert "setParticipants: setParticipants," in src
    for hub_js in HUB_PAGES.values():
        hub_src = (_WEB / hub_js).read_text(encoding="utf-8")
        assert "window.ClipgenCommandPalette.setParticipants(" in hub_src, (
            f"{hub_js} does not feed participants to the palette"
        )


def test_stay_vs_leave_wording():
    """Same-page participant commands say "Jump to … in <Page>" (stays);
    cross-page ones say "Open … in <Page>" (navigates)."""
    src = PALETTE_JS.read_text(encoding="utf-8")
    assert '"Open " + pids[i] + " in " + dest.label' in src
    for hub_js in ("screenspace.js", "transcripts.js", "composer.js"):
        hub_src = (_WEB / hub_js).read_text(encoding="utf-8")
        assert re.search(r'"Jump to " \+ p\.id \+ " in \w+"', hub_src), (
            f"{hub_js} same-page jump lacks the 'in <Page>' suffix"
        )


def test_topnav_exposes_quick_actions_getter():
    """Auto-ingest depends on the additive ClipgenTopNav.getQuickActions."""
    src = (_WEB / "topnav.js").read_text(encoding="utf-8")
    assert "getQuickActions: getQuickActions," in src


def test_all_hubs_register():
    """Each hub contributes a page command set (Composer and Overview have no
    quick actions, so without this they'd only get the built-ins)."""
    for page, hub_js in HUB_PAGES.items():
        src = (_WEB / hub_js).read_text(encoding="utf-8")
        assert "window.ClipgenCommandPalette.register(" in src, (
            f"{hub_js} does not register palette commands"
        )
