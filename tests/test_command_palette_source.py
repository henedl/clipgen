"""Static regression checks for the global command palette.

command-palette.js is a shared opt-in module (card-scrubber.js pattern): one
IIFE exposing window.ClipgenCommandPalette, loaded on all six hub pages after
topnav.js and before each page hub. These checks pin the contracts the
satellite-wiring test can't see (IIFE structure, script order, icon
references, CSS toggle completeness, the summon-chord shape).
"""

import re
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
_WEB = _ROOT / "assets" / "web"
_ICONS = _ROOT / "assets" / "icons"

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


def _strip_comments(src: str) -> str:
    src = re.sub(r"/\*.*?\*/", "", src, flags=re.S)
    return re.sub(r"^\s*//.*$", "", src, flags=re.M)


def test_es5_discipline():
    """House style: no arrows / async-await in the palette module."""
    src = PALETTE_JS.read_text(encoding="utf-8")
    assert "=>" not in src, "command-palette.js uses an arrow function"
    assert not re.search(r"\basync function\b|\bawait\s", src), (
        "command-palette.js uses async/await"
    )


def test_iife_with_single_namespace():
    """Must stay an IIFE (keeps it out of the wiring test's ambient-global
    scan) exposing exactly the ClipgenCommandPalette namespace."""
    src = _strip_comments(PALETTE_JS.read_text(encoding="utf-8")).strip()
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
        r"^\.cmdp-overlay\.hidden \{\n  display: none !important;", css, re.M
    )
    assert "z-index: var(--z-overlay);" in css
    assert not re.search(r"z-index: \d", css), "raw z-index in command-palette.css"


def test_summon_chord_and_modal_guard():
    """The chord listener is capture-phase, handles both Cmd/Ctrl+Shift+P and
    Cmd/Ctrl+K (Firefox reserves Ctrl+Shift+P), and open() refuses to steal
    an existing overlay's trap — both the settings modal (body.modal-open,
    no openBlockingModal) and the openBlockingModal holders (Studio
    gallery/status/confirm, Transcripts install dialog) that never set the
    class."""
    src = PALETTE_JS.read_text(encoding="utf-8")
    assert re.search(r"addEventListener\(\s*\"keydown\",[\s\S]*?\}, true\)", src), (
        "summon chord listener is not capture-phase"
    )
    assert 'k === "p"' in src
    assert 'k === "k"' in src
    assert 'classList.contains("modal-open")' in src
    assert "hasBlockingModal()" in src
    utils_src = (_WEB / "utils.js").read_text(encoding="utf-8")
    assert "var hasBlockingModal = function () {" in utils_src


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
