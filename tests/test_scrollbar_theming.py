"""Regression checks for the global themed-scrollbar language.

Scrollbars used to be whatever the OS drew: chunky and light-grey on Windows,
and — because the six live pages scroll an inner container that starts at
viewport y=0 under the fixed translucent chrome — a bright stripe showing
through the navbar's blur. `tokens.css` now owns one global rule plus the
`--scrollbar-*` tokens it reads, and every per-page override was retired.

The point of these tests is drift, not rendering: the retired
`start-overlay.css` block was a raw-rgba light-theme patch that outranked the
shared utility it was patching, and that is exactly the shape that comes back
if page CSS is allowed to grow its own scrollbar rules again. Actual pixels are
verified with /ui-check, since headless Chromium's scrollbar mode is not the
user's.
"""

import re

from _frontend_source import WEB, read

TOKENS = "tokens.css"

# The one sanctioned exception: Screenspace's region-chip strip is deliberately
# barless. It wins on ID specificity and is documented in place.
_HIDER_OWNER = "screenspace.css"


def _theme_block(css: str, opener: str) -> str:
    """Slice one top-level rule block out of tokens.css by its opening line."""
    start = css.index(opener)
    return css[start : css.index("\n}", start)]


def test_color_scheme_declared_in_both_theme_blocks():
    css = read(TOKENS)
    assert "color-scheme: dark;" in _theme_block(css, ":root {")
    assert "color-scheme: light;" in _theme_block(css, 'html[data-theme="light"],')


def test_scrollbar_tokens_defined_in_both_themes():
    css = read(TOKENS)
    dark = _theme_block(css, ":root {")
    light = _theme_block(css, 'html[data-theme="light"],')
    # Colors are theme-specific; size/track carry over from :root unchanged.
    for token in ("--scrollbar-thumb:", "--scrollbar-thumb-hover:"):
        assert token in dark, f"{token} missing from :root"
        assert token in light, f"{token} missing from the light theme block"
    for token in ("--scrollbar-size:", "--scrollbar-track:"):
        assert token in dark, f"{token} missing from :root"


def test_scrollbar_language_lives_only_in_tokens_css():
    """No page CSS may grow its own scrollbar rules (the start-overlay bug)."""
    for path in sorted(WEB.glob("*.css")):
        if path.name == TOKENS:
            continue
        text = path.read_text(encoding="utf-8")
        assert "::-webkit-scrollbar-thumb" not in text, (
            f"{path.name} styles a scrollbar thumb; the shared rule and the "
            f"--scrollbar-* tokens in {TOKENS} are the only place for that"
        )
        assert "scrollbar-color" not in text, (
            f"{path.name} sets scrollbar-color; use the {TOKENS} tokens instead"
        )


def test_only_sanctioned_scrollbar_hider_outside_tokens():
    for path in sorted(WEB.glob("*.css")):
        if path.name == TOKENS:
            continue
        values = re.findall(r"scrollbar-width:\s*([a-z]+)", path.read_text("utf-8"))
        if not values:
            continue
        assert path.name == _HIDER_OWNER, (
            f"{path.name} sets scrollbar-width; only {_HIDER_OWNER}'s deliberate "
            "#regionChips hider may override the global rule"
        )
        assert set(values) == {"none"}, (
            f"{path.name} may only hide a scrollbar, not re-style one: {values}"
        )


def test_cg_scroll_thin_is_retired():
    """The opt-in utility is gone; every scroller is themed by default now."""
    for path in sorted(WEB.iterdir()):
        if path.suffix not in (".css", ".js", ".html"):
            continue
        assert "cg-scroll-thin" not in path.read_text(encoding="utf-8"), (
            f"{path.name} still references .cg-scroll-thin, which no longer exists"
        )
