"""Regression checks for the global custom-styled `<select>` language.

Native dropdowns used to render with the UA's menulist appearance. Worse, the
~20 page rules that painted them with a border, a background and a radius were
all silently discarded by WebKit, which ignores those on a menulist — so the
maintainer's Safari / WKWebView build never showed the styling that Chromium
screenshots did. `tokens.css` now owns one global `select` base plus a
deliberately out-specified caret rule, and the page rules were cut back to
geometry-only deltas.

These tests guard drift, not pixels: the failure mode is a page rule reclaiming
the caret at ID specificity, or a second `--select-caret` forked into page CSS
with a stale colour (the `ecfed991` shape). Rendering is verified with
/ui-check, and the caret's actual appearance on WebKit only by a human.
"""

import re

from _frontend_source import WEB, read

TOKENS = "tokens.css"

# A page that needs a caret-less select opts out with .cg-select-nocaret (as
# Studio's 3.5rem `ƒ` row-function column does) rather than re-declaring
# `appearance` and re-deriving the whole look.
_APPEARANCE_OWNERS = {TOKENS}


def _theme_block(css: str, opener: str) -> str:
    """Slice one top-level rule block out of tokens.css by its opening line."""
    start = css.index(opener)
    return css[start : css.index("\n}", start)]


def test_select_caret_defined_in_both_themes():
    """The caret colour is baked into the data URI, so each theme needs its own."""
    css = read(TOKENS)
    assert "--select-caret:" in _theme_block(css, ":root {")
    assert "--select-caret:" in _theme_block(css, 'html[data-theme="light"],')


def test_select_caret_not_forked_by_page_css():
    for path in sorted(WEB.glob("*.css")):
        if path.name == TOKENS:
            continue
        assert "--select-caret:" not in path.read_text(encoding="utf-8"), (
            f"{path.name} redefines --select-caret; the two theme values in "
            f"{TOKENS} are the only place for that"
        )


def test_select_caret_needs_no_network_route():
    """Exported viewers inline tokens.css but serve no /icons/ route."""
    caret_lines = [
        line for line in read(TOKENS).splitlines() if "--select-caret:" in line
    ]
    assert caret_lines, "--select-caret is gone from tokens.css"
    for line in caret_lines:
        assert "data:image/svg+xml" in line, (
            "the caret must stay a data URI: a <select> cannot carry a mask or a "
            "::after, and exported viewers cannot fetch icons/chevron-down.svg"
        )
        assert "icons/" not in line


def test_select_appearance_owned_by_tokens_css():
    """Only tokens.css may clear `appearance` on a select."""
    pattern = re.compile(r"([^{}]*\bselect\b[^{}]*)\{([^}]*)\}")
    for path in sorted(WEB.glob("*.css")):
        if path.name in _APPEARANCE_OWNERS:
            continue
        text = path.read_text(encoding="utf-8")
        for selector, body in pattern.findall(text):
            assert "appearance" not in body, (
                f"{path.name} sets appearance on `{selector.strip()}`; the global "
                f"select rule in {TOKENS} owns that"
            )
