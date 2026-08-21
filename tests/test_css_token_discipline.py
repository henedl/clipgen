"""Ratchet raw design values in page CSS toward tokens.

`assets/web/tokens.css` owns shared design values. Page CSS may not
introduce new raw `px`/`rem` spacing, font sizes, radii, shadows,
transition durations, z-index layers, or hex colors — and converts
touched values to tokens when editing. Legacy counts are frozen per
file below; a count may only go down.
"""

import re

from _frontend_source import WEB

# Frozen per-file counts: (spacing_px, z_index, transition, hex).
# New violations fail; cleanups update the tuple downward.
_BASELINE = {
    "card-scrubber.css": (0, 2, 0, 0),
    "color-picker.css": (4, 0, 0, 11),
    "composer.css": (2, 4, 2, 5),
    "gallery.css": (7, 0, 0, 7),
    "overview-convergence.css": (16, 1, 0, 0),
    "overview-metadata.css": (15, 1, 0, 0),
    "overview-reports.css": (0, 0, 2, 1),
    "overview.css": (2, 0, 1, 0),
    "primitives.css": (36, 3, 0, 0),
    "screenspace.css": (52, 0, 1, 14),
    "settings-modal.css": (11, 0, 1, 4),
    "start-overlay.css": (168, 9, 34, 11),
    "studio.css": (54, 5, 9, 50),
    "topnav.css": (25, 0, 0, 7),
    "transcripts.css": (35, 2, 6, 0),
    "viewer.css": (26, 0, 2, 16),
    "workflows.css": (6, 2, 3, 1),
}

_CATEGORIES = ("raw px/rem", "raw z-index", "raw duration", "hex color")

# Scoped, intentional overrides of a tokens.css custom property.
# Each names the reason; anything else is a forked token.
_SCOPED_OVERRIDES = {
    "screenspace.css": {"--select-caret-inset"},  # tighter model-view select
    "start-overlay.css": {"--duration-veil"},  # slower start-veil reveal
    "topnav.css": {"--topnav-height"},  # desktop-chrome height override
}

_PROP = re.compile(
    r"^\s*(margin(?:-[a-z]+)?|padding(?:-[a-z]+)?|gap|row-gap|column-gap"
    r"|font-size|border-radius|box-shadow)\s*:",
    re.IGNORECASE,
)
_NUM = re.compile(r"\b\d+(?:\.\d+)?(?:px|rem)\b")
_Z = re.compile(r"^\s*z-index\s*:\s*-?\d")
_TRANS = re.compile(
    r"^\s*(transition|animation)(?:-duration|-delay)?\s*:.*?\b\d+(?:\.\d+)?m?s\b"
)
_HEX = re.compile(r"#[0-9a-fA-F]{3,8}\b")
_DEF = re.compile(r"^\s*(--[\w-]+)\s*:", re.MULTILINE)


def _strip(css: str) -> str:
    return re.sub(r"/\*.*?\*/", "", css, flags=re.DOTALL)


def _count(css: str) -> tuple[int, int, int, int]:
    px = z = tr = hx = 0
    for line in css.splitlines():
        if "url(" in line:  # data-URIs and icon refs, not design values
            continue
        if _PROP.match(line) and _NUM.search(line):
            px += 1
        if _Z.match(line):
            z += 1
        if _TRANS.match(line):
            tr += 1
        hx += len(_HEX.findall(line))
    return px, z, tr, hx


def _page_css():
    for p in sorted(WEB.glob("*.css")):
        if p.name != "tokens.css":
            yield p.name, _strip(p.read_text(encoding="utf-8"))


def test_no_new_raw_design_values():
    """Raw-value counts may not grow; use tokens instead."""
    grew, shrank = [], []
    for name, css in _page_css():
        actual = _count(css)
        frozen = _BASELINE.get(name, (0, 0, 0, 0))
        for cat, a, f in zip(_CATEGORIES, actual, frozen):
            if a > f:
                grew.append(f"{name}: {a} {cat} (baseline {f}) — use a token")
            elif a < f:
                shrank.append(f"{name}: {cat} {f} -> {a}")
    assert not grew, "\n".join(grew)
    assert not shrank, "nice — fewer raw values; ratchet _BASELINE down:\n" + "\n".join(
        shrank
    )


def test_baseline_has_no_dead_entries():
    live = {name for name, _ in _page_css()}
    dead = sorted(set(_BASELINE) - live)
    assert not dead, f"baselined files no longer exist: {dead}"


def test_shared_tokens_not_redefined():
    """A tokens.css custom property may only be re-set on the allowlist."""
    tokens = set(_DEF.findall(_strip((WEB / "tokens.css").read_text(encoding="utf-8"))))
    offenders, stale = [], []
    for name, css in _page_css():
        redefined = set(_DEF.findall(css)) & tokens
        allowed = _SCOPED_OVERRIDES.get(name, set())
        for prop in sorted(redefined - allowed):
            offenders.append(f"{name} redefines {prop} (owned by tokens.css)")
        for prop in sorted(allowed - redefined):
            stale.append(f"{name}: {prop} no longer overridden")
    assert not offenders, "\n".join(offenders)
    assert not stale, "remove stale _SCOPED_OVERRIDES entries:\n" + "\n".join(stale)
