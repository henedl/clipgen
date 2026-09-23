"""Icons are masked SVG files, not inline markup.

The canonical pattern is a `<span>` with `mask-image:
url("icons/name.svg")` and `background-color: currentColor` (see
`XREF_BADGES` in utils.js). Inline `<svg>`, data-URI SVG, and path
data in JS are exceptions, frozen below with their reasons; each site
also carries an inline comment.
"""

import re

from _frontend_source import WEB, strip_comments

# Files allowed to contain "<svg" (comment-stripped), with counts.
# Reasons: animations needing SVG children, self-contained exported
# pages (no /icons/ route offline), brand glyphs, tile artworks.
_INLINE_SVG = {
    "boot.html": 1,  # brand mark; the boot dispatcher serves no /logos/ route
    "gallery.html": 1,  # favicon data-URI (self-contained export)
    "hotkeys.css": 1,  # cheatsheet icon; hotkeys.css inlined into exports
    "start-overlay.html": 5,  # brand tab glyphs + tool-tile artworks
    "studio.html": 2,  # title-spinner animations (animateTransform)
    "studio.js": 1,  # createPulserOverlay animation
    "timeline-viewer.html": 5,  # self-contained export
    "tokens.css": 3,  # --select-caret (background-image, per theme)
    "viewer.css": 2,  # data-URI masks (self-contained export)
    "viewer.html": 4,  # self-contained export
    "workflows.html": 1,  # empty #wfWires container, holds no icon
}

# JS files allowed to define SVG path data, with counts.
_JS_PATH_DATA = {
    "viewer.js": 16,  # SS_DETECTOR_ICON_PATHS; exports have no /icons/
}

# Files allowed to embed data-URI SVG at all.
_DATA_URI = {
    "gallery.html",  # favicon
    "hotkeys.css",  # cheatsheet icon (inlined into exports)
    "timeline-viewer.html",  # favicon
    "tokens.css",  # --select-caret
    "viewer.css",  # masks
    "viewer.html",  # favicon
}

_PATH_DATA = re.compile(r"""\bd\s*[:=]\s*["'][MmLl]?\s*[\d.]""")


def _sources():
    for p in sorted(WEB.glob("*")):
        if p.suffix not in (".js", ".html", ".css"):
            continue
        src = p.read_text(encoding="utf-8")
        if p.suffix == ".html":
            src = re.sub(r"<!--.*?-->", "", src, flags=re.DOTALL)
        else:
            src = strip_comments(src)
        yield p.name, src


def _check_counts(actual: dict, frozen: dict, what: str) -> None:
    grew = [
        f"{n}: {c} {what} (allowed {frozen.get(n, 0)}) — use a masked icon span"
        for n, c in sorted(actual.items())
        if c > frozen.get(n, 0)
    ]
    stale = [
        f"{n}: {what} {c} -> {actual.get(n, 0)}"
        for n, c in sorted(frozen.items())
        if actual.get(n, 0) < c
    ]
    assert not grew, "\n".join(grew)
    assert not stale, f"ratchet the {what} allowlist down:\n" + "\n".join(stale)


def test_inline_svg_only_at_known_sites():
    actual = {n: src.count("<svg") for n, src in _sources() if "<svg" in src}
    _check_counts(actual, _INLINE_SVG, "inline <svg>")


def test_no_svg_path_data_in_js():
    actual = {
        n: len(_PATH_DATA.findall(src))
        for n, src in _sources()
        if n.endswith(".js") and _PATH_DATA.search(src)
    }
    _check_counts(actual, _JS_PATH_DATA, "path-data literals")


def test_data_uri_svg_only_at_known_sites():
    actual = sorted(n for n, src in _sources() if "data:image/svg" in src)
    new = sorted(set(actual) - _DATA_URI)
    stale = sorted(_DATA_URI - set(actual))
    assert not new, f"new data-URI SVG (use icons/ + mask-image): {new}"
    assert not stale, f"stale _DATA_URI entries: {stale}"
