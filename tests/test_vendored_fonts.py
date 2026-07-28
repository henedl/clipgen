"""Guards for the self-hosted web fonts (assets/web/fonts/).

Inter and JetBrains Mono are vendored so the desktop shell and any offline run
render with the intended typefaces instead of blocking on fonts.googleapis.com.

Both are *variable* fonts: Google serves one file per family+subset and emits a
separate @font-face per requested weight, each pinning the wght axis through its
``font-weight`` descriptor. Downloading per weight therefore yields byte-identical
copies — which is exactly what the first cut of this shipped. These tests pin the
shape so a regenerate cannot quietly reintroduce the duplication or leave a
stylesheet pointing at a file that is not there.
"""

import hashlib
import re
from collections import defaultdict

from _frontend_source import WEB

FONTS_CSS = WEB / "fonts.css"
FONTS_DIR = WEB / "fonts"


def _font_files():
    return sorted(FONTS_DIR.glob("*.woff2"))


def test_no_two_font_files_are_identical():
    """One file per family+subset — never one copy per weight."""
    by_digest = defaultdict(list)
    for path in _font_files():
        digest = hashlib.md5(path.read_bytes()).hexdigest()
        by_digest[digest].append(path.name)

    dupes = {d: names for d, names in by_digest.items() if len(names) > 1}
    assert not dupes, (
        "Byte-identical font files found — these are variable fonts, so one file "
        f"covers every weight: {dupes}"
    )


def test_every_css_reference_resolves():
    css = FONTS_CSS.read_text(encoding="utf-8")
    referenced = set(re.findall(r"url\(fonts/([^)]+)\)", css))
    assert referenced, "fonts.css references no font files at all"
    for name in sorted(referenced):
        assert (FONTS_DIR / name).is_file(), f"fonts.css points at missing {name}"


def test_every_font_file_is_referenced():
    """No orphans left behind by a regenerate."""
    css = FONTS_CSS.read_text(encoding="utf-8")
    referenced = set(re.findall(r"url\(fonts/([^)]+)\)", css))
    for path in _font_files():
        assert path.name in referenced, f"{path.name} is not referenced by fonts.css"


def test_all_used_weights_are_declared():
    """Every weight the UI asks for must have an @font-face pinning it.

    A missing declaration does not fail loudly — the browser silently synthesizes
    or falls back — so the only way to catch it is to assert the set here.
    """
    css = FONTS_CSS.read_text(encoding="utf-8")
    blocks = re.findall(
        r"font-family:\s*'([^']+)'.*?font-weight:\s*(\d+)",
        css,
        re.DOTALL | re.MULTILINE,
    )
    declared = defaultdict(set)
    for family, weight in blocks:
        declared[family].add(int(weight))

    # Mirrors the families/weights _head.html used to request from Google Fonts.
    assert declared["Inter"] == {400, 500, 600}, declared["Inter"]
    assert declared["JetBrains Mono"] == {400, 500}, declared["JetBrains Mono"]


def test_stylesheet_is_fully_offline():
    css = FONTS_CSS.read_text(encoding="utf-8")
    assert "https://" not in css.split("*/", 1)[-1], (
        "fonts.css must not reference a remote origin outside its header comment"
    )
