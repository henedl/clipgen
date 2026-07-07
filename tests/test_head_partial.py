"""Regression checks for the shared server-injected ``<head>`` partial (C6).

The four live pages embed ``<!-- CLIPGEN_HEAD_HERE -->`` where the shared
favicon + Google-fonts block belongs; ``utils.render_index_html`` expands it
from ``assets/web/_head.html``. Exported viewers stay self-contained and must
not gain the marker.
"""

from pathlib import Path

import utils

WEB = Path(__file__).resolve().parent.parent / "assets" / "web"
HEAD_MARKER = "<!-- CLIPGEN_HEAD_HERE -->"
LIVE_PAGES = ("studio.html", "screenspace.html", "transcripts.html", "workflows.html")
EXPORT_PAGES = ("viewer.html", "gallery.html", "timeline-viewer.html")
FAVICON_LINK = '<link rel="icon" type="image/svg+xml" href="logos/favicon.svg">'
FONTS_LINK = "https://fonts.googleapis.com/css2?family=Inter"


def test_head_partial_holds_shared_block():
    head = (WEB / "_head.html").read_text(encoding="utf-8")
    assert FAVICON_LINK in head
    assert FONTS_LINK in head
    # All seven favicon/apple-touch links plus preconnect/preload/stylesheet.
    assert head.count("<link") == 11


def test_live_pages_use_marker_not_inline_block():
    for page in LIVE_PAGES:
        src = (WEB / page).read_text(encoding="utf-8")
        assert HEAD_MARKER in src, page
        # The duplicated block must be gone so it can only be edited in one place.
        assert FAVICON_LINK not in src, page
        assert FONTS_LINK not in src, page


def test_export_pages_stay_self_contained():
    for page in EXPORT_PAGES:
        src = (WEB / page).read_text(encoding="utf-8")
        assert HEAD_MARKER not in src, page


def test_render_index_expands_marker():
    rendered = utils.render_index_html(WEB, "transcripts.html")
    assert HEAD_MARKER not in rendered
    assert FAVICON_LINK in rendered
    assert FONTS_LINK in rendered
    # No empty line left where the marker stood (the partial is rstripped).
    assert '\n\n  <link rel="stylesheet" href="tokens.css">' not in rendered


def test_render_index_passthrough_without_marker():
    # Export templates have no marker -> returned unchanged.
    src = (WEB / "gallery.html").read_text(encoding="utf-8")
    assert utils.render_index_html(WEB, "gallery.html") == src
