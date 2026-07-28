"""Regression checks for the shared server-injected ``<head>`` partial (C6).

The four live pages embed ``<!-- CLIPGEN_HEAD_HERE -->`` where the shared
favicon + fonts block belongs; ``utils.render_index_html`` expands it from
``assets/web/_head.html``. Exported viewers stay self-contained and must not
gain the marker.

Fonts are vendored under ``assets/web/fonts/`` rather than fetched from
fonts.googleapis.com, so the desktop shell and any offline run render with the
intended typefaces instead of blocking on a CDN. The no-remote-reference check
below is what keeps that from regressing.
"""

import utils

from _frontend_source import WEB

HEAD_MARKER = "<!-- CLIPGEN_HEAD_HERE -->"
LIVE_PAGES = ("studio.html", "screenspace.html", "transcripts.html", "workflows.html")
EXPORT_PAGES = ("viewer.html", "gallery.html", "timeline-viewer.html")
FAVICON_LINK = '<link rel="icon" type="image/svg+xml" href="logos/favicon.svg">'
FONTS_LINK = '<link rel="stylesheet" href="fonts.css">'


def test_head_partial_holds_shared_block():
    head = (WEB / "_head.html").read_text(encoding="utf-8")
    assert FAVICON_LINK in head
    assert FONTS_LINK in head
    # Seven favicon/apple-touch links plus the vendored-fonts stylesheet.
    assert head.count("<link") == 8
    # The whole point: nothing in the shared head may reach off-machine.
    assert "https://" not in head.replace("fonts.googleapis.com", "")


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


def test_no_live_page_references_a_remote_origin():
    """Every asset a live page pulls must be served locally (offline-capable)."""
    import re

    for page in LIVE_PAGES:
        rendered = utils.render_index_html(WEB, page)
        remote = re.findall(r'(?:href|src)=["\']https?://[^"\']+', rendered)
        assert not remote, f"{page} still loads remote assets: {remote}"


def test_desktop_chrome_script_is_desktop_only():
    """The native-window flag rides the head partial, and only in that launch."""
    import config

    assert "desktopChrome" not in utils.render_index_html(WEB, "studio.html")

    utils.DESKTOP_CHROME = "macos"
    rendered = utils.render_index_html(WEB, "studio.html")
    assert 'd.dataset.desktopChrome = "macos"' in rendered
    # The measurements come from config, never hand-written into the CSS.
    assert f'"{config.DESKTOP_CHROME_BAR_HEIGHT}px"' in rendered
    assert f'"{config.DESKTOP_TRAFFIC_LIGHT_INSET}px"' in rendered
    # Ahead of topnav.css, so the bar lays out inset on first paint.
    assert rendered.index("desktopChrome") < rendered.index('href="topnav.css"')

    # The render cache is keyed on mtimes; without the flag in the key too, a
    # browser render would be served into a desktop window and vice versa.
    utils.DESKTOP_CHROME = None
    assert "desktopChrome" not in utils.render_index_html(WEB, "studio.html")


def test_render_index_passthrough_without_marker():
    # Export templates have no marker -> returned unchanged.
    src = (WEB / "gallery.html").read_text(encoding="utf-8")
    assert utils.render_index_html(WEB, "gallery.html") == src
