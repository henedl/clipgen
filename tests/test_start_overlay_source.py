"""Source-level wiring assertions for the Start overlay's Google panel.

The refresh button is the only user-facing escape hatch from the server's
5-minute Drive-listing cache (``server._cached_spreadsheet_meta``), and every
part of it is invisible to the Python tests: the markup, the ``?refresh=true``
query it sends, and the fact that it lives *outside* the status node that three
code paths rewrite.
"""

from __future__ import annotations

import re

from _frontend_source import assert_es5, read, strip_comments


def test_google_refresh_button_markup():
    html = read("start-overlay.html")
    assert 'data-role="google-refresh"' in html
    assert 'data-icon="arrow-path"' in html
    # Ships hidden: the unauthenticated panel shows a Connect CTA instead.
    button = re.search(r"<button[^>]*data-role=\"google-refresh\"[^>]*>", html)
    assert button and "hidden" in button.group(0)


def test_google_refresh_button_is_not_inside_the_status_node():
    """loadGoogleSheets / renderGoogleConnectCTA rewrite the status node."""
    html = read("start-overlay.html")
    status = html.index('data-role="google-status"')
    # The status <div> is self-contained: it closes before the button opens.
    closes = html.index("</div>", status)
    assert closes < html.index('data-role="google-refresh"')


def test_refresh_click_forces_a_relist_and_always_clears_the_spinner():
    js = read("start-overlay.js")
    assert_es5(js, "start-overlay.js")
    src = strip_comments(js)
    assert '"?refresh=true"' in src
    assert "loadGoogleSheets(true)" in src
    # The spinner teardown is a final .then() after .catch(), so a failed
    # refresh can't leave the button spinning forever.
    handler = src[src.index("loadGoogleSheets(true)") :][:400]
    assert handler.index(".catch(") < handler.index('classList.remove("is-spinning")')


def test_refresh_button_has_its_css():
    css = read("start-overlay.css")
    for rule in (".sheet-panel__status-row", ".sheet-panel__refresh", "is-spinning"):
        assert rule in css, f"{rule} toggled/used in the overlay but absent from CSS"
