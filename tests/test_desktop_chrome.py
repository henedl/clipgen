"""Regression checks for the native desktop window chrome.

``desktop_chrome`` styles the macOS NSWindow so the title bar disappears and the
traffic lights float inside clipgen's topnav. None of that can run under pytest
(there is no window), so the checks here cover the two halves that *are*
testable: the platform gate — CI is Linux, where every entry point must be an
inert no-op rather than an ImportError — and the frontend contract, where the
bar has to opt into pywebview's drag selector for the desktop launch only.
"""

import re
from pathlib import Path

import desktop
import desktop_chrome
import pytest

from _frontend_source import read, strip_comments

SOURCE = Path(desktop_chrome.__file__).read_text(encoding="utf-8")
DESKTOP_SOURCE = Path(desktop.__file__).read_text(encoding="utf-8")

# Comment-stripped: the prose around this code names both the selector we use and
# the -webkit-app-region property we cannot.
TOPNAV_JS = strip_comments(read("topnav.js"))
TOPNAV_CSS = read("topnav.css")
UTILS_JS = strip_comments(read("utils.js"))


def test_chrome_is_a_no_op_off_macos(monkeypatch):
    monkeypatch.setattr(desktop_chrome.sys, "platform", "linux")
    assert desktop_chrome.is_supported() is False
    assert desktop_chrome.chrome_style() is None
    # A truthy object with no `native` — apply() must decline, not explode.
    assert desktop_chrome.apply(object()) is False


def test_apply_declines_a_window_without_a_native_handle(monkeypatch):
    """Backends other than cocoa never set `native`; that is not a crash."""
    monkeypatch.setattr(desktop_chrome.sys, "platform", "darwin")
    # apply() imports AppKit before it looks at `native`, and on a real Mac that
    # drags the whole pyobjc stack in (~2 s cold) for a branch that never touches
    # it. Stubbing the accessor also makes the test hermetic: the `native is None`
    # path is now exercised on Linux and macOS alike, rather than passing on Linux
    # only because the import happens to fail first.
    monkeypatch.setattr(desktop_chrome, "_appkit", lambda: object())

    class Window:
        native = None

    assert desktop_chrome.apply(Window()) is False


def test_pyobjc_is_imported_by_name_not_by_statement():
    """A literal `import AppKit` fails the Linux typecheck job, not the tests.

    Nothing here runs the macOS branch, so a plain import statement would sail
    through pytest on any platform and only blow up in CI's `ty` step. Assert the
    shape instead — this exact regression cost a red build once.
    """
    for statement in (
        r"^\s*import AppKit",
        r"^\s*from AppKit import",
        r"^\s*import PyObjCTools",
        r"^\s*from PyObjCTools import",
    ):
        assert not re.search(statement, SOURCE, re.MULTILINE), statement
    assert 'importlib.import_module("AppKit")' in SOURCE
    assert 'importlib.import_module("PyObjCTools.AppHelper")' in SOURCE


def test_string_imports_are_declared_to_pyinstaller():
    """importlib hides them from PyInstaller's static analysis; the spec re-adds them."""
    spec = (Path(__file__).resolve().parents[1] / "build" / "clipgen.spec").read_text(
        encoding="utf-8"
    )
    assert '"AppKit", "PyObjCTools.AppHelper"' in spec


def test_teardown_is_idempotent():
    desktop_chrome.teardown()
    desktop_chrome.teardown()
    assert desktop_chrome._observers == []
    # The layout state resets sit above the "no observers" early-out: apply()
    # runs a first layout pass before _observe() registers anything.
    assert desktop_chrome._frame_observed is None
    assert desktop_chrome._frame_token is None
    assert desktop_chrome._laying_out is False
    assert desktop_chrome._last_inventory is None


def test_the_container_is_resolved_without_the_traffic_lights():
    """During a share there may be no buttons to walk up from.

    The container used to be reached as buttons[0].superview().superview(), so
    losing the buttons lost the container — and with it the 48px band.
    """
    assert '"NSTitlebarContainerView"' in SOURCE
    # Both the styling pass and the layout pass go through the one resolver.
    assert SOURCE.count("_titlebar_container(") >= 3  # 1 def + 2 call sites


def test_the_titlebar_grows_before_the_buttons_are_looked_up():
    """The regression: macOS Sequoia swaps the lights for a sharing pill.

    Gating the growth on all three buttons being present left the bar at its
    stock 28px for the whole share, so the pill rode high in clipgen's topnav.
    """
    body = SOURCE[SOURCE.index("def _apply_titlebar_layout") :]
    body = body[: body.index("\ndef ")]
    assert body.index("setFrame_") < body.index("standardWindowButton_")


def test_a_sharing_swap_re_asserts_the_layout():
    """No window-level notification fires when the sharing pill appears.

    Resize and the two fullscreen notifications are all the module used to
    observe, which is why the lights came back at stock positions after a share
    and stayed there until the window was resized.
    """
    assert "NSViewFrameDidChangeNotification" in SOURCE
    assert "setPostsFrameChangedNotifications_" in SOURCE
    assert "NSWindowDidBecomeKeyNotification" in SOURCE


def test_the_frame_observer_cannot_recurse_into_itself():
    """Our own setFrame_ posts the notification the observer listens for."""
    assert re.search(r"^\s+_laying_out = True", SOURCE, re.MULTILINE)
    assert re.search(r"^\s+if _laying_out:", SOURCE, re.MULTILINE)


@pytest.mark.parametrize(
    "band,size,expected",
    [
        (48, 16, 16),  # the shipped bar: also the DESKTOP_TRAFFIC_LIGHT_INSET math
        (28, 16, 6),  # AppKit's stock titlebar
        # Half-pixel remainders: round() is ties-to-even, so these land on
        # different sides. Pinned because a sub-pixel origin blurs the button.
        (48, 15, 16),  # 16.5 -> 16
        (48, 13, 18),  # 17.5 -> 18
    ],
)
def test_centered_origin_matches_the_configured_inset(band, size, expected):
    assert desktop_chrome._centered_origin(band, size) == expected


def test_the_reassert_budget_bounds_an_appkit_layout_fight():
    """A container AppKit pins would otherwise be an unbounded main-thread loop."""
    desktop_chrome.teardown()
    try:
        limit = desktop_chrome._REASSERT_BURST_LIMIT
        assert all(desktop_chrome._within_reassert_budget(0.0) for _ in range(limit))
        assert desktop_chrome._within_reassert_budget(0.0) is False
        # A later burst is a fresh window, not a continuation of the fight.
        assert desktop_chrome._within_reassert_budget(
            desktop_chrome._REASSERT_WINDOW_S + 1.0
        )
    finally:
        desktop_chrome.teardown()


def test_topnav_opts_into_the_drag_selector_only_for_the_desktop_launch():
    """WKWebView ignores -webkit-app-region, so the bar uses pywebview's class."""
    assert "pywebview-drag-region" in TOPNAV_JS
    assert "-webkit-app-region" not in TOPNAV_JS
    # Gated on the server-set attribute: a browser page must be untouched.
    assert "dataset.desktopChrome" in TOPNAV_JS
    gate = TOPNAV_JS.index("dataset.desktopChrome")
    assert gate < TOPNAV_JS.index("pywebview-drag-region")


@pytest.mark.parametrize(
    "setting,expected",
    [
        ("Maximize", "zoom"),  # the macOS default
        ("Fill", "zoom"),  # Sequoia's tiling option; no public NSWindow call
        (None, "zoom"),  # never written — AppKit falls back to zoom too
        ("Minimize", "minimize"),
        ("None", None),
    ],
)
def test_double_click_honours_the_system_preference(setting, expected):
    assert desktop_chrome._double_click_action(setting) == expected


def test_double_click_is_a_no_op_without_a_window(monkeypatch):
    """Off macOS, and on a backend with no `native`, the gesture just does nothing."""
    monkeypatch.setattr(desktop_chrome.sys, "platform", "linux")
    desktop_chrome.titlebar_double_click(object())

    monkeypatch.setattr(desktop_chrome.sys, "platform", "darwin")

    class Window:
        native = None

    desktop_chrome.titlebar_double_click(Window())


def test_double_click_is_bridged_to_the_page():
    """The page can only reach the action if desktop.py exposes it to JS."""
    expose = DESKTOP_SOURCE[DESKTOP_SOURCE.index("window.expose(") :]
    assert "titlebar_double_click" in expose[: expose.index(")")]
    # Same direct-target rule as the drag, so both gestures cover the same pixels.
    handler = TOPNAV_JS[TOPNAV_JS.index('addEventListener("dblclick"') :]
    body = handler[: handler.index("});")]
    assert body.index("pywebview-drag-region") < body.index("titlebar_double_click")
    # Inside the desktop gate: a browser page must never call the bridge.
    assert TOPNAV_JS.index("dataset.desktopChrome") < TOPNAV_JS.index(
        'addEventListener("dblclick"'
    )


def test_set_appearance_is_a_no_op_without_a_window(monkeypatch):
    monkeypatch.setattr(desktop_chrome.sys, "platform", "linux")
    desktop_chrome.set_appearance(object(), "dark")

    monkeypatch.setattr(desktop_chrome.sys, "platform", "darwin")

    class Window:
        native = None

    desktop_chrome.set_appearance(Window(), "dark")


def test_window_appearance_follows_the_page_theme():
    """Light appearance over a dark page flashes white for a frame on zoom.

    Measured as one frame of (255, 255, 255) in the exposed area, so the window
    has to be told the theme — on load and on every toggle, not just one of them.
    """
    assert "NSAppearanceNameAqua" in SOURCE and "NSAppearanceNameDarkAqua" in SOURCE
    expose = DESKTOP_SOURCE[DESKTOP_SOURCE.index("window.expose(") :]
    assert "set_window_appearance" in expose[: expose.index(")")]

    for setter in ("applyStoredThemePreference", "toggleThemePreference"):
        body = UTILS_JS[UTILS_JS.index("var " + setter + " = function") :]
        assert "syncDesktopAppearance" in body[: body.index("\n};")], setter
    sync = UTILS_JS[UTILS_JS.index("var syncDesktopAppearance") :]
    sync = sync[: sync.index("\n};")]
    # Desktop-only, and the bridge does not exist yet on the first page load.
    assert "dataset.desktopChrome" in sync
    assert "pywebviewready" in sync


def test_topnav_css_insets_the_left_column_not_the_bar():
    """Padding the bar itself would shove the centered tabs half the inset right."""
    assert 'html[data-desktop-chrome="macos"] .topnav-left' in TOPNAV_CSS
    assert 'html[data-desktop-chrome="macos"] .topnav {' not in TOPNAV_CSS
    # Dragging must not sweep a text selection across the bar's labels.
    assert ".pywebview-drag-region" in TOPNAV_CSS
