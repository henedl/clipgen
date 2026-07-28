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

import desktop_chrome

from _frontend_source import read, strip_comments

SOURCE = Path(desktop_chrome.__file__).read_text(encoding="utf-8")

# Comment-stripped: the prose around this code names both the selector we use and
# the -webkit-app-region property we cannot.
TOPNAV_JS = strip_comments(read("topnav.js"))
TOPNAV_CSS = read("topnav.css")


def test_chrome_is_a_no_op_off_macos(monkeypatch):
    monkeypatch.setattr(desktop_chrome.sys, "platform", "linux")
    assert desktop_chrome.is_supported() is False
    assert desktop_chrome.chrome_style() is None
    # A truthy object with no `native` — apply() must decline, not explode.
    assert desktop_chrome.apply(object()) is False


def test_apply_declines_a_window_without_a_native_handle(monkeypatch):
    """Backends other than cocoa never set `native`; that is not a crash."""
    monkeypatch.setattr(desktop_chrome.sys, "platform", "darwin")

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


def test_topnav_opts_into_the_drag_selector_only_for_the_desktop_launch():
    """WKWebView ignores -webkit-app-region, so the bar uses pywebview's class."""
    assert "pywebview-drag-region" in TOPNAV_JS
    assert "-webkit-app-region" not in TOPNAV_JS
    # Gated on the server-set attribute: a browser page must be untouched.
    assert "dataset.desktopChrome" in TOPNAV_JS
    gate = TOPNAV_JS.index("dataset.desktopChrome")
    assert gate < TOPNAV_JS.index("pywebview-drag-region")


def test_topnav_css_insets_the_left_column_not_the_bar():
    """Padding the bar itself would shove the centered tabs half the inset right."""
    assert 'html[data-desktop-chrome="macos"] .topnav-left' in TOPNAV_CSS
    assert 'html[data-desktop-chrome="macos"] .topnav {' not in TOPNAV_CSS
    # Dragging must not sweep a text selection across the bar's labels.
    assert ".pywebview-drag-region" in TOPNAV_CSS
