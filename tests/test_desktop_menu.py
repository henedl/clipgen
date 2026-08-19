"""Regression checks for the native desktop menu bar.

The bar itself (NSMenu, key equivalents, the Window menu) cannot exist under
pytest — per the desktop_chrome precedent, native chrome is measured manually
with ``uv run clipgen.py --studio --desktop -v``. What *is* testable is the
decision layer: the Tier 1 menu tree pywebview is handed (titles, order, what
each action does), the platform gate, the topnav sync, and the source-shape
rules that keep the Linux CI typecheck green.
"""

import re
from pathlib import Path

import desktop
import desktop_menu

from _frontend_source import read, strip_comments

SOURCE = Path(desktop_menu.__file__).read_text(encoding="utf-8")
DESKTOP_SOURCE = Path(desktop.__file__).read_text(encoding="utf-8")
TOPNAV_JS = strip_comments(read("topnav.js"))
START_OVERLAY_JS = strip_comments(read("start-overlay.js"))


class FakeWindow:
    def __init__(self, url="http://127.0.0.1:8089/studio/"):
        self.scripts = []
        self.url = url

    def run_js(self, script):
        self.scripts.append(script)

    def get_current_url(self):
        return self.url


def menu_titles(menus):
    return [menu.title for menu in menus]


def walk_actions(menus):
    """Yield every MenuAction in the tree, submenus included."""
    for entry in menus:
        items = getattr(entry, "items", None)
        if items is not None:
            yield from walk_actions(items)
        elif callable(getattr(entry, "function", None)):
            yield entry


def test_menu_tree_titles_and_order():
    menus = desktop_menu.build_menus(lambda: None)
    assert menu_titles(menus) == ["__app__", "File", "Go", "Help"]
    go = menus[2]
    assert [item.title for item in go.items] == [
        "Studio",
        "Screenspace",
        "Transcripts",
        "Workflows",
        "Composer",
        "Overview",
    ]


def test_surfaces_mirror_topnav():
    """The Go menu is the topnav in native clothing; the lists must not drift."""
    for label, href in desktop_menu._SURFACES:
        assert f'label: "{label}", href: "{href}"' in TOPNAV_JS, (label, href)
    assert len(desktop_menu._SURFACES) == TOPNAV_JS.count("label:")


def test_go_actions_navigate():
    window = FakeWindow()
    menus = desktop_menu.build_menus(lambda: window)
    go = menus[2]
    for item, (_, href) in zip(go.items, desktop_menu._SURFACES, strict=True):
        item.function()
    assert [f'location.href = "{href}";' for _, href in desktop_menu._SURFACES] == (
        window.scripts
    )


def test_every_action_is_a_no_op_without_a_window(monkeypatch):
    """Clicking any item before the window exists must not raise.

    The externally-visible helpers are stubbed so the sweep stays hermetic —
    no browser tabs or Finder windows out of a test run.
    """
    opened = []
    monkeypatch.setattr(desktop_menu.webbrowser, "open", opened.append)
    monkeypatch.setattr(
        desktop_menu.subprocess, "run", lambda *a, **k: opened.append(a)
    )
    for action in walk_actions(desktop_menu.build_menus(lambda: None)):
        action.function()


def test_page_global_snippets_are_guarded():
    """The boot page loads no frontend bundle; every snippet must self-guard."""
    for name, script in vars(desktop_menu).items():
        if not name.startswith("_JS_") or not isinstance(script, str):
            continue
        if "window." in script:
            assert script.startswith("if (window."), name


def test_menus_are_gated_to_macos(monkeypatch):
    monkeypatch.setattr(desktop_menu.sys, "platform", "linux")
    assert desktop_menu.is_supported() is False
    assert desktop_menu.menus(lambda: None) == []
    # Must decline before touching pyobjc, which does not exist off macOS.
    assert desktop_menu.enhance_menu_bar(lambda: None) is None


def test_open_folder_skips_missing_directories(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(desktop_menu.subprocess, "run", lambda *a, **k: calls.append(a))
    desktop_menu._open_folder("")
    desktop_menu._open_folder(str(tmp_path / "nope"))
    assert calls == []
    desktop_menu._open_folder(str(tmp_path))
    assert len(calls) == 1


def test_open_in_browser_uses_the_current_url(monkeypatch):
    opened = []
    monkeypatch.setattr(desktop_menu.webbrowser, "open", opened.append)
    desktop_menu._open_in_browser(lambda: None)
    desktop_menu._open_in_browser(lambda: FakeWindow(url=None))
    assert opened == []
    desktop_menu._open_in_browser(lambda: FakeWindow())
    assert opened == ["http://127.0.0.1:8089/studio/"]


def test_key_equivalents_only_name_real_items():
    """A renamed menu label must not silently orphan its ⌘-shortcut."""
    titles = {a.title for a in walk_actions(desktop_menu.build_menus(lambda: None))}
    for title in desktop_menu._KEY_EQUIVALENTS:
        assert title in titles, title


def test_whats_new_opens_the_changelog_tab():
    """The menu snippet leans on start-overlay's open(tab) parameter."""
    assert 'open("updates")' in desktop_menu._JS_WHATS_NEW
    assert 'setStartTab(tab || "open")' in START_OVERLAY_JS
    assert 'data-start-tab="updates"' in read("start-overlay.html")


def test_pyobjc_is_imported_by_name_not_by_statement():
    """A literal `import AppKit` fails the Linux typecheck job, not the tests."""
    for statement in (
        r"^\s*import AppKit",
        r"^\s*from AppKit import",
        r"^\s*import PyObjCTools",
        r"^\s*from PyObjCTools import",
    ):
        assert not re.search(statement, SOURCE, re.MULTILINE), statement
    assert 'importlib.import_module("AppKit")' in SOURCE
    assert 'importlib.import_module("PyObjCTools.AppHelper")' in SOURCE


def test_desktop_hands_the_menus_to_webview():
    assert "menu=menus" in DESKTOP_SOURCE
    assert "desktop_menu.enhance_menu_bar" in DESKTOP_SOURCE
