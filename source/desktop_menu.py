"""Native menu bar for the clipgen desktop window.

pywebview's defaults give the window an App menu (About / Hide / Quit ⌘Q), Edit
(Cut / Copy / Paste / Select All) and View (Enter Full Screen) — the bare minimum
of an embedded webview. This module fills the bar out to the standard macOS set
so the app reads as a first-class citizen, in two tiers:

Tier 1 — :func:`menus` builds ``File`` / ``Go`` / ``Help`` menus (plus a
``Settings…`` item spliced into the application menu via pywebview's magic
``__app__`` title) through the public ``webview.menu`` API. Item callbacks run on
a fresh Python thread (cocoa's ``MenuHandler`` spawns one per click), so calling
``window.run_js`` from them is safe — it is only main-thread handlers that
deadlock on the run_js semaphore. Every JS snippet self-guards on the global it
pokes: the boot page loads none of the frontend bundles, and a menu click there
must be a silent no-op.

Tier 2 — :func:`enhance_menu_bar` runs one AppKit polish pass over the built
bar, because the public API cannot express any of it: custom items get no key
equivalents (cocoa hardcodes ``''``), custom menus are appended *after* View
(File belongs at index 1), the default View menu cannot receive a Reload item,
and there is no Window menu at all. The pass is queued with ``callAfter`` from
the ``shown`` event: callAfter work only executes once the run loop is live,
which is strictly after ``first_show``'s ``setMainMenu_``, and with a single
window and a static menu list pywebview never rebuilds the bar afterwards. A
failure here is cosmetic — Tier 1 keeps working — so the whole pass degrades
with a warning, matching desktop_chrome's contract that chrome polish must
never cost the user their window.
"""

import importlib
import sys
import webbrowser
from pathlib import Path
from typing import Any
from collections.abc import Callable

import config
import utils

# Mirrors SURFACES in topnav.js, same order; tests/test_desktop_menu.py checks.
_SURFACES: tuple[tuple[str, str], ...] = (
    ("Studio", "/studio/"),
    ("Screenspace", "/screenspace/"),
    ("Transcripts", "/transcripts/"),
    ("Workflows", "/workflows/"),
    ("Composer", "/composer/"),
    ("Overview", "/overview/"),
)

_HELP_URL = "https://github.com/henedl/clipgen#readme"

# Each snippet guards its global: the boot page has none, and a click must no-op.
_JS_OPEN_SETTINGS = "if (window.openSettingsModal) window.openSettingsModal({});"
_JS_NEW_SESSION = (
    "if (window.ClipgenStartOverlay && window.ClipgenStartOverlay.open)"
    " window.ClipgenStartOverlay.open();"
)
_JS_WHATS_NEW = (
    "if (window.ClipgenStartOverlay && window.ClipgenStartOverlay.open)"
    ' window.ClipgenStartOverlay.open("updates");'
)
_JS_CHECK_UPDATES = (
    "if (window.ClipgenStartOverlay && window.ClipgenStartOverlay.checkForUpdates)"
    " window.ClipgenStartOverlay.checkForUpdates();"
)
_JS_CHEATSHEET = (
    "if (window.ClipgenHotkeys && window.ClipgenHotkeys.toggleCheatsheet)"
    " window.ClipgenHotkeys.toggleCheatsheet();"
)
# No ⌘K equivalent here: the in-page registry owns Mod+K and would race it.
_JS_COMMAND_PALETTE = (
    "if (window.ClipgenCommandPalette && window.ClipgenCommandPalette.toggle)"
    " window.ClipgenCommandPalette.toggle();"
)
# Same path as the command palette: #themeToggle also syncs the native appearance.
_JS_TOGGLE_THEME = 'var t = document.getElementById("themeToggle"); if (t) t.click();'

# Title → ⌘-key for the Tier 2 pass (public API cannot). A test catches orphans.
_KEY_EQUIVALENTS: dict[str, str] = {
    "Settings…": ",",
    "New Session…": "n",
    "Studio": "1",
    "Screenspace": "2",
    "Transcripts": "3",
    "Workflows": "4",
    "Composer": "5",
    "Overview": "6",
}


def is_supported() -> bool:
    """Report whether this platform gets a native clipgen menu bar."""
    return sys.platform == "darwin"


def _appkit() -> Any:
    """Import AppKit as an opaque module.

    Imported by name rather than with a plain ``import AppKit`` because pyobjc
    only exists on macOS, and CI type-checks on Linux — a literal import is an
    ``unresolved-import`` error there. Do not "simplify" it back.

    Typed as ``Any`` on purpose: pyobjc's stubs are incomplete and every call
    below is a dynamically-bridged ObjC selector, so checking against them buys
    nothing and costs a suppression at each site.
    """
    return importlib.import_module("AppKit")


def menus(get_window: Callable[[], Any]) -> list:
    """The menu list to hand ``webview.start``; empty off macOS.

    The ``__app__`` magic title is cocoa-only (winforms would render it
    literally), so the platform gate lives here rather than in the caller.
    """
    return build_menus(get_window) if is_supported() else []


def build_menus(get_window: Callable[[], Any]) -> list:
    """Build the Tier 1 menu tree from pywebview's public menu API.

    *get_window* is a zero-arg getter (``lambda: desktop._window``) rather than
    a window: the list is built before ``webview.start``, when the module
    global is still None, and every action re-reads it at click time.
    """
    # Deferred so the module imports without pywebview installed.
    from webview.menu import Menu, MenuAction, MenuSeparator

    def js_action(title: str, script: str) -> Any:
        return MenuAction(title, lambda: _run_js(get_window, script))

    app_items = [
        js_action("Settings…", _JS_OPEN_SETTINGS),
        js_action("Check for Updates…", _JS_CHECK_UPDATES),
    ]

    file_items = [
        js_action("New Session…", _JS_NEW_SESSION),
        MenuSeparator(),
        # Read at click time: both dirs are empty until the user opens a workspace.
        MenuAction("Open Input Folder", lambda: _open_folder(config.INPUT_DIR)),
        MenuAction("Open Output Folder", lambda: _open_folder(config.OUTPUT_DIR)),
        MenuSeparator(),
        MenuAction("Open in Browser", lambda: _open_in_browser(get_window)),
    ]

    go_items = [
        js_action(label, f'location.href = "{href}";') for label, href in _SURFACES
    ]

    help_items = [
        MenuAction("clipgen Help", _open_help),
        MenuSeparator(),
        js_action("Keyboard Shortcuts", _JS_CHEATSHEET),
        js_action("Command Palette", _JS_COMMAND_PALETTE),
        MenuSeparator(),
        js_action("What's New…", _JS_WHATS_NEW),
        MenuAction("Third-Party Licenses", _open_licenses),
    ]

    return [
        Menu("__app__", app_items),
        Menu("File", file_items),
        Menu("Go", go_items),
        Menu("Help", help_items),
    ]


def _run_js(get_window: Callable[[], Any], script: str) -> None:
    """Run *script* in the page; a no-op while no window exists."""
    window = get_window()
    if window is None:
        return
    try:
        window.run_js(script)
    except Exception as exc:
        utils.warning_print(f"Menu action failed: {exc}")


def _open_folder(path: str) -> None:
    """Reveal *path* in the system file browser; a no-op unless it exists."""
    if not path or not Path(path).is_dir():
        utils.warning_print("No folder is configured yet — open a workspace first.")
        return
    utils.reveal_in_file_manager(Path(path))


def _open_in_browser(get_window: Callable[[], Any]) -> None:
    """Open the page currently shown in the window in the user's browser."""
    window = get_window()
    if window is None:
        return
    try:
        url = window.get_current_url()
    except Exception as exc:
        utils.warning_print(f"Could not read the current page URL: {exc}")
        return
    if url:
        webbrowser.open(url)


def _open_help() -> None:
    webbrowser.open(_HELP_URL)


def _open_licenses() -> None:
    path = utils.get_licenses_path()
    if path is None:
        utils.warning_print("THIRD-PARTY-LICENSES is not bundled in this build.")
        return
    utils.reveal_in_file_manager(path)


def enhance_menu_bar(get_window: Callable[[], Any]) -> None:
    """Queue the Tier 2 AppKit polish pass onto the main thread.

    Called from the ``shown`` event (a worker thread). Safe timing: callAfter
    work runs only once the run loop is live, i.e. after ``first_show`` has
    already called ``setMainMenu_`` with the Tier 1 bar.
    """
    if not is_supported():
        return
    try:
        app_helper: Any = importlib.import_module("PyObjCTools.AppHelper")
        app_helper.callAfter(lambda: _enhance_on_main(get_window))
    except Exception as exc:
        utils.warning_print(f"Could not polish the menu bar: {exc}")


def _enhance_on_main(get_window: Callable[[], Any]) -> None:
    """Everything the public menu API cannot express, in one main-thread pass."""
    try:
        AppKit = _appkit()
        app = AppKit.NSApplication.sharedApplication()
        main = app.mainMenu()
        if main is None:
            return
        _apply_key_equivalents(AppKit, main)
        _move_file_menu(main)
        _extend_view_menu(AppKit, main, get_window)
        _add_window_menu(AppKit, app, main)
        _register_help_menu(app, main)
    except Exception as exc:
        # Cosmetic: the Tier 1 menus still work without shortcuts or ordering.
        utils.warning_print(f"Could not polish the menu bar: {exc}")


def _apply_key_equivalents(AppKit: Any, main: Any) -> None:
    """Assign ⌘-shortcuts by item title; cocoa builds every custom item bare."""
    command_mask = getattr(AppKit, "NSEventModifierFlagCommand", 1 << 20)
    for holder in main.itemArray():
        submenu = holder.submenu()
        if submenu is None:
            continue
        for item in submenu.itemArray():
            key = _KEY_EQUIVALENTS.get(str(item.title()))
            if key:
                item.setKeyEquivalent_(key)
                item.setKeyEquivalentModifierMask_(command_mask)


def _move_file_menu(main: Any) -> None:
    """Move File to index 1, its native slot; cocoa appends it after View."""
    file_item = main.itemWithTitle_("File")
    if file_item is not None and main.indexOfItem_(file_item) != 1:
        # The pyobjc proxy in `file_item` keeps the item alive across the move.
        main.removeItem_(file_item)
        main.insertItem_atIndex_(file_item, 1)


def _extend_view_menu(AppKit: Any, main: Any, get_window: Callable[[], Any]) -> None:
    """Append Reload Page ⌘R and Toggle Dark Mode to the default View menu."""
    view_item = main.itemWithTitle_("View")
    if view_item is None or view_item.submenu() is None:
        return
    view = view_item.submenu()
    view.addItem_(AppKit.NSMenuItem.separatorItem())
    # Nil target: the responder chain routes reload: to the WKWebView, enabled
    # state included.
    reload_item = AppKit.NSMenuItem.alloc().initWithTitle_action_keyEquivalent_(
        "Reload Page", "reload:", "r"
    )
    view.addItem_(reload_item)
    _add_theme_toggle(AppKit, view, get_window)


def _add_theme_toggle(AppKit: Any, view: Any, get_window: Callable[[], Any]) -> None:
    """Add Toggle Dark Mode, dispatched through pywebview's menu handler.

    This is the one spot that reaches pywebview *internals*: a native item with
    a Python callback needs an ObjC-dispatchable target, and cocoa's module
    singleton ``menu_handler`` (which runs each action on a fresh thread, so
    run_js stays safe) already is one. Isolated in its own guard so a pywebview
    upgrade that renames the internals costs only this item, nothing else.
    """
    try:
        cocoa: Any = importlib.import_module("webview.platforms.cocoa")
        handler = cocoa.menu_handler
        action_id = "clipgen.toggle-dark-mode"
        handler.register_action(
            action_id, lambda: _run_js(get_window, _JS_TOGGLE_THEME)
        )
        item = AppKit.NSMenuItem.alloc().initWithTitle_action_keyEquivalent_(
            "Toggle Dark Mode", "handleMenuAction:", ""
        )
        item.setTarget_(handler)
        item.setRepresentedObject_(action_id)
        view.addItem_(item)
    except Exception as exc:
        utils.warning_print(f"Could not add the theme toggle menu item: {exc}")


def _add_window_menu(AppKit: Any, app: Any, main: Any) -> None:
    """Insert a real Window menu before Help and register it with the app.

    All standard nil-target selectors: AppKit routes them, manages their
    enabled state, and — because of ``setWindowsMenu_`` — appends the open
    window list to the bottom on its own.
    """
    window_menu = AppKit.NSMenu.alloc().init()
    window_menu.setTitle_("Window")
    window_menu.addItemWithTitle_action_keyEquivalent_(
        "Minimize", "performMiniaturize:", "m"
    )
    window_menu.addItemWithTitle_action_keyEquivalent_("Zoom", "performZoom:", "")
    window_menu.addItem_(AppKit.NSMenuItem.separatorItem())
    window_menu.addItemWithTitle_action_keyEquivalent_(
        "Bring All to Front", "arrangeInFront:", ""
    )
    holder = AppKit.NSMenuItem.alloc().init()
    holder.setSubmenu_(window_menu)
    help_item = main.itemWithTitle_("Help")
    index = (
        main.indexOfItem_(help_item) if help_item is not None else main.numberOfItems()
    )
    main.insertItem_atIndex_(holder, index)
    app.setWindowsMenu_(window_menu)


def _register_help_menu(app: Any, main: Any) -> None:
    """Tell AppKit which menu is Help so it injects the native search field."""
    help_item = main.itemWithTitle_("Help")
    if help_item is not None and help_item.submenu() is not None:
        app.setHelpMenu_(help_item.submenu())
