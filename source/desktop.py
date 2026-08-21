"""Native desktop window host for the clipgen web frontends.

Wraps :func:`server.serve_combined_app` in a pywebview window so a frozen bundle
opens like an app instead of spawning a terminal and hijacking the user's
browser. The frontend is unchanged: the server still listens on loopback, so
absolute nav links (``/studio/``), cross-blueprint fetches (``../studio/api/…``),
SSE, ndjson streaming and byte-range video all behave exactly as in a browser.
``http://127.0.0.1`` is a potentially-trustworthy origin in both WebKit and
Chromium, so ``navigator.clipboard`` keeps working too.

The engine is the OS webview — WKWebView on macOS, WebView2 on Windows — not a
bundled Chromium. See ``agents/ARCHITECTURE.md`` for the trade-offs that implies.

The window remembers its size and position across launches (persisted to
``start.json``, the per-user config file this module already keeps the webview
profile beside). Holding Shift while the app starts resets it to the default
rect — the standard escape hatch for a saved geometry that has become unusable.
"""

import ctypes
import importlib
import os
import sys
import threading
import webbrowser
from pathlib import Path
from typing import Any

import config
import desktop_chrome
import desktop_menu
import profiling
import server
import server_utils
import start_settings
import utils


# Matches --bg in the dark theme (assets/web/tokens.css), so the window does not
# flash white before the first paint.
_WINDOW_BACKGROUND = "#0a0a0b"
_WINDOW_SIZE = (1440, 900)
_WINDOW_MIN_SIZE = (960, 600)
# A restored window must expose at least this much of itself on some screen.
# Roughly the topnav's drag strip: enough to grab and pull back into view.
_MIN_VISIBLE = (200, 48)

# Rewrites target="_blank" clicks into a call back into Python. A webview has no
# back button, so letting one navigate the app window to GitHub would strand the
# user with no way home.
_EXTERNAL_LINK_SHIM = """
document.addEventListener('click', function (e) {
  var a = e.target && e.target.closest && e.target.closest('a[target="_blank"]');
  if (!a || !a.href) return;
  e.preventDefault();
  if (window.pywebview && window.pywebview.api && window.pywebview.api.open_external) {
    window.pywebview.api.open_external(a.href);
  }
});
"""


def open_external(url: str) -> None:
    """Open *url* in the user's real browser. Exposed to JS as ``open_external``."""
    if isinstance(url, str) and url.startswith(("http://", "https://")):
        webbrowser.open(url)


# Set once the window exists; save_file needs it to raise the native dialog.
_window: Any = None


def save_file(filename: str, content: str) -> str | None:
    """Write *content* to a user-chosen path. Exposed to JS as ``save_file``.

    The embedded webview cannot download anything on its own. WKWebView drops
    ``<a download>`` on the floor for ``blob:``, ``data:`` *and* HTTP responses
    carrying ``Content-Disposition: attachment`` — the request is made and the
    body discarded, with no error anywhere. Every export would look like a
    no-op. So the page hands the bytes back here and Python does the saving.

    Returns the path written, or None if the user cancelled.
    """
    import webview

    if _window is None:
        return None
    suffix = Path(filename).suffix.lstrip(".")
    file_types = (
        (f"{suffix.upper()} file (*.{suffix})", "All files (*.*)") if suffix else ()
    )
    result = _window.create_file_dialog(
        webview.FileDialog.SAVE,
        save_filename=filename,
        file_types=file_types,
    )
    if not result:
        return None
    # Documented as a sequence, but the save dialog returns a bare string on
    # some backends.
    target = Path(result if isinstance(result, str) else result[0])
    target.write_text(content, encoding="utf-8")
    return str(target)


def set_window_appearance(theme: str) -> None:
    """Follow the page theme. Exposed to JS as ``set_window_appearance``.

    Called on load and on every theme toggle; see
    ``desktop_chrome.set_appearance`` for why the window has to agree with the
    page rather than with the system.
    """
    if _window is not None and isinstance(theme, str):
        desktop_chrome.set_appearance(_window, theme)


def titlebar_double_click() -> None:
    """Zoom (or minimize) the window. Exposed to JS as ``titlebar_double_click``.

    The topnav stands in for the hidden title bar, so the page has to forward the
    gesture; AppKit never sees it. See ``desktop_chrome.titlebar_double_click``.
    """
    if _window is not None:
        desktop_chrome.titlebar_double_click(_window)


# --- Window geometry --------------------------------------------------------
#
# The last rect is persisted to start.json and restored on the next launch.
# Two things shape the implementation:
#
# - pywebview's moved/resized events carry the new values directly, each firing
#   on its own thread, per frame, for a whole drag — so the handlers only stash a
#   tuple and the disk write is debounced. Reading window.x/.width instead would
#   round-trip to the GUI thread on every one of those events.
# - x/y are reported and applied *relative to the window's current screen*
#   (cocoa.py's move() offsets from self.screen.origin, windowDidMove_ flips
#   against window.screen()), so a window dragged to a secondary display comes
#   back at the same offset on the primary one. _window_kwargs guarantees it is at
#   least visible; matching the exact display is not worth the per-backend screen
#   bookkeeping.

# Latched by _sample_reset_modifier: the user held Shift while the app started.
_reset_requested = False
# Last known rect as (x, y, width, height); None until the window is on screen.
_geometry: tuple[int, int, int, int] | None = None


def _shift_held() -> bool:
    """Whether Shift is physically down right now.

    Sampled with no running event loop — both calls below are HID queries, not
    event-stream reads, so they work before the GUI toolkit starts.
    """
    try:
        if sys.platform == "darwin":
            # Imported by name and typed Any for the two reasons in
            # desktop_chrome._appkit: pyobjc is absent on the Linux typecheck CI,
            # where a literal import is an unresolved-import error, and its stubs
            # omit half the framework (NSEvent included).
            appkit: Any = importlib.import_module("AppKit")
            flags = appkit.NSEvent.modifierFlags()
            return bool(flags & appkit.NSEventModifierFlagShift)
        if sys.platform == "win32":
            # VK_SHIFT. The high bit means "down at this moment".
            return bool(ctypes.windll.user32.GetAsyncKeyState(0x10) & 0x8000)
    except (ImportError, AttributeError, OSError):
        # Losing the reset gesture is survivable; failing to launch is not.
        pass
    return False


def _sample_reset_modifier() -> None:
    """Latch a Shift-at-launch request for the default window rect.

    Called twice, both *before* the window is created. A frozen bundle spends
    seconds on PyInstaller unpack, imports and the Flask boot before any of this
    runs, and nobody holds Shift that precisely; sampling again once the server
    is up widens the window in which the gesture counts. Sampling from
    before_show would be too late to affect the initial size.
    """
    global _reset_requested
    if _shift_held():
        _reset_requested = True


def _persist_geometry() -> None:
    """Write the tracked rect. Runs on the debounce timer thread."""
    rect = _geometry
    if rect is not None:
        start_settings.record_window_geometry(*rect)


# The lambda defers the _persist_geometry lookup to call time, per
# make_debounced_persist's contract, so tests can monkeypatch it. The cancel
# handle it returns has no caller here: every write path wants flush.
_schedule_geometry_persist, _flush_geometry, _ = server_utils.make_debounced_persist(
    lambda: _persist_geometry(),
    threading.Lock(),
    debounce_seconds=1.0,
)


def _on_shown(window: Any) -> None:
    """Seed the tracked rect from the window's real on-screen geometry.

    The moved/resized events each carry only their own axis, and the initial
    position may have been left to the OS, so this is the only way to learn the
    starting rect. If it cannot be read the rect stays None and nothing is ever
    persisted this session — better than saving a half-known frame.
    """
    global _geometry
    try:
        rect = (window.x, window.y, window.width, window.height)
    except RuntimeError:
        # get_position raises when the backend cannot resolve a screen.
        return
    if all(isinstance(value, (int, float)) for value in rect):
        _geometry = (int(rect[0]), int(rect[1]), int(rect[2]), int(rect[3]))


def _on_moved(x: float, y: float) -> None:
    global _geometry
    current = _geometry
    if current is None:
        return
    _geometry = (int(x), int(y), current[2], current[3])
    _schedule_geometry_persist()


def _on_resized(width: float, height: float) -> None:
    global _geometry
    current = _geometry
    if current is None:
        return
    _geometry = (current[0], current[1], int(width), int(height))
    _schedule_geometry_persist()


def _on_closing() -> None:
    """Flush the pending rect before the window goes away.

    Must not return False: closing is a locking event and Event.set treats a
    False from any handler as "cancel the close", which would leave the user
    with an unclosable window.
    """
    _flush_geometry()


def _is_visible_on(x: int, y: int, width: int, height: int, screens: list[Any]) -> bool:
    """Whether enough of the rect lands on some screen to be grabbable."""
    for screen in screens:
        left, top = int(screen.x), int(screen.y)
        right, bottom = left + int(screen.width), top + int(screen.height)
        if (
            min(x + width, right) - max(x, left) >= _MIN_VISIBLE[0]
            and min(y + height, bottom) - max(y, top) >= _MIN_VISIBLE[1]
        ):
            return True
    return False


def _window_kwargs(saved: dict[str, int] | None, screens: list[Any]) -> dict[str, int]:
    """Turn a persisted rect into create_window kwargs, or fall back to defaults.

    Size is clamped to the window's own minimum and to the largest screen. The
    position is only honoured when the resulting rect is actually reachable —
    dropping it (and letting the OS place the window) is the guard against a
    saved rect from a monitor that is no longer attached, which would otherwise
    open the app off-screen with no way to get it back.
    """
    if not saved:
        return {"width": _WINDOW_SIZE[0], "height": _WINDOW_SIZE[1]}
    width = max(saved["width"], _WINDOW_MIN_SIZE[0])
    height = max(saved["height"], _WINDOW_MIN_SIZE[1])
    if screens:
        width = min(width, max(int(screen.width) for screen in screens))
        height = min(height, max(int(screen.height) for screen in screens))
    kwargs = {"width": width, "height": height}
    if _is_visible_on(saved["x"], saved["y"], width, height, screens):
        kwargs["x"] = saved["x"]
        kwargs["y"] = saved["y"]
    return kwargs


def _restore_geometry() -> dict[str, int]:
    """The create_window geometry kwargs for this launch."""
    import webview

    if _reset_requested:
        # Cleared now rather than at quit so a crash still leaves defaults.
        start_settings.clear_window_geometry()
        return {"width": _WINDOW_SIZE[0], "height": _WINDOW_SIZE[1]}
    try:
        # A lazy proxy resolving the display list on access, which works before
        # webview.start(). It also initialises the GUI toolkit, so a headless box
        # raises here — swallow that and let create_window report the real problem,
        # which launch() turns into the browser fallback.
        screens = list(webview.screens)
    except (RuntimeError, OSError, AttributeError, webview.WebViewException):
        screens = []
    return _window_kwargs(start_settings.load_window_geometry(), screens)


def _hide_own_console() -> None:
    """Hide the console window on Windows, but only if this process owns it.

    A frozen Windows build stays a console app so ``clipgen.exe --gif …`` still
    prints (PyInstaller has no dual-mode build). Double-clicking from Explorer
    therefore allocates a console we want gone — but launching ``--desktop`` from
    an existing ``cmd.exe`` shares *the user's* console, and hiding that would be
    hostile. ``GetConsoleProcessList`` distinguishes the two: a console with only
    this process attached is ours to hide.
    """
    if os.name != "nt":
        return
    try:
        kernel32 = ctypes.windll.kernel32
        hwnd = kernel32.GetConsoleWindow()
        if not hwnd:
            return
        pids = (ctypes.c_uint * 2)()
        if kernel32.GetConsoleProcessList(pids, 2) > 1:
            return  # shared with the launching shell — leave it alone
        ctypes.windll.user32.ShowWindow(hwnd, 0)
    except (AttributeError, OSError):
        # Not fatal: a visible console is cosmetic, not a failure to launch.
        pass


def is_available() -> bool:
    """Report whether a desktop window can be opened in this environment."""
    try:
        import webview  # noqa: F401
    except ImportError:
        return False
    return True


def launch_desktop(
    worksheet: Any = None,
    default_page: str = "studio",
    gspread_client: Any = None,
    gspread_client_factory: Any = None,
) -> None:
    """Serve the combined app and show it in a native window until closed."""
    import webview

    _sample_reset_modifier()

    # Only clicks landing *directly* on a .pywebview-drag-region element drag the
    # window; without this a mousedown anywhere inside the topnav (a tab, the
    # settings button) would bubble up to the bar and drag it too.
    webview.settings["DRAG_REGION_DIRECT_TARGET_ONLY"] = True

    live = server.serve_combined_app(
        worksheet=worksheet,
        default_page=default_page,
        gspread_client=gspread_client,
        gspread_client_factory=gspread_client_factory,
    )
    profiling.mark("startup.server_bound")
    utils.info_print(f"clipgen running at {live.origin}")

    # Second sample: the server bind above returns fast (the heavy build runs
    # in the background behind the boot page), but sample again for safety.
    _sample_reset_modifier()

    try:
        # render_index_html() bakes this into every page it serves. Set inside the
        # try so the finally always clears it — a browser fallback launch must not
        # inherit a chrome flag. Nothing has requested a page yet: the server is up,
        # but the first GET only happens when the window loads.
        utils.DESKTOP_CHROME = desktop_chrome.chrome_style()
        geometry = _restore_geometry()
        # The cocoa/AppKit toolkit init (webview.screens inside
        # _restore_geometry) lands in this delta.
        profiling.mark("startup.gui_init")
        window = webview.create_window(
            f"clipgen v{utils.get_version()}",
            live.url,
            width=geometry["width"],
            height=geometry["height"],
            # None on both means "no saved position we trust" — pywebview then
            # centers the window, which is the pre-persistence behaviour.
            x=geometry.get("x"),
            y=geometry.get("y"),
            min_size=_WINDOW_MIN_SIZE,
            background_color=_WINDOW_BACKGROUND,
            # pywebview disables both by default. Transcripts is unusable without
            # text selection, and zoom matters for frame-level video work.
            text_select=True,
            zoomable=True,
        )
        assert window is not None  # create_window only returns None when hidden
        global _window
        _window = window
        window.expose(
            open_external, save_file, titlebar_double_click, set_window_appearance
        )
        window.events.loaded += lambda: window.run_js(_EXTERNAL_LINK_SHIM)
        # First shown-hook: the closest observable proxy for first paint.
        window.events.shown += lambda: profiling.mark("startup.window_shown")
        # before_show is a locking event, so this runs inline on AppKit's thread
        # after the backend finishes its own titlebar styling. The shown hook is
        # not optional: ordering the window front re-lays out the titlebar and
        # undoes the button placement.
        window.events.before_show += lambda: desktop_chrome.apply(window)
        window.events.shown += lambda: desktop_chrome.on_shown(window)
        window.events.loaded += lambda: desktop_chrome.reassert(window)
        window.events.shown += lambda: _on_shown(window)
        # After the chrome hooks: by shown-time first_show has already installed
        # the Tier 1 menu bar, and the polish pass queues itself onto the run
        # loop via callAfter, so it lands exactly once, after setMainMenu_.
        window.events.shown += lambda: desktop_menu.enhance_menu_bar(lambda: _window)
        window.events.moved += _on_moved
        window.events.resized += _on_resized
        window.events.closing += _on_closing

        # Defensive: a menu-construction failure must degrade to the default
        # bar, not throw — launch() demotes any launch_desktop exception to a
        # plain browser launch, which would cost the user the whole window.
        try:
            menus = desktop_menu.menus(lambda: _window)
        except Exception as exc:
            utils.warning_print(f"Could not build the menu bar: {exc}")
            menus = []

        _hide_own_console()
        profiling.mark("startup.webview_start")
        webview.start(
            menu=menus,
            debug=config.DEBUGGING,
            # pywebview defaults private_mode=True, discarding localStorage and
            # sessionStorage between runs — but the frontends keep real user state
            # there (settings toggles, viewer prefs, the Studio generate queue), so
            # persistence must be opted into explicitly.
            private_mode=False,
            storage_path=str(start_settings.config_dir() / "webview"),
        )
    finally:
        # closing already flushed on a normal quit; this covers the paths that
        # never fire it (a create_window/start failure, a forced teardown).
        _flush_geometry()
        globals()["_window"] = None
        globals()["_geometry"] = None
        globals()["_reset_requested"] = False
        utils.DESKTOP_CHROME = None
        desktop_chrome.teardown()
        server.stop_combined_app(live)


def launch(
    worksheet: Any = None,
    default_page: str = "studio",
    gspread_client: Any = None,
    gspread_client_factory: Any = None,
) -> None:
    """Open a desktop window, falling back to the browser if that is impossible."""
    if not is_available():
        utils.warning_print(
            "pywebview is unavailable — falling back to the default browser."
        )
        server.start_combined_server(
            worksheet=worksheet,
            default_page=default_page,
            gspread_client=gspread_client,
            gspread_client_factory=gspread_client_factory,
        )
        return
    try:
        launch_desktop(
            worksheet=worksheet,
            default_page=default_page,
            gspread_client=gspread_client,
            gspread_client_factory=gspread_client_factory,
        )
    except KeyboardInterrupt:
        pass
    except Exception as exc:
        # A headless box, a missing WebView2 runtime, or a GUI toolkit that
        # refuses to initialise. The work is still reachable over loopback, so
        # degrade to the browser rather than dying.
        utils.error_print(f"Could not open the desktop window: {exc}")
        utils.warning_print("Falling back to the default browser.")
        server.start_combined_server(
            worksheet=worksheet,
            default_page=default_page,
            gspread_client=gspread_client,
            gspread_client_factory=gspread_client_factory,
        )


if __name__ == "__main__":
    sys.exit(launch())
