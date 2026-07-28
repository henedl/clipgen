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
"""

import ctypes
import os
import sys
import webbrowser
from pathlib import Path
from typing import Any

import config
import server
import start_settings
import utils


# Matches --bg in the dark theme (assets/web/tokens.css), so the window does not
# flash white before the first paint.
_WINDOW_BACKGROUND = "#0a0a0b"
_WINDOW_SIZE = (1440, 900)
_WINDOW_MIN_SIZE = (960, 600)

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
) -> None:
    """Serve the combined app and show it in a native window until closed."""
    import webview

    live = server.serve_combined_app(
        worksheet=worksheet,
        default_page=default_page,
        gspread_client=gspread_client,
    )
    utils.info_print(f"clipgen running at {live.origin}")

    try:
        window = webview.create_window(
            f"clipgen v{utils.get_version()}",
            live.url,
            width=_WINDOW_SIZE[0],
            height=_WINDOW_SIZE[1],
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
        window.expose(open_external, save_file)
        window.events.loaded += lambda: window.run_js(_EXTERNAL_LINK_SHIM)

        _hide_own_console()
        webview.start(
            debug=config.DEBUGGING,
            # pywebview defaults private_mode=True, which discards localStorage
            # and sessionStorage between runs. The frontends keep real user state
            # there (settings toggles, viewer prefs, the Studio generate queue),
            # so persistence has to be opted into explicitly.
            private_mode=False,
            storage_path=str(start_settings.config_dir() / "webview"),
        )
    finally:
        globals()["_window"] = None
        server.stop_combined_app(live)


def launch(
    worksheet: Any = None,
    default_page: str = "studio",
    gspread_client: Any = None,
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
        )
        return
    try:
        launch_desktop(
            worksheet=worksheet,
            default_page=default_page,
            gspread_client=gspread_client,
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
        )


if __name__ == "__main__":
    sys.exit(launch())
