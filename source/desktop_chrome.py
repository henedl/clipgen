"""macOS window chrome for the desktop shell: hide the title bar, keep the lights.

The native title bar sits directly on top of clipgen's own 48px topnav, so the
desktop window spends ~28px on a second, redundant bar. This module removes it
*without* going frameless: the style mask gains
``NSFullSizeContentViewWindowMask`` and the titlebar is made transparent and
untitled, so web content runs to the top of the window while the traffic lights,
resize handles, snapping and green-button fullscreen all stay native.

pywebview's own ``frameless=True`` is the wrong tool here — it hides all three
standard buttons (``webview/platforms/cocoa.py``), which is exactly what we want
to keep. It does hand us everything else we need: ``window.native`` is the real
``NSWindow``, and ``events.before_show`` fires on the main thread *after* the
backend has finished its own titlebar styling, so our overrides win.

The buttons then have to move: AppKit centers them in a 28px titlebar, which
leaves them riding high in a 48px bar. ``_reposition_traffic_lights`` grows the
``NSTitlebarContainerView`` to the bar height and re-centers each button inside
it — the same approach Electron uses for ``trafficLightPosition``. AppKit resets
that layout on resize and on leaving fullscreen, hence the observers.

Dragging is *not* handled here. ``-webkit-app-region: drag`` does nothing in
WKWebView, so the topnav opts into pywebview's ``.pywebview-drag-region``
mechanism instead (see ``assets/web/topnav.js``).

Every entry point is a no-op off macOS and degrades to the standard title bar on
any AppKit surprise — losing the styling must never cost the user their window.
"""

import sys
import threading
from typing import Any

import config
import utils

# NSNotificationCenter observer tokens, held so teardown can unregister them.
_observers: list[Any] = []
# True while the window is in native fullscreen: AppKit owns the traffic lights
# then (they live in the auto-hiding overlay), so we neither move them nor ask
# the page to reserve room for them.
_fullscreen = False


def is_supported() -> bool:
    """Report whether this platform has a chrome style to apply."""
    return sys.platform == "darwin"


def chrome_style() -> str | None:
    """The value ``utils.DESKTOP_CHROME`` should take for this platform."""
    return "macos" if is_supported() else None


def _appkit() -> Any:
    """Import AppKit as an opaque module.

    Typed as ``Any`` on purpose: pyobjc's stubs are incomplete (they omit
    ``NSNotificationCenter``, among others) and every call below is a
    dynamically-bridged ObjC selector, so checking against them buys nothing and
    costs a suppression at each site.
    """
    import AppKit

    return AppKit


def apply(window: Any) -> bool:
    """Style *window*'s native chrome. Returns whether anything was applied.

    Must be called from the main thread — ``events.before_show`` is a locking
    event, so its handlers run inline on AppKit's thread, which is where all of
    the calls below have to happen anyway.
    """
    if not is_supported():
        return False
    try:
        AppKit = _appkit()
    except ImportError:  # pragma: no cover - pyobjc ships with pywebview on macOS
        utils.warning_print("AppKit unavailable — keeping the standard title bar.")
        return False

    native = getattr(window, "native", None)
    if native is None:
        utils.warning_print("No native window handle — keeping the standard title bar.")
        return False

    try:
        _style_window(AppKit, native)
        _reposition_traffic_lights(AppKit, native)
        _observe(AppKit, window, native)
    except Exception as exc:
        # Cosmetic. A window with an ordinary title bar is still a usable app.
        utils.warning_print(f"Could not style the window chrome: {exc}")
        return False
    return True


def teardown() -> None:
    """Unregister the notification observers. Safe to call more than once."""
    global _fullscreen
    _fullscreen = False
    if not _observers:
        return
    try:
        center = _appkit().NSNotificationCenter.defaultCenter()
        for token in _observers:
            center.removeObserver_(token)
    except Exception as exc:
        # The process is tearing down anyway; a leaked observer costs nothing.
        utils.verbose_print(f"Could not remove window-chrome observers: {exc}")
    _observers.clear()


def on_shown(window: Any) -> None:
    """Re-apply the button layout once the window is actually on screen.

    Ordering the window front runs a layout pass that snaps the titlebar
    container back to its stock 28px, undoing what ``apply()`` did at
    ``before_show``. Re-applying here sticks — measured: it survives
    deactivate/reactivate, and a resize is the only other thing that resets it
    (which the resize observer catches).

    ``events.shown`` is not a locking event, so this arrives on a worker thread
    and has to hop back to AppKit's.
    """
    native = getattr(window, "native", None)
    if not is_supported() or native is None:
        return
    try:
        AppKit = _appkit()
        from PyObjCTools import AppHelper

        AppHelper.callAfter(lambda: _reposition_traffic_lights(AppKit, native))
    except Exception as exc:
        utils.warning_print(f"Could not place the window buttons: {exc}")


def reassert(window: Any) -> None:
    """Re-push page-side chrome state after a navigation.

    ``document.documentElement.dataset`` is per-document, so a page load drops
    the fullscreen flag. The ``data-desktop-chrome`` attribute itself survives —
    the server writes it into every rendered page.
    """
    if _fullscreen:
        _set_page_fullscreen(window, True)


# ---- Internals ----


def _mask_bit(AppKit: Any, name: str, fallback: int) -> int:
    """Read a window-mask constant, falling back for older pyobjc builds."""
    return getattr(AppKit, name, fallback)


def _style_window(AppKit: Any, native: Any) -> None:
    full_size = _mask_bit(AppKit, "NSFullSizeContentViewWindowMask", 1 << 15)
    native.setStyleMask_(native.styleMask() | full_size)
    native.setTitlebarAppearsTransparent_(True)
    native.setTitleVisibility_(_mask_bit(AppKit, "NSWindowTitleHidden", 1))

    # pywebview paints the titlebar windowBackgroundColor for non-frameless
    # windows, which would show as an opaque strip across the top of the nav.
    try:
        titlebar = native.contentView().superview().subviews().lastObject()
        titlebar.setBackgroundColor_(AppKit.NSColor.clearColor())
    except (AttributeError, IndexError):
        pass


def _is_fullscreen(AppKit: Any, native: Any) -> bool:
    return bool(
        native.styleMask() & _mask_bit(AppKit, "NSFullScreenWindowMask", 1 << 14)
    )


def _reposition_traffic_lights(AppKit: Any, native: Any) -> None:
    """Center the three buttons in a ``DESKTOP_CHROME_BAR_HEIGHT`` band.

    Grows the titlebar container to the bar height first, so the buttons have
    somewhere to be centered *in* — moving them alone would push them past the
    28px container AppKit gives us.
    """
    if _is_fullscreen(AppKit, native):
        return
    buttons = [
        native.standardWindowButton_(getattr(AppKit, name))
        for name in (
            "NSWindowCloseButton",
            "NSWindowMiniaturizeButton",
            "NSWindowZoomButton",
        )
    ]
    if not all(buttons):
        return

    bar = config.DESKTOP_CHROME_BAR_HEIGHT
    container = buttons[0].superview().superview()
    frame = container.frame()
    frame.size.height = bar
    frame.origin.y = native.frame().size.height - bar
    container.setFrame_(frame)

    # Native x positions are tuned for a 28px titlebar and leave the row hugging
    # the window edge in a 48px one. Inset it by the same margin the centering
    # produces above and below, so the cluster sits equidistant from the top and
    # left edges. Read the pitch before moving anything — it is 20px either way,
    # which keeps repeat calls (resize, fullscreen exit) idempotent.
    pitch = buttons[1].frame().origin.x - buttons[0].frame().origin.x
    for index, button in enumerate(buttons):
        parent_height = button.superview().frame().size.height
        margin = round((parent_height - button.frame().size.height) / 2)
        origin = button.frame().origin
        origin.x = margin + index * pitch
        origin.y = margin
        button.setFrameOrigin_(origin)


def _observe(AppKit: Any, window: Any, native: Any) -> None:
    """Re-apply the button layout whenever AppKit resets it."""
    center = AppKit.NSNotificationCenter.defaultCenter()

    def on_resize(_note: Any) -> None:
        _reposition_traffic_lights(AppKit, native)

    def on_enter_fullscreen(_note: Any) -> None:
        global _fullscreen
        _fullscreen = True
        _set_page_fullscreen(window, True)

    def on_exit_fullscreen(_note: Any) -> None:
        global _fullscreen
        _fullscreen = False
        _reposition_traffic_lights(AppKit, native)
        _set_page_fullscreen(window, False)

    for name, handler in (
        ("NSWindowDidResizeNotification", on_resize),
        ("NSWindowDidEnterFullScreenNotification", on_enter_fullscreen),
        ("NSWindowDidExitFullScreenNotification", on_exit_fullscreen),
    ):
        _observers.append(
            center.addObserverForName_object_queue_usingBlock_(
                getattr(AppKit, name), native, None, handler
            )
        )


def _set_page_fullscreen(window: Any, on: bool) -> None:
    """Tell the page to drop (or restore) the traffic-light inset.

    Dispatched onto a worker thread on purpose: ``run_js`` posts to the main
    thread and blocks on a semaphore until the result comes back, so calling it
    from a notification handler — which already runs on the main thread — would
    deadlock.
    """
    script = (
        "document.documentElement.dataset.desktopFullscreen = '1';"
        if on
        else "delete document.documentElement.dataset.desktopFullscreen;"
    )

    def run() -> None:
        try:
            window.run_js(script)
        except Exception as exc:
            # The window may be closing, or no document loaded yet. Cosmetic.
            utils.verbose_print(f"Could not push the fullscreen chrome flag: {exc}")

    threading.Thread(target=run, daemon=True).start()
