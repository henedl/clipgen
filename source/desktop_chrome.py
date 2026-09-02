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
leaves them riding high in a 48px bar. ``_apply_titlebar_layout`` grows the
``NSTitlebarContainerView`` to the bar height and re-centers each button inside
it — the same approach Electron uses for ``trafficLightPosition``. AppKit resets
that layout on resize, on leaving fullscreen, and — the reason the machinery
below is larger than it looks like it needs to be — for a few hundred
milliseconds either side of macOS Sequoia laying its screen-sharing pill over
the light row. That transition fires no window-level notification, resets the
buttons *inside* a container that keeps our height, and makes its final move
after the last notification anyone can observe. So the views themselves are
watched rather than the window, and a bounded chain of deferred re-checks gets
the last word; without it the row stayed scrambled until the user resized.
Measure, do not reason: ``-v`` prints the real view shapes on every pass, and
three fixes written from plausible models of this hierarchy were all wrong.

The drag *region* is not defined here. ``-webkit-app-region: drag`` does nothing
in WKWebView, so the topnav opts into pywebview's ``.pywebview-drag-region``
mechanism instead (see ``assets/web/topnav.js``). The other half of the title-bar
contract does land here: that same surface reports double-clicks back through the
JS bridge, and ``titlebar_double_click`` performs whatever the user has configured
in System Settings → Desktop & Dock.

Every entry point is a no-op off macOS and degrades to the standard title bar on
any AppKit surprise — losing the styling must never cost the user their window.
"""

import importlib
import sys
import threading
import time
from typing import Any

import config
import utils

# NSNotificationCenter observer tokens, held so teardown can unregister them.
_observers: list[Any] = []
# Native fullscreen: AppKit owns the lights, so skip layout and the page inset.
_fullscreen = False
# Watched views and their observer tokens; re-bound when AppKit swaps views
# (see _bind_frame_observer).
_frame_observed: list[Any] = []
_frame_tokens: list[Any] = []
# Re-entrancy guard: our own setFrame_ posts the notification we listen for.
_laying_out = False
# Re-assert budget; see _within_reassert_budget. A user event re-arms it.
_REASSERT_BURST_LIMIT = 8
_REASSERT_WINDOW_S = 1.0
_reassert_count = 0
_reassert_started = 0.0
# Last verbose titlebar inventory, so a drag-resize does not print one per tick.
_last_inventory: str | None = None
# Deferred re-check chain; _schedule_settle explains why notifications alone fail.
_SETTLE_DELAY_S = 0.15
_SETTLE_ATTEMPTS = 5
_settle_timer: Any = None
# Last self-consistent button pitch; _button_pitch explains the lopsided
# mid-transition read.
_DEFAULT_PITCH = 20.0
_pitch = _DEFAULT_PITCH


def is_supported() -> bool:
    """Report whether this platform has a chrome style to apply."""
    return sys.platform == "darwin"


def chrome_style() -> str | None:
    """The value ``utils.DESKTOP_CHROME`` should take for this platform."""
    return "macos" if is_supported() else None


def _appkit() -> Any:
    """Import AppKit as an opaque module.

    Imported by name rather than with a plain ``import AppKit`` because pyobjc
    only exists on macOS, and CI type-checks on Linux — a literal import is an
    ``unresolved-import`` error there. Do not "simplify" it back.

    Typed as ``Any`` on purpose: pyobjc's stubs are incomplete (they omit
    ``NSNotificationCenter``, among others) and every call below is a
    dynamically-bridged ObjC selector, so checking against them buys nothing and
    costs a suppression at each site.
    """
    return importlib.import_module("AppKit")


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
        _apply_titlebar_layout(AppKit, native)
        _observe(AppKit, window, native)
    except Exception as exc:
        # Cosmetic. A window with an ordinary title bar is still a usable app.
        utils.warning_print(f"Could not style the window chrome: {exc}")
        return False
    return True


def teardown() -> None:
    """Unregister the notification observers. Safe to call more than once."""
    global _fullscreen, _laying_out, _settle_timer, _pitch
    global _reassert_count, _reassert_started, _last_inventory
    # Reset before the early-out: apply() lays out before _observe registers anything.
    _fullscreen = False
    _frame_observed.clear()
    _frame_tokens.clear()
    _laying_out = False
    _reassert_count = 0
    _reassert_started = 0.0
    _last_inventory = None
    _pitch = _DEFAULT_PITCH
    if _settle_timer is not None:
        _settle_timer.cancel()
        _settle_timer = None
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
        # Imported by name for the same reason as AppKit — see _appkit().
        app_helper: Any = importlib.import_module("PyObjCTools.AppHelper")
        _rearm_reassert_budget()
        app_helper.callAfter(lambda: _apply_titlebar_layout(AppKit, native))
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


def set_appearance(window: Any, theme: str) -> None:
    """Match *window*'s NSAppearance to the page's ``data-theme``.

    A window resize exposes area that AppKit fills with an appearance-derived
    colour for the one frame before WebKit repaints it. In Light appearance that
    fill is pure white, so zooming a dark-themed page flashes white — measured as
    exactly one frame of ``(255, 255, 255)``. Matching the appearance to the page
    makes the fill and the page agree, in both directions.

    Also has the side benefit of theming the webview's native furniture
    (scrollbars, form controls) to the app rather than to the system.
    """
    native = getattr(window, "native", None)
    if not is_supported() or native is None:
        return
    name = "NSAppearanceNameAqua" if theme == "light" else "NSAppearanceNameDarkAqua"
    try:
        AppKit = _appkit()
        # The constants are their own names; getattr is belt-and-braces.
        appearance = AppKit.NSAppearance.appearanceNamed_(getattr(AppKit, name, name))
        # Imported by name for the same reason as AppKit — see _appkit().
        app_helper: Any = importlib.import_module("PyObjCTools.AppHelper")

        def apply_appearance() -> None:
            native.setAppearance_(appearance)
            # Changing appearance re-lays out the titlebar and resets the lights.
            _apply_titlebar_layout(AppKit, native)

        app_helper.callAfter(apply_appearance)
    except Exception as exc:
        # Cosmetic: the worst case is the flash this exists to remove.
        utils.warning_print(f"Could not match the window appearance: {exc}")


def titlebar_double_click(window: Any) -> None:
    """Run the user's title-bar double-click action for *window*.

    Called from the page when a double-click lands on the drag region. The real
    title bar is hidden, so AppKit never sees the gesture and we have to perform
    it ourselves — reading the same preference AppKit would.
    """
    native = getattr(window, "native", None)
    if not is_supported() or native is None:
        return
    try:
        AppKit = _appkit()
        if _is_fullscreen(AppKit, native):
            return  # AppKit owns the window's size in fullscreen
        defaults = AppKit.NSUserDefaults.standardUserDefaults()
        action = _double_click_action(
            defaults.stringForKey_("AppleActionOnDoubleClick")
        )
        if action is None:
            return
        # Imported by name (see _appkit). callAfter hops off the bridge's worker thread.
        app_helper: Any = importlib.import_module("PyObjCTools.AppHelper")
        if action == "minimize":
            app_helper.callAfter(lambda: native.performMiniaturize_(None))
        else:
            app_helper.callAfter(lambda: native.performZoom_(None))
    except Exception as exc:
        # A gesture that does nothing is a cosmetic loss, not a broken window.
        utils.warning_print(f"Could not run the title-bar double-click action: {exc}")


# ---- Internals ----


def _double_click_action(setting: str | None) -> str | None:
    """Map ``AppleActionOnDoubleClick`` onto the action to perform, or None.

    Unset means macOS's own default, which is zoom. "Fill" (the tiling option
    added in Sequoia) has no public ``NSWindow`` equivalent — zoom is the closest
    honest approximation, and closer than doing nothing.
    """
    if setting == "Minimize":
        return "minimize"
    if setting == "None":
        return None
    return "zoom"


def _mask_bit(AppKit: Any, name: str, fallback: int) -> int:
    """Read a window-mask constant, falling back for older pyobjc builds."""
    return getattr(AppKit, name, fallback)


def _titlebar_container(AppKit: Any, native: Any) -> Any:
    """Resolve the window's ``NSTitlebarContainerView``, or ``None``.

    Deliberately does *not* go through ``standardWindowButton_`` first. The band
    is the one thing the page's layout depends on, and reaching it through
    ``buttons[0].superview().superview()`` made it hostage to a slot that macOS
    rearranges at will. The button route stays as the fallback because it is the
    one that has always worked; ``None`` rather than a guess, because the caller
    resizes what it is handed.
    """
    try:
        for view in native.contentView().superview().subviews():
            if str(view.className()) == "NSTitlebarContainerView":
                return view
        close = native.standardWindowButton_(AppKit.NSWindowCloseButton)
        return close.superview().superview() if close else None
    except Exception:
        return None


def _style_window(AppKit: Any, native: Any) -> None:
    full_size = _mask_bit(AppKit, "NSFullSizeContentViewWindowMask", 1 << 15)
    native.setStyleMask_(native.styleMask() | full_size)
    native.setTitlebarAppearsTransparent_(True)
    native.setTitleVisibility_(_mask_bit(AppKit, "NSWindowTitleHidden", 1))

    # pywebview paints the titlebar opaque. Guessing via lastObject() is safe here
    # only (cosmetic).
    try:
        titlebar = _titlebar_container(AppKit, native)
        if titlebar is None:
            titlebar = native.contentView().superview().subviews().lastObject()
        titlebar.setBackgroundColor_(AppKit.NSColor.clearColor())
    except (AttributeError, IndexError):
        pass


def _is_fullscreen(AppKit: Any, native: Any) -> bool:
    return bool(
        native.styleMask() & _mask_bit(AppKit, "NSFullScreenWindowMask", 1 << 14)
    )


def _centered_origin(band: float, size: float) -> int:
    """The offset that centers a *size*-long run inside a *band*-long one.

    Used for both axes: vertically it centers the control in the bar, and
    horizontally the same number becomes the left inset, so the cluster sits
    equidistant from the top and left edges of the window.
    """
    return round((band - size) / 2)


def _within_reassert_budget(now: float) -> bool:
    """Whether another re-assert is allowed, and count it.

    Nothing stops AppKit from pinning the container height and undoing our
    layout at the end of every pass. That would be an unbounded main-thread
    fight, so allow a short burst per window and then stand down until a
    user-driven event re-arms the counter. Worst case is a handful of futile
    passes and a stock-looking titlebar — cosmetic, never a hang.
    """
    global _reassert_count, _reassert_started
    if now - _reassert_started > _REASSERT_WINDOW_S:
        _reassert_started = now
        _reassert_count = 0
    _reassert_count += 1
    return _reassert_count <= _REASSERT_BURST_LIMIT


def _rearm_reassert_budget() -> None:
    """Forget the burst count. A user-driven event is never a layout fight."""
    global _reassert_count
    _reassert_count = 0


def _apply_titlebar_layout(AppKit: Any, native: Any) -> None:
    """Grow the titlebar band and place whatever macOS put in the light slot.

    The slot holds the three standard window buttons, which AppKit centers for a
    28px titlebar and which therefore ride high in our 48px one. While the screen
    is being shared, Sequoia lays an ``NSWindowSharingSessionRecipientIndicator``
    pill over them — a *sibling* of the three widgets, not a replacement: the
    buttons stay in the hierarchy, visible, and the pill simply covers them
    (measured on Sequoia: pill ``[17,14 52x20]`` against widgets at 16/36/56).
    So the pill is checked on every pass rather than as an alternative to the
    buttons — though AppKit centers it correctly by itself once the band is
    right, so that check is normally a no-op.

    The band is grown before the buttons are looked up, so a slot with no usable
    buttons still gets its height rather than being skipped entirely.

    AppKit views are unflipped, so a child's ``origin.y`` is measured from the
    container's bottom edge — which growing the band moves. Growth and placement
    have to happen in one pass, or the slot ends up worse off than untouched.
    """
    global _laying_out
    if _is_fullscreen(AppKit, native):
        return
    container = _titlebar_container(AppKit, native)
    if container is None:
        return

    bar = config.DESKTOP_CHROME_BAR_HEIGHT
    _laying_out = True
    try:
        frame = container.frame()
        frame.size.height = bar
        frame.origin.y = native.frame().size.height - bar
        container.setFrame_(frame)

        # AppKit's 28px reset never shrinks fillers back (measured +20px per pass).
        # Clamp overshoot only.
        for view in container.subviews():
            filler = view.frame()
            if filler.size.height > bar:
                filler.size.height = bar
                filler.origin.y = 0
                view.setFrame_(filler)

        buttons = _light_buttons(AppKit, native)
        if all(buttons):
            # Read the pitch first; _button_pitch ignores a lopsided mid-transition row.
            pitch = _button_pitch(buttons)
            for index, button in enumerate(buttons):
                # A reset can leave the group view short; center in the bar instead.
                parent_height = max(button.superview().frame().size.height, bar)
                margin = _centered_origin(parent_height, button.frame().size.height)
                origin = button.frame().origin
                origin.x = margin + index * pitch
                origin.y = margin
                button.setFrameOrigin_(origin)
        _place_sharing_pill(container, bar)
        _invalidate(container)
        _log_titlebar_inventory(container, buttons)
    finally:
        _laying_out = False


def _button_pitch(buttons: list[Any]) -> float:
    """The light row's x spacing, ignoring a mid-transition read.

    During a sharing transition AppKit walks the row through intermediate
    positions and briefly reports it lopsided — measured, close at 7 while the
    other two sat at 36 and 56, so gaps of 29 and 20 in a single read. Feeding
    that back through ``margin + index * pitch`` is what spread the buttons to
    16/45/74 and left them there. A row whose two gaps disagree is not a row
    worth measuring, so the last sane pitch is kept instead.
    """
    global _pitch
    first = buttons[1].frame().origin.x - buttons[0].frame().origin.x
    second = buttons[2].frame().origin.x - buttons[1].frame().origin.x
    if first > 0 and abs(first - second) <= 0.5:
        _pitch = first
    return _pitch


def _invalidate(view: Any) -> None:
    """Mark *view* and everything under it for redisplay.

    Moving a view leaves stale pixels in whatever drew behind it, and the
    titlebar is not redrawn until the window is next interacted with — which is
    how a re-placed sharing pill came to be on screen twice at once. Recursive
    on purpose: the pill and the vibrancy view that backs it are *grand*children
    of the container, so marking only its direct subviews missed both of the
    views that actually draw.
    """
    view.setNeedsDisplay_(True)
    for child in view.subviews():
        _invalidate(child)


def _place_sharing_pill(container: Any, bar: int) -> None:
    """Correct Sequoia's screen-sharing pill if it is not centered in the band.

    Measured on Sequoia the pill is an ``NSWindowSharingSessionRecipientIndicator``
    at ``[17,14 52x20]``, sitting alongside the three ``_NSTheme*Widget``s one
    level below the container's own children (the full-height ``NSTitlebarView``
    and ``_NSTitlebarDecorationView``) — hence the single level of descent.

    A correction, not a placement. Once the band is 48px AppKit centers the pill
    itself — ``y = 14`` is exactly ``(48 - 20) / 2`` — so the normal path must
    touch nothing. Nudging it the 1px from AppKit's ``x = 17`` to the close
    button's ``x = 16`` bought nothing and left a ghost of the pill's backdrop
    at the old position until the window was interacted with. Hence: vertical
    only, and only when AppKit has actually got it wrong.

    Matched by class name, not by shape. Shape matching — "short, visible, inside
    the left gutter" — was tried and reverted: the titlebar is full of short,
    left-aligned background views (``NSVisualEffectView``, the light group
    itself), and moving those drags the buttons off with them and leaves the bar
    growing by 20px a pass. If Apple renames the class this quietly stops
    running, which is the safe direction to fail in; ``_log_titlebar_inventory``
    prints the real name at ``-v``.

    Guarded separately from the growth above: a surprise here must not cost us
    the band height, which is the half that fixes the buttons.
    """
    try:
        candidates: list[Any] = []
        for view in container.subviews():
            candidates.append(view)
            candidates.extend(view.subviews())

        for view in candidates:
            if "SharingSession" not in str(view.className()):
                continue
            origin = view.frame().origin
            centered = _centered_origin(bar, view.frame().size.height)
            if abs(origin.y - centered) <= 1:
                continue
            origin.y = centered
            view.setFrameOrigin_(origin)
    except Exception as exc:
        utils.verbose_print(f"Could not place the screen-sharing indicator: {exc}")


def _log_titlebar_inventory(container: Any, buttons: list[Any]) -> None:
    """Print the titlebar's subview shapes at ``-v``, once per change.

    The sharing pill's view class is undocumented and cannot be reproduced
    headlessly, so this is how a real share gets diagnosed. Descends one level,
    matching ``_place_sharing_pill``'s candidate walk, so a log that shows no
    pill means the pill really is somewhere neither of them looks. Deduped
    because the frame observer fires on every tick of a drag-resize.
    """
    global _last_inventory

    def shape(view: Any) -> str:
        frame = view.frame()
        return (
            f"{view.className()}[{frame.origin.x:.0f},{frame.origin.y:.0f} "
            f"{frame.size.width:.0f}x{frame.size.height:.0f}"
            f"{' hidden' if view.isHidden() else ''}]"
        )

    if getattr(config, "VERBOSITY", config.STANDARD) < config.VERBOSE:
        return
    try:
        states = ",".join(
            "nil" if b is None else ("hidden" if b.isHidden() else "shown")
            for b in buttons
        )
        # AppKit reorders children between passes; sort so the dedupe sees stable lines.
        views = "; ".join(
            " > ".join(
                [shape(view)]
                + [
                    shape(child)
                    for child in sorted(
                        view.subviews(), key=lambda c: c.frame().origin.x
                    )
                ]
            )
            for view in container.subviews()
        )
        line = f"titlebar {shape(container)} buttons={states} | {views}"
    except Exception:
        return
    if line != _last_inventory:
        _last_inventory = line
        utils.verbose_print(line)


def _slot_views(AppKit: Any, native: Any) -> list[Any]:
    """The views whose geometry the layout pass owns, outermost first.

    The container, the group view the lights sit in, and the close button. All
    three are watched and all three are checked, because a screen-sharing
    transition resets them independently: measured, it puts the buttons back at
    their stock positions inside a container that still has our height, so
    watching the container alone sees nothing at all.
    """
    container = _titlebar_container(AppKit, native)
    if container is None:
        return []
    views = [container]
    close = _light_buttons(AppKit, native)[0]
    if close is not None:
        group = close.superview()
        if group is not None:
            views.append(group)
        views.append(close)
    return views


def _light_buttons(AppKit: Any, native: Any) -> list[Any]:
    """The three standard window buttons, in row order. Entries may be ``None``."""
    return [
        native.standardWindowButton_(getattr(AppKit, name))
        for name in (
            "NSWindowCloseButton",
            "NSWindowMiniaturizeButton",
            "NSWindowZoomButton",
        )
    ]


def _layout_is_current(AppKit: Any, native: Any, bar: int) -> bool:
    """Whether the band and the whole light row are already where we put them.

    The notification handler's early-out and the deferred chain's stop
    condition, so it has to test exactly what the pass writes — all three
    buttons, not just the close one. A transition scrambles the row as a whole
    (measured: 16/45/74 while close alone looked plausible), and a predicate
    that only looked at close would call that settled. Frames are CGFloats, so
    compare with a half-pixel tolerance rather than for equality.

    The group view's own height is deliberately not part of the test: the buttons
    are anchored to its bottom edge, so it can sit at AppKit's 28px without
    moving anything the user can see, and demanding 48 there would spend the
    whole budget re-asserting a difference with no visual consequence.
    """
    container = _titlebar_container(AppKit, native)
    if container is None:
        return True  # nothing to place; a pass would not change that
    if abs(container.frame().size.height - bar) > 0.5:
        return False
    buttons = _light_buttons(AppKit, native)
    if not all(buttons):
        return True
    parent_height = max(buttons[0].superview().frame().size.height, bar)
    margin = _centered_origin(parent_height, buttons[0].frame().size.height)
    for index, button in enumerate(buttons):
        origin = button.frame().origin
        if abs(origin.y - margin) > 0.5:
            return False
        if abs(origin.x - (margin + index * _pitch)) > 0.5:
            return False
    return True


def _schedule_settle(AppKit: Any, native: Any, attempts: int) -> None:
    """Arm a deferred re-check of the layout.

    Notifications alone are not enough to finish a sharing transition. AppKit
    walks the slot through several intermediate layouts and its *last* move
    lands after our final pass without posting anything we observe, so the row
    sat at 7/45/74 until the user resized the window. A short chain of timed
    re-checks gives us the last word instead; each link stops as soon as the
    layout reads correct, so a quiet window costs one no-op check.
    """
    global _settle_timer
    if attempts <= 0:
        return
    if _settle_timer is not None and _settle_timer.is_alive():
        return  # one pending check is enough; it re-arms itself if it acts

    def fire() -> None:
        try:
            # Imported by name (see _appkit). callAfter hops off the timer thread.
            app_helper: Any = importlib.import_module("PyObjCTools.AppHelper")
            app_helper.callAfter(lambda: _settle(AppKit, native, attempts))
        except Exception as exc:
            utils.verbose_print(f"Could not re-check the titlebar layout: {exc}")

    _settle_timer = threading.Timer(_SETTLE_DELAY_S, fire)
    _settle_timer.daemon = True
    _settle_timer.start()


def _settle(AppKit: Any, native: Any, attempts: int) -> None:
    """One link of the deferred chain: fix the layout, or draw and stop."""
    if _fullscreen:
        return
    if not _layout_is_current(AppKit, native, config.DESKTOP_CHROME_BAR_HEIGHT):
        _apply_titlebar_layout(AppKit, native)
        _schedule_settle(AppKit, native, attempts - 1)
        return
    # Settled. setNeedsDisplay alone leaves a ghost pill (AppKit moves it 17→8→17);
    # force a redraw.
    container = _titlebar_container(AppKit, native)
    if container is not None:
        _invalidate(container)
        container.displayIfNeeded()


def _bind_frame_observer(AppKit: Any, native: Any, handler: Any) -> None:
    """Bind (or re-bind) the frame observers, following the views' identity.

    A view's own frame change is the only signal a screen-sharing transition
    reliably produces — no window-level notification fires for it. AppKit can
    also hand the window *different* views across a transition, which would leave
    the observers bound to the old ones permanently deaf, so the binding is
    re-checked rather than made once.
    """
    views = _slot_views(AppKit, native)
    if not views:
        return
    if len(views) == len(_frame_observed) and all(
        a is b for a, b in zip(views, _frame_observed)
    ):
        return
    center = AppKit.NSNotificationCenter.defaultCenter()
    for token in _frame_tokens:
        center.removeObserver_(token)
        for index, existing in enumerate(_observers):
            if existing is token:
                del _observers[index]
                break
    _frame_tokens.clear()
    _frame_observed.clear()
    for view in views:
        # NSView default, set explicitly: this hook is the only sharing-transition signal.
        view.setPostsFrameChangedNotifications_(True)
        token = center.addObserverForName_object_queue_usingBlock_(
            AppKit.NSViewFrameDidChangeNotification, view, None, handler
        )
        _frame_tokens.append(token)
        _observers.append(token)
        _frame_observed.append(view)


def _observe(AppKit: Any, window: Any, native: Any) -> None:
    """Re-apply the button layout whenever AppKit resets it."""
    center = AppKit.NSNotificationCenter.defaultCenter()

    def on_frame_change(_note: Any) -> None:
        # Our own setFrame_/setFrameOrigin_ post this synchronously.
        if _laying_out:
            return
        # Width-only resize ticks are not a reset; filter before charging the budget.
        if _layout_is_current(AppKit, native, config.DESKTOP_CHROME_BAR_HEIGHT):
            return
        if not _within_reassert_budget(time.monotonic()):
            if _reassert_count == _REASSERT_BURST_LIMIT + 1:
                utils.verbose_print(
                    "AppKit keeps resetting the titlebar; gave up re-asserting it."
                )
            return
        _apply_titlebar_layout(AppKit, native)
        # A transition can swap the views; re-bind so the next reset is heard.
        _bind_frame_observer(AppKit, native, on_frame_change)
        # AppKit moves the slot again after this, silently; the settle chain finishes.
        _schedule_settle(AppKit, native, _SETTLE_ATTEMPTS)

    def on_become_key(_note: Any) -> None:
        _rearm_reassert_budget()
        _bind_frame_observer(AppKit, native, on_frame_change)
        _apply_titlebar_layout(AppKit, native)
        _schedule_settle(AppKit, native, _SETTLE_ATTEMPTS)

    def on_resize(_note: Any) -> None:
        _rearm_reassert_budget()
        _apply_titlebar_layout(AppKit, native)

    def on_enter_fullscreen(_note: Any) -> None:
        global _fullscreen
        _fullscreen = True
        _set_page_fullscreen(window, True)

    def on_exit_fullscreen(_note: Any) -> None:
        global _fullscreen
        _fullscreen = False
        _rearm_reassert_budget()
        _apply_titlebar_layout(AppKit, native)
        _set_page_fullscreen(window, False)

    for name, handler in (
        ("NSWindowDidResizeNotification", on_resize),
        ("NSWindowDidEnterFullScreenNotification", on_enter_fullscreen),
        ("NSWindowDidExitFullScreenNotification", on_exit_fullscreen),
        # Safety net: tabbing back into the window fixes a swap that kept the height.
        ("NSWindowDidBecomeKeyNotification", on_become_key),
    ):
        _observers.append(
            center.addObserverForName_object_queue_usingBlock_(
                getattr(AppKit, name), native, None, handler
            )
        )
    _bind_frame_observer(AppKit, native, on_frame_change)


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
