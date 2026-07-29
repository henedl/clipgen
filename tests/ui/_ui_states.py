"""Drive a page into a named UI state — the overlays and tabs the smoke never opens.

The smoke loads six pages in their boot state and clicks nothing, so most of the
frontend has never been rendered by anything automated: every modal, every tab
past the default, and (since both boot paths pre-dismiss it) the entire Start
overlay. This module is the missing half. It is a library, used by ``shot.py``
for screenshots, by the smoke for a boot-error check, and by the dead-CSS census
to reach rules that only match once something is open.

Two kinds of state, entered two different ways, for a reason each:

**Overlays — entered through the page's own public opener.** Every one of these
is already exposed on ``window`` for the command palette to call, so there is no
click choreography to reverse-engineer and nothing to keep in sync with a
toolbar's markup. They close with Escape, which every page routes through the
hotkey registry's cascade (see AGENTS.md).

**Tabs — discovered at runtime, then clicked.** Deliberately *not* a hard-coded
list. ``_ui_pages.py``'s docstring calls its per-page selector table "the part of
this harness most likely to rot", and a static list of tab names would rot the
same way but worse: a renamed tab turns into a confusing timeout, and a *new* tab
is silently never covered. Instead each page declares where its tabs live and the
state list is read off the live DOM, so adding a tab to the HTML adds a state
here for free. Clicking is also the contract a user has — calling the page's
internal switch function would bypass exactly the wiring worth testing.

A state that cannot be reached is *reported*, never skipped silently: three of
Studio's four sub-tabs are ``hidden`` in the static HTML and only appear once
intake data exists, and "we covered 1 of 4" has to be visible in the output or a
future reader will believe all four were checked.
"""

from collections.abc import Iterator
from dataclasses import dataclass
from typing import Any

import _ui_browser

# How long to wait for a state to become visible. Generous: these are local
# fixtures, so a miss is a real failure rather than a slow machine.
_STATE_TIMEOUT_MS = 5_000
# Overlays animate in (motion.js). Let the entrance settle before capture.
_SETTLE_MS = 350


@dataclass(frozen=True)
class Overlay:
    """A modal reachable from every page through a ``window`` function."""

    name: str
    # Guarded so a page that somehow lacks the opener reports "no opener"
    # instead of throwing an opaque TypeError from page.evaluate.
    probe: str
    enter: str
    ready: str


GLOBAL_OVERLAYS: tuple[Overlay, ...] = (
    Overlay(
        name="settings",
        probe="typeof window.openSettingsModal === 'function'",
        enter="window.openSettingsModal({ initialTab: 'General' })",
        ready=".settings-overlay:not(.hidden)",
    ),
    Overlay(
        name="settings-hotkeys",
        probe="typeof window.openSettingsModal === 'function'",
        enter="window.openSettingsModal({ initialTab: 'Hotkeys' })",
        ready=".settings-overlay:not(.hidden)",
    ),
    Overlay(
        name="cheatsheet",
        probe="!!(window.ClipgenHotkeys && window.ClipgenHotkeys.toggleCheatsheet)",
        enter="window.ClipgenHotkeys.toggleCheatsheet()",
        ready=".hk-overlay:not(.hidden)",
    ),
    Overlay(
        name="palette",
        probe="!!(window.ClipgenCommandPalette && window.ClipgenCommandPalette.open)",
        enter="window.ClipgenCommandPalette.open()",
        ready=".cmdp-overlay:not(.hidden)",
    ),
    Overlay(
        name="start",
        probe="!!(window.ClipgenStartOverlay && window.ClipgenStartOverlay.open)",
        enter="window.ClipgenStartOverlay.open()",
        ready="#startOverlay:not(.hidden)",
    ),
)


@dataclass(frozen=True)
class TabGroup:
    """Where one page keeps a row of tabs, and what names its states."""

    prefix: str
    selector: str
    attr: str


# Workflows is absent on purpose: it has no tab row.
PAGE_TABS: dict[str, tuple[TabGroup, ...]] = {
    "studio": (TabGroup("tab", ".preview-tab", "data-tab"),),
    "overview": (TabGroup("tab", ".preview-tab", "data-tab"),),
    "transcripts": (TabGroup("tab", ".panel-tab", "data-tab"),),
    "composer": (TabGroup("tab", ".co-panel-tab", "data-tab"),),
    "screenspace": (
        TabGroup("tool", ".wf-tab", "data-type"),
        TabGroup("pane", ".rp-tab", "data-tab"),
    ),
}


@dataclass
class StateResult:
    """What happened when we tried to reach one state."""

    name: str
    reached: bool
    detail: str = ""


@dataclass(frozen=True)
class TabState:
    """One tab button found in the live DOM."""

    name: str
    selector: str
    visible: bool


def state_names(page_name: str) -> list[str]:
    """The states this page *declares*, without booting anything.

    Tab states are discovered live, so this is only the overlay list plus a
    ``tool:``/``tab:`` hint — enough for ``--help`` and argument validation.
    """
    return [overlay.name for overlay in GLOBAL_OVERLAYS] + [
        f"{group.prefix}:<name>" for group in PAGE_TABS.get(page_name, ())
    ]


def discover_tabs(page: Any, page_name: str) -> list[TabState]:
    """Read this page's tab states off the live DOM.

    Records visibility rather than filtering on it. A tab can be in the markup
    but not on screen for two very different reasons, and only the caller can
    tell them apart: Studio's intake tabs unhide once intake data exists, while
    Screenspace's 13 ``.wf-tab``s are hidden *by design* whenever the grouped
    category nav is on — that mode keeps them in the DOM and drives them with
    ``.click()``. See :func:`enter_tab`.
    """
    found: list[TabState] = []
    for group in PAGE_TABS.get(page_name, ()):
        rows = page.evaluate(
            """(sel) => Array.from(document.querySelectorAll(sel.selector)).map(
                (el) => ({
                    key: el.getAttribute(sel.attr),
                    visible: !!(el.offsetParent || el.getClientRects().length),
                })
            )""",
            {"selector": group.selector, "attr": group.attr},
        )
        for row in rows:
            key = row.get("key")
            if not key:
                continue
            found.append(
                TabState(
                    name=f"{group.prefix}:{key}",
                    selector=f'{group.selector}[{group.attr}="{key}"]',
                    visible=bool(row.get("visible")),
                )
            )
    return found


def _wait(page: Any, selector: str) -> str:
    """Wait for ``selector``; return "" on success or the timeout's first line."""
    try:
        page.wait_for_selector(selector, state="visible", timeout=_STATE_TIMEOUT_MS)
    except _ui_browser.playwright_error() as exc:
        return str(exc).splitlines()[0]
    return ""


def enter_overlay(page: Any, overlay: Overlay) -> StateResult:
    """Open one overlay through its public opener, leaving it open."""
    if not page.evaluate(f"() => {overlay.probe}"):
        return StateResult(overlay.name, False, "page exposes no opener")
    try:
        page.evaluate(f"() => {{ {overlay.enter}; }}")
    except _ui_browser.playwright_error() as exc:
        return StateResult(overlay.name, False, str(exc).splitlines()[0])
    problem = _wait(page, overlay.ready)
    if problem:
        return StateResult(overlay.name, False, problem)
    page.wait_for_timeout(_SETTLE_MS)
    return StateResult(overlay.name, True)


def close_overlay(page: Any, overlay: Overlay) -> None:
    """Dismiss via Escape, falling back to a reload.

    Escape is the documented cascade on every page, but it is also the one thing
    a page is allowed to intercept for its own purposes. If the overlay is still
    up afterwards the next state would be captured through it, so reload rather
    than trust the keystroke.
    """
    page.keyboard.press("Escape")
    page.wait_for_timeout(_SETTLE_MS)
    if page.locator(overlay.ready).count():
        page.reload(wait_until="domcontentloaded")
        page.wait_for_timeout(_SETTLE_MS)


def enter_tab(page: Any, tab: TabState) -> StateResult:
    """Activate one discovered tab, by real click where that is possible.

    A visible tab gets a Playwright click — the actual user gesture, including
    its actionability checks. A tab that is present but hidden falls back to a
    DOM ``.click()``, which is not a shortcut around the UI so much as the same
    path the UI itself takes: Screenspace's grouped category nav hides the flat
    ``.wf-tab`` row and drives it by delegating to each tab's ``.click()``
    (screenspace.css:1401-1410). Reaching those 13 tool panels is what puts the
    largest block of ``screenspace.css`` in front of the census at all.

    The distinction is recorded in ``detail`` rather than smoothed over, because
    "clicked" and "dispatched a click at a hidden node" are not the same claim.
    """
    if tab.visible:
        try:
            page.locator(tab.selector).first.click(timeout=_STATE_TIMEOUT_MS)
        except _ui_browser.playwright_error() as exc:
            return StateResult(tab.name, False, str(exc).splitlines()[0])
        page.wait_for_timeout(_SETTLE_MS)
        return StateResult(tab.name, True)

    clicked = page.evaluate(
        """(selector) => {
            const el = document.querySelector(selector);
            if (!el) return false;
            el.click();
            return true;
        }""",
        tab.selector,
    )
    if not clicked:
        return StateResult(tab.name, False, "not in the DOM")
    page.wait_for_timeout(_SETTLE_MS)
    return StateResult(tab.name, True, "hidden; activated via DOM click")


def enter_named(page: Any, page_name: str, name: str) -> StateResult:
    """Drive into one state by name, resolving overlays and tabs alike.

    Tab names are matched against what the page actually renders, so an
    unreachable one reports the states that *were* available rather than a bare
    timeout — the difference between "you typed the wrong name" and "this tab is
    empty in the fixture".
    """
    for overlay in GLOBAL_OVERLAYS:
        if overlay.name == name:
            return enter_overlay(page, overlay)

    discovered = discover_tabs(page, page_name)
    for tab in discovered:
        if tab.name == name:
            return enter_tab(page, tab)

    available = ", ".join(
        [overlay.name for overlay in GLOBAL_OVERLAYS] + [t.name for t in discovered]
    )
    return StateResult(name, False, f"unknown state; this page has: {available}")


def each_state(page: Any, page_name: str) -> Iterator[StateResult]:
    """Walk every reachable state on an already-loaded page.

    Yields with the page *in* that state, so the caller can screenshot or probe
    before control returns. Tabs run first because they are mutually exclusive
    and need no teardown; overlays run last, each closed before the next opens.
    """
    for tab in discover_tabs(page, page_name):
        yield enter_tab(page, tab)

    for overlay in GLOBAL_OVERLAYS:
        result = enter_overlay(page, overlay)
        yield result
        if result.reached:
            close_overlay(page, overlay)
