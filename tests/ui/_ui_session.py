"""One browser context factory, and one standalone boot, shared by every UI tool.

Two things live here, and they exist for different reasons.

**:func:`new_context` — the context factory.** ``conftest.py``'s ``browser_context``
fixture and ``shot.py`` each used to build their own Chromium context, and they had
already drifted: both set the start-overlay dismissal, neither pinned locale or
timezone. Any tool that compares one run against another — the dead-CSS census
especially — needs every run to see the *same* browser, so the context arguments
have to be defined once. Both callers now route through here.

**:func:`ui_session` — the standalone boot.** ``shot.py`` inlined config
redirection, workbook opening, server start and browser launch, and its own
docstring warns not to loop it because each invocation boots a server. The census
needs all six pages from a single boot, which is the same sequence minus pytest.
So the sequence moves here as a context manager and ``shot.py`` becomes one of its
callers. ``conftest.py`` deliberately does *not* use it: its fixtures need
session scoping, ``pytest.skip`` on an unavailable browser, and ``MonkeyPatch`` so
config is restored for the rest of the suite.

What is pinned in the init script, and why each one:

* ``clipgen.startOverlayDismissed`` — ``shouldAutoOpen()`` checks this key before
  anything else, and it is the only suppression that works uniformly across all
  six pages. Carried over from both previous copies.
* ``clipgen-theme`` — ``applyStoredThemePreference`` (utils.js) reads this at boot
  and sets ``data-theme`` on ``<html>``. Setting the key rather than poking the
  attribute afterwards is the path a real user with a stored preference takes, so
  the page boots *already* in the target theme instead of repainting into it.
* ``CLIPGEN_DEV_TOKEN_TWEAK`` — already ``false`` by default (utils.js), so this
  is not a fix; it is a pin. The flag exists to be flipped on while iterating on
  the redesign, and a maintainer who left it on would otherwise silently add a
  widget plus an injected ``<style>`` to every screenshot and every census.
"""

from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import _ui_browser
import _ui_fixtures
import _ui_server
import config
import start_settings
import utils

VIEWPORT = {"width": 1600, "height": 1000}
THEMES = ("dark", "light")


@dataclass
class Session:
    """A booted server plus an open browser context."""

    context: Any
    origin: str


def build_init_script(theme: str = "dark") -> str:
    """JS run before any page script, in every context this module creates."""
    if theme not in THEMES:
        raise ValueError(f"theme must be one of {THEMES}, got {theme!r}")
    return (
        "try {"
        " sessionStorage.setItem('clipgen.startOverlayDismissed', '1');"
        f" localStorage.setItem('clipgen-theme', '{theme}');"
        "} catch (e) {}"
        "window.CLIPGEN_DEV_TOKEN_TWEAK = false;"
    )


def new_context(
    browser: Any,
    *,
    viewport: dict[str, int] | None = None,
    theme: str = "dark",
) -> Any:
    """Build the one Chromium context shape every UI tool should use.

    ``locale``/``timezone_id`` are pinned so relative-time and clock strings
    ("5 d ago", ``toLocaleTimeString``) render identically wherever the harness
    runs. That does mean screenshots show UTC rather than the reviewer's local
    time — a deliberate trade: reproducibility across machines is worth more here
    than matching one machine's clock.
    """
    context = browser.new_context(
        viewport=viewport or VIEWPORT,
        device_scale_factor=1,
        locale="en-US",
        timezone_id="UTC",
        color_scheme="light" if theme == "light" else "dark",
    )
    context.add_init_script(build_init_script(theme))
    return context


def redirect_config(settings_dir: Path | None = None) -> None:
    """Point config at the throwaway fixture tree (one-shot processes only).

    Plain assignment, no ``MonkeyPatch``: callers of :func:`ui_session` are
    standalone scripts that exit afterwards. ``conftest.py`` needs the restoring
    version and keeps its own.
    """
    config.INPUT_DIR = str(_ui_fixtures.INPUT_DIR)
    config.OUTPUT_DIR = str(_ui_fixtures.OUTPUT_DIR)
    # Sheet-loading chatter would bury the screenshot path and the --eval result.
    config.VERBOSITY = config.QUIET
    # build_combined_app records a project session, which would otherwise prepend
    # these fixture dirs to the maintainer's real "Recently opened" rail.
    # setattr rather than plain assignment because the attribute's declared type
    # is the original function; this is the same rebinding monkeypatch does.
    target = settings_dir or _ui_fixtures.SETTINGS_DIR
    setattr(  # noqa: B010
        start_settings, "_settings_path", lambda: target / "start.json"
    )
    # Only start_combined_server sets this, and we deliberately don't use it.
    utils.NO_INPUT_MODE = True


@contextmanager
def ui_session(
    *,
    viewport: dict[str, int] | None = None,
    theme: str = "dark",
) -> Iterator[Session]:
    """Boot fixtures, server and browser once; yield a context to drive.

    Raises :class:`_ui_fixtures.UiUnavailable` when playwright, a Chromium build,
    ffmpeg or the fixture workbook is missing — callers decide whether that is a
    skip or an error.

    Teardown order matters and is enforced by the nesting below: the browser
    closes before the socket does, which is what keeps the server's
    connection-thread join from hanging (see ``_ui_server``).
    """
    _ui_fixtures.ensure_inputs()
    chromium_path = _ui_browser.resolve_chromium()
    playwright_factory = _ui_browser.sync_playwright()
    _ui_fixtures.ensure_run_dirs()
    redirect_config()

    workbook, reason = _ui_fixtures.open_workbook()
    if reason:
        raise _ui_fixtures.UiUnavailable(reason)

    live = _ui_server.start(workbook)
    playwright = playwright_factory().start()
    browser = None
    context = None
    try:
        browser = playwright.chromium.launch(
            executable_path=str(chromium_path),
            headless=True,
            args=_ui_browser.LAUNCH_ARGS,
        )
        context = new_context(browser, viewport=viewport, theme=theme)
        yield Session(context=context, origin=live.origin)
    finally:
        if context is not None:
            context.close()
        if browser is not None:
            browser.close()
        playwright.stop()
        _ui_server.stop(live)
