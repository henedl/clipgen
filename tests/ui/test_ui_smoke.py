"""Load all six pages in a real browser and fail on any uncaught error.

This is the runtime half of the frontend guards. ``test_frontend_syntax.py``
proves each file parses and ``test_frontend_satellite_wiring.py`` approximates a
scope checker with regexes; neither can see a bare cross-file reference that only
throws once the page boots — the class that has shipped three times. Here it is
just a ``pageerror``.

Opt-in twice over, because a browser has no business in ``/check`` or CI: the
directory is excluded by ``norecursedirs`` in ``tests/pytest.ini``, and the
module-level gate below refuses to run without ``CLIPGEN_UI_CHECK=1``. Use
``/ui-check``, which also handles the prompted install.
"""

import json
import os
from pathlib import Path
from typing import Any

import pytest

if os.environ.get("CLIPGEN_UI_CHECK") != "1":
    pytest.skip(
        "The UI smoke harness is opt-in. Run the /ui-check skill, or:\n"
        "  CLIPGEN_UI_CHECK=1 uv run --extra dev --extra ui pytest "
        "-c tests/pytest.ini tests/ui -p no:randomly",
        allow_module_level=True,
    )

import _ui_fixtures
import _ui_pages
import _ui_states
from _ui_pages import PAGES, PageLog
from _ui_server import LiveServer

pytestmark = pytest.mark.ui

# Which page to open the shared overlays on. They are page-agnostic code
# (settings-modal.js, hotkeys.js, command-palette.js, start-overlay.js, all
# loaded by every hub), so opening them on all six would run the same files six
# times for the same assertion. Studio boots the most page code around them.
_OVERLAY_HOST = "studio"


@pytest.mark.parametrize("name", sorted(PAGES))
def test_page_loads_without_errors(
    name: str, live_server: LiveServer, browser_context: Any
) -> None:
    log = PageLog()
    shot = _ui_fixtures.SHOT_DIR / f"{name}.png"
    page = browser_context.new_page()
    _ui_pages.wire_listeners(page, log)
    try:
        _ui_pages.open_and_settle(page, live_server.origin, name, log)
    finally:
        # Screenshot even on failure — a broken page is exactly the one worth
        # looking at, and the report has to record the attempt either way.
        try:
            page.screenshot(path=str(shot), full_page=True)
        except Exception as exc:  # pragma: no cover - capture is best-effort
            log.console_errors.append(f"screenshot failed: {exc}")
        _record(name, log, shot.resolve(), live_server.origin)
        page.close()

    assert not log.fatal, _ui_pages.format_failure(
        name, log, str(shot.resolve()), str(_ui_fixtures.REPORT_PATH.resolve())
    )


@pytest.mark.parametrize("overlay", _ui_states.GLOBAL_OVERLAYS, ids=lambda o: o.name)
def test_shared_overlay_opens_without_errors(
    overlay: _ui_states.Overlay, live_server: LiveServer, browser_context: Any
) -> None:
    """Open each shared modal and fail on anything it throws on the way up.

    The page-load smoke above clicks nothing, so every modal in the app was
    unreachable by any automated check: a carve that broke ``openSettingsModal``
    would ship green. Opening is where that breaks — the overlays are built lazily
    on first open, not at boot.

    Reaching the state is itself the assertion. ``enter_overlay`` waits for the
    overlay's own visible root, so a throw during construction surfaces here as a
    failure to appear, with the ``pageerror`` that caused it reported alongside.
    """
    log = PageLog()
    shot = _ui_fixtures.SHOT_DIR / f"{_OVERLAY_HOST}-{overlay.name}.png"
    page = browser_context.new_page()
    _ui_pages.wire_listeners(page, log)
    try:
        _ui_pages.open_and_settle(page, live_server.origin, _OVERLAY_HOST, log)
        result = _ui_states.enter_overlay(page, overlay)
        try:
            page.screenshot(path=str(shot), full_page=True)
        except Exception as exc:  # pragma: no cover - capture is best-effort
            log.console_errors.append(f"screenshot failed: {exc}")
    finally:
        page.close()

    assert result.reached, (
        f"{overlay.name} did not open on {_OVERLAY_HOST}: {result.detail}\n"
        + _ui_pages.format_failure(
            overlay.name, log, str(shot), str(_ui_fixtures.REPORT_PATH)
        )
    )
    assert not log.fatal, _ui_pages.format_failure(
        overlay.name, log, str(shot), str(_ui_fixtures.REPORT_PATH)
    )


def _record(name: str, log: PageLog, shot: Path, server_url: str) -> None:
    """Merge one page's result into ``.context/ui-check/ui-report.json``.

    Read-modify-write rather than a session-scoped accumulator: pages are
    independent tests, and a run interrupted halfway should still leave the
    pages that did complete on disk.
    """
    report: dict[str, Any] = {}
    if _ui_fixtures.REPORT_PATH.is_file():
        try:
            loaded = json.loads(_ui_fixtures.REPORT_PATH.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                report = loaded
        except (OSError, json.JSONDecodeError):
            report = {}
    existing = report.get("pages")
    pages: dict[str, Any] = existing if isinstance(existing, dict) else {}
    pages[name] = {
        "screenshot": str(shot),
        "page_errors": log.page_errors,
        "console_errors": log.console_errors,
        "non_2xx": [list(item) for item in log.non_2xx],
        "request_failures": [list(item) for item in log.request_failures],
        "timeout": log.timeout,
    }
    report["server_url"] = server_url
    report["pages"] = pages
    _ui_fixtures.REPORT_PATH.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
