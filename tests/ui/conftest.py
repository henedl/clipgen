"""Session fixtures for the opt-in UI harness: fixtures -> server -> browser.

All three are session-scoped, which also settles how they interact with the
autouse fixtures in ``tests/conftest.py``. pytest instantiates by scope order, so
these run first, and none of the three conflicts:

* ``_sandbox_cwd`` chdirs into a tmp dir — harmless, because the config dirs here
  are absolute and ``utils.get_bundled_assets_root()`` is module-relative.
* ``_reset_no_input_mode`` and ``_reset_overview_observation_getter`` snapshot at
  the start of each test, by which point the session values are already live, so
  their restores are no-ops rather than teardowns of our wiring.

Fixture *teardown* is LIFO and ``browser_context`` depends on ``live_server``, so
the browser always closes before the socket does. That ordering is what keeps the
server's connection-thread join from hanging (see ``_ui_server``).
"""

from collections.abc import Iterator
from typing import Any

import pytest

import _ui_browser
import _ui_fixtures
import _ui_server
import config
import excel_io
import spreadsheet
import start_settings
import utils


@pytest.fixture(scope="session")
def ui_env() -> Iterator[None]:
    """Point config at the throwaway fixture tree and build the input fixtures."""
    try:
        _ui_fixtures.ensure_inputs()
    except _ui_fixtures.UiUnavailable as exc:
        pytest.skip(str(exc))
    _ui_fixtures.reset_run_dirs()

    with pytest.MonkeyPatch.context() as mp:
        mp.setattr(config, "INPUT_DIR", str(_ui_fixtures.INPUT_DIR))
        mp.setattr(config, "OUTPUT_DIR", str(_ui_fixtures.OUTPUT_DIR))
        # build_combined_app records a project session, which would otherwise
        # prepend the fixture dirs to the maintainer's real "Recently opened"
        # rail in ~/.config/clipgen/start.json. Redirect before the app is built.
        mp.setattr(
            start_settings,
            "_settings_path",
            lambda: _ui_fixtures.SETTINGS_DIR / "start.json",
        )
        # Only start_combined_server sets this, and we deliberately don't use it.
        mp.setattr(utils, "NO_INPUT_MODE", True)
        yield


@pytest.fixture(scope="session")
def live_server(ui_env: None) -> Iterator[_ui_server.LiveServer]:
    workbook = excel_io.open_excel_workbook(str(_ui_fixtures.workbook_path()))
    assert workbook is not None, "fixture workbook failed to open"
    # server._init_studio_state calls sys.exit(1) when build_sheet_context
    # returns None, which would abort the whole run with a bare SystemExit.
    # Fail readably instead if the generated layout ever drifts from
    # spreadsheet.py's contract.
    if spreadsheet.build_sheet_context(workbook) is None:
        pytest.fail(
            "generated fixture workbook no longer satisfies "
            "spreadsheet.build_sheet_context — see tests/ui/_ui_fixtures.py"
        )

    live = _ui_server.start(workbook)
    try:
        yield live
    finally:
        _ui_server.stop(live)


@pytest.fixture(scope="session")
def browser_context(live_server: _ui_server.LiveServer) -> Iterator[Any]:
    """A Chromium context with the blocking start overlay pre-dismissed."""
    try:
        playwright = _ui_browser.sync_playwright()().start()
    except _ui_fixtures.UiUnavailable as exc:
        pytest.skip(str(exc))
    browser = None
    context = None
    try:
        browser = playwright.chromium.launch(
            executable_path=str(_ui_browser.resolve_chromium()),
            headless=True,
            args=_ui_browser.LAUNCH_ARGS,
        )
        context = browser.new_context(
            viewport={"width": 1600, "height": 1000}, device_scale_factor=1
        )
        # shouldAutoOpen() checks this key before anything else, and it is the
        # only suppression that works uniformly across all six pages.
        context.add_init_script(
            "try { sessionStorage.setItem('clipgen.startOverlayDismissed', '1'); }"
            " catch (e) {}"
        )
        yield context
    except _ui_fixtures.UiUnavailable as exc:
        pytest.skip(str(exc))
    finally:
        if context is not None:
            context.close()
        if browser is not None:
            browser.close()
        playwright.stop()
