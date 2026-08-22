"""Five click-through journeys: click → API → DOM must still hang together.

The boot smoke (``test_ui_smoke.py``) proves the curtain goes up; it clicks
nothing. Each journey here drives one researcher path through the page's own
UI — real clicks on live selectors, never internal render calls — and asserts
both the journey's DOM contract and ``log.fatal``. Capped at five until a real
shipped bug earns a sixth (see ``plans/TEST-SUITE-PLAN.md`` /
``agents/skills/test/SKILL.md``).

Same double opt-in as the smoke: ``norecursedirs`` excludes the directory and
the module gate below refuses to run without ``CLIPGEN_UI_CHECK=1``.
"""

import json
import os
import urllib.request
from collections.abc import Iterator
from contextlib import contextmanager
from typing import Any

import pytest

if os.environ.get("CLIPGEN_UI_CHECK") != "1":
    pytest.skip(
        "The UI journey harness is opt-in. Run the /ui-check skill, or:\n"
        "  CLIPGEN_UI_CHECK=1 uv run --extra dev --extra ui pytest "
        "-c tests/pytest.ini tests/ui -p no:randomly",
        allow_module_level=True,
    )

import _ui_fixtures
import _ui_pages
import _ui_session
import _ui_states
from _ui_pages import PageLog
from _ui_server import LiveServer

import config

pytestmark = pytest.mark.ui

# The live selectors each journey leans on, in one place for the same honesty as
# ``PAGES`` in ``_ui_pages.py``: this table is the part most likely to rot on a
# redesign, and a rename should be a one-line fix here rather than a dig.
SELECTORS = {
    "studio-cell": '#sheetGrid .ts-cell.valid-ts[data-row="6"][data-participant="P01"]',
    "studio-generate": "#generateBtn",
    "studio-card-success": "#artifactsList .queue-card.queue-card-success",
    "studio-card-fail": "#artifactsList .queue-card.queue-card-fail",
    # The spinner is `.active`-class-toggled (display:none base), never removed.
    "studio-spinner-idle": "#artifactsSpinner:not(.active)",
    "studio-status-open": "#statusOverlay:not(.hidden)",
    "ss-task-card": "#taskList .task-card",
    "ss-result-row": "#resultsList .result-row",
    # Row-level clicks are a no-op; only the timestamp/text children seek.
    "ts-segment-seek": '#segmentList .segment-row[data-index="1"] .segment-timestamp',
    "settings-input": '.settings-row[data-setting="WEBP_QUALITY"] .settings-control input',
    "settings-status": ".settings-save-status",
    "start-tab-excel": '#startOverlay .sheet-card__tab[data-tab="excel"]',
    "start-picker-trigger": '[data-role="excel-picker-trigger"]',
    "start-picker-option": '[data-role="excel-picker-menu"] .sheet-picker__option',
    "start-confirm": '[data-role="confirm"]',
}

_WAIT_MS = 15_000
# The generate journey cuts a real ~3 s clip through ffmpeg; give it headroom.
_GENERATE_MS = 30_000


def _overlay(name: str) -> _ui_states.Overlay:
    return next(o for o in _ui_states.GLOBAL_OVERLAYS if o.name == name)


def _fatal_message(name: str, log: PageLog, shot_name: str) -> str:
    return _ui_pages.format_failure(
        name,
        log,
        str((_ui_fixtures.SHOT_DIR / f"{shot_name}.png").resolve()),
        str(_ui_fixtures.REPORT_PATH.resolve()),
    )


@contextmanager
def _journey(
    context: Any, live_server: LiveServer, page_name: str, shot_name: str
) -> Iterator[tuple[Any, PageLog]]:
    """Open one page with listeners wired; screenshot even on failure.

    A broken journey is exactly the one worth looking at, so the capture runs in
    the ``finally`` — same contract as the smoke.
    """
    log = PageLog()
    page = context.new_page()
    _ui_pages.wire_listeners(page, log)
    try:
        _ui_pages.open_and_settle(page, live_server.origin, page_name, log)
        yield page, log
    finally:
        try:
            shot = _ui_fixtures.SHOT_DIR / f"{shot_name}.png"
            page.screenshot(path=str(shot), full_page=True)
        except Exception as exc:  # pragma: no cover - capture is best-effort
            log.console_errors.append(f"screenshot failed: {exc}")
        page.close()


def test_studio_generate_produces_an_artifact(
    live_server: LiveServer, browser_context: Any
) -> None:
    """Queue one valid cell, click Generate, and watch a real clip land.

    Row 6 is P01 ``0:01-0:04`` — a 3 s cut of the 20 s fixture. A plain click is
    load-bearing: Shift/right-click route the cell to the reel queue instead.
    """
    with _journey(browser_context, live_server, "studio", "journey-generate") as (
        page,
        log,
    ):
        page.click(SELECTORS["studio-cell"])
        page.wait_for_selector(
            SELECTORS["studio-generate"] + ":not([disabled])", timeout=_WAIT_MS
        )
        page.click(SELECTORS["studio-generate"])
        # Terminal signals, strongest first: the "Done" status modal, then the
        # per-card success class, then the title spinner going idle.
        page.wait_for_selector(SELECTORS["studio-status-open"], timeout=_GENERATE_MS)
        assert page.locator("#statusTitle").inner_text() == "Done"
        page.wait_for_selector(SELECTORS["studio-card-success"], timeout=_WAIT_MS)
        page.wait_for_selector(
            SELECTORS["studio-spinner-idle"], state="attached", timeout=_WAIT_MS
        )
        assert page.locator(SELECTORS["studio-card-fail"]).count() == 0
    assert not log.fatal, _fatal_message("studio-generate", log, "journey-generate")


def test_screenspace_result_row_seeks_playhead(
    live_server: LiveServer, browser_context: Any
) -> None:
    """Select the seeded completed task, click its first result, land on 2.0 s.

    The assertion is the playhead state (``#timestampInput``), not
    ``video.currentTime``: Screenspace's ``loadFrame`` pauses the video and
    shows a still frame — the ``<video>`` element deliberately does not move.
    """
    with _journey(
        browser_context, live_server, "screenspace", "journey-screenspace"
    ) as (
        page,
        log,
    ):
        page.click(SELECTORS["ss-task-card"])
        page.wait_for_selector(SELECTORS["ss-result-row"], timeout=_WAIT_MS)
        page.locator(SELECTORS["ss-result-row"]).first.click()
        page.wait_for_function(
            "() => ((document.querySelector('#timestampInput') || {}).value || '')"
            ".indexOf('0:02') === 0",
            timeout=_WAIT_MS,
        )
    assert not log.fatal, _fatal_message("screenspace-seek", log, "journey-screenspace")


def test_transcripts_segment_click_seeks_video(
    live_server: LiveServer, browser_context: Any
) -> None:
    """Click segment 1's timestamp (start 2.5 s) and watch the player follow.

    Segment 0 starts at 0.0, indistinguishable from boot state. The seek
    coalescer also starts playback, so the tolerance absorbs a little drift and
    the video is paused before the screenshot.
    """
    with _journey(
        browser_context, live_server, "transcripts", "journey-transcripts"
    ) as (
        page,
        log,
    ):
        page.click(SELECTORS["ts-segment-seek"])
        page.wait_for_function(
            "() => { const v = document.querySelector('#videoPlayer');"
            " return !!v && Math.abs(v.currentTime - 2.5) < 0.75; }",
            timeout=_WAIT_MS,
        )
        page.evaluate("() => document.querySelector('#videoPlayer').pause()")
    assert not log.fatal, _fatal_message("transcripts-seek", log, "journey-transcripts")


def test_settings_change_persists_across_reload(
    live_server: LiveServer, browser_context: Any
) -> None:
    """Change WEBP_QUALITY in the modal, reload Studio, and find it stuck.

    WEBP_QUALITY is the durable setting nothing else depends on: no server-side
    capability probe (unlike the format/titlecard settings) and no effect on the
    fixture. The file under the output dir is the real store — only non-default
    values are written — so it is asserted alongside the reopened modal.
    """
    settings_path = _ui_fixtures.SETTINGS_DIR / config.STUDIO_SETTINGS_FILENAME
    try:
        with _journey(browser_context, live_server, "studio", "journey-settings") as (
            page,
            log,
        ):
            result = _ui_states.enter_overlay(page, _overlay("settings"))
            assert result.reached, result.detail
            field = page.locator(SELECTORS["settings-input"])
            field.fill("55")
            field.dispatch_event("change")
            page.wait_for_function(
                "() => (document.querySelector('.settings-save-status') || {})"
                ".textContent === 'Saved'",
                timeout=_WAIT_MS,
            )
            saved = json.loads(settings_path.read_text(encoding="utf-8"))
            assert saved.get("WEBP_QUALITY") == 55

            page.reload(wait_until="domcontentloaded")
            page.wait_for_selector("#sheetGrid tbody tr", timeout=_WAIT_MS)
            result = _ui_states.enter_overlay(page, _overlay("settings"))
            assert result.reached, result.detail
            assert page.locator(SELECTORS["settings-input"]).input_value() == "55"
        assert not log.fatal, _fatal_message(
            "settings-persist", log, "journey-settings"
        )
    finally:
        # config is a live module global in the shared server process; put the
        # default back so the rest of the session never sees the override.
        request = urllib.request.Request(
            live_server.origin + "/api/settings",
            data=json.dumps({"settings": {"WEBP_QUALITY": 80}}).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="PUT",
        )
        urllib.request.urlopen(request, timeout=10).read()


def test_start_overlay_opens_the_fixture_workbook(
    live_server: LiveServer, browser: Any
) -> None:
    """Pick the fixture workbook through the overlay's own UI, land on the grid.

    Runs in its own context without the dismissal pin so the overlay behaves as
    shipped. It still never auto-opens here (``shouldAutoOpen`` bails on
    ``sheet_loaded``), so it is opened through its public opener like every
    overlay state. The seeded recents rail is six fabricated dead projects —
    the journey goes through the Excel picker, not the rail.
    """
    context = _ui_session.new_context(browser, dismiss_start=False)
    try:
        with _journey(context, live_server, "studio", "journey-start-overlay") as (
            page,
            log,
        ):
            result = _ui_states.enter_overlay(page, _overlay("start"))
            assert result.reached, result.detail
            page.click(SELECTORS["start-tab-excel"])
            page.click(SELECTORS["start-picker-trigger"])
            option = page.locator(SELECTORS["start-picker-option"]).first
            option.wait_for(state="visible", timeout=_WAIT_MS)
            assert (option.get_attribute("data-path") or "").endswith(
                f"{_ui_fixtures.STUDY}.xlsx"
            )
            option.click()
            page.wait_for_selector(
                SELECTORS["start-confirm"] + ":not([disabled])", timeout=_WAIT_MS
            )
            # confirm() ends in window.location.reload() on success.
            with page.expect_navigation(
                wait_until="domcontentloaded", timeout=_GENERATE_MS
            ):
                page.click(SELECTORS["start-confirm"])
            page.wait_for_selector("#sheetGrid tbody tr", timeout=_WAIT_MS)
        assert not log.fatal, _fatal_message(
            "start-overlay", log, "journey-start-overlay"
        )
    finally:
        context.close()
