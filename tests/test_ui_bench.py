"""Pure parts of tests/ui/ui_bench.py and shot.py's soak math, without a browser.

The UI bench itself needs Chromium (``CLIPGEN_UI_CHECK=1``); these pin the
reduction and the slope math that turn a run artifact into a comparable
row, and the readiness-failure path that keeps a stalled page out of the
"ok" column.
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

import pytest

UI = Path(__file__).resolve().parent / "ui"
PERF = Path(__file__).resolve().parent / "perf"


def _load(path: Path, name: str) -> ModuleType:
    for extra in (UI, PERF):
        if str(extra) not in sys.path:
            sys.path.insert(0, str(extra))
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None and spec.loader is not None
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture(scope="module")
def shot() -> ModuleType:
    return _load(UI / "shot.py", "clipgen_ui_shot")


@pytest.fixture(scope="module")
def bench() -> ModuleType:
    return _load(UI / "ui_bench.py", "clipgen_ui_bench")


# ---------- soak math ---------------------------------------------------------


def test_slope_per_minute_fits_growth(shot):
    # 10 units per second is 600 per minute.
    assert shot.slope_per_minute([(0, 0), (1, 10), (2, 20)]) == pytest.approx(600)
    assert shot.slope_per_minute([(0, 5), (5, 5)]) == 0
    assert shot.slope_per_minute([(0, 5)]) is None
    assert shot.slope_per_minute([(1, 5), (1, 9)]) is None  # no time spread


def _sample(t, hidden, heap, polls):
    return {
        "t": t,
        "hidden": hidden,
        "cdp": {
            "Nodes": 100,
            "JSEventListeners": 10,
            "JSHeapUsedSize": heap,
            "Documents": 1,
        },
        "polls": polls,
        "longtasks": 0,
        "transferSize": 1000 + t,
    }


def test_soak_rates_split_polls_by_visibility(shot):
    samples = [
        _sample(0, False, 1000, {"poll.x": 0}),
        _sample(10, False, 1100, {"poll.x": 10}),
        _sample(20, True, 1200, {"poll.x": 10}),
        _sample(30, True, 1300, {"poll.x": 10}),
    ]
    rates = shot.soak_rates(samples)
    assert rates["JSHeapUsedSize"] == pytest.approx(600)
    assert rates["Nodes"] == 0
    assert rates["transferSize"] == pytest.approx(60)
    assert rates["polls"]["poll.x"]["visible"] == pytest.approx(60)
    assert rates["polls"]["poll.x"]["hidden"] == 0


def test_soak_rates_skips_unsupported_metric(shot):
    samples = [_sample(0, False, 1, {}), _sample(5, False, 2, {})]
    for s in samples:
        s["cdp"]["Documents"] = None
    assert shot.soak_rates(samples)["Documents"] is None


def test_shot_parses_soak_and_output_flags(shot):
    args = shot._parse_args(
        [
            "studio",
            "--perf",
            "--soak",
            "30",
            "--soak-interval",
            "2",
            "--soak-hidden",
            "--perf-output",
            "/tmp/p.json",
        ]
    )
    assert args.soak == 30 and args.soak_interval == 2 and args.soak_hidden
    assert args.perf_output == Path("/tmp/p.json")


# ---------- reduction ----------------------------------------------------------


def test_summarize_studio_run(bench):
    run = {
        "elapsed_s": 1.5,
        "load": {"page": {"navigation": {"domContentLoaded": 52.0}}},
        "actions": {"filter": {"ms": 15.0}, "queue": {"ms": 5.0}},
        "workload": {"rows": 200},
        "soak": {},
    }
    row = bench.summarize("studio", run)
    assert row == {
        "elapsed_s": 1.5,
        "load_ms": 52.0,
        "filter_ms": 15.0,
        "queue_ms": 5.0,
        "rows": 200,
    }
    assert bench.metrics_for("studio", [row]) == (
        "elapsed_s",
        "filter_ms",
        "load_ms",
        "queue_ms",
    )


def test_summarize_idle_run_takes_worst_page(bench):
    run = {
        "elapsed_s": 60.0,
        "load": None,
        "actions": {},
        "workload": {"pages": 6},
        "soak": {
            "studio": {
                "rates": {
                    "JSHeapUsedSize": 10.0,
                    "Nodes": 1.0,
                    "polls": {"poll.a": {"hidden": 2.0}},
                }
            },
            "composer": {"rates": {"JSHeapUsedSize": 40.0, "Nodes": 0.0, "polls": {}}},
        },
    }
    row = bench.summarize("idle", run)
    assert row["heap_per_min_max"] == 40.0
    assert row["nodes_per_min_max"] == 1.0
    assert row["hidden_polls_per_min"] == 2.0
    assert row["pages"] == 6


def test_summarize_sums_per_participant_counts(bench):
    run = {
        "elapsed_s": 1,
        "load": None,
        "actions": {},
        "workload": {"segments": {"P01": 2400, "P02": 100}},
        "soak": {},
    }
    assert bench.summarize("transcripts", run)["segments"] == 2500


def test_count_keys_cover_every_scenario(bench):
    assert set(bench.COUNT_KEY) == set(bench.SCENARIOS)


# ---------- readiness failure ------------------------------------------------------


class _StuckPage:
    """A page whose conditions never settle."""

    def wait_for_function(self, *_a, **_k):
        raise _PlaywrightError("Timeout 15000ms exceeded.")

    def evaluate(self, *_a, **_k):
        return 0


class _PlaywrightError(Exception):
    pass


def test_driver_action_reports_the_stalled_condition(bench, monkeypatch):
    monkeypatch.setattr(bench._ui_browser, "playwright_error", lambda: _PlaywrightError)
    driver = bench.Driver(_StuckPage(), None, "http://127.0.0.1:0")
    monkeypatch.setattr(driver, "server_window", dict)
    with pytest.raises(bench.ActionFailed, match="filter: Timeout"):
        driver.action("filter", lambda: None, before="tr", after="false")
    assert driver.actions == {}
