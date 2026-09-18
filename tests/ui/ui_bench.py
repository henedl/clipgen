"""Reproducible browser workload benchmarks over the headless harness.

``shot.py --perf`` measures one page as it boots. This drives a *workload*:
a fixture big enough to mean something, a readiness condition that proves
the data actually rendered, and a fixed sequence of interactions each timed
on its own, so a report can say what loading cost, what an interaction
cost, and what an idle tab churns — three different bugs.

    CLIPGEN_UI_CHECK=1 uv run --extra ui python tests/ui/ui_bench.py
    ... ui_bench.py --scenarios studio,transcripts --runs 3 --save /tmp/ui.json
    ... ui_bench.py --compare /tmp/ui.json --fail-on 15
    ... ui_bench.py --scenarios idle --soak 60

Scenarios (fixtures come from tests/perf/bench_fixtures.py):

- ``studio``       200 observations × 12 participants; load, severity filter,
                   restore, queue ten cells.
- ``transcripts``  2,400 segments for P01 (100 for P02); load, search,
                   switch participant and back.
- ``screenspace``  2,000 synthetic events with the fixture video; load,
                   open results, drain the lazy list, confidence filter,
                   switch panes.
- ``idle``         each of the six pages sits for ``--soak`` seconds sampled
                   every ``--soak-interval``, the second half emulated as a
                   background tab. Growth is a signal, not a proof.

Every action waits on a *condition* — a DOM count or a namespace state
field — never a fixed sleep; a condition that does not settle marks the
run invalid with the reason. Per run the artifact JSON under ``--work``
holds the load timing, each action's ms with the counts observed before and
after, a server-side ``/api/profile?reset=1`` window per action, the workload
counts asserted, page errors, and the screenshot path. Rows aggregate like
the backend benches (``bench_common``): every sample kept, medians compared,
``invalid`` / ``incomparable`` / ``regression`` findings and exit codes.

Not a test: ``norecursedirs = ui`` and the non-``test_`` name keep pytest
away; ``tests/test_ui_bench.py`` covers the pure parts.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
import urllib.request
from pathlib import Path
from typing import Any

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE))  # the _ui_* helpers
sys.path.insert(0, str(_HERE.parents[1] / "source"))
sys.path.insert(0, str(_HERE.parent / "perf"))  # bench_common, bench_fixtures

import _ui_browser
import _ui_fixtures
import _ui_pages
import _ui_session
import bench_common as bc
import bench_fixtures as bf
import shot

STUDY = "uibench"
SHEET_ROWS = 200
SHEET_PARTICIPANTS = 12
SEGMENTS = {"P01": 2400, "P02": 100}
EVENTS = 2000
ACTION_TIMEOUT_MS = 15_000

SCENARIOS = ("studio", "transcripts", "screenspace", "idle")
# The count that must match a baseline for rows to be comparable.
COUNT_KEY = {
    "studio": "rows",
    "transcripts": "segments",
    "screenspace": "events",
    "idle": "pages",
}


class ActionFailed(RuntimeError):
    """A readiness or action condition did not settle in time."""


# ---- Fixtures -------------------------------------------------------------------


def build_fixtures(work: Path) -> dict[str, dict[str, Any]]:
    """Write every scenario's inputs under *work*; returns per-scenario paths."""
    _ui_fixtures.ensure_inputs()
    studio_dir = work / "studio"
    sheet = studio_dir / f"{STUDY}.xlsx"
    if not sheet.is_file():
        bf.write_sheet(
            sheet, study=STUDY, rows=SHEET_ROWS, participants=SHEET_PARTICIPANTS
        )
    ts_dir = work / "transcripts"
    if not (ts_dir / "clipgen.json").is_file():
        bf.write_manifest(ts_dir, {"transcripts": bf.transcripts_section(SEGMENTS)})
    ss_dir = work / "screenspace"
    video = _ui_fixtures.video_path("P01")
    if not (ss_dir / "clipgen.json").is_file():
        bf.write_manifest(
            ss_dir,
            {
                "screenspace": bf.screenspace_section(
                    study=_ui_fixtures.STUDY,
                    participant="P01",
                    video_path=video,
                    events=EVENTS,
                    duration=float(_ui_fixtures.CLIP_SECONDS),
                )
            },
        )
    return {
        "studio": {"sheet": sheet, "input_dir": studio_dir, "output_dir": studio_dir},
        "transcripts": {"sheet": None, "input_dir": None, "output_dir": ts_dir},
        "screenspace": {"sheet": None, "input_dir": None, "output_dir": ss_dir},
        "idle": {"sheet": None, "input_dir": None, "output_dir": None},
    }


def fixture_counts(work: Path) -> dict[str, Any]:
    """What was written, so a run can prove the page loaded all of it."""
    return {
        "rows": bf.sheet_rows(work / "studio" / f"{STUDY}.xlsx"),
        "segments": bf.manifest_counts(work / "transcripts")["segments"],
        "events": bf.manifest_counts(work / "screenspace")["events"],
        "pages": len(_ui_pages.PAGES),
    }


# ---- Driving one page -----------------------------------------------------------


class Driver:
    """One page under measurement: conditions, timed actions, server windows."""

    def __init__(self, page: Any, cdp: Any, origin: str) -> None:
        self.page = page
        self.cdp = cdp
        self.origin = origin
        self.actions: dict[str, dict[str, Any]] = {}

    def wait(self, js: str, what: str) -> Any:
        """Wait until *js* (an expression) is truthy; raise with *what* on timeout."""
        try:
            return self.page.wait_for_function(
                f"() => {{ return {js}; }}", timeout=ACTION_TIMEOUT_MS
            ).json_value()
        except _ui_browser.playwright_error() as exc:
            raise ActionFailed(f"{what}: {str(exc).splitlines()[0]}") from None

    def count(self, selector: str) -> int:
        return int(
            self.page.evaluate(f"document.querySelectorAll({selector!r}).length")
        )

    def server_window(self) -> dict[str, Any]:
        """Snapshot-and-reset the server's profile; the labels since last call."""
        with urllib.request.urlopen(
            f"{self.origin}/api/profile?reset=1", timeout=5
        ) as r:
            doc = json.loads(r.read().decode("utf-8"))
        return doc.get("labels", {})

    def action(self, name: str, fn: Any, *, before: str, after: str) -> None:
        """Time *fn* plus its settle condition; record counts either side."""
        self.server_window()  # open a fresh window
        b = self.count(before)
        t0 = time.perf_counter()
        fn()
        a = self.wait(after, name)
        ms = (time.perf_counter() - t0) * 1000
        self.actions[name] = {
            "ms": ms,
            "before": b,
            "after": a,
            "server": self.server_window(),
        }

    def click_text(self, selector: str, text: str) -> None:
        """DOM-click the element under *selector* whose text is *text*."""
        hit = self.page.evaluate(
            """([sel, text]) => {
              var els = document.querySelectorAll(sel);
              for (var i = 0; i < els.length; i++) {
                if (els[i].textContent.trim().indexOf(text) === 0) { els[i].click(); return true; }
              }
              return false;
            }""",
            [selector, text],
        )
        if not hit:
            raise ActionFailed(f"no {selector!r} with text {text!r}")

    def click_all(self, selector: str, limit: int) -> int:
        return int(
            self.page.evaluate(
                """([sel, limit]) => {
                  var els = document.querySelectorAll(sel);
                  var n = Math.min(els.length, limit);
                  for (var i = 0; i < n; i++) els[i].click();
                  return n;
                }""",
                [selector, limit],
            )
        )


# ---- Scenarios ------------------------------------------------------------------

_ROWS = "#sheetGrid tbody tr:not(.empty-rows-separator)"
_SEV_ROW = '#studioSidebar [data-target="severity"] .studio-sidebar-row'


def run_studio(d: Driver, counts: dict[str, Any]) -> dict[str, Any]:
    rows = counts["rows"]
    d.wait(f"document.querySelectorAll({_ROWS!r}).length === {rows}", "grid ready")
    critical = rows // 4  # severities cycle through four labels
    d.action(
        "filter",
        lambda: d.click_text(_SEV_ROW, "Critical"),
        before=_ROWS,
        after=f"document.querySelectorAll({_ROWS!r}).length === {critical}"
        f" && document.querySelectorAll({_ROWS!r}).length",
    )
    d.action(
        "restore",
        lambda: d.click_text(_SEV_ROW, "Any severity"),
        before=_ROWS,
        after=f"document.querySelectorAll({_ROWS!r}).length === {rows}"
        f" && document.querySelectorAll({_ROWS!r}).length",
    )
    d.action(
        "queue",
        lambda: d.click_all("#sheetGrid .ts-cell.valid-ts", 10),
        before="#artifactsList .queue-card",
        after="!document.querySelector('#generateBtn').disabled"
        " && window.ClipgenStudio.state.artifactQueue.length >= 10"
        " && window.ClipgenStudio.state.artifactQueue.length",
    )
    return {"rows": d.count(_ROWS)}


_SEG = "#segmentList .segment-row"


def run_transcripts(d: Driver, counts: dict[str, Any]) -> dict[str, Any]:
    n1, n2 = counts["segments"]["P01"], counts["segments"]["P02"]
    d.wait(
        f"document.querySelectorAll({_SEG!r}).length === {n1}"
        f" && window.ClipgenTranscripts.state.segments.length === {n1}",
        "segments ready",
    )
    d.action(
        "search",
        lambda: d.page.fill("#searchInput", "segment 12"),
        before="#searchResultsList .search-result-row[data-participant]",
        after="document.querySelector('#searchCount').textContent.trim() !== ''"
        " && document.querySelectorAll('#searchResultsList .search-result-row[data-participant]').length",
    )
    d.action(
        "clear-search",
        lambda: d.page.fill("#searchInput", ""),
        before="#searchResultsList .search-result-row[data-participant]",
        after="document.querySelector('#searchResults').classList.contains('hidden')",
    )
    d.action(
        "switch",
        lambda: d.page.click('#participantPills .pill-wrap[data-pid="P02"] .pill-id'),
        before=_SEG,
        after="window.ClipgenTranscripts.state.selectedParticipant === 'P02'"
        f" && document.querySelectorAll({_SEG!r}).length === {n2}"
        f" && document.querySelectorAll({_SEG!r}).length",
    )
    d.action(
        "restore",
        lambda: d.page.click('#participantPills .pill-wrap[data-pid="P01"] .pill-id'),
        before=_SEG,
        after="window.ClipgenTranscripts.state.selectedParticipant === 'P01'"
        f" && document.querySelectorAll({_SEG!r}).length === {n1}"
        f" && document.querySelectorAll({_SEG!r}).length",
    )
    return {"segments": d.count(_SEG)}


_RESULT = "#resultsList .result-row"
_DRAINED = (
    "window.ClipgenScreenspace.state.resultsLazyObserver === null"
    " && !document.querySelector('#resultsList .results-lazy-sentinel')"
)


def run_screenspace(d: Driver, counts: dict[str, Any]) -> dict[str, Any]:
    events = counts["events"]
    d.wait(
        "document.querySelector('#videoPlayer[src]')"
        " && window.ClipgenScreenspace.state.tasks.length === 1",
        "screenspace ready",
    )

    def open_results() -> None:
        d.page.click('.rp-tab[data-tab="queue"]')
        d.page.click('#taskList .task-card[data-task-id="bench-task-1"]')

    d.action(
        "open-results",
        open_results,
        before=_RESULT,
        after="window.ClipgenScreenspace.state.resultsLoading === false"
        f" && document.querySelectorAll({_RESULT!r}).length > 0"
        f" && document.querySelectorAll({_RESULT!r}).length",
    )

    def drain() -> None:
        deadline = time.perf_counter() + ACTION_TIMEOUT_MS / 1000
        while time.perf_counter() < deadline:
            if d.page.evaluate(f"() => {_DRAINED}"):
                return
            d.page.evaluate(
                "() => { var el = document.querySelector('#resultsList');"
                " el.scrollTop = el.scrollHeight; }"
            )
            d.page.wait_for_timeout(50)

    d.action(
        "drain",
        drain,
        before=_RESULT,
        after=f"{_DRAINED} && document.querySelectorAll({_RESULT!r}).length === {events}"
        f" && document.querySelectorAll({_RESULT!r}).length",
    )

    def filter_results() -> None:
        d.page.fill("#certaintyCutoff", "60")
        d.page.dispatch_event("#certaintyCutoff", "input")

    d.action(
        "filter",
        filter_results,
        before=_RESULT,
        after="window.ClipgenScreenspace.state.certaintyCutoff === 0.6"
        f" && document.querySelectorAll({_RESULT!r}).length < {events}"
        f" && document.querySelectorAll({_RESULT!r}).length",
    )
    d.action(
        "switch-pane",
        lambda: d.page.click('.rp-tab[data-tab="preview"]'),
        before=_RESULT,
        after="document.querySelector('#resultsPanel').classList.contains('hidden')",
    )
    d.action(
        "switch-back",
        lambda: d.page.click('.rp-tab[data-tab="results"]'),
        before=_RESULT,
        after="!document.querySelector('#resultsPanel').classList.contains('hidden')"
        f" && document.querySelectorAll({_RESULT!r}).length",
    )
    return {"events": events}


SCENARIO_FN = {
    "studio": run_studio,
    "transcripts": run_transcripts,
    "screenspace": run_screenspace,
}
SCENARIO_PAGE = {
    "studio": "studio",
    "transcripts": "transcripts",
    "screenspace": "screenspace",
}


# ---- One run --------------------------------------------------------------------


def _open(
    context: Any, origin: str, name: str, log: _ui_pages.PageLog
) -> tuple[Any, Any, dict[str, Any]]:
    """New page + CDP session, navigated and settled; returns the load capture."""
    page = context.new_page()
    _ui_pages.wire_listeners(page, log)
    cdp = context.new_cdp_session(page)
    cdp.send("Performance.enable")
    _ui_pages.open_and_settle(page, origin, name, log)
    if log.timeout:
        raise ActionFailed(f"readiness: {log.timeout}")
    return page, cdp, shot._collect_perf(page, cdp)


def run_once(
    scenario: str,
    session: Any,
    counts: dict[str, Any],
    shot_path: Path,
    soak: float,
    soak_interval: float,
) -> dict[str, Any]:
    """One fresh browser context through the scenario; never raises."""
    log = _ui_pages.PageLog()
    context = _ui_session.new_context(session.context.browser)
    result: dict[str, Any] = {
        "valid": True,
        "reasons": [],
        "elapsed_s": 0.0,
        "load": None,
        "actions": {},
        "soak": {},
        "workload": {},
        "errors": {},
        "screenshot": str(shot_path),
    }
    t0 = time.perf_counter()
    page = None
    try:
        if scenario == "idle":
            for name in sorted(_ui_pages.PAGES):
                page, cdp, load = _open(context, session.origin, name, log)
                result["soak"][name] = shot.run_soak(
                    page, cdp, soak, soak_interval, True
                )
                page.close()
                page = None
            result["workload"] = {"pages": len(_ui_pages.PAGES)}
        else:
            page, cdp, load = _open(
                context, session.origin, SCENARIO_PAGE[scenario], log
            )
            result["load"] = load
            driver = Driver(page, cdp, session.origin)
            result["workload"] = SCENARIO_FN[scenario](driver, counts)
            result["actions"] = driver.actions
            result["after"] = shot._collect_perf(page, cdp)
    except ActionFailed as exc:
        result["reasons"].append(str(exc))
    except _ui_browser.playwright_error() as exc:
        result["reasons"].append(str(exc).splitlines()[0])
    finally:
        result["elapsed_s"] = time.perf_counter() - t0
        if page is not None:
            try:
                page.screenshot(path=str(shot_path), full_page=True)
            except _ui_browser.playwright_error():
                pass
        context.close()
    if log.page_errors:
        result["reasons"].append(f"{len(log.page_errors)} page error(s)")
    result["errors"] = {
        "page_errors": log.page_errors,
        "console_errors": log.console_errors,
        "timeout": log.timeout,
    }
    result["valid"] = not result["reasons"]
    return result


# ---- Reduction --------------------------------------------------------------------


def summarize(scenario: str, run: dict[str, Any]) -> dict[str, Any]:
    """The per-run sample: load, each action's ms, workload counts, soak rates."""
    out: dict[str, Any] = {"elapsed_s": run["elapsed_s"]}
    load = run.get("load") or {}
    nav = (load.get("page") or {}).get("navigation") or {}
    out["load_ms"] = float(nav.get("domContentLoaded") or 0.0)
    for name, act in (run.get("actions") or {}).items():
        out[f"{name}_ms"] = act["ms"]
    if scenario == "idle":
        heap = [s["rates"].get("JSHeapUsedSize") or 0.0 for s in run["soak"].values()]
        nodes = [s["rates"].get("Nodes") or 0.0 for s in run["soak"].values()]
        hidden_ticks = 0.0
        for s in run["soak"].values():
            for by_vis in s["rates"].get("polls", {}).values():
                hidden_ticks += by_vis.get("hidden") or 0.0
        out["heap_per_min_max"] = max(heap) if heap else 0.0
        out["nodes_per_min_max"] = max(nodes) if nodes else 0.0
        out["hidden_polls_per_min"] = hidden_ticks
    counts = run.get("workload") or {}
    for key, value in counts.items():
        out[key] = sum(value.values()) if isinstance(value, dict) else value
    return out


def metrics_for(scenario: str, samples: list[dict[str, Any]]) -> tuple[str, ...]:
    keys = {k for s in samples for k in s if k.endswith("_ms") or k == "elapsed_s"}
    return tuple(sorted(keys))


def print_table(
    rows: dict[str, dict[str, Any]],
    findings: list[tuple[str, str, str]],
    metric: str,
) -> None:
    print(
        f"{'scenario':<13}{'elapsed':>9}{'min':>9}{'mad':>8}  actions (median ms)  status"
    )
    for name, row in rows.items():
        status = bc.status_of(name, findings)
        if not row["samples"]:
            print(f"{name:<13}{'':>26}  -  {status}")
            continue
        parts = []
        for key in sorted(row["stats"]):
            if key.endswith("_ms"):
                parts.append(
                    f"{key.removesuffix('_ms')}={row['stats'][key]['median']:.0f}"
                )
        print(
            f"{name:<13}{bc.stat_cells(row, 'elapsed_s')}  {' '.join(parts)}  {status}"
        )


# ---- main -------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument(
        "--scenarios",
        default=",".join(s for s in SCENARIOS if s != "idle"),
        help="comma-separated scenario list (default: all but idle)",
    )
    ap.add_argument(
        "--work",
        type=Path,
        default=_ui_fixtures.ROOT / "ui-bench",
        help="fixture and artifact root (default: .context/ui-check/ui-bench)",
    )
    ap.add_argument("--soak", type=float, default=60.0, help="idle window in seconds")
    ap.add_argument("--soak-interval", type=float, default=5.0)
    bc.add_common_args(ap, "load")
    args = ap.parse_args(argv)
    bc.validate_common_args(ap, args)
    scenarios = [s.strip() for s in args.scenarios.split(",") if s.strip()]
    unknown = [s for s in scenarios if s not in SCENARIOS]
    if unknown:
        print(f"unknown scenarios: {', '.join(unknown)} (know: {', '.join(SCENARIOS)})")
        return bc.EXIT_USAGE

    import profiling

    profiling.enable()  # the page reads CLIPGEN_CONFIG.profiling from the config payload
    work = args.work
    try:
        fixtures = build_fixtures(work)
    except _ui_fixtures.UiUnavailable as exc:
        print(exc, file=sys.stderr)
        return bc.EXIT_USAGE
    counts = fixture_counts(work)
    print(
        f"fixtures: {counts['rows']} rows × {SHEET_PARTICIPANTS}, "
        f"{sum(counts['segments'].values())} segments, {counts['events']} events"
    )
    (work / "runs").mkdir(parents=True, exist_ok=True)
    (work / "shots").mkdir(parents=True, exist_ok=True)

    rows: dict[str, dict[str, Any]] = {}
    for scenario in scenarios:
        fx = fixtures[scenario]
        results: list[dict[str, Any]] = []
        try:
            with _ui_session.ui_session(
                input_dir=fx["input_dir"],
                output_dir=fx["output_dir"],
                sheet=fx["sheet"],
            ) as session:
                for i in range(max(1, args.runs)):
                    shot_path = work / "shots" / f"{scenario}-{i + 1}.png"
                    run = run_once(
                        scenario,
                        session,
                        counts,
                        shot_path,
                        args.soak,
                        args.soak_interval,
                    )
                    (work / "runs" / f"{scenario}-{i + 1}.json").write_text(
                        json.dumps(run, indent=2, default=str), encoding="utf-8"
                    )
                    results.append(run)
                    if not run["valid"]:
                        for reason in run["reasons"]:
                            print(f"  ! {scenario}: {reason}")
                        break
        except _ui_fixtures.UiUnavailable as exc:
            print(exc, file=sys.stderr)
            return bc.EXIT_USAGE
        samples = [summarize(scenario, r) for r in results if r["valid"]]
        metrics = metrics_for(scenario, samples) or ("elapsed_s",)
        row = bc.build_row(
            [dict(r, doc=r) for r in results],
            lambda doc, s=scenario: summarize(s, doc),
            metrics,
        )
        rows[scenario] = row
        if row["samples"]:
            print(
                f"  {scenario}: elapsed {row['stats']['elapsed_s']['median']:.2f}s median"
            )

    meta = {
        "fixture": {
            "spec": {
                "rows": SHEET_ROWS,
                "participants": SHEET_PARTICIPANTS,
                "segments": SEGMENTS,
                "events": EVENTS,
            },
            "probed": counts,
        },
        "workload": {"soak": args.soak, "soak_interval": args.soak_interval},
        "env": profiling.environment(),
        "runs": args.runs,
        "metric": args.fail_metric,
    }
    metric = "elapsed_s" if args.fail_metric == "elapsed" else "load_ms"
    findings: list[tuple[str, str, str]] = []
    if args.compare:
        baseline, problems = bc.load_baseline(
            args.compare, "scenarios", meta, allow_env=args.allow_env_mismatch
        )
        if baseline is None:
            for line in problems:
                print(f"incomparable: {line}")
            return bc.EXIT_INCOMPARABLE
        for scenario, row in rows.items():
            findings += bc.compare(
                {scenario: row},
                baseline,
                metric=metric,
                pct=args.fail_on,
                abs_s=args.fail_abs,
                count_key=COUNT_KEY[scenario],
            )
    else:
        findings = [
            ("invalid", name, "; ".join(row["reasons"]))
            for name, row in rows.items()
            if not row["valid"]
        ]
    print()
    print_table(rows, findings, metric)
    bc.report_findings(findings)
    if args.save:
        bc.save_results(args.save, meta, "scenarios", rows)
    return bc.exit_code(findings)


if __name__ == "__main__":
    raise SystemExit(main())
