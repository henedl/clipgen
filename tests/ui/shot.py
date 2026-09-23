"""Open one page, photograph it, and optionally run JavaScript inside it.

The six-page smoke costs ~30 s cold, which is the wrong loop for iterating on a
single surface. This does one page in a few seconds, and adds the thing the old
"paste this DevTools snippet to the human" workflow was standing in for:
``--eval`` runs arbitrary JS in the live page and prints what it returns. Read
computed styles, count rendered nodes, dump ``state``, call
``el.getAnimations()`` — whatever you would have asked someone else to type.

    uv run --extra ui python tests/ui/shot.py studio
    uv run --extra ui python tests/ui/shot.py studio --selector "#sheetGrid"
    uv run --extra ui python tests/ui/shot.py studio --theme light
    uv run --extra ui python tests/ui/shot.py studio --state settings
    uv run --extra ui python tests/ui/shot.py screenspace --all-states
    uv run --extra ui python tests/ui/shot.py transcripts \
        --eval "return document.querySelectorAll('.pill-wrap').length"
    uv run --extra ui python tests/ui/shot.py overview --eval-file /tmp/probe.js --wait 2000
    uv run --extra ui python tests/ui/shot.py studio --perf --wait 5000
    uv run --extra ui python tests/ui/shot.py studio --perf --trace /tmp/studio.trace.json
    uv run --extra ui python tests/ui/shot.py studio --perf --sheet /tmp/gridbench.xlsx
    uv run --extra ui python tests/ui/shot.py transcripts --perf --output /tmp/tsbench

``--perf`` is the DevTools half of the profiling story (backend half:
``--profile`` / agents/skills/profile/SKILL.md). It turns on server-side
profiling for the in-process Flask app (so ``CLIPGEN_CONFIG.profiling`` reaches
the page and ``window.__clipgenPerf`` populates), captures CDP
``Performance.getMetrics`` plus navigation/resource timing, and prints one
grep-able ``perf | `` line per fact — including one ``perf | api <path>`` line
per API fetch in start order, the boot waterfall — plus a single
``perf-json:`` line for machine parsing; the server's own ``profile | ``
report (route timings) prints at process exit. ``--trace`` writes a Chrome trace-event JSON (open in
Perfetto, or parse it) covering navigation through capture. ``--full-chromium``
prefers the full Chromium build over the headless shell, whose paint metrics
are only indicative.

``--theme light`` is worth reaching for: light is a shipped feature reachable from
every page's theme toggle, and until it was added here nothing in the harness had
ever rendered it. ``--state`` and ``--all-states`` (see ``_ui_states``) reach the
modals and tabs the six-page smoke never opens — ``--all-states`` drives them all
from a single boot, which is the one thing looping this script cannot do.

Not a test: ``norecursedirs = ui`` plus the non-``test_`` filename keep pytest
away from it. The JS it runs only ever reaches a loopback server serving
generated fixture data.
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))  # the _ui_* helpers
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "source"))

import _ui_browser
import _ui_fixtures
import _ui_pages
import _ui_session
import _ui_states


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="shot.py",
        description="Screenshot and probe one clipgen page in a headless browser.",
    )
    parser.add_argument("page", choices=sorted(_ui_pages.PAGES))
    parser.add_argument(
        "--selector",
        help="Clip the screenshot to this element instead of the full page.",
    )
    parser.add_argument(
        "--eval",
        dest="evaluate",
        help="JavaScript body to run in the page; its return value is printed as JSON.",
    )
    parser.add_argument(
        "--eval-file",
        type=Path,
        help="Read the --eval body from a file (for anything longer than a shell line).",
    )
    parser.add_argument(
        "--wait", type=int, default=0, help="Extra settle milliseconds before capture."
    )
    parser.add_argument("--out", type=Path, help="Override the screenshot path.")
    parser.add_argument(
        "--viewport", default="1600x1000", help="Viewport size, e.g. 1280x800."
    )
    parser.add_argument(
        "--theme",
        choices=_ui_session.THEMES,
        default="dark",
        help="Boot the page in this theme. Light is a shipped feature the "
        "harness otherwise never renders.",
    )
    parser.add_argument(
        "--state",
        help="Drive into one UI state before capturing, e.g. settings, "
        "cheatsheet, palette, start, tab:map, tool:color. Pass an unknown name "
        "to be told which states this page has.",
    )
    parser.add_argument(
        "--all-states",
        action="store_true",
        help="Capture every reachable state on this page from one boot, to "
        "<page>-<state>.png. Unreachable states are reported, not skipped.",
    )
    parser.add_argument(
        "--perf",
        action="store_true",
        help="Capture performance data: CDP Performance.getMetrics, navigation/"
        "resource timing, and the page's clipgenPerf accumulator (server-side "
        "profiling is enabled so it populates). Prints 'perf | ' lines plus one "
        "'perf-json:' line.",
    )
    parser.add_argument(
        "--soak",
        type=float,
        default=0.0,
        metavar="SECONDS",
        help="With --perf: keep the page open this long after the first capture, "
        "sampling every --soak-interval seconds. Prints one 'perf | soak.*' "
        "trajectory per metric (heap, DOM nodes, listeners, transfer bytes, "
        "poll ticks) with its slope per minute. Sustained growth is a signal "
        "to investigate, not proof of a leak.",
    )
    parser.add_argument(
        "--soak-interval",
        type=float,
        default=5.0,
        metavar="SECONDS",
        help="Seconds between soak samples (default 5).",
    )
    parser.add_argument(
        "--soak-hidden",
        action="store_true",
        help="Run the second half of the soak with the page emulated as a "
        "background tab, so pollers that should pause are counted separately.",
    )
    parser.add_argument(
        "--perf-output",
        type=Path,
        metavar="PATH",
        help="Write the --perf data (the perf-json object plus page errors and "
        "the screenshot path) as JSON to PATH.",
    )
    parser.add_argument(
        "--trace",
        type=Path,
        help="Write a Chrome trace-event JSON (Perfetto-compatible) covering "
        "navigation through capture to this path.",
    )
    parser.add_argument(
        "--full-chromium",
        action="store_true",
        help="Prefer the full Chromium build over the headless shell (paint/"
        "compositor metrics on the shell are only indicative).",
    )
    parser.add_argument(
        "--sheet",
        type=Path,
        help="Workbook to load instead of the 6-row UI fixture. Use the "
        "gridbench recipe in agents/skills/profile/SKILL.md — the fixture "
        "is too small for studio.renderGrid to mean anything.",
    )
    parser.add_argument(
        "--input",
        dest="input_dir",
        type=Path,
        help="Override config.INPUT_DIR (source videos). Defaults to the "
        "fixture input dir, or --sheet's parent.",
    )
    parser.add_argument(
        "--output",
        dest="output_dir",
        type=Path,
        help="Override config.OUTPUT_DIR (manifests). Point this at a "
        "synthesized clipgen.json or a --ss-task output dir "
        "to measure transcripts.renderSegments / screenspace.renderResults "
        "on a real-sized list.",
    )
    return parser.parse_args(argv)


# Explicit field mapping throughout: Playwright's JS→Python serializer walks own
# enumerable properties only, so a raw PerformanceEntry would come back as {}.
_PERF_SNIPPET = """
var nav = performance.getEntriesByType("navigation").map(function (e) {
  return {
    duration: e.duration,
    domContentLoaded: e.domContentLoadedEventEnd,
    loadEvent: e.loadEventEnd,
    transferSize: e.transferSize,
  };
});
var resources = performance.getEntriesByType("resource");
var totalTransfer = 0;
for (var i = 0; i < resources.length; i++) {
  totalTransfer += resources[i].transferSize || 0;
}
var slowest = resources
  .slice()
  .sort(function (a, b) { return b.duration - a.duration; })
  .slice(0, 5)
  .map(function (e) {
    return { name: e.name.split("/").slice(-2).join("/"), durationMs: e.duration };
  });
// The boot waterfall: every API fetch in start order, so a serial chain or a
// slow cold route shows as a line instead of hiding inside resources.count.
var api = resources
  .filter(function (e) { return e.name.indexOf("/api/") >= 0; })
  .sort(function (a, b) { return a.startTime - b.startTime; })
  .map(function (e) {
    var path = e.name.replace(/^[a-z]+:[/][/][^/]+/, "");
    return {
      path: path,
      startMs: e.startTime,
      durationMs: e.duration,
      transferSize: e.transferSize || 0,
    };
  });
return {
  navigation: nav.length ? nav[0] : null,
  resources: { count: resources.length, transferSize: totalTransfer, slowest: slowest, api: api },
  clipgenPerf: window.clipgenPerf ? window.clipgenPerf.snapshot() : null,
};
"""

# CDP metrics worth a line each; the rest still land in perf-json.
_CDP_METRIC_ALLOWLIST = {
    "LayoutCount",
    "RecalcStyleCount",
    "LayoutDuration",
    "RecalcStyleDuration",
    "ScriptDuration",
    "TaskDuration",
    "JSHeapUsedSize",
    "JSHeapTotalSize",
    "Nodes",
    "JSEventListeners",
}


# Metrics the soak and the per-line report depend on; a build without one
# reports it as missing rather than as 0.
_EXPECTED_CDP = ("Nodes", "JSEventListeners", "JSHeapUsedSize", "Documents")


def _collect_perf(page: Any, cdp: Any) -> dict[str, Any]:
    """Gather CDP metrics + page timing into one plain dict."""
    data: dict[str, Any] = {}
    metrics = cdp.send("Performance.getMetrics").get("metrics", [])
    data["cdp"] = {m["name"]: m["value"] for m in metrics}
    data["page"] = page.evaluate(f"() => {{ {_PERF_SNIPPET} }}")
    data["unsupported"] = [name for name in _EXPECTED_CDP if name not in data["cdp"]]
    cg = (data["page"] or {}).get("clipgenPerf") or {}
    if (cg.get("supported") or {}).get("longtask") is False:
        data["unsupported"].append("longtask")
    return data


def _soak_sample(page: Any, cdp: Any, t: float, hidden: bool) -> dict[str, Any]:
    """One soak sample: the leak-prone counters plus poll ticks and transfer."""
    snap = _collect_perf(page, cdp)
    cg = (snap.get("page") or {}).get("clipgenPerf") or {}
    measures = cg.get("measures") or {}
    return {
        "t": t,
        "hidden": hidden,
        "cdp": {name: snap["cdp"].get(name) for name in _EXPECTED_CDP},
        "polls": {
            label: m["n"] for label, m in measures.items() if label.startswith("poll.")
        },
        "longtasks": (cg.get("longtasks") or {}).get("count", 0),
        "transferSize": ((snap.get("page") or {}).get("resources") or {}).get(
            "transferSize", 0
        ),
    }


def slope_per_minute(points: list[tuple[float, float]]) -> float | None:
    """Least-squares slope of *points* (seconds, value) scaled to per minute."""
    if len(points) < 2:
        return None
    n = len(points)
    mean_t = sum(t for t, _ in points) / n
    mean_v = sum(v for _, v in points) / n
    var = sum((t - mean_t) ** 2 for t, _ in points)
    if not var:
        return None
    cov = sum((t - mean_t) * (v - mean_v) for t, v in points)
    return cov / var * 60.0


def soak_rates(samples: list[dict[str, Any]]) -> dict[str, Any]:
    """Per-minute slopes for every counter, and poll ticks split by visibility."""
    rates: dict[str, Any] = {}
    for name in _EXPECTED_CDP:
        points = [
            (s["t"], s["cdp"][name]) for s in samples if s["cdp"].get(name) is not None
        ]
        rates[name] = slope_per_minute(points)
    rates["transferSize"] = slope_per_minute(
        [(s["t"], s["transferSize"]) for s in samples]
    )
    rates["longtasks"] = slope_per_minute([(s["t"], s["longtasks"]) for s in samples])
    labels = sorted({label for s in samples for label in s["polls"]})
    polls: dict[str, dict[str, float | None]] = {}
    for label in labels:
        by_vis: dict[str, float | None] = {}
        for hidden in (False, True):
            points = [
                (s["t"], s["polls"].get(label, 0))
                for s in samples
                if s["hidden"] == hidden
            ]
            by_vis["hidden" if hidden else "visible"] = slope_per_minute(points)
        polls[label] = by_vis
    rates["polls"] = polls
    return rates


def run_soak(
    page: Any, cdp: Any, seconds: float, interval: float, hidden_half: bool
) -> dict[str, Any]:
    """Sample every *interval* s for *seconds*; optionally hide the second half."""
    import time

    import _ui_pages

    samples: list[dict[str, Any]] = []
    t0 = time.perf_counter()
    hidden = False
    samples.append(_soak_sample(page, cdp, 0.0, hidden))
    while True:
        elapsed = time.perf_counter() - t0
        if elapsed >= seconds:
            break
        if hidden_half and not hidden and elapsed >= seconds / 2:
            hidden = _ui_pages.set_hidden(page, True)
        page.wait_for_timeout(min(interval, seconds - elapsed) * 1000)
        samples.append(_soak_sample(page, cdp, time.perf_counter() - t0, hidden))
    if hidden:
        _ui_pages.set_hidden(page, False)
    return {
        "seconds": seconds,
        "interval": interval,
        "hidden_half": hidden_half,
        "samples": samples,
        "rates": soak_rates(samples),
    }


def _print_perf(data: dict[str, Any]) -> None:
    for name in sorted(_CDP_METRIC_ALLOWLIST & set(data["cdp"])):
        print(f"perf | cdp.{name} {data['cdp'][name]}")
    page_data = data.get("page") or {}
    nav = page_data.get("navigation") or {}
    for key in ("domContentLoaded", "loadEvent", "duration"):
        if key in nav:
            print(f"perf | nav.{key} {nav[key]:.1f}ms")
    res = page_data.get("resources") or {}
    if res:
        print(f"perf | resources.count {res.get('count', 0)}")
        print(f"perf | resources.transferSize {res.get('transferSize', 0)}")
    for entry in res.get("api") or []:
        print(
            f"perf | api {entry['path']} start={entry['startMs']:.0f}ms "
            f"dur={entry['durationMs']:.1f}ms bytes={entry['transferSize']}"
        )
    cg = page_data.get("clipgenPerf") or {}
    for label, m in sorted((cg.get("measures") or {}).items()):
        print(f"perf | {label} {m['totalMs']:.1f}ms n={m['n']} max={m['maxMs']:.1f}ms")
    lt = cg.get("longtasks") or {}
    if lt.get("count"):
        print(
            f"perf | longtasks {lt['totalMs']:.1f}ms n={lt['count']} "
            f"max={lt['maxMs']:.1f}ms"
        )
    for name in data.get("unsupported") or []:
        print(f"perf | {name} unsupported")
    if data.get("soak"):
        _print_soak(data["soak"])
    print("perf-json: " + json.dumps(data, ensure_ascii=False, default=str))


def _fmt_rate(value: float | None) -> str:
    return "n/a" if value is None else f"{value:+.1f}/min"


def _print_soak(soak: dict[str, Any]) -> None:
    """Trajectory lines: first -> last with the fitted slope per minute.

    Growth here is a signal to investigate — a poller re-rendering into the
    DOM, listeners re-bound per tick, retained payloads — not proof of a leak;
    a toast or a clock moves a few nodes too.
    """
    samples = soak["samples"]
    rates = soak["rates"]
    first, last = samples[0], samples[-1]
    secs = soak["seconds"]
    for name in _EXPECTED_CDP:
        b, a = first["cdp"].get(name), last["cdp"].get(name)
        if b is None or a is None:
            continue
        print(
            f"perf | soak.{name} {b:.0f} -> {a:.0f} ({_fmt_rate(rates.get(name))}) "
            f"over {secs:g}s"
        )
    print(
        f"perf | soak.transferSize {first['transferSize']} -> {last['transferSize']} "
        f"({_fmt_rate(rates.get('transferSize'))})"
    )
    print(
        f"perf | soak.longtasks {first['longtasks']} -> {last['longtasks']} "
        f"({_fmt_rate(rates.get('longtasks'))})"
    )
    for label, by_vis in sorted(rates.get("polls", {}).items()):
        n0, n1 = first["polls"].get(label, 0), last["polls"].get(label, 0)
        parts = [f"visible {_fmt_rate(by_vis.get('visible'))}"]
        if soak.get("hidden_half"):
            parts.append(f"hidden {_fmt_rate(by_vis.get('hidden'))}")
        print(f"perf | soak.{label} n={n0} -> {n1} ({', '.join(parts)})")
    print(
        "perf | soak.samples "
        + " ".join(
            f"{s['t']:.0f}s:{'H' if s['hidden'] else 'V'}:"
            f"{s['cdp'].get('JSHeapUsedSize') or 0:.0f}"
            for s in samples
        )
    )


def _viewport(raw: str) -> dict[str, int]:
    try:
        width, height = (int(part) for part in raw.lower().split("x", 1))
    except ValueError:
        raise SystemExit(f"--viewport expects WIDTHxHEIGHT, got {raw!r}") from None
    return {"width": width, "height": height}


def _capture(page: Any, args: argparse.Namespace, out: Path) -> None:
    if args.selector:
        page.locator(args.selector).first.screenshot(path=str(out))
    else:
        page.screenshot(path=str(out), full_page=True)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    viewport = _viewport(args.viewport)
    body = args.evaluate
    if args.eval_file is not None:
        body = args.eval_file.read_text(encoding="utf-8")

    out = args.out or (_ui_fixtures.SHOT_DIR / f"{args.page}.png")
    out.parent.mkdir(parents=True, exist_ok=True)

    if args.perf:
        # Server-side profiling must be on before the app is built: the page
        # reads CLIPGEN_CONFIG.profiling from the config payload, and the atexit
        # report prints the route timings when this process exits.
        import profiling

        profiling.enable()

    log = _ui_pages.PageLog()
    result: Any = None
    eval_error: str | None = None
    perf_data: dict[str, Any] | None = None
    shots: list[Path] = []
    states: list[_ui_states.StateResult] = []
    state_problem: str = ""
    try:
        with _ui_session.ui_session(
            viewport=viewport,
            theme=args.theme,
            full_chromium=args.full_chromium,
            input_dir=args.input_dir,
            output_dir=args.output_dir,
            sheet=args.sheet,
        ) as session:
            page = session.context.new_page()
            _ui_pages.wire_listeners(page, log)
            cdp = None
            if args.perf:
                # Enable before navigation so getMetrics covers the page load.
                cdp = session.context.new_cdp_session(page)
                cdp.send("Performance.enable")
            if args.trace is not None:
                args.trace.parent.mkdir(parents=True, exist_ok=True)
                session.context.browser.start_tracing(
                    page=page, path=str(args.trace), screenshots=False
                )
            try:
                _ui_pages.open_and_settle(
                    page, session.origin, args.page, log, args.wait
                )
                if args.all_states:
                    # Capture the boot state too: --all-states should mean all.
                    _capture(page, args, out)
                    shots.append(out)
                    for state in _ui_states.each_state(page, args.page):
                        states.append(state)
                        if not state.reached:
                            continue
                        shot = out.with_name(
                            f"{out.stem}-{state.name.replace(':', '-')}{out.suffix}"
                        )
                        _capture(page, args, shot)
                        shots.append(shot)
                elif args.state:
                    state = _ui_states.enter_named(page, args.page, args.state)
                    states.append(state)
                    if not state.reached:
                        state_problem = state.detail
                if body:
                    try:
                        result = page.evaluate(f"() => {{ {body} }}")
                    except _ui_browser.playwright_error() as exc:
                        # A throwing snippet is a normal outcome when probing, not
                        # a harness crash. Report the JS error, keep the screenshot.
                        eval_error = str(exc).splitlines()[0]
                if cdp is not None:
                    try:
                        perf_data = _collect_perf(page, cdp)
                        if args.soak > 0:
                            perf_data["soak"] = run_soak(
                                page,
                                cdp,
                                args.soak,
                                args.soak_interval,
                                args.soak_hidden,
                            )
                    except _ui_browser.playwright_error() as exc:
                        eval_error = eval_error or str(exc).splitlines()[0]
            finally:
                if not args.all_states:
                    _capture(page, args, out)
                    shots.append(out)
                if args.trace is not None:
                    try:
                        session.context.browser.stop_tracing()
                    except _ui_browser.playwright_error():
                        pass  # trace loss is a diagnostics gap, not a run failure
    except _ui_fixtures.UiUnavailable as exc:
        print(exc, file=sys.stderr)
        return 1

    for shot in shots:
        print(f"screenshot: {shot.resolve()}")
    # Report every state we tried, reached or not. A silently skipped state reads
    # as covered, which is the failure mode this whole harness exists to avoid.
    for state in states:
        mark = "ok  " if state.reached else "MISS"
        suffix = f" — {state.detail}" if state.detail else ""
        print(f"state: {mark} {state.name}{suffix}")
    if body and eval_error is None:
        print("eval: " + json.dumps(result, ensure_ascii=False, default=str))
    if perf_data is not None:
        _print_perf(perf_data)
        if args.perf_output is not None:
            perf_data["errors"] = {
                "page_errors": log.page_errors,
                "console_errors": log.console_errors,
                "timeout": log.timeout,
            }
            perf_data["screenshots"] = [str(shot.resolve()) for shot in shots]
            args.perf_output.parent.mkdir(parents=True, exist_ok=True)
            args.perf_output.write_text(
                json.dumps(perf_data, indent=2, ensure_ascii=False, default=str),
                encoding="utf-8",
            )
            print(f"perf-output: {args.perf_output.resolve()}")
    if args.trace is not None and args.trace.exists():
        print(f"trace: {args.trace.resolve()}")
    if eval_error is not None:
        print(f"eval failed: {eval_error}", file=sys.stderr)
    for text in log.page_errors:
        print(f"[pageerror] {text}", file=sys.stderr)
    for text in log.console_errors:
        print(f"[console]   {text}", file=sys.stderr)
    if log.timeout:
        print(f"[timeout]   {log.timeout}", file=sys.stderr)
    if state_problem:
        print(f"state failed: {state_problem}", file=sys.stderr)
    # An explicitly requested --state that could not be reached is a failure; an
    # unreachable state during --all-states is reported above and is not, since
    # some states legitimately do not exist in this fixture.
    return 1 if (log.fatal or eval_error is not None or state_problem) else 0


if __name__ == "__main__":
    raise SystemExit(main())
