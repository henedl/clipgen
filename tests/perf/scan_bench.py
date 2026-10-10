"""Repeatable Screenspace scan benchmark over the standard tool sweep.

Runs each tool through the real CLI (`uv run clipgen.py --ss-task ...
--profile-output`) against the deterministic testsrc fixture from
agents/skills/profile/SKILL.md, reads the JSON report, and prints one table
row per tool. This replaces the hand-rolled per-tool command loops that
every profiling session rebuilt (and that zsh's no-word-split quoting broke
mid-session at least once) and makes before/after comparison a diff of two
JSON files instead of an eyeball job.

Usage (from the repo root):

    uv run python tests/perf/scan_bench.py                     # sweep + table
    uv run python tests/perf/scan_bench.py --save base.json    # snapshot
    uv run python tests/perf/scan_bench.py --compare base.json --fail-on 10
    uv run python tests/perf/scan_bench.py --tools color,text --runs 5
    uv run python tests/perf/scan_bench.py --tier session --tools color,change
    uv run python tests/perf/scan_bench.py --source ~/rec/P07.mp4 --tools change
    uv run python tests/perf/scan_bench.py --scale 600,1800,3600 --tier session

Tiers: ``smoke`` (default) is the short regression gate — 120 s 720p
testsrc, 1 s keyframes, 0.1 s interval. ``session`` is what users run: an
hour of OBS-like 1080p (8.3 s keyframes, screen-recording content) at the
default 1.0 s interval, one repetition. ``--source`` benches a real
recording in place; ``--scale`` fits cost against fixture length. Every row
reports ``xrt`` (media seconds per wall second) and the decode path taken
(``full`` vs ``key``frame-only), the two numbers a 120 s run cannot predict.

The fixture video is built on demand (ffmpeg testsrc: constant motion, so
phash-skip never hides the callback) and rebuilt when its probed duration,
size, frame rate or audio presence differs from `FIXTURE_SPEC` — a leftover
120 s clip would otherwise ignore `--duration 15` and poison `--compare`.
Each tool writes to its own wiped output dir so a cached manifest can never
absorb the scan. `--runs N` (default 3) keeps every fresh-process sample and
compares medians; a repetition that exits non-zero, records no callback, or
leaves the task short of `completed` invalidates the tool's row. Template
and Shape use a fixed top-left 20% region for a non-degenerate reference and
bounded search workload. `text` is excluded from the default sweep: OCR is an
order of magnitude slower than every other tool and pins ~0.8 GB of RSS per
pooled OCR engine (measured 3.3 GB at the default auto pool of 4).

Row reduction and validity are unit-tested in test_scan_bench_parse.py; the
shared compare/aggregate logic lives in bench_common.py.
"""

from __future__ import annotations

import argparse
import functools
import json
import shutil
import sys
from pathlib import Path
from typing import Any

import bench_common as bc
import bench_fixtures as bf

# Missing required flags refuse the task and report only ffprobe.run.
TOOL_FLAGS: dict[str, list[str]] = {
    "color": ["--ss-target-color", "#FF0000", "--ss-tolerance", "20,30,30"],
    "change": ["--ss-threshold", "0.05"],
    "similarity": ["--ss-reference-timestamp", "1", "--ss-threshold", "0.5"],
    "inactivity": ["--ss-threshold", "10"],
    "scene": ["--ss-scene-ref", "menu:1"],
    "flow": ["--ss-threshold", "2"],
    "template": ["--ss-reference-timestamp", "1", "--ss-threshold", "0.7"],
    "shape": ["--ss-reference-timestamp", "1", "--ss-threshold", "0.55"],
    "attention": [],
    "text": ["--ss-text", "00:00:01"],
}
DEFAULT_TOOLS = [t for t in TOOL_FLAGS if t != "text"]

BENCH_REGION = {
    "x": 0.0,
    "y": 0.0,
    "w": 0.2,
    "h": 0.2,
    "source_width": 1280,
    "source_height": 720,
}
REGION_TOOLS = {"template", "shape"}
METRICS = ("elapsed_s", "callback_s")

TIERS: dict[str, dict[str, Any]] = {
    "smoke": {"duration": 120, "interval": 0.1, "input": Path("/tmp/ssbench")},
    "session": {
        "duration": 3600,
        "interval": 1.0,
        "runs": 1,
        "input": Path("/tmp/ssbench-session"),
    },
}


def fixture_spec(duration: int, tier: str = "smoke") -> dict[str, Any]:
    if tier == "session":
        return bc.session_spec(duration, audio="none")
    return {
        "duration": duration,
        "width": 1280,
        "height": 720,
        "fps": 30.0,
        "has_audio": False,
    }


def region_for(probed: dict[str, Any] | None) -> dict[str, Any]:
    """The bench region, sized to the fixture; smoke keeps the 1280x720 literal."""
    if not probed or (probed["width"], probed["height"]) == (1280, 720):
        return BENCH_REGION
    return dict(
        BENCH_REGION, source_width=probed["width"], source_height=probed["height"]
    )


def summarize(tool: str, doc: dict[str, Any]) -> dict[str, float]:
    """Reduce one run's JSON report to the per-tool comparison row."""
    labels = doc.get("labels", {})

    def get(label: str) -> dict[str, float]:
        return labels.get(label, {"seconds": 0.0, "count": 0})

    callback = get(f"scan.callback.{tool}")
    frames = int(callback["count"])
    heatmap = sum(
        get(label)["seconds"] for label in ("heatmap.gifs", "heatmap.grid_layers")
    )
    return {
        "callback_s": callback["seconds"],
        "callback_avg_ms": callback["seconds"] / frames * 1000 if frames else 0.0,
        "frames": frames,
        "decode_s": get("scan.decode_wait")["seconds"],
        "filter_s": get("scan.fast_filter")["seconds"],
        "heatmap_s": heatmap,
        "peak_rss_mb": doc.get("peak_rss_mb") or 0.0,
        "keyframe_scans": int(get("scan.decode_mode.keyframe")["count"]),
        "full_scans": int(get("scan.decode_mode.full")["count"]),
    }


def decode_mode(sample: dict[str, Any]) -> str:
    """``key``, ``full``, ``mixed`` or ``-`` (a build without the label)."""
    key, full = sample.get("keyframe_scans", 0), sample.get("full_scans", 0)
    if key and full:
        return "mixed"
    return "key" if key else "full" if full else "-"


def task_completed(out_dir: Path) -> bool:
    """True when the run's manifest holds a task that reached ``completed``."""
    try:
        doc = json.loads((out_dir / "clipgen.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return False
    tasks = doc.get("screenspace", {}).get("tasks", [])
    return any(task.get("status") == "completed" for task in tasks)


def ensure_fixture(input_dir: Path, duration: int, tier: str = "smoke") -> Path:
    """Build the benchmark video, rebuilding if it does not match the spec."""
    video = input_dir / "bench_P01.mp4"

    def build() -> None:
        if tier == "session":
            bf.make_long_video(
                video, duration=duration, audio="none", **bc.SESSION_VIDEO
            )
        else:
            bf.make_testsrc_video(video, duration=duration)

    bc.ensure_video(video, fixture_spec(duration, tier), build)
    return video


def run_tool(
    tool: str,
    input_dir: Path,
    out_root: Path,
    interval: float,
    deep: bool = False,
    *,
    region: dict[str, Any] | None = None,
    timeout: float = 1800,
) -> dict[str, Any]:
    """One fresh-process run of *tool* into a wiped output dir.

    With *deep*, attach ``--profile-deep scan.callback.<tool>`` and print the
    pstats block verbatim — the harness owns the canonical flag set, so this
    replaces hand-rolling the tool's CLI command just to drill into it. A
    deep run is a diagnostic and is never saved or compared.
    """
    out_dir = out_root / f"bench-{tool}"
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)
    region_args: list[str] = []
    if tool in REGION_TOOLS:
        manifest = {
            "screenspace": {
                "regions": {"bench": region or BENCH_REGION},
                "tasks": [],
                "events": [],
                "stashes": [],
            }
        }
        (out_dir / "clipgen.json").write_text(json.dumps(manifest))
        region_args = ["bench"]
    args = [
        "--ss-task",
        tool,
        "P01",
        *region_args,
        *TOOL_FLAGS[tool],
        "--ss-interval",
        str(interval),
        "-i",
        str(input_dir),
        "-o",
        str(out_dir),
    ]
    if deep:
        args += ["--profile-deep", f"scan.callback.{tool}"]
    result = bc.run_clipgen(
        args, out_dir, required_labels=[f"scan.callback.{tool}"], timeout=timeout
    )
    if deep:
        result["reasons"] = [r for r in result["reasons"] if r != "deep run"]
        result["valid"] = not result["reasons"]
    if result["valid"] and not task_completed(out_dir):
        result["valid"] = False
        result["reasons"].append("task did not complete")
    if not result["valid"]:
        bc.print_failure(tool, result)
    if deep:
        in_block = False
        for line in result["output"].splitlines():
            if line.startswith("profile-deep |"):
                in_block = True
            if in_block:
                print(line)
    return result


def print_table(
    rows: dict[str, dict[str, Any]],
    baseline: dict[str, dict[str, Any]] | None,
    findings: list[tuple[str, str, str]],
    metric: str = "elapsed_s",
) -> None:
    delta_hdr = f"  Δ{metric.removesuffix('_s'):<8}" if baseline else ""
    print(
        f"{'tool':<12}{'elapsed':>9}{'min':>9}{'mad':>8}{'callback':>10}{'avg':>9}"
        f"{'frames':>8}{'decode':>9}{'heatmap':>9}{'rss':>9}{'xrt':>8}{'path':>6}"
        f"{delta_hdr}  status"
    )
    for tool, row in rows.items():
        status = bc.status_of(tool, findings)
        if not row["samples"]:
            print(f"{tool:<12}{'':>106}  {status}")
            continue

        def med(key: str, samples: list[dict[str, Any]] = row["samples"]) -> float:
            return bc.aggregate(samples, key)["median"]

        cb = med("callback_s")
        frames = int(med("frames"))
        line = (
            f"{tool:<12}{bc.stat_cells(row, 'elapsed_s')}"
            f"{cb:>9.3f}s{(cb / frames * 1000 if frames else 0.0):>7.1f}ms"
            f"{frames:>8d}{med('decode_s'):>8.3f}s{med('heatmap_s'):>8.3f}s"
            f"{med('peak_rss_mb'):>7.0f}MB{med('xrt'):>7.1f}x"
            f"{decode_mode(row['samples'][0]):>6}"
        )
        if baseline:
            base = baseline.get(tool)
            if base and base.get("samples"):
                pct = bc.delta_pct(med(metric), base["stats"][metric]["median"])
                line += f"  {pct:>+9.1f}%" if pct is not None else "  (no base)"
            else:
                line += "   (no base)"
        print(f"{line}  {status}")


def sweep(
    args: argparse.Namespace,
    tools: list[str],
    input_dir: Path,
    out_root: Path,
    media_s: float,
    region: dict[str, Any],
) -> dict[str, Any]:
    """Every tool's row over one fixture; the first run's ``env`` rides as ``_env``."""
    rows: dict[str, Any] = {}
    env: dict[str, Any] = {}
    timeout = bc.timeout_for(media_s, args.timeout)
    for tool in tools:
        results = []
        for _ in range(max(1, args.runs)):
            result = run_tool(
                tool,
                input_dir,
                out_root,
                args.interval,
                args.deep,
                region=region,
                timeout=timeout,
            )
            results.append(result)
            env = env or result["doc"].get("env", {})
            if not result["valid"]:
                break  # one bad repetition voids the row; no point repeating
        rows[tool] = bc.build_row(
            results, functools.partial(summarize, tool), METRICS, media_s
        )
        if rows[tool]["samples"]:
            st = rows[tool]["stats"]
            sample = rows[tool]["samples"][0]
            print(
                f"  {tool}: elapsed {st['elapsed_s']['median']:.3f}s median"
                f" (callback {st['callback_s']['median']:.3f}s,"
                f" {sample['xrt']:.1f}x realtime, {decode_mode(sample)} decode)"
            )
    rows["_env"] = env
    return rows


def run_scale(args: argparse.Namespace, tools: list[str]) -> int:
    """Run the sweep at each ``--scale`` duration and fit cost against length."""
    try:
        durations = bc.parse_scale(args.scale)
    except ValueError as exc:
        print(exc)
        return bc.EXIT_USAGE
    points: dict[str, list[tuple[float, float]]] = {t: [] for t in tools}
    bad = False
    for duration in durations:
        input_dir = args.input / f"scale-{duration}"
        video = ensure_fixture(input_dir, duration, args.tier)
        probed = bc.probe_fixture(video)
        media_s = float(probed["duration"]) if probed else float(duration)
        print(f"fixture {video.name}  {media_s:.0f}s  interval={args.interval}")
        rows = sweep(
            args, tools, input_dir, input_dir / "bench-out", media_s, region_for(probed)
        )
        rows.pop("_env", None)
        for tool, row in rows.items():
            if row["samples"]:
                points[tool].append((media_s, row["stats"]["elapsed_s"]["median"]))
            else:
                bad = True
    bc.print_scaling(points)
    return bc.EXIT_INVALID if bad else bc.EXIT_OK


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument(
        "--tools",
        default=",".join(DEFAULT_TOOLS),
        help="comma-separated tool list (default: all but text)",
    )
    ap.add_argument(
        "--input",
        default=None,
        type=Path,
        help="fixture dir; bench_P01.mp4 is built or rebuilt here "
        "(default: /tmp/ssbench, or /tmp/ssbench-session)",
    )
    ap.add_argument(
        "--output",
        default=None,
        type=Path,
        help="output root (default: <input>/bench-out)",
    )
    ap.add_argument(
        "--interval",
        default=None,
        type=float,
        help="sampling interval in seconds (default: 0.1, or 1.0 for session)",
    )
    ap.add_argument(
        "--duration",
        default=None,
        type=int,
        help="fixture length in seconds; rebuilds when the existing file differs "
        "(default: 120, or 3600 for session)",
    )
    bc.add_common_args(ap, "callback")
    bc.add_long_args(ap, TIERS)
    ap.add_argument(
        "--deep",
        action="store_true",
        help="attach --profile-deep scan.callback.<tool> and print each pstats block",
    )
    args = ap.parse_args()
    bc.validate_common_args(ap, args)
    bc.apply_tier(ap, args, TIERS)

    tools = [t.strip() for t in args.tools.split(",") if t.strip()]
    unknown = [t for t in tools if t not in TOOL_FLAGS]
    if unknown:
        print(f"unknown tools: {', '.join(unknown)} (know: {', '.join(TOOL_FLAGS)})")
        return bc.EXIT_USAGE
    if args.scale:
        return run_scale(args, tools)

    if args.source is not None:
        input_dir = args.input / "source"
        video = bc.link_source(args.source, input_dir, "bench_P01.mp4")
        probed = bc.probe_fixture(video, keyframes=True)
        if probed is None:
            print(f"cannot probe {args.source}")
            return bc.EXIT_USAGE
        spec: dict[str, Any] = {"source": bc.source_identity(args.source)}
    else:
        input_dir = args.input
        video = ensure_fixture(input_dir, args.duration, args.tier)
        probed = bc.probe_fixture(video, keyframes=args.tier != "smoke")
        spec = fixture_spec(args.duration, args.tier)
    out_root = args.output or (args.input / "bench-out")
    media_s = float(probed["duration"]) if probed else 0.0
    region = region_for(probed)
    gap = (probed or {}).get("keyframe_gap")
    print(
        f"fixture {video.name}  {media_s:.0f}s  interval={args.interval}"
        + (f"  keyframes every {gap:.1f}s" if gap else "")
    )

    rows = sweep(args, tools, input_dir, out_root, media_s, region)
    meta = {
        "fixture": {"spec": spec, "probed": probed},
        "workload": {
            "interval": args.interval,
            "region": region,
            **({"tier": args.tier} if args.tier != "smoke" else {}),
        },
        "env": rows.pop("_env", {}),
        "runs": args.runs,
        "metric": args.fail_metric,
    }
    baseline = None
    findings: list[tuple[str, str, str]] = []
    metric = "elapsed_s" if args.fail_metric == "elapsed" else "callback_s"
    if args.compare:
        baseline, problems = bc.load_baseline(
            args.compare, "tools", meta, allow_env=args.allow_env_mismatch
        )
        if baseline is None:
            for line in problems:
                print(f"incomparable: {line}")
            return bc.EXIT_INCOMPARABLE
        findings = bc.compare(
            rows,
            baseline,
            metric=metric,
            pct=args.fail_on,
            abs_s=args.fail_abs,
            count_key="frames",
        )
    else:
        findings = [
            ("invalid", name, "; ".join(row["reasons"]))
            for name, row in rows.items()
            if not row["valid"]
        ]

    print()
    print_table(rows, baseline, findings, metric)
    bc.report_findings(findings)
    if args.save:
        bc.save_results(args.save, meta, "tools", rows)
    return bc.exit_code(findings)


if __name__ == "__main__":
    sys.exit(main())
