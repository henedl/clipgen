"""Repeatable clip-pipeline benchmark: clips, carded clips, and a reel.

The scan side has scan_bench.py; this is the same treatment for the other
half of the product — cutting. Each scenario runs the real CLI against a
deterministic fixture (testsrc video + generated Excel sheet), reads the
`--profile-output` JSON, and prints one row per scenario with the ratio the
pool knob is otherwise tuned blind against (pipeline.clip / pool_wall).

Usage (from the repo root):

    uv run python tests/perf/clip_bench.py                     # sweep + table
    uv run python tests/perf/clip_bench.py --save base.json    # snapshot
    uv run python tests/perf/clip_bench.py --compare base.json --fail-on 10
    uv run python tests/perf/clip_bench.py --scenarios clips --runs 5
    uv run python tests/perf/clip_bench.py --tier session
    uv run python tests/perf/clip_bench.py --tier session --scenarios clips-compress
    uv run python tests/perf/clip_bench.py --source ~/rec/P07.mp4 --clip-seconds 240
    uv run python tests/perf/clip_bench.py --scale 600,3600,10800 --clip-seconds 60

Fixtures are built on demand in --input: a 120 s testsrc+sine video named
`clipbench_P01.mp4` and `clipbench.xlsx` with 30 one-participant rows
(varied 3-7 s ranges inside the first minute, so encodes are not all
identical). The video is rebuilt when its probed duration, size, frame rate
or audio presence differs from the spec.

Tiers: ``smoke`` (default) is that short regression gate. ``session`` cuts
20 three-minute clips spread across an hour of OBS-like 1080p, so deep
seeks and long bodies are measured; ``--clip-seconds`` spreads clips of any
length across any fixture or ``--source`` recording. ``clips-reencode`` and
``clips-compress`` (opt-in; ``REENCODING`` / ``MAX_FILESIZE_MB`` set before
``cli.main``) cover the paths whose cost grows with clip length. ``xrt`` is
clip seconds cut per wall second. ``--scale`` holds the clips fixed and
grows the source: cutting should read ``flat``. Each scenario writes to its own wiped output dir so
reservations and manifests never leak between runs. `--runs N` (default 3)
keeps every fresh-process sample and compares medians; a repetition that
exits non-zero, cuts fewer than the expected clips, or leaves no output
files invalidates the scenario.

Row reduction and validity are unit-tested in test_clip_bench_parse.py; the
shared compare/aggregate logic lives in bench_common.py.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path
from typing import Any

import bench_common as bc
import bench_fixtures as bf

DATA_ROWS = 30  # sheet rows with timestamps; row 6 is the first
RANGE = "6-25"
EXPECTED_CLIPS = 20  # rows 6-25 inclusive

# Same 20 rows (6-25) so clip counts match across scenarios.
SCENARIOS: dict[str, list[str]] = {
    "clips": ["-r", RANGE, "--no-titlecards"],
    "clips-cards": ["-r", RANGE, "--titlecards"],
    "reel": ["-R", RANGE, "--titlecards"],
}
# Opt-in: no CLI flag reaches these settings, so the run sets config first.
EXTRA_SCENARIOS: dict[str, list[str]] = {
    "clips-reencode": ["-r", RANGE, "--no-titlecards"],
    "clips-compress": ["-r", RANGE, "--no-titlecards"],
}
OVERRIDES: dict[str, dict[str, Any]] = {
    "clips-reencode": {"REENCODING": True},
    "clips-compress": {"MAX_FILESIZE_MB": 0.1},  # every real clip exceeds 100 KB
}
ALL_SCENARIOS = {**SCENARIOS, **EXTRA_SCENARIOS}
METRICS = ("elapsed_s", "clip_s")

TIERS: dict[str, dict[str, Any]] = {
    "smoke": {"duration": 120, "input": Path("/tmp/clipbench")},
    "session": {
        "duration": 3600,
        "clip_seconds": 180,
        "runs": 1,
        "input": Path("/tmp/clipbench-session"),
    },
}


def fixture_spec(duration: int, tier: str = "smoke") -> dict[str, Any]:
    if tier == "session":
        return bc.session_spec(duration, audio="sine")
    return {
        "duration": duration,
        "width": 1280,
        "height": 720,
        "fps": 30.0,
        "has_audio": True,
    }


def summarize(doc: dict[str, Any]) -> dict[str, Any]:
    """Reduce one run's JSON report to the per-scenario comparison row."""
    labels = doc.get("labels", {})

    def get(label: str) -> dict[str, float]:
        return labels.get(label, {"seconds": 0.0, "count": 0})

    clip = get("pipeline.clip")
    pool = get("pipeline.pool_wall")
    wrap = get("titlecard.wrap")
    # One label per subprocess kind; bare ffmpeg.run is the untagged fallback.
    ffmpeg = {
        label: row for label, row in labels.items() if label.startswith("ffmpeg.run")
    }
    return {
        "clip_s": clip["seconds"],
        "clips": int(clip["count"]),
        "pool_wall_s": pool["seconds"],
        "parallelism": clip["seconds"] / pool["seconds"] if pool["seconds"] else 0.0,
        "ffmpeg_s": sum(r["seconds"] for r in ffmpeg.values()),
        "ffmpeg_n": sum(int(r["count"]) for r in ffmpeg.values()),
        "ffmpeg_kinds": {
            label.removeprefix("ffmpeg.run.") or "other": {
                "s": r["seconds"],
                "n": int(r["count"]),
            }
            for label, r in ffmpeg.items()
        },
        "ffprobe_n": int(get("ffprobe.run")["count"]),
        "cards_s": wrap["seconds"],
        "cards_copy": int(get("titlecard.copy")["count"]),
        "cards_reencode": int(get("titlecard.reencode")["count"]),
        "peak_rss_mb": doc.get("peak_rss_mb") or 0.0,
        "compress_n": int(get("ffmpeg.run.compress")["count"]),
    }


def check_work(scenario: str, doc: dict[str, Any], out_dir: Path) -> list[str]:
    """Reasons the run did not do the expected work; empty when it did."""
    reasons: list[str] = []
    labels = doc.get("labels", {})
    clips = int(labels.get("pipeline.clip", {}).get("count", 0))
    if clips != EXPECTED_CLIPS:
        reasons.append(f"{clips} clips cut, expected {EXPECTED_CLIPS}")
    if "--titlecards" in ALL_SCENARIOS[scenario]:
        wraps = int(labels.get("titlecard.wrap", {}).get("count", 0))
        if wraps != EXPECTED_CLIPS:
            # clipgen disables cards when ffmpeg lacks drawtext; that is not a speedup.
            reasons.append(f"{wraps} clips carded, expected {EXPECTED_CLIPS}")
    if scenario == "clips-compress":
        compressed = int(labels.get("ffmpeg.run.compress", {}).get("count", 0))
        if compressed < EXPECTED_CLIPS:
            # Two passes per clip; clips already under the cap are skipped.
            reasons.append(
                f"{compressed} compress passes, expected >= {EXPECTED_CLIPS}"
                " (clips under the 100 KB cap?)"
            )
    outputs = [p for p in out_dir.rglob("*.mp4") if not p.name.startswith("_")]
    expected = 1 if scenario == "reel" else EXPECTED_CLIPS
    if len(outputs) < expected:
        reasons.append(f"{len(outputs)} output files, expected {expected}")
    return reasons


def clip_cell_range(row: int, duration: int) -> str:
    """MM:SS window for data row *row* that stays inside *duration* seconds."""
    start = 2 + (row * 3) % 49
    end = start + 3 + row % 5
    last = max(2, min(duration - 1, 57))
    if end > last:
        end = last
        start = max(1, end - (3 + row % 5))
    if start >= end:
        start = max(0, end - 1)
    return f"0:{start:02d}-0:{end:02d}"


def sheet_cells(duration: int, clip_seconds: int | None) -> list[str]:
    """Every data row's cell: the legacy first-minute windows, or spread clips."""
    if clip_seconds:
        # Spread the benched rows over the whole file; the spare rows repeat them.
        cells = bf.clip_cells(
            EXPECTED_CLIPS, duration=duration, clip_seconds=clip_seconds
        )
        return cells + cells[: DATA_ROWS - EXPECTED_CLIPS]
    return [clip_cell_range(r, duration) for r in range(DATA_ROWS)]


def clip_media_seconds(cells: list[str]) -> float:
    """Summed length of the benched rows' clips (``RANGE``), the ``xrt`` numerator."""

    def secs(text: str) -> int:
        total = 0
        for part in text.split(":"):
            total = total * 60 + int(part)
        return total

    total = 0
    for cell in cells[:EXPECTED_CLIPS]:
        start, _, end = cell.partition("-")
        total += secs(end) - secs(start)
    return float(total)


def _write_clip_sheet(sheet: Path, cells: list[str]) -> None:
    """Write clipbench.xlsx with one participant column holding *cells*."""
    bf.write_sheet(
        sheet,
        study="clipbench",
        rows=DATA_ROWS,
        participants=1,
        cell=lambda r, _p: cells[r],
    )


def ensure_fixtures(
    input_dir: Path,
    duration: int,
    tier: str = "smoke",
    clip_seconds: int | None = None,
) -> Path:
    """Build the bench video (when it misses the spec) and rewrite the sheet."""
    input_dir.mkdir(parents=True, exist_ok=True)
    video = input_dir / "clipbench_P01.mp4"
    sheet = input_dir / "clipbench.xlsx"

    def build() -> None:
        if tier == "session":
            bf.make_long_video(
                video, duration=duration, audio="sine", **bc.SESSION_VIDEO
            )
        else:
            bf.make_testsrc_video(video, duration=duration, audio=True)

    bc.ensure_video(video, fixture_spec(duration, tier), build)
    _write_clip_sheet(sheet, sheet_cells(duration, clip_seconds))
    return sheet


def run_scenario(
    scenario: str, input_dir: Path, out_root: Path, *, timeout: float = 1800
) -> dict[str, Any]:
    """One fresh-process run of *scenario* into a wiped output dir."""
    out_dir = out_root / f"bench-{scenario}"
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)
    args = [
        "-s",
        str(input_dir / "clipbench.xlsx"),
        *ALL_SCENARIOS[scenario],
        "-i",
        str(input_dir),
        "-o",
        str(out_dir),
        "--no-input",
    ]
    extra: dict[str, Any] = {"timeout": timeout}
    if scenario in OVERRIDES:
        extra["overrides"] = OVERRIDES[scenario]
    result = bc.run_clipgen(args, out_dir, required_labels=["pipeline.clip"], **extra)
    if result["valid"]:
        result["reasons"].extend(check_work(scenario, result["doc"], out_dir))
        result["valid"] = not result["reasons"]
    if not result["valid"]:
        bc.print_failure(scenario, result)
    return result


def print_table(
    rows: dict[str, dict[str, Any]],
    baseline: dict[str, dict[str, Any]] | None,
    findings: list[tuple[str, str, str]],
    metric: str = "elapsed_s",
) -> None:
    delta_hdr = f"  Δ{metric.removesuffix('_s'):<8}" if baseline else ""
    print(
        f"{'scenario':<14}{'elapsed':>9}{'min':>9}{'mad':>8}{'clip':>9}{'n':>4}"
        f"{'pool':>9}{'par':>6}{'ffmpeg':>9}{'cards':>9}{'copy/re':>9}{'rss':>8}"
        f"{'xrt':>8}{delta_hdr}  status"
    )
    for scenario, row in rows.items():
        status = bc.status_of(scenario, findings)
        if not row["samples"]:
            print(f"{scenario:<14}{'':>106}  {status}")
            continue

        def med(key: str, samples: list[dict[str, Any]] = row["samples"]) -> float:
            return bc.aggregate(samples, key)["median"]

        clip = med("clip_s")
        line = (
            f"{scenario:<14}{bc.stat_cells(row, 'elapsed_s')}"
            f"{clip:>8.3f}s{int(med('clips')):>4d}"
            f"{med('pool_wall_s'):>8.3f}s{med('parallelism'):>6.1f}"
            f"{med('ffmpeg_s'):>8.3f}s{med('cards_s'):>8.3f}s"
            f"{int(med('cards_copy')):>4d}/{int(med('cards_reencode')):<4d}"
            f"{med('peak_rss_mb'):>6.0f}MB{med('xrt'):>7.1f}x"
        )
        if baseline:
            base = baseline.get(scenario)
            if base and base.get("samples"):
                pct = bc.delta_pct(med(metric), base["stats"][metric]["median"])
                line += f"  {pct:>+9.1f}%" if pct is not None else "  (no base)"
            else:
                line += "   (no base)"
        print(f"{line}  {status}")
        kinds = row["samples"][0].get("ffmpeg_kinds") or {}
        if kinds:
            parts = "  ".join(
                f"{kind} {v['s']:.3f}s/n={v['n']}"
                for kind, v in sorted(kinds.items(), key=lambda kv: -kv[1]["s"])
            )
            print(f"{'':<14}ffmpeg (first run): {parts}")


def sweep(
    args: argparse.Namespace,
    scenarios: list[str],
    input_dir: Path,
    out_root: Path,
    media_s: float,
) -> dict[str, Any]:
    """Every scenario's row; the first run's ``env`` rides along as ``_env``."""
    rows: dict[str, Any] = {}
    env: dict[str, Any] = {}
    timeout = bc.timeout_for(media_s, args.timeout)
    for scenario in scenarios:
        results = []
        for _ in range(max(1, args.runs)):
            result = run_scenario(scenario, input_dir, out_root, timeout=timeout)
            results.append(result)
            env = env or result["doc"].get("env", {})
            if not result["valid"]:
                break  # one bad repetition voids the row; no point repeating
        rows[scenario] = bc.build_row(results, summarize, METRICS, media_s)
        if rows[scenario]["samples"]:
            st = rows[scenario]["stats"]
            print(
                f"  {scenario}: elapsed {st['elapsed_s']['median']:.3f}s median"
                f" (clip {st['clip_s']['median']:.3f}s,"
                f" {rows[scenario]['samples'][0]['xrt']:.1f}x realtime)"
            )
    rows["_env"] = env
    return rows


def run_scale(args: argparse.Namespace, scenarios: list[str]) -> int:
    """Fixed clips over a growing source: the fit should read ``flat``."""
    try:
        durations = bc.parse_scale(args.scale)
    except ValueError as exc:
        print(exc)
        return bc.EXIT_USAGE
    clip_seconds = args.clip_seconds or 60
    points: dict[str, list[tuple[float, float]]] = {s: [] for s in scenarios}
    bad = False
    for duration in durations:
        input_dir = args.input / f"scale-{duration}"
        ensure_fixtures(input_dir, duration, args.tier, clip_seconds)
        cells = sheet_cells(duration, clip_seconds)
        print(f"fixture {duration}s  clips {cells[0]} … {cells[EXPECTED_CLIPS - 1]}")
        rows = sweep(
            args,
            scenarios,
            input_dir,
            input_dir / "bench-out",
            clip_media_seconds(cells),
        )
        rows.pop("_env", None)
        for scenario, row in rows.items():
            if row["samples"]:
                points[scenario].append(
                    (float(duration), row["stats"]["elapsed_s"]["median"])
                )
            else:
                bad = True
    bc.print_scaling(points)
    return bc.EXIT_INVALID if bad else bc.EXIT_OK


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument(
        "--scenarios",
        default=",".join(SCENARIOS),
        help="comma-separated scenario list (default: clips, clips-cards, reel; "
        "opt-in: clips-reencode, clips-compress)",
    )
    ap.add_argument(
        "--input",
        default=None,
        type=Path,
        help="fixture dir; video and sheet are built or rebuilt here "
        "(default: /tmp/clipbench, or /tmp/clipbench-session)",
    )
    ap.add_argument(
        "--output",
        default=None,
        type=Path,
        help="output root (default: <input>/bench-out)",
    )
    ap.add_argument(
        "--duration",
        default=None,
        type=int,
        help="fixture length in seconds; rebuilds when the existing file differs "
        "(default: 120, or 3600 for session)",
    )
    ap.add_argument(
        "--clip-seconds",
        default=None,
        type=int,
        help="spread clips of this length across the whole source "
        "(default: 3-7 s clips in the first minute; 180 for session)",
    )
    bc.add_common_args(ap, "clip")
    bc.add_long_args(ap, TIERS)
    args = ap.parse_args()
    bc.validate_common_args(ap, args)
    bc.apply_tier(ap, args, TIERS)

    scenarios = [s.strip() for s in args.scenarios.split(",") if s.strip()]
    unknown = [s for s in scenarios if s not in ALL_SCENARIOS]
    if unknown:
        known = ", ".join(ALL_SCENARIOS)
        print(f"unknown scenarios: {', '.join(unknown)} (know: {known})")
        return bc.EXIT_USAGE
    if args.scale:
        return run_scale(args, scenarios)

    if args.source is not None:
        input_dir = args.input / "source"
        video = bc.link_source(args.source, input_dir, "clipbench_P01.mp4")
        probed = bc.probe_fixture(video, keyframes=True)
        if probed is None:
            print(f"cannot probe {args.source}")
            return bc.EXIT_USAGE
        args.clip_seconds = args.clip_seconds or 120
        cells = sheet_cells(int(probed["duration"]), args.clip_seconds)
        _write_clip_sheet(input_dir / "clipbench.xlsx", cells)
        spec: dict[str, Any] = {"source": bc.source_identity(args.source)}
    else:
        input_dir = args.input
        ensure_fixtures(input_dir, args.duration, args.tier, args.clip_seconds)
        video = input_dir / "clipbench_P01.mp4"
        probed = bc.probe_fixture(video, keyframes=args.tier != "smoke")
        cells = sheet_cells(args.duration, args.clip_seconds)
        spec = fixture_spec(args.duration, args.tier)
    out_root = args.output or (args.input / "bench-out")
    media_s = clip_media_seconds(cells)
    print(
        f"fixture {video.name}  {probed['duration']:.0f}s"
        f"  clips {cells[0]} … {cells[EXPECTED_CLIPS - 1]}"
        if probed is not None
        else f"fixture {video.name}"
    )

    rows = sweep(args, scenarios, input_dir, out_root, media_s)
    workload: dict[str, Any] = {"range": RANGE, "rows": DATA_ROWS}
    if args.clip_seconds:
        workload["clip_seconds"] = args.clip_seconds
    if args.tier != "smoke":
        workload["tier"] = args.tier
    meta = {
        "fixture": {"spec": spec, "probed": probed},
        "workload": workload,
        "env": rows.pop("_env", {}),
        "runs": args.runs,
        "metric": args.fail_metric,
    }
    baseline = None
    findings: list[tuple[str, str, str]] = []
    metric = "elapsed_s" if args.fail_metric == "elapsed" else "clip_s"
    if args.compare:
        baseline, problems = bc.load_baseline(
            args.compare, "scenarios", meta, allow_env=args.allow_env_mismatch
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
            count_key="clips",
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
        bc.save_results(args.save, meta, "scenarios", rows)
    return bc.exit_code(findings)


if __name__ == "__main__":
    sys.exit(main())
