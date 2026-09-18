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

Fixtures are built on demand in --input: a 120 s testsrc+sine video named
`clipbench_P01.mp4` and `clipbench.xlsx` with 30 one-participant rows
(varied 3-9 s ranges so encodes are not all identical). Both are rebuilt
when the video's probed duration, size, frame rate or audio presence differs
from `FIXTURE_SPEC`. Each scenario writes to its own wiped output dir so
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
import subprocess
import sys
from pathlib import Path
from typing import Any

import bench_common as bc

DATA_ROWS = 30  # sheet rows with timestamps; row 6 is the first
RANGE = "6-25"
EXPECTED_CLIPS = 20  # rows 6-25 inclusive

# Same 20 rows (6-25) so clip counts match across scenarios.
SCENARIOS: dict[str, list[str]] = {
    "clips": ["-r", RANGE, "--no-titlecards"],
    "clips-cards": ["-r", RANGE, "--titlecards"],
    "reel": ["-R", RANGE, "--titlecards"],
}
METRICS = ("elapsed_s", "clip_s")


def fixture_spec(duration: int) -> dict[str, Any]:
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
    }


def check_work(scenario: str, doc: dict[str, Any], out_dir: Path) -> list[str]:
    """Reasons the run did not do the expected work; empty when it did."""
    reasons: list[str] = []
    labels = doc.get("labels", {})
    clips = int(labels.get("pipeline.clip", {}).get("count", 0))
    if clips != EXPECTED_CLIPS:
        reasons.append(f"{clips} clips cut, expected {EXPECTED_CLIPS}")
    if "--titlecards" in SCENARIOS[scenario]:
        wraps = int(labels.get("titlecard.wrap", {}).get("count", 0))
        if wraps != EXPECTED_CLIPS:
            # clipgen disables cards when ffmpeg lacks drawtext; that is not a speedup.
            reasons.append(f"{wraps} clips carded, expected {EXPECTED_CLIPS}")
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


def _write_clip_sheet(sheet: Path, duration: int) -> None:
    """Write clipbench.xlsx with timestamps that fit *duration*."""
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Observations"
    ws["A1"] = "clipbench"
    ws["F2"] = "ID"
    ws["G2"] = "P01"
    for col, header in enumerate(
        ("Count", "Reported", "Severity", "Category", "Observation", "Summary"), 1
    ):
        ws.cell(5, col, header)
    for r in range(DATA_ROWS):
        ws.cell(6 + r, 3, ("Critical", "Serious", "Moderate", "Minor")[r % 4])
        ws.cell(6 + r, 4, "Onboarding")
        ws.cell(6 + r, 5, f"Observation {r}")
        ws.cell(6 + r, 7, clip_cell_range(r, duration))
    wb.save(sheet)


def ensure_fixtures(input_dir: Path, duration: int) -> Path:
    """Build the bench video and sheet, rebuilding if the video misses the spec."""
    input_dir.mkdir(parents=True, exist_ok=True)
    video = input_dir / "clipbench_P01.mp4"
    sheet = input_dir / "clipbench.xlsx"
    probed = bc.probe_fixture(video) if video.is_file() else None
    if not bc.fixture_matches(probed, fixture_spec(duration)):
        if video.is_file():
            was = f"{probed['duration']:.0f}s" if probed else "unreadable"
            print(f"rebuilding {video.name} ({was} → {duration}s)")
            video.unlink()
        if sheet.is_file():
            sheet.unlink()
        subprocess.run(
            [
                "ffmpeg",
                "-y",
                "-loglevel",
                "error",
                "-f",
                "lavfi",
                "-i",
                f"testsrc=duration={duration}:size=1280x720:rate=30",
                "-f",
                "lavfi",
                "-i",
                f"sine=frequency=220:duration={duration}",
                "-pix_fmt",
                "yuv420p",
                "-c:v",
                "libx264",
                "-g",
                "30",
                "-c:a",
                "aac",
                "-shortest",
                str(video),
            ],
            check=True,
        )
    if not sheet.is_file():
        _write_clip_sheet(sheet, duration)
    return sheet


def run_scenario(scenario: str, input_dir: Path, out_root: Path) -> dict[str, Any]:
    """One fresh-process run of *scenario* into a wiped output dir."""
    out_dir = out_root / f"bench-{scenario}"
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)
    args = [
        "-s",
        str(input_dir / "clipbench.xlsx"),
        *SCENARIOS[scenario],
        "-i",
        str(input_dir),
        "-o",
        str(out_dir),
        "--no-input",
    ]
    result = bc.run_clipgen(args, out_dir, required_labels=["pipeline.clip"])
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
        f"{delta_hdr}  status"
    )
    for scenario, row in rows.items():
        status = bc.status_of(scenario, findings)
        if not row["samples"]:
            print(f"{scenario:<14}{'':>98}  {status}")
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
            f"{med('peak_rss_mb'):>6.0f}MB"
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


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument(
        "--scenarios",
        default=",".join(SCENARIOS),
        help="comma-separated scenario list (default: all)",
    )
    ap.add_argument(
        "--input",
        default="/tmp/clipbench",
        type=Path,
        help="fixture dir; video and sheet are built or rebuilt here",
    )
    ap.add_argument(
        "--output",
        default=None,
        type=Path,
        help="output root (default: <input>/bench-out)",
    )
    ap.add_argument(
        "--duration",
        default=120,
        type=int,
        help="fixture length in seconds; rebuilds when the existing file differs",
    )
    bc.add_common_args(ap, "clip")
    args = ap.parse_args()
    bc.validate_common_args(ap, args)

    scenarios = [s.strip() for s in args.scenarios.split(",") if s.strip()]
    unknown = [s for s in scenarios if s not in SCENARIOS]
    if unknown:
        print(f"unknown scenarios: {', '.join(unknown)} (know: {', '.join(SCENARIOS)})")
        return bc.EXIT_USAGE
    out_root = args.output or (args.input / "bench-out")
    ensure_fixtures(args.input, args.duration)
    video = args.input / "clipbench_P01.mp4"
    probed = bc.probe_fixture(video)
    print(
        f"fixture {video.name}  {probed['duration']:.0f}s"
        if probed is not None
        else f"fixture {video.name}"
    )

    rows: dict[str, dict[str, Any]] = {}
    env: dict[str, Any] = {}
    for scenario in scenarios:
        results = []
        for _ in range(max(1, args.runs)):
            result = run_scenario(scenario, args.input, out_root)
            results.append(result)
            env = env or result["doc"].get("env", {})
            if not result["valid"]:
                break  # one bad repetition voids the row; no point repeating
        rows[scenario] = bc.build_row(results, summarize, METRICS)
        if rows[scenario]["samples"]:
            st = rows[scenario]["stats"]
            print(
                f"  {scenario}: elapsed {st['elapsed_s']['median']:.3f}s median"
                f" (clip {st['clip_s']['median']:.3f}s)"
            )

    meta = {
        "fixture": {"spec": fixture_spec(args.duration), "probed": probed},
        "workload": {"range": RANGE, "rows": DATA_ROWS},
        "env": env,
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
