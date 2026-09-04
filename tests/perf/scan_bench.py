"""Repeatable Screenspace scan benchmark over the standard tool sweep.

Runs each tool through the real CLI (`uv run clipgen.py --ss-task ...
--profile`) against the deterministic testsrc fixture from
agents/skills/profile/SKILL.md, parses the `profile |` report, and prints one
table row per tool. This replaces the hand-rolled per-tool command loops that
every profiling session rebuilt (and that zsh's no-word-split quoting broke
mid-session at least once) and makes before/after comparison a diff of two
JSON files instead of an eyeball job.

Usage (from the repo root):

    uv run python tests/perf/scan_bench.py                     # sweep + table
    uv run python tests/perf/scan_bench.py --save base.json    # snapshot
    uv run python tests/perf/scan_bench.py --compare base.json --fail-on 10
    uv run python tests/perf/scan_bench.py --tools color,text --runs 2

The fixture video is built on demand (ffmpeg testsrc: constant motion, so
phash-skip never hides the callback) and rebuilt when `--duration` does not
match the file already on disk — a leftover 120 s clip would otherwise
ignore `--duration 15` and poison `--compare`. Each tool writes to its own
wiped output dir so a cached manifest can never absorb the scan. `--runs N`
keeps the fastest run per tool (minimum callback seconds), the standard
treatment for scheduler noise. Template and Shape use a fixed top-left 20%
region for a non-degenerate reference and bounded search workload. `text` is
excluded from the default sweep: OCR is an order of magnitude slower than
every other tool and pins ~0.8 GB of RSS per pooled OCR engine (measured
3.3 GB at the default auto pool of 4).

The parser (`parse_profile`) is unit-tested in test_scan_bench_parse.py;
everything else is a thin subprocess driver kept dependency-free on purpose.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]

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

# Label is padded to 32 chars and may contain spaces (`route /api/foo`).
_PROFILE_RE = re.compile(r"^profile \| (.+?)\s+([\d.]+)s\s+n=(\d+)(.*)$", re.MULTILINE)
_RSS_RE = re.compile(r"^profile \| peak_rss\s+([\d.]+)MB", re.MULTILINE)
_MS_RE = re.compile(r"\b(avg|max|first)=([\d.]+)ms")
_BYTES_RE = re.compile(r"\bbytes=(\S+)")


def _parse_bytes(token: str) -> float:
    """Invert profiling.format_bytes; one-decimal sizes are approximate."""
    if token.endswith("MB"):
        return float(token[:-2]) * 1024 * 1024
    if token.endswith("KB"):
        return float(token[:-2]) * 1024
    if token.endswith("B"):
        return float(token[:-1])
    return 0.0


def parse_profile(text: str) -> dict[str, dict[str, float]]:
    """Parse `profile |` report lines into {label: {seconds, n, ...}} (+peak_rss)."""
    out: dict[str, dict[str, float]] = {}
    for label, seconds, n, rest in _PROFILE_RE.findall(text):
        row: dict[str, float] = {"seconds": float(seconds), "n": int(n)}
        for key, ms in _MS_RE.findall(rest):
            row[key] = float(ms) / 1000.0
        nbytes = _BYTES_RE.search(rest)
        if nbytes:
            row["bytes"] = _parse_bytes(nbytes.group(1))
        out[label] = row
    rss = _RSS_RE.search(text)
    if rss:
        out["peak_rss"] = {"seconds": 0.0, "n": 0, "mb": float(rss.group(1))}
    return out


def probe_duration(path: Path) -> float | None:
    """Seconds of *path*, or None if ffprobe cannot say."""
    proc = subprocess.run(
        [
            "ffprobe",
            "-v",
            "error",
            "-show_entries",
            "format=duration",
            "-of",
            "csv=p=0",
            str(path),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    try:
        return float(proc.stdout.strip())
    except ValueError:
        return None


def delta_pct(current: float, base: float) -> float | None:
    """Percent change, or None when *base* is 0."""
    if not base:
        return None
    return (current / base - 1.0) * 100.0


def regressions(
    rows: dict[str, dict[str, float]],
    baseline: dict[str, dict[str, float]],
    key: str,
    limit: float,
) -> list[tuple[str, float]]:
    """Names whose *key* rose by more than *limit* percent."""
    hit: list[tuple[str, float]] = []
    for name, row in rows.items():
        pct = delta_pct(row.get(key, 0.0), (baseline.get(name) or {}).get(key, 0.0))
        if pct is not None and pct > limit:
            hit.append((name, pct))
    return hit


def summarize(tool: str, profile: dict[str, dict[str, float]]) -> dict[str, float]:
    """Reduce one run's parsed report to the per-tool comparison row."""
    callback = profile.get(f"scan.callback.{tool}", {"seconds": 0.0, "n": 0})
    decode = profile.get("scan.decode_wait", {"seconds": 0.0, "n": 0})
    filt = profile.get("scan.fast_filter", {"seconds": 0.0, "n": 0})
    heatmap = sum(
        profile.get(label, {}).get("seconds", 0.0)
        for label in ("heatmap.gifs", "heatmap.grid_layers")
    )
    frames = int(callback["n"])
    return {
        "callback_s": callback["seconds"],
        "callback_avg_ms": callback["seconds"] / frames * 1000 if frames else 0.0,
        "frames": frames,
        "decode_s": decode["seconds"],
        "filter_s": filt["seconds"],
        "heatmap_s": heatmap,
        "peak_rss_mb": profile.get("peak_rss", {}).get("mb", 0.0),
    }


def keep_best(
    best: dict[str, float] | None, summary: dict[str, float]
) -> dict[str, float]:
    """Pick the run to keep across --runs: fastest *successful* callback.

    A failed run parses to callback_s == 0.0, which naive min-keeping would
    hold onto forever (0 < anything), poisoning --compare baselines with a 0s
    row. A successful run always beats a failed one; among successes the
    minimum callback wins (standard treatment for scheduler noise).
    """
    if best is None:
        return summary
    if not summary["callback_s"]:
        return best
    if not best["callback_s"] or summary["callback_s"] < best["callback_s"]:
        return summary
    return best


def ensure_fixture(input_dir: Path, duration: int) -> Path:
    """Build the benchmark video, rebuilding if its length does not match."""
    video = input_dir / "bench_P01.mp4"
    existing = probe_duration(video) if video.is_file() else None
    if existing is not None and abs(existing - duration) < 0.5:
        return video
    if video.is_file():
        was = f"{existing:.0f}s" if existing is not None else "unreadable"
        print(f"rebuilding {video.name} ({was} → {duration}s)")
        video.unlink()
    input_dir.mkdir(parents=True, exist_ok=True)
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
            "-pix_fmt",
            "yuv420p",
            "-c:v",
            "libx264",
            "-g",
            "30",
            str(video),
        ],
        check=True,
    )
    return video


def run_tool(
    tool: str,
    input_dir: Path,
    out_root: Path,
    interval: float,
    deep: bool = False,
) -> dict[str, dict[str, float]]:
    """Run one tool through the CLI into a wiped output dir; return the parse.

    With *deep*, attach ``--profile-deep scan.callback.<tool>`` and print the
    pstats block verbatim — the harness owns the canonical flag set, so this
    replaces hand-rolling the tool's CLI command just to drill into it.
    """
    out_dir = out_root / f"bench-{tool}"
    if out_dir.exists():
        shutil.rmtree(out_dir)
    region_args: list[str] = []
    if tool in REGION_TOOLS:
        out_dir.mkdir(parents=True)
        manifest = {
            "screenspace": {
                "regions": {"bench": BENCH_REGION},
                "tasks": [],
                "events": [],
                "stashes": [],
            }
        }
        (out_dir / "clipgen.json").write_text(json.dumps(manifest))
        region_args = ["bench"]
    cmd = [
        "uv",
        "run",
        "clipgen.py",
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
        "--profile",
    ]
    if deep:
        cmd += ["--profile-deep", f"scan.callback.{tool}"]
    proc = subprocess.run(
        cmd, cwd=REPO_ROOT, capture_output=True, text=True, timeout=1800, check=False
    )
    output = proc.stdout + proc.stderr
    parsed = parse_profile(output)
    if proc.returncode:
        print(f"  ! {tool}: clipgen exit {proc.returncode}")
    if f"scan.callback.{tool}" not in parsed:
        print(f"  ! {tool}: no scan.callback.{tool} in report (task refused?)")
        tail = "\n".join(output.strip().splitlines()[-5:])
        print("    " + tail.replace("\n", "\n    "))
    if deep:
        in_block = False
        for line in output.splitlines():
            if line.startswith("profile-deep |"):
                in_block = True
            if in_block:
                print(line)
    return parsed


def print_table(
    rows: dict[str, dict[str, float]], baseline: dict[str, dict[str, float]] | None
) -> None:
    delta_hdr = "  Δcallback" if baseline else ""
    print(
        f"{'tool':<12}{'callback':>10}{'avg':>9}{'frames':>8}"
        f"{'decode':>9}{'filter':>9}{'heatmap':>9}{'rss':>9}{delta_hdr}"
    )
    for tool, row in rows.items():
        line = (
            f"{tool:<12}{row['callback_s']:>9.3f}s{row['callback_avg_ms']:>7.1f}ms"
            f"{row['frames']:>8d}{row['decode_s']:>8.3f}s{row['filter_s']:>8.3f}s"
            f"{row['heatmap_s']:>8.3f}s{row['peak_rss_mb']:>7.0f}MB"
        )
        if baseline:
            base = baseline.get(tool)
            if base and base.get("callback_s"):
                pct = delta_pct(row["callback_s"], base["callback_s"])
                line += f"  {pct:>+8.1f}%" if pct is not None else "  (no base)"
                if (
                    row.get("frames")
                    and base.get("frames")
                    and row["frames"] != base["frames"]
                ):
                    line += " frames≠"
            else:
                line += "  (no base)"
        print(line)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument(
        "--tools",
        default=",".join(DEFAULT_TOOLS),
        help="comma-separated tool list (default: all but text)",
    )
    ap.add_argument(
        "--input",
        default="/tmp/ssbench",
        type=Path,
        help="fixture dir; bench_P01.mp4 is built or rebuilt here",
    )
    ap.add_argument(
        "--output",
        default=None,
        type=Path,
        help="output root (default: <input>/bench-out)",
    )
    ap.add_argument("--interval", default=0.1, type=float)
    ap.add_argument(
        "--duration",
        default=120,
        type=int,
        help="fixture length in seconds; rebuilds when the existing file differs",
    )
    ap.add_argument(
        "--runs",
        default=1,
        type=int,
        help="runs per tool; the fastest (min callback) is kept",
    )
    ap.add_argument("--save", type=Path, help="write results JSON here")
    ap.add_argument("--compare", type=Path, help="baseline JSON to diff against")
    ap.add_argument(
        "--fail-on",
        type=float,
        default=None,
        help="with --compare, exit 1 if any Δcallback %% exceeds this",
    )
    ap.add_argument(
        "--deep",
        action="store_true",
        help="attach --profile-deep scan.callback.<tool> and print each pstats block",
    )
    args = ap.parse_args()
    if args.fail_on is not None and not args.compare:
        ap.error("--fail-on requires --compare")

    tools = [t.strip() for t in args.tools.split(",") if t.strip()]
    unknown = [t for t in tools if t not in TOOL_FLAGS]
    if unknown:
        print(f"unknown tools: {', '.join(unknown)} (know: {', '.join(TOOL_FLAGS)})")
        return 2
    out_root = args.output or (args.input / "bench-out")
    video = ensure_fixture(args.input, args.duration)
    probed = probe_duration(video)
    print(
        f"fixture {video.name}  {probed:.0f}s  interval={args.interval}"
        if probed is not None
        else f"fixture {video.name}  interval={args.interval}"
    )

    baseline = None
    if args.compare:
        baseline = json.loads(args.compare.read_text())["tools"]

    rows: dict[str, dict[str, float]] = {}
    for tool in tools:
        best: dict[str, float] | None = None
        for _ in range(max(1, args.runs)):
            summary = summarize(
                tool, run_tool(tool, args.input, out_root, args.interval, args.deep)
            )
            best = keep_best(best, summary)
        assert best is not None  # loop runs at least once
        rows[tool] = best
        print(
            f"  {tool}: callback {best['callback_s']:.3f}s over {best['frames']} frames"
        )

    print()
    print_table(rows, baseline)

    if args.save:
        args.save.write_text(
            json.dumps(
                {
                    "meta": {
                        "interval": args.interval,
                        "video": str(args.input / "bench_P01.mp4"),
                        "duration": args.duration,
                        "runs": args.runs,
                    },
                    "tools": rows,
                },
                indent=2,
            )
        )
        print(f"\nsaved -> {args.save}")
    if args.fail_on is not None and baseline is not None:
        hit = regressions(rows, baseline, "callback_s", args.fail_on)
        if hit:
            for name, pct in hit:
                print(f"fail-on: {name} {pct:+.1f}% (limit {args.fail_on:g}%)")
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
