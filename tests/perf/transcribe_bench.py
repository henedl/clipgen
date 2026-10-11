"""Repeatable transcription benchmark: Whisper, then Whisper plus the speaker pass.

The scan and clip halves have their benches; this is the third cost every
long session pays, and the one that grows with the *whole* recording rather
than with clip count. Each scenario runs the real CLI (`--pre-transcribe`)
against a speech fixture, reads the `--profile-output` JSON, and prints one
row per scenario with the realtime factor and peak memory.

Usage (from the repo root):

    uv run python tests/perf/transcribe_bench.py                    # 5 min speech
    uv run python tests/perf/transcribe_bench.py --save base.json
    uv run python tests/perf/transcribe_bench.py --compare base.json --fail-on 10
    uv run python tests/perf/transcribe_bench.py --tier session      # one hour
    uv run python tests/perf/transcribe_bench.py --source ~/rec/P07.mp4
    uv run python tests/perf/transcribe_bench.py --scale 300,900,1800

The fixture is real speech, not a sine: a two-voice dialogue rendered with
macOS ``say`` and looped (``bench_fixtures.make_speech_audio``), so VAD,
Whisper and the speaker pass all do their actual work — a sine track is
silence to VAD and measures nothing. Off macOS, pass ``--source``. The
video track is a small 320x240 still; only the audio matters here.

Rows report ``xrt`` (audio seconds per wall second, end to end including
model load), ``whisper_xrt`` (audio over ``prepare + decode``, the profiler's
own ratio), the PCM decode, the speaker pass, and ``peak_rss_mb`` — the
whole track is held in memory as 16 kHz float32, which is what grows with a
three-hour session. Same `--save`/`--compare`/`--runs` rules as the other
benches; the count that must match a baseline is the segment count.
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path
from typing import Any

import bench_common as bc
import bench_fixtures as bf

STUDY = "tsbench"
SCENARIOS: dict[str, list[str]] = {
    "whisper": [],
    "speakers": ["--speakers"],
}
METRICS = ("elapsed_s", "whisper_s")

# A small still video: the bench measures audio work, not video decode.
VIDEO: bc.VideoShape = {"size": "320x240", "rate": 10, "gop": 100, "content": "screen"}

TIERS: dict[str, dict[str, Any]] = {
    "smoke": {"duration": 300, "input": Path("/tmp/tsbench")},
    "session": {"duration": 3600, "runs": 1, "input": Path("/tmp/tsbench-session")},
}


def fixture_spec(duration: int) -> dict[str, Any]:
    width, height = (int(v) for v in VIDEO["size"].split("x"))
    return {
        "duration": duration,
        "width": width,
        "height": height,
        "fps": float(VIDEO["rate"]),
        "has_audio": True,
        "content": VIDEO["content"],
        "audio": "speech",
    }


def ensure_fixtures(input_dir: Path, duration: int) -> Path:
    """Build the speech video (when it misses the spec) and the one-column sheet."""
    video = input_dir / f"{STUDY}_P01.mp4"
    bc.ensure_video(
        video,
        fixture_spec(duration),
        lambda: bf.make_long_video(video, duration=duration, audio="speech", **VIDEO),
    )
    return write_sheet(input_dir)


def write_sheet(input_dir: Path) -> Path:
    sheet = input_dir / f"{STUDY}.xlsx"
    bf.write_sheet(sheet, study=STUDY, rows=1, participants=1)
    return sheet


def summarize(doc: dict[str, Any]) -> dict[str, Any]:
    """Reduce one run's JSON report to the per-scenario comparison row."""
    labels = doc.get("labels", {})

    def get(label: str) -> dict[str, float]:
        return labels.get(label, {"seconds": 0.0, "count": 0})

    prepare = get("transcribe.prepare")["seconds"]
    decode = get("transcribe.decode")
    return {
        "whisper_s": prepare + decode["seconds"],
        "prepare_s": prepare,
        "decode_s": decode["seconds"],
        "segments": int(decode["count"]),
        "model_load_s": get("transcribe.model_load")["seconds"],
        "pcm_s": get("transcribe.decode_audio")["seconds"],
        "speakers_s": get("speakers.diarize")["seconds"],
        "speakers_pcm_s": get("speakers.decode")["seconds"],
        "speakers_embed_s": get("speakers.embed")["seconds"],
        "speakers_cluster_s": get("speakers.cluster")["seconds"],
        "peak_rss_mb": doc.get("peak_rss_mb") or 0.0,
    }


def check_work(scenario: str, doc: dict[str, Any], out_dir: Path) -> list[str]:
    """Reasons the run did not do the expected work; empty when it did."""
    reasons: list[str] = []
    labels = doc.get("labels", {})
    if not labels.get("transcribe.decode", {}).get("count"):
        reasons.append("no segments decoded (VAD found no speech?)")
    if scenario == "speakers" and "speakers.cluster" not in labels:
        reasons.append("speaker pass did not run (model missing?)")
    try:
        manifest = json.loads((out_dir / "clipgen.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        manifest = {}
    stored = manifest.get("transcripts", {}).get("source_transcripts", {})
    if not stored.get("P01", {}).get("segments"):
        reasons.append("no transcript stored for P01")
    return reasons


def run_scenario(
    scenario: str,
    input_dir: Path,
    out_root: Path,
    model: str,
    *,
    timeout: float = 1800,
) -> dict[str, Any]:
    """One fresh-process run into a wiped output dir (a stored transcript is skipped)."""
    out_dir = out_root / f"bench-{scenario}"
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)
    args = [
        "-s",
        str(input_dir / f"{STUDY}.xlsx"),
        "--pre-transcribe",
        "P01",
        "--whisper-model",
        model,
        *SCENARIOS[scenario],
        "-i",
        str(input_dir),
        "-o",
        str(out_dir),
    ]
    result = bc.run_clipgen(
        args, out_dir, required_labels=["transcribe.prepare"], timeout=timeout
    )
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
        f"{'scenario':<10}{'elapsed':>9}{'min':>9}{'mad':>8}{'whisper':>10}"
        f"{'w-xrt':>8}{'segs':>6}{'load':>8}{'pcm':>8}{'speakers':>10}"
        f"{'cluster':>9}{'rss':>9}{'xrt':>8}{delta_hdr}  status"
    )
    for name, row in rows.items():
        status = bc.status_of(name, findings)
        if not row["samples"]:
            print(f"{name:<10}{'':>102}  {status}")
            continue

        def med(key: str, samples: list[dict[str, Any]] = row["samples"]) -> float:
            return bc.aggregate(samples, key)["median"]

        whisper = med("whisper_s")
        media = med("media_s")
        line = (
            f"{name:<10}{bc.stat_cells(row, 'elapsed_s')}"
            f"{whisper:>9.2f}s{(media / whisper if whisper else 0.0):>7.1f}x"
            f"{int(med('segments')):>6d}{med('model_load_s'):>7.2f}s"
            f"{med('pcm_s'):>7.2f}s{med('speakers_s'):>9.2f}s"
            f"{med('speakers_cluster_s'):>8.2f}s{med('peak_rss_mb'):>7.0f}MB"
            f"{med('xrt'):>7.1f}x"
        )
        if baseline:
            base = baseline.get(name)
            if base and base.get("samples"):
                pct = bc.delta_pct(med(metric), base["stats"][metric]["median"])
                line += f"  {pct:>+9.1f}%" if pct is not None else "  (no base)"
            else:
                line += "   (no base)"
        print(f"{line}  {status}")


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
            result = run_scenario(
                scenario, input_dir, out_root, args.model, timeout=timeout
            )
            results.append(result)
            env = env or result["doc"].get("env", {})
            if not result["valid"]:
                break  # one bad repetition voids the row; no point repeating
        rows[scenario] = bc.build_row(results, summarize, METRICS, media_s)
        if rows[scenario]["samples"]:
            sample = rows[scenario]["samples"][0]
            print(
                f"  {scenario}: elapsed {rows[scenario]['stats']['elapsed_s']['median']:.2f}s"
                f" median ({sample['segments']} segments, {sample['xrt']:.1f}x realtime,"
                f" {sample['peak_rss_mb']:.0f} MB peak)"
            )
    rows["_env"] = env
    return rows


def run_scale(args: argparse.Namespace, scenarios: list[str]) -> int:
    """Run every scenario at each ``--scale`` duration and fit cost against length."""
    try:
        durations = bc.parse_scale(args.scale)
    except ValueError as exc:
        print(exc)
        return bc.EXIT_USAGE
    points: dict[str, list[tuple[float, float]]] = {}
    bad = False
    for duration in durations:
        input_dir = args.input / f"scale-{duration}"
        ensure_fixtures(input_dir, duration)
        print(f"fixture {STUDY}_P01.mp4  {duration}s speech")
        rows = sweep(args, scenarios, input_dir, input_dir / "bench-out", duration)
        rows.pop("_env", None)
        for scenario, row in rows.items():
            if not row["samples"]:
                bad = True
                continue
            for key in ("elapsed_s", "speakers_cluster_s", "peak_rss_mb"):
                value = bc.aggregate(row["samples"], key)["median"]
                if key == "speakers_cluster_s" and scenario != "speakers":
                    continue
                points.setdefault(f"{scenario}:{key}", []).append(
                    (float(duration), value)
                )
    bc.print_scaling(points)
    return bc.EXIT_INVALID if bad else bc.EXIT_OK


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument(
        "--scenarios",
        default=",".join(SCENARIOS),
        help="comma-separated scenario list (default: whisper, speakers)",
    )
    ap.add_argument(
        "--input",
        default=None,
        type=Path,
        help="fixture dir (default: /tmp/tsbench, or /tmp/tsbench-session)",
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
        help="speech fixture length in seconds (default: 300, or 3600 for session)",
    )
    ap.add_argument(
        "--model",
        default="base",
        choices=("tiny", "base", "small", "medium", "large-v3"),
        help="Whisper model (default: base, clipgen's default)",
    )
    bc.add_common_args(ap, "whisper")
    bc.add_long_args(ap, TIERS)
    args = ap.parse_args()
    bc.validate_common_args(ap, args)
    bc.apply_tier(ap, args, TIERS)

    scenarios = [s.strip() for s in args.scenarios.split(",") if s.strip()]
    unknown = [s for s in scenarios if s not in SCENARIOS]
    if unknown:
        print(f"unknown scenarios: {', '.join(unknown)} (know: {', '.join(SCENARIOS)})")
        return bc.EXIT_USAGE
    if args.scale:
        return run_scale(args, scenarios)

    if args.source is not None:
        input_dir = args.input / "source"
        video = bc.link_source(args.source, input_dir, f"{STUDY}_P01.mp4")
        probed = bc.probe_fixture(video)
        if probed is None or not probed["has_audio"]:
            print(f"cannot probe an audio track in {args.source}")
            return bc.EXIT_USAGE
        write_sheet(input_dir)
        spec: dict[str, Any] = {"source": bc.source_identity(args.source)}
    else:
        input_dir = args.input
        try:
            ensure_fixtures(input_dir, args.duration)
        except RuntimeError as exc:
            print(exc)
            return bc.EXIT_USAGE
        video = input_dir / f"{STUDY}_P01.mp4"
        probed = bc.probe_fixture(video)
        spec = fixture_spec(args.duration)
    out_root = args.output or (args.input / "bench-out")
    media_s = float(probed["duration"]) if probed else 0.0
    print(f"fixture {video.name}  {media_s:.0f}s  model={args.model}")

    rows = sweep(args, scenarios, input_dir, out_root, media_s)
    workload: dict[str, Any] = {"model": args.model}
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
    metric = "elapsed_s" if args.fail_metric == "elapsed" else "whisper_s"
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
            count_key="segments",
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
