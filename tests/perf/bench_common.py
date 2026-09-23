"""Shared plumbing for scan_bench.py and clip_bench.py.

Both benches drive the real CLI as a fresh subprocess per repetition, read
the ``--profile-output`` JSON it writes (never the console lines), keep every
sample, compare medians, and refuse comparisons whose workload or
environment differ. This module holds the parts they share:

- ``run_clipgen`` — one fresh-process run with an explicit validity verdict.
  "Fresh process" means a new interpreter and a new ffmpeg; the operating
  system's file cache is whatever it is, so the first repetition is not
  "cold" in any meaningful sense and is kept like the others.
- ``probe_fixture`` / ``fixture_matches`` — the fixture identity a
  comparison must hold constant (duration alone let a 120 s leftover pass
  for a 15 s one).
- ``aggregate`` — median / min / max / MAD over retained samples.
- ``compare`` — a typed list of findings: ``invalid`` (a run failed),
  ``incomparable`` (different work), ``regression`` (slower). A failed or
  partial run can never read as a speedup.
- ``env_diff`` / ``meta_diff`` — what differs between two snapshots.

Exit codes: 0 ok, 1 regression, 2 usage, 3 invalid run, 4 incomparable.
Unit-tested in test_scan_bench_parse.py / test_clip_bench_parse.py.
"""

from __future__ import annotations

import argparse
import json
import statistics
import subprocess
import time
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[2]

EXIT_OK = 0
EXIT_REGRESSION = 1
EXIT_USAGE = 2
EXIT_INVALID = 3
EXIT_INCOMPARABLE = 4

# Ordered by severity: the worst finding decides the exit code.
_EXIT_BY_KIND = {
    "invalid": EXIT_INVALID,
    "incomparable": EXIT_INCOMPARABLE,
    "regression": EXIT_REGRESSION,
}

_ENV_KEYS = ("platform", "chip", "arch", "cpu_count", "memory_mb", "ffmpeg")


def run_clipgen(
    args: list[str],
    out_dir: Path,
    *,
    required_labels: list[str],
    timeout: float = 1800,
) -> dict[str, Any]:
    """Run ``clipgen.py`` once and judge the run; never raises on failure.

    Returns ``{valid, reasons, elapsed_s, returncode, doc, output}``. The run
    is invalid on a non-zero exit, a missing or unreadable ``profile.json``, a
    deep-profile run (its totals carry cProfile's overhead), or any missing
    required label. ``elapsed_s`` is the subprocess wall time, including the
    ``uv run`` launch — the end-to-end number a user waits for.
    """
    profile_path = out_dir / "profile.json"
    cmd = [
        "uv",
        "run",
        "clipgen.py",
        *args,
        "--profile-output",
        str(profile_path),
    ]
    t0 = time.perf_counter()
    proc = subprocess.run(
        cmd, cwd=REPO_ROOT, capture_output=True, text=True, timeout=timeout, check=False
    )
    elapsed = time.perf_counter() - t0
    reasons: list[str] = []
    if proc.returncode:
        reasons.append(f"exit {proc.returncode}")
    doc: dict[str, Any] = {}
    try:
        doc = json.loads(profile_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        reasons.append("no profile.json")
    if doc.get("mode") == "deep":
        reasons.append("deep run")
    labels = doc.get("labels", {})
    reasons.extend(
        f"missing {label}" for label in required_labels if label not in labels
    )
    return {
        "valid": not reasons,
        "reasons": reasons,
        "elapsed_s": elapsed,
        "returncode": proc.returncode,
        "doc": doc,
        "output": proc.stdout + proc.stderr,
    }


def print_failure(name: str, result: dict[str, Any]) -> None:
    """One ``!`` line per reason plus the run's last five output lines."""
    for reason in result["reasons"]:
        print(f"  ! {name}: {reason}")
    tail = "\n".join(result["output"].strip().splitlines()[-5:])
    if tail:
        print("    " + tail.replace("\n", "\n    "))


def probe_fixture(path: Path) -> dict[str, Any] | None:
    """``{duration, width, height, fps, has_audio}`` of *path*, or None."""
    try:
        proc = subprocess.run(
            [
                "ffprobe",
                "-v",
                "error",
                "-show_streams",
                "-show_format",
                "-of",
                "json",
                str(path),
            ],
            capture_output=True,
            text=True,
            check=False,
        )
        info = json.loads(proc.stdout)
    except (OSError, ValueError):
        return None
    video = next(
        (s for s in info.get("streams", []) if s.get("codec_type") == "video"), None
    )
    if video is None:
        return None
    num, _, den = str(video.get("r_frame_rate", "0/1")).partition("/")
    try:
        fps = float(num) / float(den or 1)
        duration = float(info.get("format", {}).get("duration", 0.0))
    except ValueError:
        return None
    return {
        "duration": duration,
        "width": int(video.get("width", 0)),
        "height": int(video.get("height", 0)),
        "fps": fps,
        "has_audio": any(
            s.get("codec_type") == "audio" for s in info.get("streams", [])
        ),
    }


def fixture_matches(probed: dict[str, Any] | None, spec: dict[str, Any]) -> bool:
    """True when every spec field matches; duration within half a second."""
    if probed is None:
        return False
    for key, want in spec.items():
        have = probed.get(key)
        if key == "duration":
            if have is None or abs(float(have) - float(want)) >= 0.5:
                return False
        elif have != want:
            return False
    return True


def aggregate(samples: list[dict[str, Any]], key: str) -> dict[str, float]:
    """``{median, min, max, mad, n}`` of *key* across *samples*."""
    values = [float(s[key]) for s in samples if key in s]
    if not values:
        return {"median": 0.0, "min": 0.0, "max": 0.0, "mad": 0.0, "n": 0}
    med = statistics.median(values)
    mad = statistics.median(abs(v - med) for v in values)
    return {
        "median": med,
        "min": min(values),
        "max": max(values),
        "mad": mad,
        "n": len(values),
    }


def build_row(
    results: list[dict[str, Any]],
    summarize: Any,
    metrics: tuple[str, ...],
) -> dict[str, Any]:
    """Reduce one scenario's repetitions to the persisted row.

    Every sample is kept; ``stats`` aggregates *metrics*. Any invalid
    repetition invalidates the row — a good run must never hide a bad one.
    """
    samples = [
        dict(summarize(r["doc"]), elapsed_s=r["elapsed_s"])
        for r in results
        if r["valid"]
    ]
    reasons = [reason for r in results for reason in r["reasons"]]
    return {
        "valid": all(r["valid"] for r in results) and bool(results),
        "reasons": reasons,
        "samples": samples,
        "stats": {metric: aggregate(samples, metric) for metric in metrics},
    }


def delta_pct(current: float, base: float) -> float | None:
    """Percent change, or None when *base* is 0."""
    if not base:
        return None
    return (current / base - 1.0) * 100.0


def compare(
    rows: dict[str, dict[str, Any]],
    baseline: dict[str, dict[str, Any]],
    *,
    metric: str,
    pct: float | None,
    abs_s: float | None,
    count_key: str,
) -> list[tuple[str, str, str]]:
    """Findings as ``(kind, name, detail)``; empty means every row is fine.

    Medians are compared. With both thresholds given, a regression needs
    both to be exceeded. A row whose work count differs from the baseline's
    is ``incomparable``, never a regression or a speedup.
    """
    findings: list[tuple[str, str, str]] = []
    for name, row in rows.items():
        if not row["valid"]:
            findings.append(("invalid", name, "; ".join(row["reasons"]) or "no runs"))
            continue
        base = baseline.get(name)
        if base is None or not base.get("valid"):
            findings.append(("incomparable", name, "no valid baseline row"))
            continue
        have = aggregate(row["samples"], count_key)["median"]
        want = aggregate(base["samples"], count_key)["median"]
        if have != want:
            findings.append(
                ("incomparable", name, f"{count_key} {have:g} vs baseline {want:g}")
            )
            continue
        current = row["stats"][metric]["median"]
        ref = base["stats"][metric]["median"]
        change = delta_pct(current, ref)
        over_pct = pct is not None and change is not None and change > pct
        over_abs = abs_s is not None and (current - ref) > abs_s
        if pct is not None and abs_s is not None:
            hit = over_pct and over_abs
        else:
            hit = over_pct or over_abs
        if hit:
            shown = f"{change:+.1f}%" if change is not None else "n/a"
            findings.append(
                ("regression", name, f"{metric} {ref:.3f}s -> {current:.3f}s ({shown})")
            )
    return findings


def env_diff(a: dict[str, Any], b: dict[str, Any]) -> list[str]:
    """Environment facts that differ between two ``env`` blocks.

    The commit is deliberately not compared: it always differs between a
    baseline and the branch under test, which is the point of comparing.
    """
    out: list[str] = []
    for key in _ENV_KEYS:
        if a.get(key) != b.get(key):
            out.append(f"{key}: {a.get(key)} -> {b.get(key)}")
    py_a = ".".join(str(a.get("python", "")).split(".")[:2])
    py_b = ".".join(str(b.get("python", "")).split(".")[:2])
    if py_a != py_b:
        out.append(f"python: {py_a} -> {py_b}")
    sa, sb = a.get("settings", {}), b.get("settings", {})
    for key in sorted(set(sa) | set(sb)):
        if sa.get(key) != sb.get(key):
            out.append(f"{key}: {sa.get(key)} -> {sb.get(key)}")
    return out


def meta_diff(base: dict[str, Any], meta: dict[str, Any]) -> list[str]:
    """Fixture-spec and workload fields that differ; these void a compare."""
    out: list[str] = []
    for section in ("fixture", "workload"):
        left = base.get(section, {})
        right = meta.get(section, {})
        if section == "fixture":
            left, right = left.get("spec", {}), right.get("spec", {})
        for key in sorted(set(left) | set(right)):
            if left.get(key) != right.get(key):
                out.append(f"{section}.{key}: {left.get(key)} -> {right.get(key)}")
    return out


def status_of(name: str, findings: list[tuple[str, str, str]]) -> str:
    """The row's status word for the table."""
    for kind in ("invalid", "incomparable", "regression"):
        if any(k == kind and n == name for k, n, _ in findings):
            return kind
    return "ok"


def exit_code(findings: list[tuple[str, str, str]]) -> int:
    """Worst finding wins: invalid over incomparable over regression."""
    for kind, code in _EXIT_BY_KIND.items():
        if any(k == kind for k, _, _ in findings):
            return code
    return EXIT_OK


def add_common_args(ap: argparse.ArgumentParser, breakdown: str) -> None:
    """The flags both benches share; *breakdown* names the secondary metric."""
    ap.add_argument(
        "--runs",
        default=3,
        type=int,
        help="fresh-process repetitions per scenario; every sample is kept",
    )
    ap.add_argument("--save", type=Path, help="write results JSON here")
    ap.add_argument("--compare", type=Path, help="baseline JSON to diff against")
    ap.add_argument(
        "--fail-on",
        type=float,
        default=None,
        help="with --compare, flag a median that rose by more than this percent",
    )
    ap.add_argument(
        "--fail-metric",
        choices=("elapsed", breakdown),
        default="elapsed",
        help="metric the thresholds apply to (default: subprocess elapsed)",
    )
    ap.add_argument(
        "--fail-abs",
        type=float,
        default=None,
        help="absolute seconds a median may rise; with --fail-on both must trip",
    )
    ap.add_argument(
        "--allow-env-mismatch",
        action="store_true",
        help="compare across differing machines/settings anyway",
    )


def validate_common_args(ap: argparse.ArgumentParser, args: argparse.Namespace) -> None:
    if (args.fail_on is not None or args.fail_abs is not None) and not args.compare:
        ap.error("--fail-on/--fail-abs require --compare")
    if getattr(args, "deep", False) and (args.save or args.compare):
        ap.error("--deep runs are diagnostics; drop --save/--compare")


def load_baseline(
    path: Path, section: str, meta: dict[str, Any], *, allow_env: bool
) -> tuple[dict[str, Any] | None, list[str]]:
    """``(rows, problems)``; rows is None when the snapshot cannot be compared."""
    try:
        snapshot = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        return None, [f"cannot read {path}: {exc}"]
    base_meta = snapshot.get("meta", {})
    problems = meta_diff(base_meta, meta)
    env_problems = env_diff(base_meta.get("env", {}), meta.get("env", {}))
    if env_problems:
        for line in env_problems:
            print(f"env: {line}")
        if not allow_env:
            problems.extend(env_problems)
    if problems:
        return None, problems
    return snapshot.get(section, {}), []


def save_results(
    path: Path, meta: dict[str, Any], section: str, rows: dict[str, Any]
) -> None:
    path.write_text(
        json.dumps({"meta": meta, section: rows}, indent=2), encoding="utf-8"
    )
    print(f"\nsaved -> {path}")


def report_findings(findings: list[tuple[str, str, str]]) -> None:
    for kind, name, detail in findings:
        print(f"{kind}: {name} {detail}")


def stat_cells(row: dict[str, Any], metric: str) -> str:
    """``median   min   mad`` cells for the table."""
    st = row["stats"][metric]
    return f"{st['median']:>8.3f}s{st['min']:>8.3f}s{st['mad']:>7.3f}s"
