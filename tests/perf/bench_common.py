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
import itertools
import json
import math
import statistics
import subprocess
import time
from pathlib import Path
from typing import Any, TypedDict

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


# Sets config attributes, then runs the launcher exactly as `uv run clipgen.py` would.
_OVERRIDE_LAUNCHER = """
import json, runpy, sys
sys.path.insert(0, "source")
import config
for key, value in json.loads(sys.argv[1]).items():
    setattr(config, key, value)
sys.argv = ["clipgen.py", *sys.argv[2:]]
runpy.run_path("clipgen.py", run_name="__main__")
"""


def run_clipgen(
    args: list[str],
    out_dir: Path,
    *,
    required_labels: list[str],
    timeout: float = 1800,
    overrides: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Run ``clipgen.py`` once and judge the run; never raises on failure.

    Returns ``{valid, reasons, elapsed_s, returncode, doc, output}``. The run
    is invalid on a non-zero exit, a timeout, a missing or unreadable
    ``profile.json``, a deep-profile run (its totals carry cProfile's
    overhead), or any missing required label. ``elapsed_s`` is the subprocess
    wall time, including the ``uv run`` launch — the end-to-end number a user
    waits for.

    *overrides* sets ``config`` attributes before ``cli.main`` runs, for
    settings no CLI flag reaches (``REENCODING``, ``MAX_FILESIZE_MB``).
    """
    profile_path = out_dir / "profile.json"
    tail = [*args, "--profile-output", str(profile_path)]
    if overrides:
        cmd = ["uv", "run", "python", "-c", _OVERRIDE_LAUNCHER, json.dumps(overrides)]
        cmd += tail
    else:
        cmd = ["uv", "run", "clipgen.py", *tail]
    t0 = time.perf_counter()
    reasons: list[str] = []
    try:
        proc = subprocess.run(
            cmd,
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
        returncode, output = proc.returncode, proc.stdout + proc.stderr
    except subprocess.TimeoutExpired as exc:
        returncode = -1
        output = _as_text(exc.stdout) + _as_text(exc.stderr)
        reasons.append(f"timeout after {timeout:.0f}s")
    elapsed = time.perf_counter() - t0
    if returncode and not reasons:
        reasons.append(f"exit {returncode}")
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
        "returncode": returncode,
        "doc": doc,
        "output": output,
    }


def _as_text(data: str | bytes | None) -> str:
    if isinstance(data, bytes):
        return data.decode("utf-8", "replace")
    return data or ""


def print_failure(name: str, result: dict[str, Any]) -> None:
    """One ``!`` line per reason plus the run's last five output lines."""
    for reason in result["reasons"]:
        print(f"  ! {name}: {reason}")
    tail = "\n".join(result["output"].strip().splitlines()[-5:])
    if tail:
        print("    " + tail.replace("\n", "\n    "))


def probe_fixture(path: Path, *, keyframes: bool = False) -> dict[str, Any] | None:
    """``{duration, width, height, fps, has_audio}`` of *path*, or None.

    With *keyframes*, also ``keyframe_gap``: the widest keyframe spacing in
    the first minute, so a 1 s-GOP leftover cannot pass for an OBS one.
    """
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
    probed: dict[str, Any] = {
        "duration": duration,
        "width": int(video.get("width", 0)),
        "height": int(video.get("height", 0)),
        "fps": fps,
        "has_audio": any(
            s.get("codec_type") == "audio" for s in info.get("streams", [])
        ),
    }
    if keyframes:
        probed["keyframe_gap"] = probe_keyframe_gap(path)
    return probed


def probe_keyframe_gap(path: Path, window: float = 60.0) -> float | None:
    """Widest keyframe spacing in the first *window* seconds, or None."""
    try:
        proc = subprocess.run(
            ["ffprobe", "-v", "error", "-select_streams", "v:0"]
            + ["-read_intervals", f"%+{window}", "-show_entries"]
            + ["packet=pts_time,flags", "-of", "csv=p=0", str(path)],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError:
        return None
    times = []
    for line in proc.stdout.splitlines():
        pts, _, flags = line.partition(",")
        if "K" in flags:
            try:
                times.append(float(pts))
            except ValueError:
                continue
    times.sort()
    if len(times) < 2:
        return None
    return round(max(b - a for a, b in itertools.pairwise(times)), 3)


# Probed fields compared within a tolerance (seconds) instead of exactly.
_TOLERANCE = {"duration": 0.5, "keyframe_gap": 0.5}


def fixture_matches(probed: dict[str, Any] | None, spec: dict[str, Any]) -> bool:
    """True when every spec field matches; durations within half a second."""
    if probed is None:
        return False
    for key, want in spec.items():
        have = probed.get(key)
        if key in _TOLERANCE:
            if have is None or abs(float(have) - float(want)) >= _TOLERANCE[key]:
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
    media_s: float | None = None,
) -> dict[str, Any]:
    """Reduce one scenario's repetitions to the persisted row.

    Every sample is kept; ``stats`` aggregates *metrics*. Any invalid
    repetition invalidates the row — a good run must never hide a bad one.
    With *media_s*, each sample also carries :func:`add_rates`' fields.
    """
    samples = [
        dict(summarize(r["doc"]), elapsed_s=r["elapsed_s"])
        for r in results
        if r["valid"]
    ]
    if media_s:
        samples = [add_rates(sample, media_s) for sample in samples]
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
        default=None,
        type=int,
        help="fresh-process repetitions per scenario; every sample is kept "
        "(default: 3, or the tier's)",
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


# ---- Long media: tiers, real sources, normalized rates, scaling -----------------


class VideoShape(TypedDict):
    """The ``make_long_video`` keywords that define a fixture's look."""

    size: str
    rate: int
    gop: int
    content: str


# The `session` fixture: an OBS 1080p recording with fixed 250-frame keyframes.
SESSION_VIDEO: VideoShape = {
    "size": "1920x1080",
    "rate": 30,
    "gop": 250,
    "content": "screen",
}


def session_spec(duration: int, *, audio: str) -> dict[str, Any]:
    """Fixture identity for a ``session``-tier :func:`bench_fixtures.make_long_video`."""
    width, height = (int(v) for v in SESSION_VIDEO["size"].split("x"))
    rate = SESSION_VIDEO["rate"]
    return {
        "duration": duration,
        "width": width,
        "height": height,
        "fps": float(rate),
        "has_audio": audio != "none",
        "keyframe_gap": round(SESSION_VIDEO["gop"] / rate, 3),
        "content": SESSION_VIDEO["content"],
        "audio": audio,
    }


def ensure_video(video: Path, spec: dict[str, Any], build: Any) -> None:
    """Run *build* unless *video* already matches *spec*.

    Keys ffprobe cannot see (``content``, ``audio``) are recorded in a hidden
    sidecar at build time and merged under the probe, so a ``testsrc`` leftover
    never passes for a ``screen`` fixture of the same geometry.
    """
    sidecar = video.with_name(f".{video.name}.spec.json")
    probed = None
    if video.is_file():
        probed = (
            probe_fixture(video, keyframes=True)
            if "keyframe_gap" in spec
            else probe_fixture(video)
        )
    have = None
    if probed is not None:
        try:
            recorded = json.loads(sidecar.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            recorded = {}
        have = {**recorded, **probed}
    if fixture_matches(have, spec):
        return
    if video.is_file():
        was = f"{probed['duration']:.0f}s" if probed else "unreadable"
        print(f"rebuilding {video.name} ({was} → {spec['duration']}s)")
        video.unlink()
    sidecar.unlink(missing_ok=True)
    build()
    sidecar.write_text(json.dumps(spec), encoding="utf-8")


def source_identity(path: Path) -> dict[str, Any]:
    """What a ``--source`` compare must hold constant; basename only, never the path."""
    st = path.stat()
    return {"name": path.name, "bytes": st.st_size, "mtime": int(st.st_mtime)}


def link_source(source: Path, dest_dir: Path, name: str) -> Path:
    """Symlink a real recording into *dest_dir* as *name*; the media is never copied."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    link = dest_dir / name
    if link.is_symlink() or link.exists():
        link.unlink()
    link.symlink_to(source.resolve())
    return link


def add_long_args(ap: argparse.ArgumentParser, tiers: dict[str, Any]) -> None:
    """``--tier``, ``--source``, ``--scale`` and ``--timeout``, shared by every bench."""
    ap.add_argument(
        "--tier",
        choices=tuple(tiers),
        default="smoke",
        help="smoke: short regression gate (default); "
        "session: hour-long OBS-like 1080p fixture, one run",
    )
    ap.add_argument(
        "--source",
        type=Path,
        default=None,
        help="a real local recording to bench instead of a fixture "
        "(symlinked, never copied or saved; JSON keeps only its basename)",
    )
    ap.add_argument(
        "--scale",
        default=None,
        help="comma-separated fixture durations in seconds; runs each and "
        "fits cost against length (diagnostic: no --save/--compare)",
    )
    ap.add_argument(
        "--timeout",
        type=float,
        default=None,
        help="per-run subprocess timeout in seconds (default: max(1800, 3× media))",
    )


def apply_tier(
    ap: argparse.ArgumentParser, args: argparse.Namespace, tiers: dict[str, Any]
) -> None:
    """Fill every flag left unset (None) from the chosen tier; reject bad combos."""
    for key, value in tiers[args.tier].items():
        if getattr(args, key, None) is None:
            setattr(args, key, value)
    if args.runs is None:
        args.runs = 3
    if args.scale and (args.save or args.compare or args.source):
        ap.error("--scale is a diagnostic; drop --save/--compare/--source")
    if args.source is not None and not args.source.is_file():
        ap.error(f"--source {args.source} is not a file")


def parse_scale(text: str) -> list[int]:
    """``"300,1800,3600"`` → sorted unique durations; at least two."""
    durations = sorted({int(float(v)) for v in text.split(",") if v.strip()})
    if len(durations) < 2 or durations[0] <= 0:
        raise ValueError("--scale needs at least two positive durations")
    return durations


def timeout_for(media_s: float, override: float | None) -> float:
    """Long runs get room: three times the media length, never under 30 min."""
    if override:
        return float(override)
    return max(1800.0, 3.0 * float(media_s))


def add_rates(sample: dict[str, Any], media_s: float) -> dict[str, Any]:
    """Annotate *sample* with ``media_s``, ``xrt`` and ``s_per_media_hour``.

    *media_s* is the media the run actually processed — the scanned or
    transcribed source, or the summed clip lengths — so runs over different
    durations read side by side, and a 120 s number is never silently
    extrapolated to a three-hour session.
    """
    elapsed = float(sample.get("elapsed_s") or 0.0)
    sample["media_s"] = float(media_s)
    sample["xrt"] = media_s / elapsed if elapsed else 0.0
    sample["s_per_media_hour"] = elapsed / media_s * 3600 if media_s else 0.0
    return sample


def fit_scaling(points: list[tuple[float, float]]) -> dict[str, Any]:
    """Fit cost against media length; *points* are ``(seconds_of_media, cost_s)``.

    Returns the least-squares ``intercept`` and ``slope_per_hour`` plus the
    ``exponent`` between the two longest points (log-log), where fixed
    start-up overhead no longer masks the growth: ~0 is flat, ~1 linear,
    > 1.15 superlinear.
    """
    pts = sorted((float(x), float(y)) for x, y in points if x > 0)
    if len(pts) < 2:
        return {
            "intercept": 0.0,
            "slope_per_hour": 0.0,
            "exponent": None,
            "verdict": "n/a",
        }
    n = len(pts)
    mx = sum(x for x, _ in pts) / n
    my = sum(y for _, y in pts) / n
    var = sum((x - mx) ** 2 for x, _ in pts)
    slope = sum((x - mx) * (y - my) for x, y in pts) / var if var else 0.0
    (x1, y1), (x2, y2) = pts[-2], pts[-1]
    exponent = None
    if y1 > 0 and y2 > 0 and x2 > x1:
        exponent = math.log(y2 / y1) / math.log(x2 / x1)
    if exponent is None:
        verdict = "n/a"
    elif exponent > 1.15:
        verdict = "superlinear"
    elif exponent < 0.15:
        verdict = "flat"
    elif exponent < 0.85:
        verdict = "sublinear"
    else:
        verdict = "linear"
    return {
        "intercept": my - slope * mx,
        "slope_per_hour": slope * 3600,
        "exponent": exponent,
        "verdict": verdict,
    }


def print_scaling(
    points: dict[str, list[tuple[float, float]]],
    *,
    unit: str = "media hour",
    per: float = 3600.0,
) -> None:
    """One fit line per row name: cost per *unit* (*per* x-units) and the verdict."""
    print(
        f"\n{'scaling':<24}{'points':>7}{'per ' + unit:>16}{'base':>9}{'exp':>7}  verdict"
    )
    for name, pts in points.items():
        fit = fit_scaling(pts)
        exp = f"{fit['exponent']:.2f}" if fit["exponent"] is not None else "-"
        slope = fit["slope_per_hour"] / 3600.0 * per
        print(
            f"{name:<24}{len(pts):>7}{slope:>15.3f}s"
            f"{fit['intercept']:>8.2f}s{exp:>7}  {fit['verdict']}"
        )
        for x, y in sorted(pts):
            print(f"{'':<24}{x:>7.0f} -> {y:.3f}")
