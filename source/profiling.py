"""Opt-in performance instrumentation, gated on ``config.PROFILING``.

Recording functions (``add``, ``count``, ``span``, ``scan_summary``) are no-ops
when ``config.PROFILING`` is False, so instrumented call sites cost one boolean
check when profiling is off. Unlike ``config.DEBUGGING``, enabling profiling
must never change what work runs — it only measures.

Hot loops (per-frame scan callbacks) must NOT call into this module per
iteration: hoist ``if config.PROFILING:`` before the loop, accumulate into
locals, and flush once per scan via ``add()``. ``span()`` is for coarse work
only — one ffmpeg subprocess, one cache compute, one HTTP request.
``stream_span()`` is the equivalent for a streaming response body, which
``span()`` cannot reach (see its docstring).

Output is one grep-able line per label with a fixed ``profile | `` prefix,
printed when the process exits (``enable()`` registers an atexit hook) and
regardless of verbosity. Report lines use bare ``print`` deliberately, not the
``utils`` helpers: Rich's console wraps at ~80 columns when stdout is piped,
which would split long route labels mid-line and break ``grep "profile |"`` —
the exact consumer this output exists for. ``--profile-output PATH`` writes the
same data as full-precision JSON (:func:`export`), which the benchmark runners
consume instead of parsing the console lines. On a live server, ``GET
/api/profile`` returns :func:`export` and ``?reset=1`` brackets a measurement
window. See agents/skills/profile/SKILL.md.

A window is defined by *completion*: every span records in its ``finally``, so
work that started before a reset and finishes after it lands in the new
window. :func:`active_spans` names what is in flight at snapshot time so a
reader can tell which totals a boundary cut through.
"""

import atexit
import cProfile
import functools
import io
import itertools
import platform
import pstats
import subprocess
import sys
import threading
import time
from collections.abc import Callable, Generator, Iterable, Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

import config

_LOCK = threading.Lock()
# label -> [total_seconds, count, max_seconds, bytes, first_seconds]. Cap is a safety net.
_TOTALS: dict[str, list[float]] = {}
_MAX_LABELS = 1024
_DROPPED = 0  # labels refused at the cap; reported, never silent
_REPORT_REGISTERED = False
_WINDOW_T0: float | None = None

# Spans in flight: id -> (label, start, thread ident). Listed, never summed.
_ACTIVE: dict[int, tuple[str, float, int]] = {}
_ACTIVE_IDS = itertools.count(1)
_MAX_ACTIVE = 256

# --profile-deep cProfile per (label, thread): one Profile cannot enable from two threads.
_DEEP: dict[tuple[str, int], cProfile.Profile] = {}
_DEEP_DONE: set[str] = set()  # labels whose deep-profiled span finished at least once
_MAX_DEEP = 16
_DEEP_TOP = 15  # rows of pstats output per label

# Recorded even with profiling off (first mark precedes argv parsing); reported only when on.
_STARTUP_T0: float | None = None
_STARTUP_MARKS: list[tuple[str, float]] = []
_MAX_STARTUP_MARKS = 64

# The tuning knobs a comparison must hold constant (agents/PERFORMANCE.md).
_ENV_SETTINGS = (
    "CLIP_PARALLEL_WORKERS",
    "SCREENSPACE_PARALLEL_WORKERS",
    "SCREENSPACE_OCR_POOL_SIZE",
    "SCREENSPACE_CV_RESOLUTION_SCALE",
    "SCREENSPACE_FAST_SCAN_SKIP_NONKEY",
    "FFMPEG_VIDEO_ENCODER",
    "TITLECARD_ENCODE_PRESET",
    "TITLECARD_ENCODE_CRF",
    "TRANSCRIBE_DEVICE",
    "TRANSCRIBE_CPU_THREADS",
    "TRANSCRIBE_BEAM_SIZE",
    "TRANSCRIBE_VAD_FILTER",
    "TRANSCRIBE_WORD_TIMESTAMPS",
    "WORKFLOWS_BATCH_WORKERS",
)
_ENV_PACKAGES = (
    "faster-whisper",
    "ctranslate2",
    "rapidocr",
    "onnxruntime",
    "opencv-python-headless",
    "numpy",
    "flask",
)


def set_process_start(t0: float) -> None:
    """Anchor startup marks to *t0* (a ``time.perf_counter()`` reading).

    Captured as the first statement of clipgen.py so the anchor predates every
    clipgen import; without it :func:`startup_snapshot` returns nothing.
    """
    global _STARTUP_T0
    _STARTUP_T0 = t0


def mark(label: str) -> None:
    """Record a startup milestone unconditionally (see the module-level note)."""
    now = time.perf_counter()
    with _LOCK:
        if len(_STARTUP_MARKS) < _MAX_STARTUP_MARKS:
            _STARTUP_MARKS.append((label, now))


def startup_snapshot() -> list[dict[str, Any]]:
    """Return ``[{label, at_ms, delta_ms}]`` in record order; empty without T0.

    ``at_ms`` is time since process start, ``delta_ms`` since the previous mark.
    Marks from concurrent threads (the boot-build phases vs. the AppKit
    window-shown hook) interleave chronologically, so a delta spanning a thread
    boundary attributes wall time, not per-thread work.
    """
    with _LOCK:
        t0 = _STARTUP_T0
        marks = list(_STARTUP_MARKS)
    if t0 is None or not marks:
        return []
    out: list[dict[str, Any]] = []
    prev = t0
    for label, at in marks:
        out.append(
            {
                "label": label,
                "at_ms": (at - t0) * 1000,
                "delta_ms": (at - prev) * 1000,
            }
        )
        prev = at
    return out


def report_startup() -> None:
    """Print one ``startup | `` line per recorded mark; silent when empty."""
    for entry in startup_snapshot():
        # bare print: see module docstring (grep-ability over Rich wrapping)
        print(
            f"startup | {entry['label']:<32} +{entry['delta_ms']:8.1f}ms"
            f"  t={entry['at_ms']:8.1f}ms"
        )


def enable() -> None:
    """Turn profiling on and register the end-of-process report once."""
    global _REPORT_REGISTERED, _WINDOW_T0
    config.PROFILING = True
    with _LOCK:
        if _WINDOW_T0 is None:
            _WINDOW_T0 = time.perf_counter()
        if _REPORT_REGISTERED:
            return
        _REPORT_REGISTERED = True
    atexit.register(report)


def add(
    label: str,
    seconds: float = 0.0,
    n: int = 1,
    *,
    peak: float | None = None,
    nbytes: int = 0,
) -> None:
    """Accumulate *seconds*, *n* occurrences and *nbytes* under *label*.

    *nbytes* is a payload size — a response body, a streamed drain, a manifest
    section — summed like *seconds*. It is the second axis a poll can go wrong
    on: ``route`` timing alone hides an endpoint that answers in 2 ms with a
    5 MB body every tick.

    *peak* is the largest single occurrence in this contribution. A batched
    flush (``n > 1``) has no per-item max in *seconds* — that is a sum — so the
    max is left alone unless the caller tracked one in its own loop and passed
    it. Without that kwarg the tail would be permanently invisible on exactly
    the labels where it matters most (``scan.callback``, ``transcribe.decode``);
    one compare per iteration in an already-accumulating loop is not a cost.

    Percentiles are deliberately not offered: p95 needs retained samples, i.e.
    unbounded per-label memory, which this accumulator exists to refuse. Max is
    the only tail statistic that is O(1).

    The first single-occurrence contribution is also kept as ``first``: the
    first observation. A route whose first call pays a lazy import or a cache
    fill reads as a modest ``avg`` and a large ``max`` — indistinguishable
    from an occasional slow request — unless the report can say the slow one
    was the first. Batched flushes (``n > 1``) carry no per-item first and
    leave it 0. Resetting the window clears ``first`` but not the caches that
    made it slow, so a second window's first observation is usually warm.

    A label past ``_MAX_LABELS`` is refused and counted in ``dropped_labels``
    rather than silently lost.
    """
    global _DROPPED
    if not config.PROFILING:
        return
    if peak is None and n == 1:
        peak = seconds
    with _LOCK:
        entry = _TOTALS.get(label)
        if entry is None:
            if len(_TOTALS) >= _MAX_LABELS:
                _DROPPED += 1
                return
            first = seconds if n == 1 else 0.0
            _TOTALS[label] = [seconds, float(n), peak or 0.0, float(nbytes), first]
        else:
            entry[0] += seconds
            entry[1] += n
            if peak is not None and peak > entry[2]:
                entry[2] = peak
            entry[3] += nbytes


def count(label: str, n: int = 1) -> None:
    """Accumulate *n* occurrences under *label* with no time component."""
    add(label, 0.0, n)


def dropped_labels() -> int:
    """Labels refused at the cap since the last reset."""
    with _LOCK:
        return _DROPPED


def _begin_active(label: str) -> int | None:
    """Register an in-flight span; ``None`` past the cap (still timed)."""
    with _LOCK:
        if len(_ACTIVE) >= _MAX_ACTIVE:
            return None
        sid = next(_ACTIVE_IDS)
        _ACTIVE[sid] = (label, time.perf_counter(), threading.get_ident())
        return sid


def _end_active(sid: int | None) -> None:
    if sid is None:
        return
    with _LOCK:
        _ACTIVE.pop(sid, None)


def active_spans() -> list[dict[str, Any]]:
    """Spans in flight right now: ``[{label, elapsed_s, thread}]``, oldest first."""
    now = time.perf_counter()
    with _LOCK:
        items = sorted(_ACTIVE.values(), key=lambda item: item[1])
    return [
        {"label": label, "elapsed_s": now - start, "thread": tid}
        for label, start, tid in items
    ]


def deep_profiler(label: str) -> cProfile.Profile | None:
    """Per-(label, thread) cProfile when *label* matches ``config.PROFILE_DEEP``.

    Returns ``None`` unless profiling is on **and** ``PROFILE_DEEP`` is a
    non-empty substring of *label* — the common case is two string checks with
    no lock. Callers bracket exactly the hot work with ``enable()``/
    ``disable()``; the same object is handed back for repeated spans on one
    thread, so stats accumulate across frames. The stopwatch totals recorded
    while a deep profile is attached include cProfile's own overhead — use a
    deep run to find function names, never to compare against a plain run.
    """
    target = config.PROFILE_DEEP
    if not config.PROFILING or not target or target not in label:
        return None
    key = (label, threading.get_ident())
    with _LOCK:
        prof = _DEEP.get(key)
        if prof is None:
            if len(_DEEP) >= _MAX_DEEP:
                return None
            prof = cProfile.Profile()
            _DEEP[key] = prof
        return prof


def deep_enable(prof: cProfile.Profile) -> bool:
    """Enable *prof* on this thread; ``False`` if a profiler is already active.

    A broad ``--profile-deep`` substring can match both an outer span and work
    nested inside it on the same thread (``heatmap.gifs`` wraps an inline
    ``heatmap.gif`` encode when there is no rolling pair). cProfile raises on
    the second ``enable()``, and letting that propagate aborts the instrumented
    work — profiling must never change what runs. Nested matches keep the
    outermost profiler; the inner work is still attributed to it.
    """
    try:
        prof.enable()
        return True
    except ValueError:
        return False


def _deep_finish(label: str, deep: cProfile.Profile | None, deep_on: bool) -> None:
    """Disable a span's deep profiler and mark the label as having completed."""
    if deep_on and deep is not None:
        deep.disable()
        with _LOCK:
            _DEEP_DONE.add(label)


@contextmanager
def span(label: str) -> Iterator[None]:
    """Time the enclosed block under *label*; passthrough when profiling is off."""
    if not config.PROFILING:
        yield
        return
    deep = deep_profiler(label)
    deep_on = deep is not None and deep_enable(deep)
    sid = _begin_active(label)
    start = time.perf_counter()
    try:
        yield
    finally:
        _deep_finish(label, deep, deep_on)
        _end_active(sid)
        add(label, time.perf_counter() - start)


def timed(label: str) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Decorator form of ``span`` for functions with many return paths."""

    def decorate(fn: Callable[..., Any]) -> Callable[..., Any]:
        @functools.wraps(fn)
        def wrapper(*args: Any, **kwargs: Any) -> Any:
            if not config.PROFILING:
                return fn(*args, **kwargs)
            # Not span(): the profiler must enable on the executor thread doing the work.
            deep = deep_profiler(label)
            deep_on = deep is not None and deep_enable(deep)
            sid = _begin_active(label)
            start = time.perf_counter()
            try:
                return fn(*args, **kwargs)
            finally:
                _deep_finish(label, deep, deep_on)
                _end_active(sid)
                add(label, time.perf_counter() - start)

        return wrapper

    return decorate


def stream_span(body: Iterable[Any], *, rule: str) -> Generator[Any, None, None]:
    """Yield through *body*, recording its drain under the ``stream`` family.

    Four labels per *rule* (a Flask URL rule): ``stream <rule>`` is the total
    generation time with the encoded body size as ``bytes``; ``stream.first
    <rule>`` is the time to the first chunk — the perceived latency of a
    drain, which the total cannot show: a generate that streams its first
    cached clip at 10 ms and finishes at 40 s and one that sits silent for
    20 s then floods have the same total; ``stream.complete`` /
    ``stream.error`` / ``stream.disconnect <rule>`` count how the drain ended.
    Text chunks are counted as UTF-8 bytes, the size that crossed the wire.

    Streaming responses are invisible to the ``after_request`` timing that
    produces ``route <rule>``: Flask runs that hook in ``finalize_request``, on
    the ``Response`` *object*, before the WSGI server iterates the body. So a
    streamed endpoint records only the time to *construct* its generator —
    measured, a body taking 0.6 s reports 0.0 s. Every ndjson/SSE route in
    clipgen was therefore reporting ~0 ms, including the four longest
    operations in the product (``/studio/api/generate``, ``/api/reel``,
    ``/api/generate-intake``, ``/api/reel-direct``).

    Deliberately a separate ``stream`` label family rather than folding into
    ``route``: a drain's duration is dominated by server-side job execution,
    not request handling, and an 8-minute reel build sorted next to a 40 ms
    route total would only move the confusion.

    A client disconnect mid-stream arrives as ``GeneratorExit``; the body's
    ``close()`` runs inside the span so its cleanup is measured. If that
    cleanup itself fails while another exception is already unwinding, the
    original exception wins — a broken teardown must not mask the real error.
    """
    if not config.PROFILING:
        yield from body
        return
    label = f"stream {rule}"
    sid = _begin_active(label)
    start = time.perf_counter()
    total = 0
    first = True
    outcome = "error"
    try:
        for chunk in body:
            if first:
                first = False
                add(f"stream.first {rule}", time.perf_counter() - start)
            total += (
                len(chunk.encode("utf-8")) if isinstance(chunk, str) else len(chunk)
            )
            yield chunk
        outcome = "complete"
    except GeneratorExit:
        outcome = "disconnect"
        raise
    finally:
        # A plain loop never forwards close(); call it so cleanup lands inside the span.
        close = getattr(body, "close", None)
        try:
            if close is not None:
                close()
        except Exception:
            if outcome == "complete":
                raise
        finally:
            _end_active(sid)
            add(label, time.perf_counter() - start, nbytes=total)
            count(f"stream.{outcome} {rule}")


def snapshot(*, reset: bool = False) -> dict[str, dict[str, float]]:
    """Return ``{label: {seconds, count, max, bytes, first}}`` sorted by seconds desc.

    With *reset*, the copy and the clear happen under one lock acquisition, so
    a contribution can never land between them and vanish from both windows.
    ``max`` is 0.0 for labels fed only by batched flushes that supplied no
    ``peak=`` — see :func:`add`. ``bytes`` is 0 for labels that never passed
    ``nbytes=``; ``first`` (the first observation) is 0.0 for batched labels.
    Sorting stays on seconds so none of them reorders the report.
    """
    global _DROPPED, _WINDOW_T0
    with _LOCK:
        items = [
            (label, entry[0], int(entry[1]), entry[2], int(entry[3]), entry[4])
            for label, entry in _TOTALS.items()
        ]
        if reset:
            _TOTALS.clear()
            _DROPPED = 0
            _WINDOW_T0 = time.perf_counter()
    items.sort(key=lambda item: (-item[1], item[0]))
    return {
        label: {"seconds": secs, "count": n, "max": mx, "bytes": nb, "first": first}
        for label, secs, n, mx, nb, first in items
    }


def format_bytes(nbytes: int) -> str:
    """``1.2MB``-style size for report lines; bytes below 1 KB stay exact."""
    if nbytes < 1024:
        return f"{nbytes}B"
    if nbytes < 1024 * 1024:
        return f"{nbytes / 1024:.1f}KB"
    return f"{nbytes / (1024 * 1024):.1f}MB"


def peak_rss_mb() -> float | None:
    """Process peak RSS in MB; ``None`` where the platform will not say.

    ``ru_maxrss`` units are platform-defined and nothing reports which: macOS
    gives bytes, Linux/BSD kilobytes. Branch on the platform rather than infer
    from magnitude — a 3 GB Linux process and a 3 MB macOS one produce the same
    integer. Windows has no ``resource`` module and is skipped rather than
    pulling in psutil for one number (a C-extension wheel PyInstaller would have
    to collect into every bundle).

    Deliberately process-global and monotonic: it puts a number behind the
    "multiplies peak RAM/VRAM" claims on Whisper model size and
    ``SCREENSPACE_OCR_POOL_SIZE`` without pretending to attribute memory to
    labels — so ``?reset=1`` does not and cannot reset it. ``RUSAGE_SELF``
    excludes ffmpeg subprocesses; for clipgen the memory that hurts (Whisper
    weights, OCR engines, decoded frames) is all in-process.
    """
    try:
        import resource  # POSIX only
    except ImportError:
        return None
    raw = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    scale = 1 if sys.platform == "darwin" else 1024  # bytes vs kilobytes
    return raw * scale / (1024 * 1024)


def reset() -> None:
    """Clear the accumulated totals (brackets a measurement window).

    Deep profilers are left alone: a thread may be inside one right now, and
    ``--profile-deep`` output is a diagnostic artifact, not a window total.
    """
    snapshot(reset=True)


def deep_reset() -> None:
    """Drop every deep profiler; only for tests that share a process."""
    with _LOCK:
        _DEEP.clear()
        _DEEP_DONE.clear()


def _run_text(argv: list[str], timeout: float) -> str | None:
    """First stdout line of *argv*, or ``None`` when it cannot run."""
    try:
        proc = subprocess.run(
            argv, capture_output=True, text=True, timeout=timeout, check=False
        )
    except (OSError, subprocess.SubprocessError):
        return None
    if proc.returncode:
        return None
    return proc.stdout.strip().splitlines()[0] if proc.stdout.strip() else ""


def _git_state() -> tuple[str | None, bool | None]:
    """``(commit, dirty)`` of the source checkout; both ``None`` when frozen."""
    if getattr(sys, "frozen", False):
        return None, None
    import utils

    root = str(utils.get_bundled_assets_root())
    commit = _run_text(["git", "-C", root, "rev-parse", "HEAD"], 2.0)
    if not commit:
        return None, None
    status = _run_text(["git", "-C", root, "status", "--porcelain"], 2.0)
    return commit, None if status is None else bool(status)


def environment() -> dict[str, Any]:
    """Facts a comparison must hold constant; every probe degrades to ``None``.

    Commit and dirty state say *what* ran; platform, chip, core count, memory,
    Python, ffmpeg and the package versions say *where*; ``settings`` are the
    tuning knobs from agents/PERFORMANCE.md read at call time, so a run with
    ``CLIP_PARALLEL_WORKERS`` overridden is visibly not comparable to one
    without.
    """
    from importlib import metadata

    import hardware
    import utils

    commit, dirty = _git_state()
    packages: dict[str, str | None] = {}
    for name in _ENV_PACKAGES:
        try:
            packages[name] = metadata.version(name)
        except metadata.PackageNotFoundError:
            packages[name] = None
    return {
        "version": utils.get_version(),
        "commit": commit,
        "dirty": dirty,
        "python": platform.python_version(),
        "platform": platform.platform(),
        **hardware.profile(),
        "ffmpeg": _run_text(["ffmpeg", "-version"], 10.0),
        "packages": packages,
        "settings": {name: getattr(config, name, None) for name in _ENV_SETTINGS},
    }


def export(*, reset: bool = False) -> dict[str, Any]:
    """The full report as data: labels, startup marks, mode, window, environment.

    ``mode`` is ``"deep"`` when a cProfile is attached (``--profile-deep``);
    the benches refuse to compare such a run, since its totals carry the
    profiler's own overhead. ``window_seconds`` runs from enable or the last
    reset to now.
    """
    with _LOCK:
        t0 = _WINDOW_T0
    now = time.perf_counter()
    return {
        "mode": "deep" if config.PROFILE_DEEP else "profile",
        "deep_label": config.PROFILE_DEEP,
        "window_seconds": (now - t0) if t0 is not None else 0.0,
        "labels": snapshot(reset=reset),
        "startup": startup_snapshot(),
        "peak_rss_mb": peak_rss_mb(),
        "dropped_labels": dropped_labels(),
        "active_spans": active_spans(),
        "env": environment(),
    }


def write_export(path: Path) -> Path | None:
    """Write :func:`export` to *path* as JSON; ``None`` when the write fails."""
    import utils

    return utils.write_json_atomic(path, export(), "profile output")


def _deep_report() -> None:
    """Print one pstats block per deep-profiled label; silent when none ran."""
    with _LOCK:
        groups: dict[str, list[cProfile.Profile]] = {}
        for (label, _tid), prof in _DEEP.items():
            groups.setdefault(label, []).append(prof)
    for label, profs in sorted(groups.items()):
        out = io.StringIO()
        # One active cProfile per interpreter; losing threads raise here, the winner must survive.
        stats = None
        for prof in profs:
            try:
                if stats is None:
                    stats = pstats.Stats(prof, stream=out)
                else:
                    stats.add(prof)
            except (TypeError, ValueError):
                continue
        if stats is None or not getattr(stats, "total_calls", 0):
            continue
        stats.strip_dirs().sort_stats("tottime").print_stats(_DEEP_TOP)
        print(f"profile-deep | {label}")  # bare print: see module docstring
        for line in out.getvalue().splitlines():
            if line.strip():
                print(f"  {line}")


def report() -> None:
    """Print one ``profile | `` line per label plus peak RSS; silent when empty."""
    report_startup()
    snap = snapshot()
    for label, entry in snap.items():
        secs, n, mx = entry["seconds"], int(entry["count"]), entry["max"]
        line = f"profile | {label:<32} {secs:8.3f}s  n={n}"
        if n and secs:
            line += f"  avg={secs / n * 1000:.1f}ms"
        if mx:
            line += f"  max={mx * 1000:.1f}ms"
        nb = int(entry.get("bytes", 0))
        if nb:
            line += f"  bytes={format_bytes(nb)}"
        first = entry.get("first", 0.0)
        # Name the first observation only when it doubles the warm average; sub-5ms is noise.
        if n > 1 and first >= 0.005 and first >= 2 * (secs - first) / (n - 1):
            line += f"  first={first * 1000:.1f}ms"
        print(line)  # bare print: Rich would wrap piped output (see module docstring)
    # Gate on output, not config.PROFILING: report() stays silent when nothing was recorded.
    if snap:
        peak = peak_rss_mb()
        if peak is not None:
            print(f"profile | {'peak_rss':<32} {peak:8.1f}MB")
        dropped = dropped_labels()
        if dropped:
            print(f"profile | {'labels_dropped':<32} {0.0:8.3f}s  n={dropped}")
    _deep_report()
    if config.PROFILE_OUTPUT:
        written = write_export(Path(config.PROFILE_OUTPUT))
        if written is not None:
            print(f"profile | written {written}")


def scan_summary(
    name: str,
    parts: list[tuple[str, float, int]],
    *,
    kind: str = "scan",
    extra: str = "",
) -> None:
    """Print a single per-run line so concurrent runs keep attribution.

    The totals table aggregates across every scan/transcription in the process;
    this is the per-input view. ``kind`` names the family (``scan``, ``whisper``)
    and ``extra`` carries derived figures that must not become labels — a
    realtime factor or an audio duration in the ``seconds`` column would sort to
    the top of the report in the slot that means "wall time this consumed".
    """
    if not config.PROFILING:
        return
    joined = "  ".join(f"{label}={seconds:.3f}s/n={n}" for label, seconds, n in parts)
    if extra:
        joined += "  " + extra
    print(f"profile | {kind} {name}: {joined}")  # bare print: see module docstring
