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
the exact consumer this output exists for. On a live server, ``GET
/api/profile`` returns ``snapshot()`` and ``?reset=1`` brackets a measurement
window. See agents/skills/profile/SKILL.md.
"""

import atexit
import functools
import sys
import threading
import time
from collections.abc import Callable, Generator, Iterable, Iterator
from contextlib import contextmanager
from typing import Any

import config

_LOCK = threading.Lock()
# label -> [total_seconds, count, max_seconds]. Labels are static strings (or
# Flask url_rule strings, ~200 of them); the cap is a safety net, not an LRU.
_TOTALS: dict[str, list[float]] = {}
_MAX_LABELS = 1024
_REPORT_REGISTERED = False


def enable() -> None:
    """Turn profiling on and register the end-of-process report once."""
    global _REPORT_REGISTERED
    config.PROFILING = True
    with _LOCK:
        if _REPORT_REGISTERED:
            return
        _REPORT_REGISTERED = True
    atexit.register(report)


def add(
    label: str, seconds: float = 0.0, n: int = 1, *, peak: float | None = None
) -> None:
    """Accumulate *seconds* and *n* occurrences under *label*.

    *peak* is the largest single occurrence in this contribution. A batched
    flush (``n > 1``) has no per-item max in *seconds* — that is a sum — so the
    max is left alone unless the caller tracked one in its own loop and passed
    it. Without that kwarg the tail would be permanently invisible on exactly
    the labels where it matters most (``scan.callback``, ``transcribe.decode``);
    one compare per iteration in an already-accumulating loop is not a cost.

    Percentiles are deliberately not offered: p95 needs retained samples, i.e.
    unbounded per-label memory, which this accumulator exists to refuse. Max is
    the only tail statistic that is O(1).
    """
    if not config.PROFILING:
        return
    if peak is None and n == 1:
        peak = seconds
    with _LOCK:
        entry = _TOTALS.get(label)
        if entry is None:
            if len(_TOTALS) >= _MAX_LABELS:
                return
            _TOTALS[label] = [seconds, float(n), peak or 0.0]
        else:
            entry[0] += seconds
            entry[1] += n
            if peak is not None and peak > entry[2]:
                entry[2] = peak


def count(label: str, n: int = 1) -> None:
    """Accumulate *n* occurrences under *label* with no time component."""
    add(label, 0.0, n)


@contextmanager
def span(label: str) -> Iterator[None]:
    """Time the enclosed block under *label*; passthrough when profiling is off."""
    if not config.PROFILING:
        yield
        return
    start = time.perf_counter()
    try:
        yield
    finally:
        add(label, time.perf_counter() - start)


def timed(label: str) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Decorator form of ``span`` for functions with many return paths."""

    def decorate(fn: Callable[..., Any]) -> Callable[..., Any]:
        @functools.wraps(fn)
        def wrapper(*args: Any, **kwargs: Any) -> Any:
            if not config.PROFILING:
                return fn(*args, **kwargs)
            start = time.perf_counter()
            try:
                return fn(*args, **kwargs)
            finally:
                add(label, time.perf_counter() - start)

        return wrapper

    return decorate


def stream_span(label: str, body: Iterable[str]) -> Generator[str, None, None]:
    """Yield through *body*, recording total generation time under *label*.

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

    The ``finally`` also covers a client disconnect mid-stream (``GeneratorExit``),
    so an abandoned generation still records what it spent.
    """
    if not config.PROFILING:
        yield from body
        return
    start = time.perf_counter()
    try:
        yield from body
    finally:
        add(label, time.perf_counter() - start)


def snapshot() -> dict[str, dict[str, float]]:
    """Return ``{label: {"seconds": s, "count": n, "max": m}}`` sorted by seconds desc.

    ``max`` is 0.0 for labels fed only by batched flushes that supplied no
    ``peak=`` — see :func:`add`. Sorting stays on seconds so the max never
    reorders the report.
    """
    with _LOCK:
        items = [
            (label, entry[0], int(entry[1]), entry[2])
            for label, entry in _TOTALS.items()
        ]
    items.sort(key=lambda item: (-item[1], item[0]))
    return {
        label: {"seconds": secs, "count": n, "max": mx} for label, secs, n, mx in items
    }


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
    weights, EasyOCR Readers, decoded frames) is all in-process.
    """
    try:
        import resource  # POSIX only
    except ImportError:
        return None
    raw = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    scale = 1 if sys.platform == "darwin" else 1024  # bytes vs kilobytes
    return raw * scale / (1024 * 1024)


def reset() -> None:
    """Clear all accumulated totals (brackets a measurement window)."""
    with _LOCK:
        _TOTALS.clear()


def report() -> None:
    """Print one ``profile | `` line per label plus peak RSS; silent when empty."""
    snap = snapshot()
    for label, entry in snap.items():
        secs, n, mx = entry["seconds"], int(entry["count"]), entry["max"]
        line = f"profile | {label:<32} {secs:8.3f}s  n={n}"
        if n and secs:
            line += f"  avg={secs / n * 1000:.1f}ms"
        if mx:
            line += f"  max={mx * 1000:.1f}ms"
        print(line)  # bare print: Rich would wrap piped output (see module docstring)
    # Gated on "did we print anything", not on config.PROFILING: report()'s
    # documented contract is silence when nothing was recorded, and it has never
    # read the flag itself.
    if snap:
        peak = peak_rss_mb()
        if peak is not None:
            print(f"profile | {'peak_rss':<32} {peak:8.1f}MB")


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
