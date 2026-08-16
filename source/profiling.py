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
import threading
import time
from collections.abc import Callable, Generator, Iterable, Iterator
from contextlib import contextmanager
from typing import Any

import config

_LOCK = threading.Lock()
# label -> [total_seconds, count]. Labels are static strings (or Flask url_rule
# strings, ~200 of them); the cap is a safety net, not an LRU.
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


def add(label: str, seconds: float = 0.0, n: int = 1) -> None:
    """Accumulate *seconds* and *n* occurrences under *label*."""
    if not config.PROFILING:
        return
    with _LOCK:
        entry = _TOTALS.get(label)
        if entry is None:
            if len(_TOTALS) >= _MAX_LABELS:
                return
            _TOTALS[label] = [seconds, float(n)]
        else:
            entry[0] += seconds
            entry[1] += n


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
    """Return ``{label: {"seconds": s, "count": n}}`` sorted by seconds desc."""
    with _LOCK:
        items = [(label, entry[0], int(entry[1])) for label, entry in _TOTALS.items()]
    items.sort(key=lambda item: (-item[1], item[0]))
    return {label: {"seconds": secs, "count": n} for label, secs, n in items}


def reset() -> None:
    """Clear all accumulated totals (brackets a measurement window)."""
    with _LOCK:
        _TOTALS.clear()


def report() -> None:
    """Print one ``profile | `` line per label; silent when nothing recorded."""
    for label, entry in snapshot().items():
        secs, n = entry["seconds"], int(entry["count"])
        line = f"profile | {label:<32} {secs:8.3f}s  n={n}"
        if n and secs:
            line += f"  avg={secs / n * 1000:.1f}ms"
        print(line)  # bare print: Rich would wrap piped output (see module docstring)


def scan_summary(name: str, parts: list[tuple[str, float, int]]) -> None:
    """Print a single per-scan line so concurrent scans keep attribution."""
    if not config.PROFILING:
        return
    joined = "  ".join(f"{label}={seconds:.3f}s/n={n}" for label, seconds, n in parts)
    print(f"profile | scan {name}: {joined}")  # bare print: see module docstring
