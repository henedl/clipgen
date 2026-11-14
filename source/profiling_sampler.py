"""Resource probes and the stack sampler behind ``--profile-sample``.

The stopwatch in ``profiling.py`` answers "which labelled span took the wall
time". It cannot say three things, and this module exists for them:

* **CPU.** A span that waits on ffmpeg and one that burns a core read the
  same. :func:`process_cpu` and :func:`thread_cpu` put CPU seconds beside the
  wall seconds, for the process, its reaped children and each Python thread.
* **Memory.** ``ru_maxrss`` is a peak with no timeline, and on macOS it
  under-reports once the compressor evicts pages. :func:`os_usage` reads the
  physical footprint (what Activity Monitor shows) and the sampler keeps a
  bounded timeline plus what was running at the peak.
* **Unlabelled time.** ``--profile-deep`` needs a label to attach to. The
  sampler needs none: it reads every Python thread's stack ``hz`` times a
  second and weights each stack by the CPU that thread burned since the last
  tick, so idle threads cost nothing in the CPU view however often they are
  seen.

No project imports: ``profiling`` imports this module and hands it callbacks,
so it stays below ``profiling`` in the import layering. Everything here
degrades to "not reported" rather than raising; a probe must never break a run.

Sampling is in-process, so it shares the GIL with the work it watches. A tick
waits for the interpreter like any other thread, which biases *where* a sample
lands toward points that release it. The CPU weight is the OS's own number and
is not biased. Native pool threads (OpenCV, onnxruntime, CTranslate2) have no
Python frame: their CPU is the process total minus the Python threads.
"""

import os
import re
import sys
import threading
import time
from collections.abc import Callable
from typing import Any

_MB = 1024 * 1024
_LOCK = threading.Lock()
_THREAD: threading.Thread | None = None
_STOP = threading.Event()
_HZ = 0
_T0 = 0.0
_TICKS = 0
_OVERRUNS = 0
_FAULTS = 0  # ticks that raised; reported, never silent
_TICK_COST = 0.0

# (thread label, code objects root to leaf) -> [wall samples, cpu seconds].
_STACKS: dict[tuple[Any, ...], list[float]] = {}
_MAX_STACKS = 20000
_MAX_DEPTH = 64
_THREADS: dict[str, list[float]] = {}  # thread label -> [cpu seconds, samples]
# thread ident -> (cumulative cpu, stack key) as of its last tick.
_LAST: dict[int, tuple[float, tuple[Any, ...]]] = {}

# Memory timeline: [seconds since start, footprint MB], halved when full.
_MEM: list[tuple[float, float]] = []
_MEM_GAP = 0.2
_MAX_MEM = 512
_MEM_PEAK: dict[str, Any] = {}
_ON_MEMORY: Callable[[float], None] | None = None
_CONTEXT: Callable[[], list[str]] | None = None

_OWN_DIR = os.path.dirname(os.path.abspath(__file__))
_DIGITS = re.compile(r"\d+")
_LIBPROC: Any = None  # False once loading failed


def process_cpu() -> tuple[float, float]:
    """``(self, children)`` CPU seconds, user + system.

    Children are only the ones already reaped, which is every finished ffmpeg
    or ffprobe. It is the one view of what the encodes cost: their time never
    appears in this process's own counters.
    """
    try:
        import resource
    except ImportError:
        # Windows: no child accounting, and os.times ticks at 64 Hz.
        return time.process_time(), 0.0
    child = resource.getrusage(resource.RUSAGE_CHILDREN)
    return time.process_time(), child.ru_utime + child.ru_stime


def child_peak_rss_mb() -> float | None:
    """Largest resident size of any reaped child (one ffmpeg), in MB."""
    try:
        import resource
    except ImportError:
        return None
    raw = resource.getrusage(resource.RUSAGE_CHILDREN).ru_maxrss
    return raw * (1 if sys.platform == "darwin" else 1024) / _MB


def _libproc() -> Any:
    """macOS libproc, called with the GIL held so a probe never yields it."""
    global _LIBPROC
    if _LIBPROC is None:
        try:
            import ctypes
            import ctypes.util

            lib = ctypes.PyDLL(ctypes.util.find_library("proc"))
            lib.proc_pidinfo.argtypes = [
                ctypes.c_int,
                ctypes.c_int,
                ctypes.c_uint64,
                ctypes.c_void_p,
                ctypes.c_int,
            ]
            _LIBPROC = lib
        except (OSError, AttributeError, TypeError):
            _LIBPROC = False
    return _LIBPROC


# rusage_info_v4 after its 16-byte uuid: field index -> name (libproc.h).
_RUSAGE_V4_WORDS = 35
_RUSAGE_FIELDS = {
    2: "idle_wakeups",
    3: "wakeups",
    6: "rss",
    7: "footprint",
    16: "disk_read_bytes",
    17: "disk_write_bytes",
    28: "peak_footprint",
}


def _darwin_usage() -> dict[str, float]:
    import ctypes

    lib = _libproc()
    if not lib:
        return {}
    buf = (ctypes.c_uint64 * (2 + _RUSAGE_V4_WORDS))()
    if lib.proc_pid_rusage(os.getpid(), 4, ctypes.byref(buf)) != 0:
        return {}
    raw = {name: buf[2 + index] for index, name in _RUSAGE_FIELDS.items()}
    return {
        "footprint_mb": raw["footprint"] / _MB,
        "peak_footprint_mb": raw["peak_footprint"] / _MB,
        "rss_mb": raw["rss"] / _MB,
        "wakeups": raw["wakeups"],
        "idle_wakeups": raw["idle_wakeups"],
        "disk_read_bytes": raw["disk_read_bytes"],
        "disk_write_bytes": raw["disk_write_bytes"],
    }


def _linux_usage() -> dict[str, float]:
    out: dict[str, float] = {}
    with open("/proc/self/status", encoding="ascii") as handle:
        for line in handle:
            if line.startswith("VmRSS:"):
                out["footprint_mb"] = out["rss_mb"] = int(line.split()[1]) / 1024
            elif line.startswith("VmHWM:"):
                out["peak_footprint_mb"] = int(line.split()[1]) / 1024
    with open("/proc/self/io", encoding="ascii") as handle:
        for line in handle:
            key, _, value = line.partition(":")
            if key == "read_bytes":
                out["disk_read_bytes"] = int(value)
            elif key == "write_bytes":
                out["disk_write_bytes"] = int(value)
    return out


def os_usage() -> dict[str, float]:
    """What the OS says this process holds and has done; ``{}`` where it will not.

    ``footprint_mb`` is the number to read for memory. On macOS it is the
    physical footprint, which counts compressed pages; ``rss_mb`` drops them,
    so a process under memory pressure looks smaller the worse off it is.
    ``peak_footprint_mb`` is the lifetime high-water mark of the same number.
    ``wakeups`` and ``idle_wakeups`` (macOS) count how often the process woke
    a core, the cost of a poller that never pauses. Disk counters are bytes
    that reached the device. (``ri_billed_energy`` stays 0 for an ordinary
    process, measured, so no energy figure is offered.)
    Windows reports nothing: no ``resource``, no ``/proc``, and no psutil.
    """
    try:
        if sys.platform == "darwin":
            return _darwin_usage()
        if sys.platform.startswith("linux"):
            return _linux_usage()
    except (OSError, ValueError, IndexError):
        return {}
    return {}


def _darwin_thread_cpu(native_id: int) -> float | None:
    import ctypes

    lib = _libproc()
    if not lib:
        return None
    # proc_threadinfo opens with user and system time in nanoseconds.
    buf = (ctypes.c_uint64 * 14)()
    size = ctypes.sizeof(buf)
    # 15 is PROC_PIDTHREADID64INFO: look a thread up by its 64-bit id.
    if lib.proc_pidinfo(os.getpid(), 15, native_id, ctypes.byref(buf), size) < 16:
        return None
    return (buf[0] + buf[1]) / 1e9


def _linux_thread_cpu(native_id: int) -> float:
    """On-CPU seconds from ``schedstat``, whose first field is nanoseconds.

    A file read on purpose: ``pthread_getcpuclockid`` on a thread that exited
    since it was listed is undefined behaviour, and a probe must not crash.
    """
    with open(f"/proc/self/task/{native_id}/schedstat", encoding="ascii") as handle:
        return int(handle.read().split()[0]) / 1e9


def thread_cpu(thread: threading.Thread) -> float | None:
    """Cumulative CPU seconds of a live *thread*; ``None`` where unsupported."""
    native = thread.native_id
    if native is None:
        return None
    try:
        if sys.platform == "darwin":
            return _darwin_thread_cpu(native)
        if sys.platform.startswith("linux"):
            return _linux_thread_cpu(native)
    except (OSError, ValueError, IndexError):
        return None
    return None


def _thread_label(name: str) -> str:
    """Thread name with its counters dropped, so a pool's threads share a row."""
    return _DIGITS.sub("N", name)


def running() -> bool:
    return _THREAD is not None


def start(
    hz: int,
    *,
    on_memory: Callable[[float], None] | None = None,
    context: Callable[[], list[str]] | None = None,
) -> bool:
    """Start the sampler thread at *hz*; ``False`` when it is already running.

    *on_memory* receives every footprint reading (profiling folds it into the
    running operations). *context* names what is in flight; it is called only
    when a reading sets a new peak.
    """
    global _THREAD, _HZ, _T0, _ON_MEMORY, _CONTEXT
    # Load the probes now; importing ctypes inside a tick stalls it 100 ms.
    thread_cpu(threading.current_thread())
    os_usage()
    with _LOCK:
        if _THREAD is not None or hz <= 0:
            return False
        _LAST.clear()
        _HZ = hz
        _T0 = time.perf_counter()
        _ON_MEMORY = on_memory
        _CONTEXT = context
        _STOP.clear()
        _THREAD = threading.Thread(
            target=_run, args=(1.0 / hz,), name="clipgen-profile-sampler", daemon=True
        )
        thread = _THREAD
    thread.start()
    return True


def stop() -> None:
    """Stop the sampler and wait for its last tick; the data stays readable."""
    global _THREAD
    with _LOCK:
        thread = _THREAD
        _THREAD = None
    if thread is not None:
        _STOP.set()
        thread.join(timeout=2.0)


def _run(interval: float) -> None:
    global _OVERRUNS, _TICK_COST, _FAULTS
    me = threading.get_ident()
    next_at = time.perf_counter() + interval
    last_mem = 0.0
    cost0 = time.thread_time()
    while True:
        # sleep, not Event.wait: half the CPU per wake on an idle machine.
        time.sleep(max(0.0, next_at - time.perf_counter()))
        if _STOP.is_set():
            return
        began = time.perf_counter()
        try:
            _tick(me)
            if began - last_mem >= _MEM_GAP:
                last_mem = began
                _sample_memory(began)
        except Exception:
            faulted = True
        else:
            faulted = False
        now = time.perf_counter()
        next_at += interval
        with _LOCK:
            _FAULTS += faulted
            cost = time.thread_time()
            _TICK_COST += cost - cost0
            cost0 = cost
            if next_at < now:
                missed = int((now - next_at) / interval) + 1
                _OVERRUNS += missed
                next_at += missed * interval


def _tick(me: int) -> None:
    """Read every Python thread's stack and CPU, once.

    A thread that burned no CPU since the last tick cannot have moved, so its
    previous stack is reused without walking it. That keeps an idle server's
    dozen parked threads at one syscall each.
    """
    global _TICKS
    frames = sys._current_frames()
    threads = {thread.ident: thread for thread in threading.enumerate()}
    seen: list[tuple[tuple[Any, ...], float]] = []
    for ident, frame in frames.items():
        if ident == me:
            continue
        thread = threads.get(ident)
        cpu = thread_cpu(thread) if thread is not None else None
        last = _LAST.get(ident)
        if last is not None and cpu is not None and cpu == last[0]:
            seen.append((last[1], 0.0))
            continue
        codes = []
        cursor: Any = frame
        while cursor is not None and len(codes) < _MAX_DEPTH:
            codes.append(cursor.f_code)
            cursor = cursor.f_back
        codes.reverse()
        label = _thread_label(thread.name) if thread is not None else "(foreign)"
        key = (label, *codes)
        delta = cpu - last[0] if cpu is not None and last is not None else 0.0
        if cpu is not None:
            _LAST[ident] = (cpu, key)
        seen.append((key, max(delta, 0.0)))
    del frames
    with _LOCK:
        _TICKS += 1
        for key, delta in seen:
            entry = _STACKS.get(key)
            if entry is None:
                if len(_STACKS) >= _MAX_STACKS:
                    key = (key[0], "(stacks past the cap)")
                    entry = _STACKS.setdefault(key, [0.0, 0.0])
                else:
                    entry = _STACKS[key] = [0.0, 0.0]
            entry[0] += 1
            entry[1] += delta
            totals = _THREADS.setdefault(key[0], [0.0, 0.0])
            totals[0] += delta
            totals[1] += 1
        if len(_LAST) > 4 * len(threads) + 64:
            for ident in [i for i in _LAST if i not in threads]:
                del _LAST[ident]


def _sample_memory(now: float) -> None:
    """One footprint reading: timeline point, peak with context, callback."""
    global _MEM_GAP
    footprint = os_usage().get("footprint_mb")
    if footprint is None:
        return
    is_peak = footprint > _MEM_PEAK.get("mb", 0.0)
    active = _CONTEXT() if is_peak and _CONTEXT is not None else []
    with _LOCK:
        _MEM.append((now - _T0, footprint))
        if len(_MEM) > _MAX_MEM:
            del _MEM[1::2]
            _MEM_GAP *= 2
        if is_peak:
            _MEM_PEAK.update(mb=footprint, at_s=now - _T0, active=active)
    if _ON_MEMORY is not None:
        _ON_MEMORY(footprint)


def frame_name(code: Any) -> str:
    """``module:qualname`` for clipgen's own code, ``package/module:qualname`` otherwise."""
    if isinstance(code, str):
        return code
    path = code.co_filename
    name = getattr(code, "co_qualname", code.co_name)
    if path.startswith("<"):
        return f"{path}:{name}"
    stem = os.path.splitext(os.path.basename(path))[0]
    if _is_own(code):
        return f"{stem}:{name}"
    parent = os.path.basename(os.path.dirname(path))
    return f"{parent}/{stem}:{name}" if parent else f"{stem}:{name}"


def _is_own(code: Any) -> bool:
    """Whether *code* is clipgen source (frozen builds keep bare file names)."""
    if isinstance(code, str):
        return False
    path = code.co_filename
    if path.startswith("<"):
        return False
    directory = os.path.dirname(path)
    return directory == _OWN_DIR or not directory


def _ranked(totals: dict[Any, list[float]], top: int) -> list[dict[str, Any]]:
    """Top rows by CPU, then by samples where CPU is unavailable."""
    rows = sorted(totals.items(), key=lambda item: (-item[1][1], -item[1][0]))
    return [
        {"frame": frame_name(code), "cpu_s": cpu, "samples": int(samples)}
        for code, (samples, cpu) in rows[:top]
        if cpu > 0 or samples > 0
    ]


def snapshot(*, reset: bool = False, top: int = 15) -> dict[str, Any] | None:
    """The sampler's findings as data; ``None`` when it never ran.

    ``top_self`` ranks the innermost frame of each sample: where the CPU was
    actually spent, usually library code. ``top_own`` ranks the innermost
    *clipgen* frame instead, which answers the more useful question: which of
    our functions is the CPU going through. ``threads`` is CPU per Python
    thread with pool counters folded together. ``cost_s`` is the sampler
    thread's own CPU, so a reader can see what the measurement cost.
    """
    global _TICKS, _OVERRUNS, _FAULTS, _TICK_COST, _T0, _MEM_GAP
    with _LOCK:
        if not _HZ:
            return None
        now = time.perf_counter()
        stacks = dict(_STACKS)
        out: dict[str, Any] = {
            "hz": _HZ,
            "seconds": now - _T0,
            "ticks": _TICKS,
            "overruns": _OVERRUNS,
            "faults": _FAULTS,
            "cost_s": _TICK_COST,
            "stacks": len(stacks),
            "threads": {
                label: {"cpu_s": cpu, "samples": int(samples)}
                for label, (cpu, samples) in sorted(
                    _THREADS.items(), key=lambda item: -item[1][0]
                )
            },
            "memory": {"timeline": [list(point) for point in _MEM], **_MEM_PEAK},
        }
        if reset:
            _STACKS.clear()
            _THREADS.clear()
            _MEM.clear()
            _MEM_PEAK.clear()
            _TICKS = _OVERRUNS = _FAULTS = 0
            _TICK_COST = 0.0
            _MEM_GAP = 0.2
            _T0 = now
    self_totals: dict[Any, list[float]] = {}
    own_totals: dict[Any, list[float]] = {}
    for key, (samples, cpu) in stacks.items():
        leaf = key[-1]
        own = next((code for code in reversed(key[1:]) if _is_own(code)), None)
        for totals, code in ((self_totals, leaf), (own_totals, own)):
            if code is None:
                continue
            entry = totals.setdefault(code, [0.0, 0.0])
            entry[0] += samples
            entry[1] += cpu
    out["cpu_s"] = sum(cpu for _samples, cpu in stacks.values())
    out["top_self"] = _ranked(self_totals, top)
    out["top_own"] = _ranked(own_totals, top)
    return out


def folded(kind: str = "cpu") -> str:
    """Collapsed stacks, one ``thread;root;...;leaf weight`` line each.

    The format flamegraph.pl and speedscope read. ``cpu`` weights are CPU
    microseconds, so the graph shows work; ``wall`` weights are sample counts,
    so it shows where threads sit, waiting included.
    """
    with _LOCK:
        stacks = list(_STACKS.items())
    lines = []
    for key, (samples, cpu) in stacks:
        weight = int(cpu * 1e6) if kind == "cpu" else int(samples)
        if weight > 0:
            lines.append(f"{';'.join(frame_name(part) for part in key)} {weight}")
    lines.sort()
    return "\n".join(lines) + ("\n" if lines else "")
