"""Shared Flask scaffolding for the server blueprints.

Every blueprint returns the same JSON envelope — ``{"ok": True, ...}`` on success,
``{"ok": False, "error": msg}`` plus an HTTP status on failure — and repeats the
same numeric-arg parse-and-validate block dozens of times. Collapsed here:

- :func:`ok` / :func:`err` build either envelope in one call.
- :class:`ApiError` + :func:`json_endpoint` let a handler ``raise`` a uniform
  4xx instead of threading an ``err(...)`` tuple back through every guard.
- :func:`parse_number_arg` parses + bound-checks one numeric value;
  :func:`opt_number` is its lenient fall-back-don't-fail sibling.
- :func:`find_by_id` / :func:`remove_by_id` are the manifest-collection CRUD
  lookups (stashes, blueprints, cuts, annotations).
- :func:`make_debounced_persist` builds the manifest-write debounce.
- :func:`make_participant_cache` builds the mtime-guarded participant cache
  (Transcripts + Screenspace).
- :func:`make_sse_channel` builds one SSE pub/sub channel (bounded per-client
  queue + coalesce-on-overflow + keepalive + cleanup).
- :class:`MediaCache` + :func:`parse_clip_window` + :func:`clip_media_response`
  + :func:`mtime_or_zero` back the hover-scrubber media routes (sprite sheets /
  audio snippets).

Deliberately tiny and Flask-only (no ``config``/``utils`` imports) so it stays
import-clean: ``utils`` is Flask-free on purpose and imported by non-server
modules, so these helpers must not live there.
"""

from __future__ import annotations

import math
import queue
import threading
from collections import OrderedDict
from collections.abc import Callable
from functools import wraps
from pathlib import Path
from typing import Any

from flask import Response, jsonify, request

import config
import profiling


def ok(**fields: Any):
    """Success envelope: ``jsonify({"ok": True, **fields})``."""
    return jsonify({"ok": True, **fields})


def err(message: str, code: int = 400):
    """Error envelope: ``(jsonify({"ok": False, "error": message}), code)``."""
    return jsonify({"ok": False, "error": message}), code


def err_no_video(participant: str, code: int = 404):
    """The shared "No video for participant <id>" error envelope."""
    return err(f"No video for participant {participant}", code)


class ApiError(Exception):
    """Raised inside a :func:`json_endpoint` handler to short-circuit to ``err``.

    Carries the user-facing message and the HTTP status to emit. Handlers raise
    it (directly or via :func:`parse_number_arg`) instead of returning an
    ``err(...)`` tuple, so deeply-nested validation guards stay one-liners.
    """

    def __init__(self, message: str, code: int = 400):
        super().__init__(message)
        self.message = message
        self.code = code


def json_endpoint(fn: Callable[..., Any]) -> Callable[..., Any]:
    """Wrap a route so a raised :class:`ApiError` becomes an ``err`` response.

    Catches **only** ``ApiError`` — never bare ``Exception`` — so it never
    swallows real 500s or interferes with routes that keep their own
    try/except / resource cleanup. Pairs with :func:`parse_number_arg`.
    """

    @wraps(fn)
    def wrapper(*args: Any, **kwargs: Any):
        try:
            return fn(*args, **kwargs)
        except ApiError as exc:
            return err(exc.message, exc.code)

    return wrapper


def parse_number_arg(
    raw: Any,
    name: str,
    *,
    int_only: bool = False,
    min_: float | None = None,
    max_: float | None = None,
    finite: bool = False,
) -> Any:
    """Parse + bound-check a single numeric value, raising ``ApiError`` on failure.

    The caller passes the already-fetched raw value (``request.args.get(...)``,
    a JSON-body ``dict.get(...)``, or a route param). Returns an ``int`` when
    ``int_only`` else a ``float``. ``int_only`` parses via ``int(float(raw))`` so
    ``"3.0"`` and ``3`` both work. Bounds are inclusive. ``finite`` rejects
    ``inf``/``nan``. Every failure raises :class:`ApiError` (HTTP 400) with a
    uniform message, caught by :func:`json_endpoint`.
    """
    try:
        value: float = float(raw)
    except (TypeError, ValueError):
        raise ApiError(f"{name} must be a number")
    if (finite or int_only) and not math.isfinite(value):
        raise ApiError(f"{name} must be a finite number")
    if min_ is not None and value < min_:
        raise ApiError(f"{name} must be >= {min_}")
    if max_ is not None and value > max_:
        raise ApiError(f"{name} must be <= {max_}")
    return int(value) if int_only else value


def opt_number(args: Any, name: str, default: float | None = None) -> float | None:
    """Lenient optional float from a request-args mapping.

    Missing or unparseable returns *default* — never raises. For preview-style
    override knobs where a malformed value silently falls back; the strict,
    raising sibling is :func:`parse_number_arg`.
    """
    raw = args.get(name)
    if raw is None:
        return default
    try:
        return float(raw)
    except (TypeError, ValueError):
        return default


def find_by_id(items: Any, id_: Any) -> dict[str, Any] | None:
    """First item whose ``"id"`` equals *id_*, else None."""
    return next((it for it in items if it.get("id") == id_), None)


def remove_by_id(items: list[dict[str, Any]], id_: Any) -> dict[str, Any] | None:
    """Pop and return the first item whose ``"id"`` equals *id_*; None when absent."""
    for i, it in enumerate(items):
        if it.get("id") == id_:
            return items.pop(i)
    return None


class MediaCache:
    """Bounded-LRU byte cache with single-flight compute.

    Fast path takes the main lock only for the dict get/reorder. On a miss, a
    per-key lock serializes concurrent identical misses so the expensive
    producer (ffmpeg) runs once, not once per waiting request — the others wake
    to the freshly cached bytes. The producer runs holding no main lock.
    """

    def __init__(self, max_entries: int) -> None:
        self._store: OrderedDict[tuple, bytes] = OrderedDict()
        self._max = max_entries
        self._lock = threading.Lock()
        self._inflight: dict[tuple, threading.Lock] = {}

    def get_or_compute(
        self, key: tuple, compute: Callable[[], bytes | None]
    ) -> bytes | None:
        # Fast path: cache hit.
        with self._lock:
            cached = self._store.get(key)
            if cached is not None:
                self._store.move_to_end(key)
                if config.PROFILING:
                    profiling.count("media_cache.hit")
                return cached
            keylock = self._inflight.get(key)
            if keylock is None:
                keylock = threading.Lock()
                self._inflight[key] = keylock

        with keylock:
            # Re-check: another thread may have produced it while we waited.
            with self._lock:
                cached = self._store.get(key)
                if cached is not None:
                    self._store.move_to_end(key)
                    if config.PROFILING:
                        profiling.count("media_cache.hit")
                    return cached

            if config.PROFILING:
                profiling.count("media_cache.miss")
            with profiling.span("media_cache.compute"):
                value = compute()  # expensive; no main lock held

            with self._lock:
                if value is not None:
                    self._store[key] = value
                    while len(self._store) > self._max:
                        self._store.popitem(last=False)
                # Drop our in-flight marker so the dict can't grow unbounded.
                if self._inflight.get(key) is keylock:
                    del self._inflight[key]
            return value

    def clear(self) -> None:
        with self._lock:
            self._store.clear()
            self._inflight.clear()


def mtime_or_zero(path: str | Path) -> float:
    """A file's ``st_mtime`` for cache keys, or 0.0 when it can't be stat'd."""
    try:
        return Path(path).stat().st_mtime
    except OSError:
        return 0.0


def parse_clip_window() -> tuple[float, float] | None:
    """Parse + validate the ``?start=&end=`` seconds shared by the scrubber
    media routes. Returns ``(start_seconds, duration_seconds)`` or ``None`` when
    the params are missing/non-numeric or the range is empty."""
    try:
        start_sec = max(0.0, float(request.args.get("start", "")))
        end_sec = float(request.args.get("end", ""))
    except (ValueError, TypeError):
        return None
    duration = end_sec - start_sec
    if duration <= 0:
        return None
    return start_sec, duration


def clip_media_response(
    *,
    cache: MediaCache,
    resolve: Callable[[float, float], tuple[str, float, float] | None],
    produce: Callable[[str, float, float], bytes | None],
    mimetype: str,
    kind_label: str,
    key_extras: tuple[Any, ...] = (),
    invalid_message: str = "Invalid clip range",
):
    """Shared guts of the hover-scrubber sprite / audio routes.

    Parses the ``?start=&end=`` window, maps it through *resolve* →
    ``(path, local_start, duration)`` (None → 404), then serves *produce*'s
    bytes from *cache* keyed on
    ``(path, start, duration, *key_extras, mtime)`` — the mtime so a replaced
    source file invalidates stale media. *kind_label* names the 404 when
    extraction fails (``"Sprite extraction failed"``).
    """
    window = parse_clip_window()
    if window is None:
        return err(invalid_message)
    start_sec, duration = window

    resolved = resolve(start_sec, duration)
    if resolved is None:
        return err("Source video not found", 404)
    path, local_start, duration = resolved

    key = (
        path,
        round(local_start, 3),
        round(duration, 3),
        *key_extras,
        mtime_or_zero(path),
    )
    media_bytes = cache.get_or_compute(
        key, lambda: produce(path, local_start, duration)
    )
    if media_bytes is None:
        return err(f"{kind_label} extraction failed", 404)
    return Response(
        media_bytes,
        mimetype=mimetype,
        headers={"Cache-Control": "public, max-age=86400"},
    )


def make_debounced_persist(
    persist: Callable[[], None],
    manifest_lock: threading.Lock,
    *,
    debounce_seconds: float = 2.0,
) -> tuple[Callable[[], None], Callable[[], None], Callable[[], bool]]:
    """Build a manifest-write debounce; returns ``(schedule, flush, cancel)``.

    Rapid UI mutations coalesce into one disk write after a ``debounce_seconds``
    quiet period instead of blocking each request on a full manifest save.
    ``schedule_persist()`` marks dirty and (re)arms the timer;
    ``flush_pending_persist()`` cancels the timer and persists immediately if
    dirty (register it with ``atexit``); ``cancel_pending_persist_timer()``
    cancels and returns the prior dirty state, for synchronous save paths that
    supersede a pending debounced write.

    ``persist`` runs with ``manifest_lock`` held. Callers must pass a lambda
    that looks up their ``_do_persist`` module global at call time (not the
    function object) so tests monkeypatching ``<module>._do_persist`` are seen.
    """
    timer: threading.Timer | None = None
    timer_lock = threading.Lock()
    dirty = False

    def cancel_pending_persist_timer() -> bool:
        """Cancel the debounce timer and clear the dirty flag. Returns prior dirty state."""
        nonlocal timer, dirty
        with timer_lock:
            pending = timer
            timer = None
            was_dirty = dirty
            dirty = False
        if pending is not None:
            pending.cancel()
        return was_dirty

    def _on_persist_timer() -> None:
        nonlocal timer, dirty
        with timer_lock:
            timer = None
            if not dirty:
                return
            dirty = False
        with manifest_lock:
            persist()

    def schedule_persist() -> None:
        """Mark the manifest dirty and (re)arm the debounce timer."""
        nonlocal timer, dirty
        with timer_lock:
            dirty = True
            if timer is not None:
                timer.cancel()
            timer = threading.Timer(debounce_seconds, _on_persist_timer)
            timer.daemon = True
            timer.start()

    def flush_pending_persist() -> None:
        """Cancel any pending debounced write and persist immediately if dirty."""
        if cancel_pending_persist_timer():
            with manifest_lock:
                persist()

    return schedule_persist, flush_pending_persist, cancel_pending_persist_timer


def make_participant_cache(
    module: Any,
    *,
    input_dir_getter: Callable[[], Any],
    resolve: Callable[[Any], list[dict[str, Any]]],
) -> tuple[Callable[[], None], Callable[[str], dict[str, Any] | None]]:
    """Build the mtime-guarded participant cache shared by the Transcripts and
    Screenspace blueprints; returns ``(refresh, find)``.

    Operates on ``module._participants`` / ``module._participant_source`` under
    ``module._participants_lock`` — module attributes, not closure state — so
    the blueprints' init/set-source functions, the remux ``sheet_context``
    getters, and tests that monkeypatch those globals all keep working (the
    same late-binding contract as :func:`make_debounced_persist`).

    ``refresh()`` rebuilds ``module._participants`` when the input directory
    changed since the last build. Keyed on the dir's ``st_mtime_ns`` (which
    advances on add/remove/rename), mirroring
    ``utils.discover_participant_videos``' own memo — the steady-state cost is
    one ``stat()``. This is what lets a video dropped into ``-i`` mid-session
    show up without a server restart. No-op while ``_participant_source`` is
    None (blueprint not configured yet). The rebuild rebinds ``_participants``
    (atomic under the GIL), so a concurrent reader sees either the old list or
    the new one, never a torn one.

    ``find(pid)`` refreshes, then returns the cached record or None. The
    ``input_dir_getter`` / ``resolve`` callables keep ``utils``/``files``
    imports out of this module.
    """

    def refresh() -> None:
        source = module._participant_source
        if source is None:
            return
        input_dir = str(Path(input_dir_getter()))
        try:
            mtime: int | None = Path(input_dir).stat().st_mtime_ns
        except OSError:
            mtime = None
        if source["dir"] == input_dir and source["mtime"] == mtime:
            return
        with module._participants_lock:
            # A racing request may have rebuilt, or a sheet swap may have
            # replaced the source entirely, while we waited on the lock.
            if module._participant_source is not source:
                return
            if source["dir"] == input_dir and source["mtime"] == mtime:
                return
            module._participants = resolve(source["sheet_context"])
            source["dir"] = input_dir
            source["mtime"] = mtime

    def find(participant_id: str) -> dict[str, Any] | None:
        refresh()
        return find_by_id(module._participants, participant_id)

    return refresh, find


def make_sse_channel(
    *, maxsize: int = 64, keepalive_seconds: float = 15.0
) -> tuple[
    Callable[..., None],
    Callable[..., Response],
    list[tuple[Any, queue.Queue[str]]],
]:
    """Build one SSE pub/sub channel; returns ``(notify, stream, clients)``.

    Collapses the bounded-queue + coalesce-on-overflow + keepalive + cleanup
    boilerplate otherwise duplicated across the run / batch / task SSE endpoints.

    - ``notify(key=None, marker="update")`` wakes every client registered with a
      matching ``key`` (``key=None`` = broadcast). On a full queue it coalesces:
      drop one stale entry, re-push ``marker`` (dropped silently if still full).
      The queued token is never inspected by the streamer — it only triggers a
      full payload rebuild — so ``marker``'s value is cosmetic.
    - ``stream(payload, key=None) -> Response`` registers a client keyed by
      ``key``, returns a ``text/event-stream`` Response that emits ``payload()``
      immediately, re-emits it on each wake (draining the backlog first), sends a
      keepalive comment after a quiet ``keepalive_seconds`` timeout, and
      deregisters in ``finally``.
    - ``clients`` is the live registry list (``(key, queue)`` tuples), exposed so
      tests can inject/clear entries; mutate it in place, never rebind.
    """
    clients: list[tuple[Any, queue.Queue[str]]] = []
    lock = threading.Lock()

    def notify(key: Any = None, marker: str = "update") -> None:
        with lock:
            for ckey, cq in clients:
                if ckey != key:
                    continue
                try:
                    cq.put_nowait(marker)
                except queue.Full:
                    try:
                        cq.get_nowait()
                    except queue.Empty:
                        pass
                    try:
                        cq.put_nowait(marker)
                    except queue.Full:
                        pass

    def stream(payload: Callable[[], str], key: Any = None) -> Response:
        client_q: queue.Queue[str] = queue.Queue(maxsize=maxsize)
        entry = (key, client_q)
        with lock:
            clients.append(entry)

        def generate():  # type: ignore[no-untyped-def]
            try:
                yield payload()
                while True:
                    try:
                        client_q.get(timeout=keepalive_seconds)
                        # Drain any backlog (coalesce rapid updates) before emitting.
                        while not client_q.empty():
                            try:
                                client_q.get_nowait()
                            except queue.Empty:
                                break
                        yield payload()
                    except queue.Empty:
                        yield ": keepalive\n\n"
            except GeneratorExit:
                pass
            finally:
                with lock:
                    try:
                        clients.remove(entry)
                    except ValueError:
                        pass

        return Response(
            generate(),
            mimetype="text/event-stream",
            headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
        )

    return notify, stream, clients
