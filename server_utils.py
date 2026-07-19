"""Shared Flask helpers for the four server blueprints.

All four server modules (``server``, ``screenspace_server``,
``transcripts_server``, ``workflows_server``) return the same JSON envelope —
``{"ok": True, ...}`` on success and ``{"ok": False, "error": msg}`` with an
HTTP status on failure — and repeat the same numeric-arg parse-and-validate
block dozens of times. This module collapses that scaffolding:

- :func:`ok` / :func:`err` build the success / error envelope in one call.
- :class:`ApiError` + :func:`json_endpoint` let a handler ``raise`` a uniform
  400/4xx instead of threading an ``err(...)`` tuple back through every guard.
- :func:`parse_number_arg` parses + bound-checks one numeric value, raising
  :class:`ApiError` on bad input (caught by :func:`json_endpoint`).
- :func:`make_debounced_persist` builds the manifest-write debounce shared by
  the screenspace and transcripts blueprints.
- :func:`make_sse_channel` builds one Server-Sent-Events pub/sub channel
  (bounded per-client queue + coalesce-on-overflow + keepalive + cleanup),
  shared by the run / batch / task streaming endpoints.
- :class:`MediaCache` + :func:`parse_clip_window` back the hover-scrubber
  media routes (sprite sheets / audio snippets) on the Studio and Composer
  blueprints.

Kept deliberately tiny and Flask-only (no ``config``/``utils`` imports) so it
stays import-clean — ``utils`` is Flask-free on purpose and imported by
non-server modules, so these helpers must not live there.
"""

from __future__ import annotations

import math
import queue
import threading
from collections import OrderedDict
from functools import wraps
from typing import Any, Callable

from flask import Response, jsonify, request


def ok(**fields: Any):
    """Success envelope: ``jsonify({"ok": True, **fields})``."""
    return jsonify({"ok": True, **fields})


def err(message: str, code: int = 400):
    """Error envelope: ``(jsonify({"ok": False, "error": message}), code)``."""
    return jsonify({"ok": False, "error": message}), code


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


class MediaCache:
    """Bounded-LRU byte cache with single-flight compute.

    Fast path takes the main lock only for the dict get/reorder. On a miss, a
    per-key lock serializes concurrent identical misses so the expensive
    producer (ffmpeg) runs once, not once per waiting request — the others wake
    to the freshly cached bytes. The producer runs holding no main lock.
    """

    def __init__(self, max_entries: int) -> None:
        self._store: "OrderedDict[tuple, bytes]" = OrderedDict()
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
                    return cached

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


def make_sse_channel(
    *, maxsize: int = 64, keepalive_seconds: float = 15.0
) -> tuple[
    Callable[..., None],
    Callable[..., Response],
    list[tuple[Any, "queue.Queue[str]"]],
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
