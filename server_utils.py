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

Kept deliberately tiny and Flask-only (no ``config``/``utils`` imports) so it
stays import-clean — ``utils`` is Flask-free on purpose and imported by
non-server modules, so these helpers must not live there.
"""

from __future__ import annotations

import math
import threading
from functools import wraps
from typing import Any, Callable

from flask import jsonify


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
