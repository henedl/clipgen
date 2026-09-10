"""Shared Flask scaffolding for the server blueprints.

Every blueprint returns the same JSON envelope — ``{"ok": True, ...}`` on success,
``{"ok": False, "error": msg}`` plus an HTTP status on failure — and repeats the
same numeric-arg parse-and-validate block dozens of times. Collapsed here:

- :func:`ok` / :func:`err` build either envelope in one call;
  :func:`pending` ("still generating") and :func:`refused` ("declined") are
  the two ``ok: False`` states served at HTTP 200.
- :class:`ApiError` + :func:`json_endpoint` let a handler ``raise`` a uniform
  4xx instead of threading an ``err(...)`` tuple back through every guard.
- :func:`parse_number_arg` parses + bound-checks one numeric value;
  :func:`opt_number` is its lenient fall-back-don't-fail sibling;
  :func:`require_json_body` is the JSON-object guard.
- :func:`find_by_id` / :func:`remove_by_id` are the manifest-collection CRUD
  lookups (stashes, blueprints, cuts, annotations).
- :func:`make_debounced_persist` builds the manifest-write debounce.
- :func:`make_participant_cache` builds the mtime-guarded participant cache
  (Transcripts + Screenspace).
- :class:`JobSlot` + :func:`ndjson_batch_response` run one cancellable batch
  at a time and stream it as NDJSON (subtitle embed, audio normalize);
  :class:`JobRegistry` tracks keyed background jobs (remux, model downloads).
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

import json
import math
import queue
import threading
import uuid
from collections import OrderedDict
from collections.abc import Callable, Iterator
from functools import wraps
from pathlib import Path
from typing import Any, cast

from flask import Response, jsonify, request, send_from_directory

import config
import profiling
import utils


def ok(**fields: Any):
    """Success envelope: ``jsonify({"ok": True, **fields})``."""
    return jsonify({"ok": True, **fields})


def err(message: str, code: int = 400, **fields: Any):
    """Error envelope: ``(jsonify({"ok": False, "error": message, **fields}), code)``."""
    return jsonify({"ok": False, "error": message, **fields}), code


def pending(**fields: Any):
    """Third envelope state: ``{"ok": False, "generating": True, ...}`` at HTTP 200.

    Polling routes use it for "not ready yet". The client resolves it (2xx) and
    must branch with ``isPending(data)`` in ``utils.js`` rather than assume
    success.
    """
    return jsonify({"ok": False, "generating": True, **fields})


def refused(reason: str, **fields: Any):
    """Declined at HTTP 200: ``{"ok": False, "reason": reason, ...}``.

    The request was understood and not acted on (cancelled, model not cached).
    Clients branch on ``reason`` or the extra fields, never on the status.
    """
    return jsonify({"ok": False, "reason": reason, **fields})


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


def require_json_body(message: str = "JSON body required") -> dict[str, Any]:
    """The request's JSON object; ``ApiError`` (400) when absent, empty, or not an object."""
    data = request.get_json(silent=True)
    if not data or not isinstance(data, dict):
        raise ApiError(message)
    return cast(dict[str, Any], data)


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
    """Bounded-LRU cache with single-flight compute.

    Fast path takes the main lock only for the dict get/reorder. On a miss, a
    per-key lock serializes concurrent identical misses so the expensive
    producer (ffmpeg, OCR) runs once, not once per waiting request — the others
    wake to the freshly cached value. The producer runs holding no main lock.
    *stat_prefix* names the profiling counters (``<prefix>.hit`` / ``.miss``).
    """

    def __init__(self, max_entries: int, stat_prefix: str = "media_cache") -> None:
        self._store: OrderedDict[tuple, Any] = OrderedDict()
        self._max = max_entries
        self._lock = threading.Lock()
        self._inflight: dict[tuple, threading.Lock] = {}
        self._stat = stat_prefix

    def get_or_compute(self, key: tuple, compute: Callable[[], Any]) -> Any:
        # Fast path: cache hit.
        with self._lock:
            cached = self._store.get(key)
            if cached is not None:
                self._store.move_to_end(key)
                if config.PROFILING:
                    profiling.count(self._stat + ".hit")
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
                        profiling.count(self._stat + ".hit")
                    return cached

            if config.PROFILING:
                profiling.count(self._stat + ".miss")
            with profiling.span(self._stat + ".compute"):
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


class JobRegistry:
    """Keyed background jobs: one running job per key, progress under one lock.

    ``start(key, target, *args)`` is an atomic check-and-set that hands the
    thread a fresh token dict (``fresh()``) and returns it, or ``None`` while
    ``running(token)`` still holds for that key. Threads report through
    ``publish(key, token, **fields)``, which ignores a token the registry has
    since replaced, so a dead run never clobbers its successor. Pollers read
    ``get``/``snapshot`` copies. Remux jobs and LLM downloads use it; the
    Screenspace worker keeps its own priority queue.
    """

    def __init__(
        self,
        fresh: Callable[[], dict[str, Any]],
        running: Callable[[dict[str, Any]], bool],
    ) -> None:
        self._lock = threading.Lock()
        self._jobs: dict[str, dict[str, Any]] = {}
        self._fresh = fresh
        self._running = running

    def start(
        self, key: str, target: Callable[..., None], *args: Any, name: str = ""
    ) -> dict[str, Any] | None:
        with self._lock:
            existing = self._jobs.get(key)
            if existing is not None and self._running(existing):
                return None
            token = self._fresh()
            self._jobs[key] = token
        threading.Thread(
            target=target, args=(token, *args), daemon=True, name=name or f"job-{key}"
        ).start()
        return token

    def seed(self, key: str, token: dict[str, Any]) -> None:
        """Register *token* under *key* without a thread (tests, resumed state)."""
        with self._lock:
            self._jobs[key] = token

    def publish(self, key: str, token: dict[str, Any], **fields: Any) -> bool:
        with self._lock:
            if self._jobs.get(key) is not token:
                return False
            token.update(fields)
            return True

    def get(self, key: str) -> dict[str, Any] | None:
        with self._lock:
            job = self._jobs.get(key)
            return dict(job) if job is not None else None

    def snapshot(self) -> dict[str, dict[str, Any]]:
        with self._lock:
            return {key: dict(job) for key, job in self._jobs.items()}

    def pop(self, key: str) -> None:
        with self._lock:
            self._jobs.pop(key, None)

    def clear(self) -> None:
        """Forget every job (tests and shutdown)."""
        with self._lock:
            self._jobs.clear()


class JobSlot:
    """One in-flight run at a time, with a token-scoped release and cancel.

    ``claim()`` is an atomic check-and-set. Its token is echoed to the client so
    a late cancel cannot stop a successor run; ``release(token)`` is a no-op
    unless that token still owns the slot (``None`` releases unconditionally —
    tests and teardown), which makes the stream's ``finally`` plus the
    response's ``call_on_close`` a harmless double release.
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self.busy = False
        self.owner: str | None = None
        self.cancel_event = threading.Event()

    def claim(self) -> str | None:
        with self._lock:
            if self.busy:
                return None
            self.busy = True
            self.owner = uuid.uuid4().hex
            self.cancel_event.clear()
            return self.owner

    def release(self, token: str | None = None) -> None:
        with self._lock:
            if token is not None and self.owner != token:
                return
            self.busy = False
            self.owner = None

    def cancel(self, token: Any) -> None:
        """Set the cancel event only if *token* owns the running slot."""
        with self._lock:
            if self.busy and token == self.owner:
                self.cancel_event.set()


def ndjson_batch_response(
    slot: JobSlot,
    token: str,
    total: int,
    run_one: Callable[[int], dict[str, Any]],
    header: dict[str, Any] | None = None,
) -> Response:
    """Stream one NDJSON line per item under *slot*, then release it.

    Lines: ``{"total", "token", **header}``; ``{"index", ...run_one(i)}`` per
    item, stopping between items once the slot's cancel event is set (then
    ``{"cancelled": true}``); the terminal ``{"done": true}`` sentinel, whose
    absence tells the client the stream was truncated.
    """
    cancelled = slot.cancel_event.is_set

    def stream() -> Iterator[str]:
        try:
            yield json.dumps({"total": total, "token": token, **(header or {})}) + "\n"
            for idx in range(total):
                if cancelled():
                    break
                outcome = run_one(idx)
                outcome["index"] = idx
                yield json.dumps(outcome) + "\n"
            if cancelled():
                yield json.dumps({"cancelled": True}) + "\n"
            yield json.dumps({"done": True}) + "\n"
        finally:
            # Also runs on client disconnect, so a closed tab cannot wedge the slot.
            slot.release(token)

    response = ndjson_response(stream())
    # An unstarted generator never runs its finally; the token makes this harmless.
    response.call_on_close(lambda: slot.release(token))
    return response


def mtime_or_zero(path: str | Path) -> int:
    """A file's ``st_mtime_ns`` for cache keys, or 0 when it can't be stat'd."""
    try:
        return Path(path).stat().st_mtime_ns
    except OSError:
        return 0


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
            # A racing request or sheet swap may have replaced the source meanwhile.
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


def _profiled_rule(prefix: str) -> str:
    """``"<prefix> <url_rule>"`` for the request in flight; call inside the view.

    Must be resolved eagerly, not from inside a streaming generator: the request
    context is gone by the time the WSGI server iterates the body (that is the
    whole reason ``after_request`` cannot time these — see
    ``profiling.stream_span``).
    """
    rule = request.url_rule.rule if request.url_rule is not None else "?"
    return f"{prefix} {rule}"


def ndjson_response(body: Any) -> Response:
    """Streaming NDJSON response with proxy buffering off; *body* is profiled."""
    return Response(
        profiled_stream(body),
        mimetype="application/x-ndjson",
        headers={"X-Accel-Buffering": "no"},
    )


def profiled_stream(body: Any) -> Any:
    """Wrap a streaming response body so its wall time lands under ``stream <rule>``.

    Passthrough when profiling is off. See ``profiling.stream_span`` for why
    ``route <rule>`` cannot measure these.
    """
    if not config.PROFILING:
        return body
    return profiling.stream_span(
        _profiled_rule("stream"), body, first_label=_profiled_rule("stream.first")
    )


# notify() default: wake every client regardless of key.
_BROADCAST = object()


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

    - ``notify(key=_BROADCAST, marker="update")`` wakes every client registered
      with a matching ``key``; omitting ``key`` wakes every client. On a full queue it coalesces:
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

    def notify(key: Any = _BROADCAST, marker: str = "update") -> None:
        with lock:
            for ckey, cq in clients:
                if key is not _BROADCAST and ckey != key:
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
        # Count only: EventSource auto-reconnects, so a duration would shrink as the stream breaks.
        if config.PROFILING:
            profiling.count(_profiled_rule("sse.open"))
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


# ---- Flask blueprint helpers ----

# render_index_html() expands this from assets/web/_head.html; exported viewers
# are self-contained and skip it.
_HEAD_MARKER = "<!-- CLIPGEN_HEAD_HERE -->"

# Keyed str(index path) -> (index_mtime_ns, head_mtime_ns|None, desktop_chrome,
# rendered); chrome varies per launch, not per file.
_index_html_cache: dict[str, tuple[int | None, int | None, str | None, str]] = {}
_index_html_lock = threading.Lock()


def _desktop_chrome_head(chrome: str) -> str:
    """Inline ``<script>`` telling the page it is hosted in a native window.

    Runs in ``<head>``, so it lands before the deferred ``topnav.js`` reads the
    attribute — the bar lays out inset for the traffic lights on first paint rather
    than jumping. The two measurements come from config so AppKit (which positions
    the real buttons) and CSS (which reserves the space) cannot drift apart.
    """
    return (
        "\n  <script>(function () {\n"
        "    var d = document.documentElement;\n"
        f'    d.dataset.desktopChrome = "{chrome}";\n'
        f'    d.style.setProperty("--desktop-chrome-height", "{config.DESKTOP_CHROME_BAR_HEIGHT}px");\n'
        f'    d.style.setProperty("--desktop-traffic-inset", "{config.DESKTOP_TRAFFIC_LIGHT_INSET}px");\n'
        "  })();</script>"
    )


def render_index_html(assets_dir: Path, index_html: str) -> str:
    """Read an index page, expanding the shared ``<head>`` marker if present.

    Pages without the marker are returned unchanged, so this stays safe for any
    current or future index page. Results are memoized per index path and
    invalidated when the page (or, for marker pages, ``_head.html``) mtime changes.
    """
    index_path = assets_dir / index_html
    head_path = assets_dir / "_head.html"
    chrome = utils.DESKTOP_CHROME
    try:
        index_mtime: int | None = index_path.stat().st_mtime_ns
    except OSError:
        index_mtime = None

    with _index_html_lock:
        cached = _index_html_cache.get(str(index_path))
        if cached is not None and cached[0] == index_mtime and cached[2] == chrome:
            head_mtime_cached = cached[1]
            if head_mtime_cached is None:
                return cached[3]
            try:
                head_mtime: int | None = head_path.stat().st_mtime_ns
            except OSError:
                head_mtime = None
            if head_mtime == head_mtime_cached:
                return cached[3]

        html = index_path.read_text(encoding="utf-8")
        head_mtime_used: int | None = None
        if _HEAD_MARKER in html:
            head = head_path.read_text(encoding="utf-8").rstrip("\n")
            if chrome:
                head += _desktop_chrome_head(chrome)
            html = html.replace(_HEAD_MARKER, head)
            try:
                head_mtime_used = head_path.stat().st_mtime_ns
            except OSError:
                head_mtime_used = None
        _index_html_cache[str(index_path)] = (
            index_mtime,
            head_mtime_used,
            chrome,
            html,
        )
        return html


def register_static_routes(
    bp: Any,
    index_html: str,
    *,
    media_dir_getter: Any = None,
    media_error: str = "Media directory not configured",
    icons: bool = False,
    logos: bool = True,
) -> None:
    """Register standard static-file serving routes on a Flask Blueprint.

    Always registers ``/`` (index) and ``/<path:filename>`` (static assets).
    Optionally registers ``/icons/<path:filename>``, ``/logos/<path:filename>``,
    and ``/media/<path:filename>``.

    Args:
        bp: Flask Blueprint to register routes on.
        index_html: Filename of the HTML page served at ``/``.
        media_dir_getter: Callable returning the current media directory path.
            When provided, a ``/media/<path:filename>`` route is registered.
        media_error: Error message returned (500) when the media dir is falsy.
        icons: When True, registers ``/icons/<path:filename>`` from ``assets/icons/``.
        logos: When True (default), registers ``/logos/<path:filename>`` from
            ``assets/logos/`` so favicons and the brand mark are available to
            every served page.
    """
    from flask import Response, jsonify

    assets_dir = utils.get_bundled_assets_root() / "assets" / "web"

    @bp.route("/")
    def serve_index() -> Response:
        return Response(render_index_html(assets_dir, index_html), mimetype="text/html")

    @bp.route("/<path:filename>")
    def serve_static(filename: str) -> Response:
        return send_from_directory(assets_dir, filename)

    if icons:
        icons_dir = utils.get_bundled_assets_root() / "assets" / "icons"

        @bp.route("/icons/<path:filename>")
        def serve_icons(filename: str) -> Response:
            return send_from_directory(icons_dir, filename)

    if logos:
        logos_dir = utils.get_bundled_assets_root() / "assets" / "logos"

        @bp.route("/logos/<path:filename>")
        def serve_logos(filename: str) -> Response:
            return send_from_directory(logos_dir, filename)

    if media_dir_getter is not None:

        @bp.route("/media/<path:filename>")
        def serve_media(filename: str) -> Response | tuple[Response, int]:
            d = media_dir_getter()
            if not d:
                return jsonify({"ok": False, "error": media_error}), 500
            return send_from_directory(d, filename)
