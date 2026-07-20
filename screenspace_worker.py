# -*- coding: utf-8 -*-
"""Screenspace background worker.

A daemon-thread task queue that runs analysis tasks sequentially with
pause/resume/cancel, multi-video timeline mapping, heatmap rendering, and event
generation. Imports the tool registry, heatmap/manifest helpers, the multitool
offset probe, and the ffprobe metadata helper from sibling modules.
"""

import copy
import queue
import threading
import time
from concurrent.futures import Future, ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

import config
import utils
import video
from screenspace_tools import TOOLS
from screenspace_heatmap import (
    generate_attention_heatmap,
    generate_change_heatmap,
    generate_flow_heatmap,
    generate_heatmap_gif,
    generate_rolling_heatmap_gif,
    generate_template_heatmap,
)
from screenspace_manifest import (
    TASK_STATUS_CANCELLED,
    TASK_STATUS_COMPLETED,
    TASK_STATUS_FAILED,
    TASK_STATUS_PAUSED,
    TASK_STATUS_QUEUED,
    TASK_STATUS_RUNNING,
    _SENTINEL,
    _offset_result_times,
    generate_events_from_results,
)
from screenspace_multitool import _multitool_has_offset
from screenspace_frames import _probe_video_meta

# Per-frame grid payloads consumed only by heatmap generation and never sent to
# the client (unlike flow_grid, which the flow overlay needs). Stripped from
# external reads and dropped from results once heatmaps are written.
_SERVER_ONLY_GRID_KEYS = ("change_grid", "saliency_grid")


def _copy_task_for_read(
    task: dict[str, Any], include_results: bool = True
) -> dict[str, Any]:
    """Deep-copy a task for external reads, omitting server-only grid payloads.

    ``change_grid``/``saliency_grid`` are large per-frame data consumed once by
    heatmap generation and never sent to the client (unlike ``flow_grid``, which
    the flow overlay needs). Excluding them before the deep copy keeps
    progress-driven SSE reads (~every 0.5s on long Change/Attention scans) from
    repeatedly duplicating per-frame grids that are immediately discarded by the
    API layer.

    ``include_results=False`` drops the ever-growing ``result``/``_raw_results``
    lists entirely (reporting ``result_count`` instead), so status ticks stop
    deep-copying the whole detection list every 0.5s; clients pull new results
    via :meth:`ScreenspaceWorker.get_task_result_tail`.
    """
    slim = dict(task)
    if include_results:
        for key in ("result", "_raw_results"):
            seq = task.get(key)
            if isinstance(seq, list) and seq:
                slim[key] = [
                    {k: v for k, v in r.items() if k not in _SERVER_ONLY_GRID_KEYS}
                    if isinstance(r, dict)
                    else r
                    for r in seq
                ]
    else:
        res = task.get("result")
        # Mirror the frontend's old count logic: list → len, truthy non-list
        # (e.g. a single string artifact path) → 1, empty/None → 0.
        slim["result_count"] = len(res) if isinstance(res, list) else (1 if res else 0)
        slim.pop("result", None)
        slim.pop("_raw_results", None)
    return copy.deepcopy(slim)


def dispatch_tool_scan(
    tool: Any,
    video_paths: list[str],
    region_coords: dict[str, int],
    params: dict[str, Any],
    *,
    task_id: str,
    scan_mode: str,
    on_progress: Callable[[float], None],
    cancel_flag: Callable[[], bool],
    on_result: Callable[[dict[str, Any]], None] | None,
    fast_opts: dict[str, Any] | None,
) -> Any:
    """Run one tool's ``scan`` over a participant's source video(s).

    For a single-video participant, scans ``video_paths[0]`` directly. For a
    multi-video participant (a session split across files), maps the task's
    GLOBAL ``[start, end]`` window (``params['start_seconds']`` /
    ``['end_seconds']``) onto the concatenated timeline, scans each spanned
    sub-video at its local offsets, and shifts emitted result times back onto the
    global timeline (tagging each with its ``_source_video``). ``params`` is
    expected pre-resolved (fast-scan interval scaling already applied).

    Shared by :meth:`ScreenspaceWorker._dispatch` and the Workflows ``ss_scan``
    node so both get identical single/multi-video behavior.
    """
    if not video_paths:
        return []
    timeline = video.timeline_or_none(video_paths)
    if timeline is None:
        return tool.scan(
            video_paths[0],
            region_coords,
            params,
            task_id=task_id,
            scan_mode=scan_mode,
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )

    # Multi-video: map the task's GLOBAL [start, end] range onto the timeline
    # and scan each spanned sub-video at its local offsets. Emitted result
    # times are shifted back to the global timeline and tagged with the
    # sub-video they came from so events line up with clips/transcripts.
    total = timeline[-1][1] + timeline[-1][2]
    global_start = params.get("start_seconds", 0.0) or 0.0
    global_end = params.get("end_seconds")
    if global_end is None:
        global_end = total
    pieces = utils.map_global_range_to_segments(timeline, global_start, global_end)
    if not pieces:
        return []
    piece_durations = [le - ls for _index, ls, le in pieces]
    span = sum(piece_durations) or 1.0
    accumulated = 0.0
    all_results: list[dict[str, Any]] = []
    for (index, local_start, local_end), piece_dur in zip(pieces, piece_durations):
        if cancel_flag():
            break
        cumulative = timeline[index][2]
        source_name = Path(timeline[index][0]).name
        frac_start = accumulated / span
        frac_end = (accumulated + piece_dur) / span

        def piece_progress(
            p: float, _a: float = frac_start, _b: float = frac_end
        ) -> None:
            on_progress(_a + p * (_b - _a))

        def piece_on_result(
            rd: dict[str, Any],
            _cum: int = cumulative,
            _src: str = source_name,
        ) -> None:
            if on_result is not None:
                shifted = dict(rd)
                _offset_result_times(shifted, _cum)
                shifted["_source_video"] = _src
                on_result(shifted)

        piece_params = {
            **params,
            "start_seconds": local_start,
            "end_seconds": local_end,
        }
        piece_results = tool.scan(
            timeline[index][0],
            region_coords,
            piece_params,
            task_id=task_id,
            scan_mode=scan_mode,
            on_progress=piece_progress,
            cancel_flag=cancel_flag,
            on_result=piece_on_result,
            fast_opts=fast_opts,
        )
        for rd in piece_results or []:
            _offset_result_times(rd, cumulative)
            rd["_source_video"] = source_name
            all_results.append(rd)
        accumulated += piece_dur
    on_progress(1.0)
    return all_results


class ScreenspaceWorker:
    """Background thread that processes analysis tasks sequentially."""

    def __init__(self) -> None:
        self._queue: queue.PriorityQueue[Any] = queue.PriorityQueue()
        self._tasks: dict[str, dict[str, Any]] = {}
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._running = False
        self._paused = threading.Event()
        self.on_task_complete: Callable[[], None] | None = None
        self.on_progress_update: Callable[[], None] | None = None

    def start(self) -> None:
        """Start the worker thread."""
        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def restore_tasks(self, tasks: list[dict[str, Any]]) -> None:
        """Load historical tasks into the worker (completed/failed/cancelled).

        Tasks the manifest froze mid-flight (``running``/``queued``/``paused``
        — a previous session crashed or was quit during a scan) are demoted to
        ``failed``: no thread will ever continue them after a restart, and
        restoring them verbatim left the task list (and the Overview
        "analysis is still running" banner) claiming a scan was in progress
        forever.
        """
        in_flight = (TASK_STATUS_RUNNING, TASK_STATUS_QUEUED, TASK_STATUS_PAUSED)
        with self._lock:
            for t in tasks:
                if not t.get("id"):
                    continue
                restored = copy.deepcopy(t)
                old_status = restored.get("status")
                if old_status in in_flight:
                    restored["status"] = TASK_STATUS_FAILED
                    restored["error"] = (
                        f"Interrupted — clipgen exited while this scan was "
                        f"{old_status}. Re-run the task to get results."
                    )
                self._tasks[restored["id"]] = restored

    def stop(self) -> None:
        """Signal the worker thread to stop."""
        self._running = False
        self._queue.put((0, "", _SENTINEL))
        if self._thread is not None:
            self._thread.join(timeout=15)

    def enqueue(self, task: dict[str, Any]) -> str:
        """Add a task to the queue. Returns the task ID."""
        task_id = task["id"]
        with self._lock:
            self._tasks[task_id] = task
        self._queue.put((task.get("priority", 100), task["created_at"], task_id))
        return task_id

    def cancel(self, task_id: str) -> bool:
        """Cancel a queued, running, or paused task. Returns True if cancelled."""
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return False
            if task["status"] in (TASK_STATUS_QUEUED, TASK_STATUS_PAUSED):
                task["status"] = TASK_STATUS_CANCELLED
                return True
            if task["status"] == TASK_STATUS_RUNNING:
                task["_cancelled"] = True
                return True
        return False

    def get_task(self, task_id: str) -> dict[str, Any] | None:
        """Return a task dict by ID (thread-safe copy)."""
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return None
            return _copy_task_for_read(task)

    def get_all_tasks(self, include_results: bool = True) -> list[dict[str, Any]]:
        """Return all tasks (thread-safe copies).

        Tasks flagged ``_remove_on_finish`` (dismissed while running) are hidden:
        they are gone as far as the UI and manifest are concerned, even though
        they linger in ``_tasks`` until their worker thread notices the cancel.

        ``include_results=False`` returns slim status copies (no ``result``/
        ``_raw_results``, just ``result_count``) for the ≈0.5s SSE/poll ticks.
        """
        with self._lock:
            return [
                _copy_task_for_read(t, include_results=include_results)
                for t in self._tasks.values()
                if not t.get("_remove_on_finish")
            ]

    def get_task_result_tail(
        self, task_id: str, since: int = 0
    ) -> tuple[list[dict[str, Any]], int] | None:
        """Thread-safe copy of a task's result tail from index ``since``.

        Results are appended in scan order during a run (see ``_on_result``), so a
        count cursor yields exactly the new detections and keeps the live-results
        fetch flat. Strips server-only grid payloads like :func:`_copy_task_for_read`.
        Returns ``(tail, total)``, or ``None`` if the task is unknown.
        """
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return None
            res = task.get("result") or []
            total = len(res)
            start = max(since, 0)
            tail = res[start:] if start < total else []
            stripped = [
                {k: v for k, v in r.items() if k not in _SERVER_ONLY_GRID_KEYS}
                if isinstance(r, dict)
                else r
                for r in tail
            ]
            return copy.deepcopy(stripped), total

    def reorder(self, task_ids: list[str]) -> bool:
        """Reorder queued tasks by the given ID sequence.

        Assigns new priorities so that earlier IDs in the list have
        lower (higher-priority) values.  Drains and re-inserts queue
        items so the PriorityQueue heap reflects the new ordering.
        """
        with self._lock:
            for i, tid in enumerate(task_ids):
                task = self._tasks.get(tid)
                if task and task["status"] == TASK_STATUS_QUEUED:
                    task["priority"] = i + 1

            # Drain the queue and re-insert with updated priorities
            pending_items: list[Any] = []
            while not self._queue.empty():
                try:
                    pending_items.append(self._queue.get_nowait())
                except queue.Empty:
                    break
            for priority, created_at, tid in pending_items:
                task = self._tasks.get(tid) if isinstance(tid, str) else None
                new_priority = task["priority"] if task else priority
                self._queue.put((new_priority, created_at, tid))
        return True

    @property
    def is_alive(self) -> bool:
        """Return whether the worker thread is alive."""
        return self._thread is not None and self._thread.is_alive()

    @property
    def is_paused(self) -> bool:
        """Return whether the queue is paused."""
        return self._paused.is_set()

    def pause(self) -> None:
        """Pause the queue. Stops the running task so it yields partial results."""
        self._paused.set()
        with self._lock:
            for task in self._tasks.values():
                if task["status"] == TASK_STATUS_RUNNING:
                    task["_paused_flag"] = True

    def resume(self) -> None:
        """Resume the queue. Re-enqueues paused tasks from where they left off."""
        self._paused.clear()
        to_resume: list[dict[str, Any]] = []
        with self._lock:
            for task in self._tasks.values():
                if task["status"] == TASK_STATUS_PAUSED:
                    to_resume.append(task)

        for task in to_resume:
            progress = task.get("progress", 0.0)
            params = task.get("parameters", {})

            # Offset multitool chains run the two-phase scan: phase 2 joins
            # across every frame from the original start, so they cannot resume
            # mid-stream (advancing start_seconds would drop the pre-pause
            # frames the join depends on). Restart these from scratch instead —
            # ``start_seconds`` is left untouched and no partial results carry
            # over, so the re-run reproduces the full result set.
            if task.get("type") == "multitool" and _multitool_has_offset(
                params.get("steps", [])
            ):
                with self._lock:
                    task.pop("_partial_results", None)
                    task.pop("_progress_offset", None)
                    task.pop("_progress_scale", None)
                    task.pop("_paused_flag", None)
                    task["result"] = []
                    task["progress"] = 0.0
                    task["status"] = TASK_STATUS_QUEUED
                self._queue.put(
                    (task.get("priority", 100), task["created_at"], task["id"])
                )
                continue

            start = params.get("start_seconds", 0.0)
            end = params.get("end_seconds")
            if end is None:
                # Resume math is on the GLOBAL timeline; multi-video end is the
                # total across all parts.
                timeline = video.timeline_or_none(task["video_paths"])
                if timeline is not None:
                    end = timeline[-1][1] + timeline[-1][2]
                else:
                    _, end = _probe_video_meta(task["video_paths"][0])
            # ``progress`` is a GLOBAL fraction of the original scan range
            # (mapped via _progress_offset/_progress_scale in _on_progress),
            # but ``start`` is the CURRENT segment start — already advanced by
            # any previous resume. Convert back to the current segment's local
            # fraction before projecting, or a 2nd+ resume overshoots the true
            # stop point and skips frames.
            prev_offset = task.get("_progress_offset", 0.0)
            prev_scale = task.get("_progress_scale", 1.0)
            local = max(0.0, min(1.0, (progress - prev_offset) / prev_scale))
            resume_at = start + local * (end - start)

            with self._lock:
                task["_partial_results"] = list(task.get("result") or [])
                task["_progress_offset"] = progress
                task["_progress_scale"] = max(1.0 - progress, 0.001)
                task.pop("_paused_flag", None)
                params["start_seconds"] = resume_at
                task["status"] = TASK_STATUS_QUEUED

            self._queue.put((task.get("priority", 100), task["created_at"], task["id"]))

    def remove_task(self, task_id: str) -> bool:
        """Cancel (if active) and fully remove a task."""
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return False
            if task["status"] == TASK_STATUS_RUNNING:
                # The running scan's cancel_flag looks the task up by id, so the
                # task must stay in ``_tasks`` for the cancel to land — popping it
                # here would strand the worker thread, which would run the scan to
                # completion and keep streaming progress (CPU + SSE spam). Flag it
                # cancelled + remove-on-finish; ``get_all_tasks`` hides it from the
                # UI/manifest immediately and ``_execute_task`` pops it once the
                # scan unwinds.
                task["_cancelled"] = True
                task["_remove_on_finish"] = True
                return True
            self._tasks.pop(task_id, None)
            return True

    def drain_new_events(self) -> list[dict[str, Any]]:
        """Collect and clear ``_generated_events`` from all tasks. Thread-safe."""
        events: list[dict[str, Any]] = []
        with self._lock:
            for t in self._tasks.values():
                events.extend(t.pop("_generated_events", []))
        return events

    def _generate_events_from_results(
        self, task: dict[str, Any], raw_results: list[dict[str, Any]]
    ) -> list[dict[str, Any]]:
        return generate_events_from_results(task, raw_results)

    def _write_heatmap_gifs(
        self,
        task_id: str,
        results: list[dict[str, Any]],
        width: int,
        height: int,
        heatmap_type: str,
        *,
        rolling: bool,
    ) -> dict[str, str]:
        """Write the cumulative (and optionally rolling-window) heatmap GIFs.

        Returns the generated GIF filenames keyed by attachment name.
        """
        out_dir = Path(utils.get_effective_output_dir())
        attachments: dict[str, str] = {}
        gp = generate_heatmap_gif(
            results,
            width,
            height,
            str(out_dir / f"heatmap_{task_id}.gif"),
            heatmap_type=heatmap_type,
        )
        if gp:
            attachments["heatmap_gif"] = Path(gp).name
        if rolling:
            rp = generate_rolling_heatmap_gif(
                results,
                width,
                height,
                str(out_dir / f"heatmap_rolling_{task_id}.gif"),
                heatmap_type=heatmap_type,
                window_frames=config.SCREENSPACE_HEATMAP_ROLLING_WINDOW,
            )
            if rp:
                attachments["heatmap_rolling_gif"] = Path(rp).name
        return attachments

    def _generate_heatmap(
        self,
        task_type: str,
        task_id: str,
        video_paths: list[str],
        region_coords: dict[str, Any],
        results: list[dict[str, Any]],
    ) -> dict[str, str]:
        """Generate heatmap artifacts for template, flow, change, or attention tasks.

        Pure compute + file I/O — takes primitives instead of the task dict and
        returns the generated filenames keyed by attachment name, so it can run
        *outside* the worker lock (it touches no shared task state). Honors the
        per-tool ``SCREENSPACE_GENERATE_*_HEATMAP`` settings: when a tool's
        heatmaps are disabled, returns ``{}``.
        """
        heatmap_enabled = {
            "template": config.SCREENSPACE_GENERATE_TEMPLATE_HEATMAP,
            "flow": config.SCREENSPACE_GENERATE_FLOW_HEATMAP,
            "change": config.SCREENSPACE_GENERATE_CHANGE_HEATMAP,
            "attention": config.SCREENSPACE_GENERATE_ATTENTION_HEATMAP,
        }
        if not heatmap_enabled.get(task_type, False):
            return {}
        attachments: dict[str, str] = {}
        heatmap_path = str(
            Path(utils.get_effective_output_dir()) / f"heatmap_{task_id}.png"
        )
        if task_type == "template":
            props = video.probe_video_properties(video_paths[0])
            fw = props.get("width", 1920) if props else 1920
            fh = props.get("height", 1080) if props else 1080
            hp = generate_template_heatmap(results, fw, fh, heatmap_path)
            if hp:
                attachments["heatmap"] = Path(hp).name
            attachments.update(
                self._write_heatmap_gifs(
                    task_id, results, fw, fh, "template", rolling=True
                )
            )
        elif task_type == "attention":
            # Full-frame tool: region_coords is {0,0,0,0}, so size to the video
            # frame like template. The rolling GIF is the eye-tracking-style
            # gaze-replay deliverable.
            props = video.probe_video_properties(video_paths[0])
            fw = props.get("width", 1920) if props else 1920
            fh = props.get("height", 1080) if props else 1080
            hp = generate_attention_heatmap(results, fw, fh, heatmap_path)
            if hp:
                attachments["heatmap"] = Path(hp).name
            attachments.update(
                self._write_heatmap_gifs(
                    task_id, results, fw, fh, "attention", rolling=True
                )
            )
        elif task_type in ("flow", "change"):
            rw = region_coords.get("w", 256)
            rh = region_coords.get("h", 256)
            if task_type == "flow":
                hp = generate_flow_heatmap(results, rw, rh, heatmap_path)
            else:
                hp = generate_change_heatmap(results, rw, rh, heatmap_path)
            if hp:
                attachments["heatmap"] = Path(hp).name
            attachments.update(
                self._write_heatmap_gifs(
                    task_id, results, rw, rh, task_type, rolling=task_type == "change"
                )
            )
        return attachments

    def _run(self) -> None:
        """Worker loop with concurrent task execution via ThreadPoolExecutor."""
        max_workers = config.SCREENSPACE_PARALLEL_WORKERS
        with ThreadPoolExecutor(max_workers=max_workers) as pool:
            active: dict[str, Future[None]] = {}

            while self._running:
                try:
                    # 1. Collect completed futures
                    done_ids = [tid for tid, f in active.items() if f.done()]
                    for tid in done_ids:
                        future = active.pop(tid)
                        try:
                            future.result()
                        except Exception as exc:
                            utils.warning_print(f"Worker task {tid} raised: {exc}")
                        if self.on_task_complete:
                            try:
                                self.on_task_complete()
                            except Exception as exc:
                                utils.warning_print(
                                    f"on_task_complete callback failed: {exc}"
                                )
                        if self.on_progress_update:
                            try:
                                self.on_progress_update()
                            except Exception as exc:
                                utils.warning_print(
                                    f"on_progress_update callback failed: {exc}"
                                )

                    # 2. If paused, wait
                    if self._paused.is_set():
                        time.sleep(0.25)
                        continue

                    # 3. Submit new tasks if capacity available
                    if len(active) >= max_workers:
                        time.sleep(0.1)
                        continue

                    try:
                        item = self._queue.get(timeout=0.25)
                    except queue.Empty:
                        continue

                    priority, created_at, task_id = item
                    if task_id is _SENTINEL:
                        # Wait for active tasks to finish
                        for drain_tid, f in active.items():
                            try:
                                f.result(timeout=10)
                            except Exception as exc:
                                utils.debug_print(
                                    f"Task {drain_tid} raised during shutdown: {exc}"
                                )
                        if self.on_task_complete:
                            try:
                                self.on_task_complete()
                            except Exception as exc:
                                utils.warning_print(
                                    f"on_task_complete callback failed during shutdown: {exc}"
                                )
                        break

                    if self._paused.is_set():
                        self._queue.put(item)
                        time.sleep(0.25)
                        continue

                    with self._lock:
                        task = self._tasks.get(task_id)
                        if task is None or task["status"] != TASK_STATUS_QUEUED:
                            continue
                        task["status"] = TASK_STATUS_RUNNING

                    future = pool.submit(self._execute_task, task)
                    active[task_id] = future

                except Exception as exc:
                    utils.warning_print(f"Worker loop error: {exc}")

    def _execute_task(self, task: dict[str, Any]) -> None:
        """Dispatch task to the appropriate workflow function."""
        task_id = task["id"]

        # Seed incremental results (includes partial results on resume)
        with self._lock:
            t = self._tasks.get(task_id)
            if t:
                t["result"] = list(t.get("_partial_results", []))
                t["_raw_results"] = list(t.get("_partial_results", []))

        # Per-task throttle for SSE notifications (avoids cross-task race)
        last_notify: list[float] = [0.0]

        def _throttled_notify() -> None:
            now = time.monotonic()
            if now - last_notify[0] >= 0.5:
                last_notify[0] = now
                if self.on_progress_update:
                    self.on_progress_update()

        def _on_progress(progress: float) -> None:
            with self._lock:
                t = self._tasks.get(task_id)
                if t and not t.get("_paused_flag"):
                    offset = t.get("_progress_offset", 0.0)
                    scale = t.get("_progress_scale", 1.0)
                    t["progress"] = min(offset + progress * scale, 1.0)
            _throttled_notify()

        def _cancel_flag() -> bool:
            with self._lock:
                t = self._tasks.get(task_id)
                return bool(t and (t.get("_cancelled") or t.get("_paused_flag")))

        def _on_result(result_dict: dict[str, Any]) -> None:
            with self._lock:
                t = self._tasks.get(task_id)
                if t:
                    if isinstance(t.get("result"), list):
                        t["result"].append(result_dict)
                    if isinstance(t.get("_raw_results"), list):
                        t["_raw_results"].append(result_dict)
                    if t.get("parameters", {}).get("detect_first"):
                        t["_cancelled"] = True
            _throttled_notify()

        try:
            result = self._dispatch(task, _on_progress, _cancel_flag, _on_result)

            def _visible_results(seq: Any) -> Any:
                # Attention's visible results are the confirmed shifts only;
                # the full per-sample stream (one entry per 0.5s, thousands on
                # long videos) exists solely to feed heatmap dwell weighting
                # and must never reach the results/timeline API — neither in
                # the completed→heatmap-attached window nor while paused.
                if task.get("type") == "attention" and isinstance(seq, list):
                    return [r for r in seq if isinstance(r, dict) and r.get("shift")]
                return seq

            # Inputs for deferred (lock-free) heatmap generation, captured under
            # the lock and consumed after it's released.
            heatmap_inputs: tuple[str, list[str], dict[str, Any]] | None = None
            with self._lock:
                t = self._tasks.get(task_id)
                if t:
                    if t.get("_paused_flag"):
                        t["status"] = TASK_STATUS_PAUSED
                        # ``result`` holds only THIS invocation's detections;
                        # prepend the pre-pause carry-over (like the completed
                        # branch below) or a pause on a resumed scan drops the
                        # earlier results for good — the next resume re-seeds
                        # _partial_results from t["result"]. For attention that
                        # re-seed is shift-only, so a resumed scan's heatmap
                        # under-accumulates the pre-pause segment (accepted;
                        # see plans/ATTENTION-PLAN.md).
                        partial = t.get("_partial_results")
                        if partial and isinstance(result, list):
                            result = partial + result
                        t["result"] = _visible_results(result)
                    elif t.get("_cancelled") and not t.get("parameters", {}).get(
                        "detect_first"
                    ):
                        t["status"] = TASK_STATUS_CANCELLED
                    else:
                        # Normal completion or detect_first early stop
                        partial = t.pop("_partial_results", None)
                        if partial and isinstance(result, list):
                            result = partial + result
                        t["status"] = TASK_STATUS_COMPLETED
                        t["result"] = _visible_results(result)
                        t["progress"] = 1.0
                        t.pop("_progress_offset", None)
                        t.pop("_progress_scale", None)
                        raw = t.pop("_raw_results", [])
                        t["_generated_events"] = self._generate_events_from_results(
                            t, raw
                        )
                        if isinstance(result, list) and result:
                            # Defer heatmap/GIF generation to outside the lock —
                            # it's heavy I/O (PNG + cumulative/rolling GIFs) that
                            # would otherwise block status reads, cancellation,
                            # dismissal, and SSE snapshots after scan completion.
                            heatmap_inputs = (
                                t.get("type", ""),
                                list(t.get("video_paths", [])),
                                dict(t.get("region_coords", {})),
                            )
                    t["completed_at"] = datetime.now(timezone.utc).isoformat()

            # Generate heatmaps without holding the lock, then briefly reacquire
            # it to attach the filenames. Safe because the scan has finished
            # (no more _on_result appends) and nothing else mutates `result`.
            if heatmap_inputs is not None and isinstance(result, list) and result:
                task_type, video_paths, region_coords = heatmap_inputs
                attachments = self._generate_heatmap(
                    task_type, task_id, video_paths, region_coords, result
                )
                # change_grid/saliency_grid are consumed only by heatmap
                # generation; drop them so completed tasks don't retain
                # per-frame grids in memory until dismissal. Attention's
                # shift-only t["result"] shares these dicts, so its visible
                # entries are stripped by the same pass.
                for r in result:
                    if isinstance(r, dict):
                        for key in _SERVER_ONLY_GRID_KEYS:
                            r.pop(key, None)
                with self._lock:
                    t = self._tasks.get(task_id)
                    if t is not None:
                        t.update(attachments)
                # The task was already marked completed before these heatmap
                # filenames were attached, so emit an SSE update now — otherwise
                # the frontend (which may already have seen the completed task via
                # another worker's progress push) never re-renders the heatmap
                # section until a page reload.
                if attachments and self.on_progress_update:
                    self.on_progress_update()
        except Exception as exc:
            with self._lock:
                t = self._tasks.get(task_id)
                if t:
                    t["status"] = TASK_STATUS_FAILED
                    t["error"] = str(exc)
                    t["completed_at"] = datetime.now(timezone.utc).isoformat()
        finally:
            # A task dismissed while running was kept in _tasks so its cancel
            # could land; now that the scan has unwound, drop it for good.
            with self._lock:
                t = self._tasks.get(task_id)
                if t is not None and t.get("_remove_on_finish"):
                    self._tasks.pop(task_id, None)

    def _dispatch(
        self,
        task: dict[str, Any],
        on_progress: Callable[[float], None],
        cancel_flag: Callable[[], bool],
        on_result: Callable[[dict[str, Any]], None] | None = None,
    ) -> Any:
        """Route task to the correct tool via the ``TOOLS`` registry.

        Builds the shared fast-scan ``fast_opts`` payload here so individual
        tool classes only need to declare ``fast_scan_region_dim`` /
        ``supports_fast_scan`` / ``fast_scan_extra_opts`` as class attributes,
        then delegates the single/multi-video scan to :func:`dispatch_tool_scan`.
        """
        task_type = task["type"]
        tool = TOOLS.get(task_type)
        if tool is None:
            raise ValueError(f"Unknown task type: {task_type}")

        if task_type in config.SCREENSPACE_MASK_FALLBACK_TOOLS and task.get(
            "region_coords", {}
        ).get("mask_points"):
            utils.warning_print(
                f"{task_type}: shaped region — analyzing its bounding rect "
                "(this tool cannot honor a polygon mask)"
            )

        # Shallow copy so fast-scan interval scaling is not persisted on the
        # task dict (pause/resume re-dispatches the same parameters).
        params = dict(task.get("parameters", {}))
        scan_mode = params.get("scan_mode", "normal")
        fast_opts: dict[str, Any] | None = None
        if scan_mode == "fast" and tool.supports_fast_scan:
            multiplier = config.SCREENSPACE_FAST_SCAN_INTERVAL_MULTIPLIER
            if params.get("interval", 0) > 0:
                params["interval"] = params["interval"] * multiplier
            fast_opts = {
                "phash_skip": True,
                "max_region_dim": tool.fast_scan_region_dim,
                **tool.fast_scan_extra_opts,
            }

        return dispatch_tool_scan(
            tool,
            task["video_paths"],
            task["region_coords"],
            params,
            task_id=task["id"],
            scan_mode=scan_mode,
            on_progress=on_progress,
            cancel_flag=cancel_flag,
            on_result=on_result,
            fast_opts=fast_opts,
        )


# ---------------------------------------------------------------------------
# Manifest I/O
# ---------------------------------------------------------------------------
