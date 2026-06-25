"""Tests for task creation and the ScreenspaceWorker queue."""

import threading
import time
from unittest import mock


import config
import screenspace


class TestCreateTask:
    def test_creates_valid_task(self):
        task = screenspace.create_task(
            task_type="color",
            participant="P01",
            source_video="study_P01.mp4",
            video_paths=["/path/study_P01.mp4"],
            region_name="healthbar",
            region_coords={"x": 0, "y": 0, "w": 100, "h": 50},
        )
        assert task["id"].startswith("ss_")
        assert task["type"] == "color"
        assert task["status"] == "queued"
        assert task["progress"] == 0.0


class TestScreenspaceWorker:
    def test_enqueue_and_get(self):
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "color",
            "P01",
            "s_P01.mp4",
            ["/v.mp4"],
            "hb",
            {"x": 0, "y": 0, "w": 10, "h": 10},
        )
        tid = worker.enqueue(task)
        assert tid == task["id"]
        retrieved = worker.get_task(tid)
        assert retrieved is not None
        assert retrieved["id"] == tid

    def test_get_all_tasks(self):
        worker = screenspace.ScreenspaceWorker()
        t1 = screenspace.create_task(
            "color", "P01", "s.mp4", ["/v.mp4"], "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        t2 = screenspace.create_task(
            "change", "P02", "s.mp4", ["/v.mp4"], "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        worker.enqueue(t1)
        worker.enqueue(t2)
        all_tasks = worker.get_all_tasks()
        assert len(all_tasks) == 2
        ids = {t["id"] for t in all_tasks}
        assert t1["id"] in ids and t2["id"] in ids

    def test_cancel_queued_task(self):
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "color", "P01", "s.mp4", ["/v.mp4"], "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        worker.enqueue(task)
        assert worker.cancel(task["id"]) is True
        t = worker.get_task(task["id"])
        assert t is not None
        assert t["status"] == "cancelled"

    def test_cancel_nonexistent(self):
        worker = screenspace.ScreenspaceWorker()
        assert worker.cancel("ss_nonexist") is False

    def test_worker_processes_and_fails_bad_video(self):
        worker = screenspace.ScreenspaceWorker()
        worker.start()
        try:
            task = screenspace.create_task(
                "color",
                "P01",
                "nope.mp4",
                ["/nonexistent/nope.mp4"],
                "r",
                {"x": 0, "y": 0, "w": 10, "h": 10},
                parameters={
                    "target_color": {"h": 0, "s": 0, "v": 0},
                    "tolerance": {"h": 10, "s": 10, "v": 10},
                },
            )
            worker.enqueue(task)
            for _ in range(50):
                t = worker.get_task(task["id"])
                if t and t["status"] in ("completed", "failed"):
                    break
                time.sleep(0.05)
            t = worker.get_task(task["id"])
            assert t is not None
            assert t["status"] in ("completed", "failed")
        finally:
            worker.stop()

    def test_worker_survives_on_task_complete_exception(self):
        """Worker continues processing tasks after on_task_complete raises."""
        worker = screenspace.ScreenspaceWorker()
        call_count = {"n": 0}

        def bad_callback():
            call_count["n"] += 1
            if call_count["n"] == 1:
                raise TypeError("simulated persistence failure")

        worker.on_task_complete = bad_callback
        worker.start()
        try:
            t1 = screenspace.create_task(
                "color",
                "P01",
                "s.mp4",
                ["/nonexistent/v.mp4"],
                "r",
                {"x": 0, "y": 0, "w": 10, "h": 10},
                parameters={
                    "target_color": {"h": 0, "s": 0, "v": 0},
                    "tolerance": {"h": 10, "s": 10, "v": 10},
                },
            )
            worker.enqueue(t1)
            for _ in range(50):
                t = worker.get_task(t1["id"])
                if t and t["status"] in ("completed", "failed"):
                    break
                time.sleep(0.05)

            # Second task should still process even though first callback raised
            t2 = screenspace.create_task(
                "color",
                "P02",
                "s.mp4",
                ["/nonexistent/v.mp4"],
                "r",
                {"x": 0, "y": 0, "w": 10, "h": 10},
                parameters={
                    "target_color": {"h": 0, "s": 0, "v": 0},
                    "tolerance": {"h": 10, "s": 10, "v": 10},
                },
            )
            worker.enqueue(t2)
            for _ in range(50):
                t = worker.get_task(t2["id"])
                if t and t["status"] in ("completed", "failed"):
                    break
                time.sleep(0.05)
            t = worker.get_task(t2["id"])
            assert t is not None
            assert t["status"] in ("completed", "failed")
            # on_task_complete fires in the _run loop after the future is
            # collected, which may lag behind the task status change.
            # Wait for the callback to actually fire for task 2.
            for _ in range(40):
                if call_count["n"] >= 2:
                    break
                time.sleep(0.05)
            assert call_count["n"] >= 2
        finally:
            worker.stop()

    def test_text_task_easyocr_importable(self):
        import easyocr  # noqa: F401

    def test_reorder(self):
        worker = screenspace.ScreenspaceWorker()
        t1 = screenspace.create_task(
            "color", "P01", "s.mp4", ["/v.mp4"], "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        t2 = screenspace.create_task(
            "change", "P01", "s.mp4", ["/v.mp4"], "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        worker.enqueue(t1)
        worker.enqueue(t2)
        assert worker.reorder([t2["id"], t1["id"]]) is True
        got = worker.get_task(t2["id"])
        assert got is not None
        assert got["priority"] == 1

    def test_get_task_returns_none_for_unknown(self):
        worker = screenspace.ScreenspaceWorker()
        assert worker.get_task("ss_unknown") is None

    def test_remove_queued_task(self):
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "color", "P01", "s.mp4", ["/v.mp4"], "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        worker.enqueue(task)
        assert worker.remove_task(task["id"]) is True
        assert worker.get_task(task["id"]) is None

    def test_remove_nonexistent(self):
        worker = screenspace.ScreenspaceWorker()
        assert worker.remove_task("ss_nonexist") is False

    def test_remove_running_task_cancels_and_hides(self):
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "color", "P01", "s.mp4", ["/v.mp4"], "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        worker.enqueue(task)
        # Simulate running status
        with worker._lock:
            worker._tasks[task["id"]]["status"] = screenspace.TASK_STATUS_RUNNING
        assert worker.remove_task(task["id"]) is True
        # The task must stay in _tasks (the scan's cancel_flag looks it up by id)
        # but be flagged for cancellation + removal, and hidden from the UI.
        with worker._lock:
            kept = worker._tasks[task["id"]]
            assert kept["_cancelled"] is True
            assert kept["_remove_on_finish"] is True
        assert worker.get_all_tasks() == []

    def test_pause_resume_flags(self):
        worker = screenspace.ScreenspaceWorker()
        assert worker.is_paused is False
        worker.pause()
        assert worker.is_paused is True
        worker.resume()
        assert worker.is_paused is False

    def test_pause_sets_paused_flag_on_running_task(self):
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "color", "P01", "s.mp4", ["/v.mp4"], "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        worker.enqueue(task)
        with worker._lock:
            worker._tasks[task["id"]]["status"] = screenspace.TASK_STATUS_RUNNING
        worker.pause()
        with worker._lock:
            assert worker._tasks[task["id"]].get("_paused_flag") is True

    def test_resume_requeues_paused_task(self):
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "color",
            "P01",
            "s.mp4",
            ["/v.mp4"],
            "r",
            {"x": 0, "y": 0, "w": 1, "h": 1},
            parameters={"start_seconds": 0.0, "end_seconds": 100.0},
        )
        worker.enqueue(task)
        with worker._lock:
            worker._tasks[task["id"]]["status"] = screenspace.TASK_STATUS_PAUSED
            worker._tasks[task["id"]]["progress"] = 0.5
            worker._tasks[task["id"]]["result"] = [{"timestamp": 5.0}]
        worker.resume()
        t = worker.get_task(task["id"])
        assert t is not None
        assert t["status"] == "queued"
        assert t.get("parameters", {}).get("start_seconds") == 50.0

    def test_resume_restarts_offset_multitool_from_scratch(self):
        # Offset chains need every frame from the original start to join, so a
        # paused offset task must restart wholesale — start_seconds is NOT
        # advanced and no partial results carry over.
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "multitool",
            "P01",
            "s.mp4",
            ["/v.mp4"],
            "r",
            {"x": 0, "y": 0, "w": 1, "h": 1},
            parameters={
                "start_seconds": 10.0,
                "end_seconds": 100.0,
                "steps": [
                    {"type": "color"},
                    {"type": "change", "offset": {"min": 0.0, "max": 3.0}},
                ],
            },
        )
        worker.enqueue(task)
        with worker._lock:
            worker._tasks[task["id"]]["status"] = screenspace.TASK_STATUS_PAUSED
            worker._tasks[task["id"]]["progress"] = 0.5
            worker._tasks[task["id"]]["result"] = [{"timestamp": 5.0}]
        worker.resume()
        t = worker.get_task(task["id"])
        assert t is not None
        assert t["status"] == "queued"
        assert t["parameters"]["start_seconds"] == 10.0  # unchanged
        assert t["result"] == []
        assert t["progress"] == 0.0
        assert "_partial_results" not in t

    def test_resume_advances_non_offset_multitool(self):
        # A multitool with no offsets resumes incrementally like any other task.
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "multitool",
            "P01",
            "s.mp4",
            ["/v.mp4"],
            "r",
            {"x": 0, "y": 0, "w": 1, "h": 1},
            parameters={
                "start_seconds": 10.0,
                "end_seconds": 100.0,
                "steps": [{"type": "color"}, {"type": "change"}],
            },
        )
        worker.enqueue(task)
        with worker._lock:
            worker._tasks[task["id"]]["status"] = screenspace.TASK_STATUS_PAUSED
            worker._tasks[task["id"]]["progress"] = 0.5
            worker._tasks[task["id"]]["result"] = [{"timestamp": 5.0}]
        worker.resume()
        t = worker.get_task(task["id"])
        assert t is not None
        assert t["parameters"]["start_seconds"] == 55.0  # 10 + 0.5*(100-10)
        assert t.get("_partial_results") == [{"timestamp": 5.0}]


class TestWorkerParallel:
    def test_two_tasks_run_concurrently(self, monkeypatch):
        """With PARALLEL_WORKERS=2, two tasks reach RUNNING simultaneously."""
        monkeypatch.setattr(config, "SCREENSPACE_PARALLEL_WORKERS", 2)

        barrier = threading.Barrier(2, timeout=5)
        reached_running = {"count": 0}

        def slow_dispatch(self, task, on_progress, cancel_flag, on_result=None):
            reached_running["count"] += 1
            barrier.wait()  # both tasks must reach here
            time.sleep(0.05)
            return []

        monkeypatch.setattr(screenspace.ScreenspaceWorker, "_dispatch", slow_dispatch)

        worker = screenspace.ScreenspaceWorker()
        worker.start()

        t1 = screenspace.create_task(
            "color",
            "P01",
            "s.mp4",
            ["/v.mp4"],
            "r1",
            {"x": 0, "y": 0, "w": 1, "h": 1},
        )
        t2 = screenspace.create_task(
            "color",
            "P02",
            "s.mp4",
            ["/v2.mp4"],
            "r2",
            {"x": 0, "y": 0, "w": 1, "h": 1},
        )
        worker.enqueue(t1)
        worker.enqueue(t2)

        # Wait for both tasks to complete
        for _ in range(100):
            tasks = worker.get_all_tasks()
            statuses = [t["status"] for t in tasks]
            if all(s in ("completed", "failed") for s in statuses):
                break
            time.sleep(0.05)

        worker.stop()
        # Both tasks reached the barrier (ran concurrently)
        assert reached_running["count"] == 2

    def test_sequential_when_workers_1(self, monkeypatch):
        """With PARALLEL_WORKERS=1, tasks execute one at a time."""
        monkeypatch.setattr(config, "SCREENSPACE_PARALLEL_WORKERS", 1)

        max_concurrent = {"value": 0, "current": 0}
        lock = threading.Lock()

        def counting_dispatch(self, task, on_progress, cancel_flag, on_result=None):
            with lock:
                max_concurrent["current"] += 1
                max_concurrent["value"] = max(
                    max_concurrent["value"], max_concurrent["current"]
                )
            time.sleep(0.1)
            with lock:
                max_concurrent["current"] -= 1
            return []

        monkeypatch.setattr(
            screenspace.ScreenspaceWorker, "_dispatch", counting_dispatch
        )

        worker = screenspace.ScreenspaceWorker()
        worker.start()

        for i in range(3):
            t = screenspace.create_task(
                "color",
                f"P0{i}",
                "s.mp4",
                [f"/v{i}.mp4"],
                f"r{i}",
                {"x": 0, "y": 0, "w": 1, "h": 1},
            )
            worker.enqueue(t)

        for _ in range(100):
            tasks = worker.get_all_tasks()
            if all(t["status"] in ("completed", "failed") for t in tasks):
                break
            time.sleep(0.05)

        worker.stop()
        assert max_concurrent["value"] == 1

    def test_parallel_pause_flags_all_running(self, monkeypatch):
        """Pausing flags all running tasks."""
        monkeypatch.setattr(config, "SCREENSPACE_PARALLEL_WORKERS", 2)

        gate = threading.Event()

        def blocking_dispatch(self, task, on_progress, cancel_flag, on_result=None):
            gate.wait(timeout=5)
            return []

        monkeypatch.setattr(
            screenspace.ScreenspaceWorker, "_dispatch", blocking_dispatch
        )

        worker = screenspace.ScreenspaceWorker()
        worker.start()

        t1 = screenspace.create_task(
            "color",
            "P01",
            "s.mp4",
            ["/v.mp4"],
            "r1",
            {"x": 0, "y": 0, "w": 1, "h": 1},
        )
        t2 = screenspace.create_task(
            "color",
            "P02",
            "s.mp4",
            ["/v2.mp4"],
            "r2",
            {"x": 0, "y": 0, "w": 1, "h": 1},
        )
        worker.enqueue(t1)
        worker.enqueue(t2)

        # Wait for both to be RUNNING
        for _ in range(50):
            tasks = worker.get_all_tasks()
            running = [t for t in tasks if t["status"] == "running"]
            if len(running) == 2:
                break
            time.sleep(0.05)

        worker.pause()
        # Check both have _paused_flag
        with worker._lock:
            for tid in [t1["id"], t2["id"]]:
                assert worker._tasks[tid].get("_paused_flag") is True

        gate.set()
        worker.stop()

    def test_dismiss_running_task_propagates_cancel(self, monkeypatch):
        """Dismissing a running task must actually stop its scan.

        Regression: remove_task used to pop the task from _tasks immediately, but
        the scan's cancel_flag looks the task up by id — so the cancel never
        landed, the worker ran to completion, and it kept streaming progress
        (pinned CPU + a flood of SSE pushes / icon re-fetches).
        """
        saw_cancel = threading.Event()

        def cancellable_dispatch(self, task, on_progress, cancel_flag, on_result=None):
            for _ in range(200):  # spin up to ~10s waiting for the dismiss
                if cancel_flag():
                    saw_cancel.set()
                    return []
                time.sleep(0.05)
            return []

        monkeypatch.setattr(
            screenspace.ScreenspaceWorker, "_dispatch", cancellable_dispatch
        )

        worker = screenspace.ScreenspaceWorker()
        worker.start()
        try:
            task = screenspace.create_task(
                "color",
                "P01",
                "s.mp4",
                ["/v.mp4"],
                "r",
                {"x": 0, "y": 0, "w": 1, "h": 1},
            )
            worker.enqueue(task)
            # Wait for it to start running.
            for _ in range(50):
                t = worker.get_task(task["id"])
                if t and t["status"] == screenspace.TASK_STATUS_RUNNING:
                    break
                time.sleep(0.05)

            assert worker.remove_task(task["id"]) is True
            # Gone from the UI immediately...
            assert worker.get_all_tasks() == []
            # ...the running scan's cancel_flag fires (the core regression)...
            assert saw_cancel.wait(timeout=5)
            # ...and the task is fully evicted once the scan unwinds.
            for _ in range(50):
                if worker.get_task(task["id"]) is None:
                    break
                time.sleep(0.05)
            assert worker.get_task(task["id"]) is None
        finally:
            worker.stop()


class TestOcrReaderLock:
    def test_ocr_lock_prevents_duplicate_creation(self, monkeypatch):
        """Only one EasyOCR Reader is created even with concurrent calls."""
        call_count = {"n": 0}
        fake_reader = mock.MagicMock()

        def fake_reader_init(languages, verbose=False):
            call_count["n"] += 1
            time.sleep(0.05)  # simulate slow init
            return fake_reader

        # Clear cache
        screenspace._ocr_readers.clear()

        mock_easyocr = mock.MagicMock()
        mock_easyocr.Reader = fake_reader_init
        monkeypatch.setitem(__import__("sys").modules, "easyocr", mock_easyocr)

        threads = []
        for _ in range(4):
            t = threading.Thread(target=screenspace._get_ocr_reader, args=(["en"],))
            threads.append(t)
            t.start()
        for t in threads:
            t.join()

        assert call_count["n"] == 1
        screenspace._ocr_readers.clear()


# ---------------------------------------------------------------------------
# Timelapse bug fixes
# ---------------------------------------------------------------------------
