"""Tests for Screenspace analysis primitives, manifest I/O, and worker."""

import json
import time

import numpy as np
import pytest

import config
import screenspace


# ---------------------------------------------------------------------------
# Analysis primitives
# ---------------------------------------------------------------------------


class TestExtractRegion:
    def test_basic_crop(self):
        frame = np.zeros((100, 200, 3), dtype=np.uint8)
        frame[10:30, 20:70] = 128
        region = {"x": 20, "y": 10, "w": 50, "h": 20}
        cropped = screenspace.extract_region(frame, region)
        assert cropped.shape == (20, 50, 3)
        assert np.all(cropped == 128)

    def test_clamps_to_bounds(self):
        frame = np.zeros((50, 50, 3), dtype=np.uint8)
        region = {"x": 40, "y": 40, "w": 100, "h": 100}
        cropped = screenspace.extract_region(frame, region)
        assert cropped.shape[0] <= 50
        assert cropped.shape[1] <= 50

    def test_region_at_origin(self):
        frame = np.ones((100, 100, 3), dtype=np.uint8) * 42
        region = {"x": 0, "y": 0, "w": 10, "h": 10}
        cropped = screenspace.extract_region(frame, region)
        assert cropped.shape == (10, 10, 3)
        assert np.all(cropped == 42)


class TestAverageColorHsv:
    def test_returns_correct_keys(self):
        region = np.full((10, 10, 3), 128, dtype=np.uint8)
        result = screenspace.average_color_hsv(region)
        assert "h" in result and "s" in result and "v" in result

    def test_pure_blue(self):
        # BGR pure blue: (255, 0, 0) -> HSV ~(120, 255, 255)
        region = np.full((10, 10, 3), [255, 0, 0], dtype=np.uint8)
        result = screenspace.average_color_hsv(region)
        assert abs(result["h"] - 120) < 2
        assert result["s"] > 250
        assert result["v"] > 250


class TestColorMatches:
    def test_exact_match(self):
        region = np.full((10, 10, 3), [255, 0, 0], dtype=np.uint8)
        hsv = screenspace.average_color_hsv(region)
        assert screenspace.color_matches(region, hsv, {"h": 5, "s": 10, "v": 10})

    def test_mismatch(self):
        blue = np.full((10, 10, 3), [255, 0, 0], dtype=np.uint8)
        red_target = {"h": 0.0, "s": 255.0, "v": 255.0}
        assert not screenspace.color_matches(
            blue, red_target, {"h": 5, "s": 10, "v": 10}
        )

    def test_hue_wraparound(self):
        # BGR red: (0, 0, 255) -> HSV ~(0, 255, 255)
        region = np.full((10, 10, 3), [0, 0, 255], dtype=np.uint8)
        # Target near the wrap boundary
        target = {"h": 175.0, "s": 255.0, "v": 255.0}
        assert screenspace.color_matches(region, target, {"h": 10, "s": 10, "v": 10})


class TestComputeFrameDiff:
    def test_identical_frames(self):
        frame = np.random.randint(0, 255, (50, 50, 3), dtype=np.uint8)
        diff = screenspace.compute_frame_diff(frame, frame.copy())
        assert diff == 0.0

    def test_completely_different(self):
        black = np.zeros((50, 50, 3), dtype=np.uint8)
        white = np.full((50, 50, 3), 255, dtype=np.uint8)
        diff = screenspace.compute_frame_diff(black, white)
        assert diff > 0.9


class TestRegionsAreSimilar:
    def test_identical(self):
        frame = np.random.randint(50, 200, (50, 50, 3), dtype=np.uint8)
        is_similar, score = screenspace.regions_are_similar(frame, frame.copy())
        assert is_similar is True
        assert score >= 0.99

    def test_different(self):
        a = np.zeros((50, 50, 3), dtype=np.uint8)
        b = np.full((50, 50, 3), 255, dtype=np.uint8)
        is_similar, score = screenspace.regions_are_similar(a, b)
        assert is_similar is False
        assert score < 0.5


class TestComputePhash:
    def test_deterministic(self):
        region = np.random.randint(0, 255, (64, 64, 3), dtype=np.uint8)
        hash1 = screenspace.compute_phash(region)
        hash2 = screenspace.compute_phash(region.copy())
        assert hash1 == hash2

    def test_different_images(self):
        a = np.zeros((64, 64, 3), dtype=np.uint8)
        b = np.full((64, 64, 3), 255, dtype=np.uint8)
        hash_a = screenspace.compute_phash(a)
        hash_b = screenspace.compute_phash(b)
        assert hash_a != hash_b


class TestBuildTimelapseCommand:
    def test_mp4_output(self):
        cmd = screenspace.build_timelapse_command(
            "/video.mp4", {"x": 10, "y": 20, "w": 300, "h": 100}, 10.0, "/out.mp4"
        )
        assert "ffmpeg" in cmd[0]
        assert "/video.mp4" in cmd
        assert "/out.mp4" in cmd
        vf_idx = cmd.index("-vf")
        assert "crop=300:100:10:20" in cmd[vf_idx + 1]
        assert "setpts=PTS/10.0" in cmd[vf_idx + 1]

    def test_gif_output(self):
        cmd = screenspace.build_timelapse_command(
            "/v.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            5.0,
            "/out.gif",
            "gif",
        )
        assert "-loop" in cmd


class TestMergeTimestampSpans:
    def test_empty(self):
        assert screenspace._merge_timestamp_spans([], 1.0) == []

    def test_single(self):
        spans = screenspace._merge_timestamp_spans([5.0], 1.0)
        assert len(spans) == 1
        assert spans[0]["start"] == 5.0

    def test_consecutive_merged(self):
        spans = screenspace._merge_timestamp_spans([1.0, 2.0, 3.0, 10.0], 1.0)
        assert len(spans) == 2
        assert spans[0]["start"] == 1.0
        assert spans[0]["end"] == 3.0
        assert spans[1]["start"] == 10.0


# ---------------------------------------------------------------------------
# Manifest I/O
# ---------------------------------------------------------------------------


class TestManifest:
    def test_empty_manifest_structure(self):
        m = screenspace._empty_screenspace_manifest()
        assert m == {"regions": {}, "tasks": []}

    def test_roundtrip(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        regions = {
            "healthbar": {
                "x": 10,
                "y": 20,
                "w": 100,
                "h": 30,
                "description": "Health",
            }
        }
        tasks = [
            {
                "id": "ss_abcd1234",
                "type": "color",
                "participant": "P01",
                "status": "completed",
                "result": [{"start": 10.0, "end": 15.0}],
                "_cancelled": False,
            }
        ]
        path = screenspace.save_screenspace_manifest(regions, tasks)
        assert path is not None
        assert path.is_file()

        loaded = screenspace.load_screenspace_manifest()
        assert loaded["regions"]["healthbar"]["x"] == 10
        assert len(loaded["tasks"]) == 1
        assert "_cancelled" not in loaded["tasks"][0]

    def test_load_missing_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        result = screenspace.load_screenspace_manifest()
        assert result == {"regions": {}, "tasks": []}

    def test_load_malformed_json(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        manifest_path = tmp_path / config.SCREENSPACE_MANIFEST_FILENAME
        manifest_path.write_text("not json")
        result = screenspace.load_screenspace_manifest()
        assert result == {"regions": {}, "tasks": []}

    def test_load_non_dict(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        manifest_path = tmp_path / config.SCREENSPACE_MANIFEST_FILENAME
        manifest_path.write_text(json.dumps([1, 2, 3]))
        result = screenspace.load_screenspace_manifest()
        assert result == {"regions": {}, "tasks": []}


# ---------------------------------------------------------------------------
# Task queue and worker
# ---------------------------------------------------------------------------


class TestCreateTask:
    def test_creates_valid_task(self):
        task = screenspace.create_task(
            task_type="color",
            participant="P01",
            source_video="study_P01.mp4",
            video_path="/path/study_P01.mp4",
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
            "color", "P01", "s_P01.mp4", "/v.mp4", "hb", {"x": 0, "y": 0, "w": 10, "h": 10}
        )
        tid = worker.enqueue(task)
        assert tid == task["id"]
        retrieved = worker.get_task(tid)
        assert retrieved is not None
        assert retrieved["id"] == tid

    def test_get_all_tasks(self):
        worker = screenspace.ScreenspaceWorker()
        t1 = screenspace.create_task(
            "color", "P01", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        t2 = screenspace.create_task(
            "change", "P02", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
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
            "color", "P01", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        worker.enqueue(task)
        assert worker.cancel(task["id"]) is True
        t = worker.get_task(task["id"])
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
                "/nonexistent/nope.mp4",
                "r",
                {"x": 0, "y": 0, "w": 10, "h": 10},
                parameters={"target_color": {"h": 0, "s": 0, "v": 0}, "tolerance": {"h": 10, "s": 10, "v": 10}},
            )
            worker.enqueue(task)
            for _ in range(50):
                t = worker.get_task(task["id"])
                if t and t["status"] in ("completed", "failed"):
                    break
                time.sleep(0.1)
            t = worker.get_task(task["id"])
            assert t["status"] in ("completed", "failed")
        finally:
            worker.stop()

    def test_text_task_easyocr_importable(self):
        import easyocr  # noqa: F401

    def test_reorder(self):
        worker = screenspace.ScreenspaceWorker()
        t1 = screenspace.create_task(
            "color", "P01", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        t2 = screenspace.create_task(
            "change", "P01", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        worker.enqueue(t1)
        worker.enqueue(t2)
        assert worker.reorder([t2["id"], t1["id"]]) is True
        got = worker.get_task(t2["id"])
        assert got["priority"] == 1

    def test_get_task_returns_none_for_unknown(self):
        worker = screenspace.ScreenspaceWorker()
        assert worker.get_task("ss_unknown") is None

    def test_remove_queued_task(self):
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "color", "P01", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        worker.enqueue(task)
        assert worker.remove_task(task["id"]) is True
        assert worker.get_task(task["id"]) is None

    def test_remove_nonexistent(self):
        worker = screenspace.ScreenspaceWorker()
        assert worker.remove_task("ss_nonexist") is False

    def test_remove_running_task_sets_cancelled(self):
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "color", "P01", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
        )
        worker.enqueue(task)
        # Simulate running status
        with worker._lock:
            worker._tasks[task["id"]]["status"] = screenspace.TASK_STATUS_RUNNING
        assert worker.remove_task(task["id"]) is True
        assert worker.get_task(task["id"]) is None


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
            "color", "P01", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
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
            "/v.mp4",
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
        assert t["status"] == "queued"
        assert t.get("parameters", {}).get("start_seconds") == 50.0


class TestScanNumbers:
    def test_unknown_operator_raises(self):
        with pytest.raises(ValueError, match="Unknown.*operator"):
            screenspace.scan_numbers(
                "/nonexistent.mp4",
                {"x": 0, "y": 0, "w": 10, "h": 10},
                operator="invalid",
                target_value=100,
            )

    def test_valid_operators_accepted(self):
        for op in ("eq", "gt", "lt", "gte", "lte", "range"):
            # Should not raise ValueError -- returns [] because video doesn't exist
            result = screenspace.scan_numbers(
                "/nonexistent.mp4",
                {"x": 0, "y": 0, "w": 10, "h": 10},
                operator=op,
                target_value=100,
                range_min=0,
                range_max=200,
            )
            assert result == []

    def test_dispatch_routes_numbers(self):
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "numbers",
            "P01",
            "s_P01.mp4",
            "/nonexistent.mp4",
            "r",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            parameters={"operator": "gt", "target_value": 50},
        )
        # Should not raise -- dispatches to scan_numbers, returns [] for bad video
        result = worker._dispatch(task, lambda p: None, lambda: False)
        assert result == []

    def test_dispatch_unknown_type_raises(self):
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "bogus",
            "P01",
            "s.mp4",
            "/v.mp4",
            "r",
            {"x": 0, "y": 0, "w": 1, "h": 1},
        )
        with pytest.raises(ValueError, match="Unknown task type"):
            worker._dispatch(task, lambda p: None, lambda: False)
