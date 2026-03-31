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


class TestMatchTemplate:
    def test_exact_match(self):
        # Use a textured pattern so template matching works after blur
        rng = np.random.RandomState(42)
        frame = rng.randint(0, 255, (100, 200, 3), dtype=np.uint8)
        template = frame[30:60, 80:140].copy()
        results = screenspace.match_template(frame, template, threshold=0.9)
        assert len(results) >= 1
        assert results[0]["score"] >= 0.9

    def test_no_match(self):
        frame = np.zeros((100, 200, 3), dtype=np.uint8)
        template = np.full((20, 40, 3), 128, dtype=np.uint8)
        results = screenspace.match_template(frame, template, threshold=0.9)
        assert len(results) == 0

    def test_template_larger_than_frame(self):
        frame = np.zeros((20, 20, 3), dtype=np.uint8)
        template = np.zeros((50, 50, 3), dtype=np.uint8)
        results = screenspace.match_template(frame, template, threshold=0.5)
        assert results == []

    def test_match_with_mask(self):
        rng = np.random.RandomState(42)
        frame = rng.randint(0, 255, (100, 200, 3), dtype=np.uint8)
        template = frame[30:60, 80:140].copy()
        # Full opaque mask — should behave like no mask
        mask = np.full((30, 60), 255, dtype=np.uint8)
        results = screenspace.match_template(frame, template, threshold=0.9, mask=mask)
        assert len(results) >= 1

    def test_match_with_none_mask(self):
        rng = np.random.RandomState(42)
        frame = rng.randint(0, 255, (100, 200, 3), dtype=np.uint8)
        template = frame[30:60, 80:140].copy()
        results = screenspace.match_template(frame, template, threshold=0.9, mask=None)
        assert len(results) >= 1


class TestComputeOpticalFlow:
    def test_no_motion(self):
        gray = np.full((50, 50), 128, dtype=np.uint8)
        result = screenspace.compute_optical_flow(gray, gray.copy())
        assert result["magnitude"] < 0.5
        assert "angle" in result

    def test_motion_detected(self):
        prev = np.zeros((80, 80), dtype=np.uint8)
        curr = np.zeros((80, 80), dtype=np.uint8)
        prev[20:40, 20:40] = 255
        curr[30:50, 30:50] = 255
        result = screenspace.compute_optical_flow(prev, curr)
        assert result["magnitude"] > 0


class TestSceneFingerprint:
    def test_same_frame_similar(self):
        frame = np.random.randint(50, 200, (50, 50, 3), dtype=np.uint8)
        fp1 = screenspace.compute_scene_fingerprint(frame)
        fp2 = screenspace.compute_scene_fingerprint(frame.copy())
        score = screenspace.compare_scene_fingerprints(fp1, fp2)
        assert score >= 0.99

    def test_different_frames_dissimilar(self):
        a = np.zeros((50, 50, 3), dtype=np.uint8)
        b = np.full((50, 50, 3), 255, dtype=np.uint8)
        fp_a = screenspace.compute_scene_fingerprint(a)
        fp_b = screenspace.compute_scene_fingerprint(b)
        score = screenspace.compare_scene_fingerprints(fp_a, fp_b)
        assert score < 0.8

    def test_fingerprint_has_expected_keys(self):
        frame = np.random.randint(0, 255, (30, 30, 3), dtype=np.uint8)
        fp = screenspace.compute_scene_fingerprint(frame)
        assert "histogram" in fp
        assert "edge_density" in fp
        assert "color_stats" in fp


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
        assert m == {"regions": {}, "tasks": [], "events": [], "stashes": []}

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
        assert result == {"regions": {}, "tasks": [], "events": [], "stashes": []}

    def test_load_malformed_json(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        manifest_path = tmp_path / config.SCREENSPACE_MANIFEST_FILENAME
        manifest_path.write_text("not json")
        result = screenspace.load_screenspace_manifest()
        assert result == {"regions": {}, "tasks": [], "events": [], "stashes": []}

    def test_load_non_dict(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        manifest_path = tmp_path / config.SCREENSPACE_MANIFEST_FILENAME
        manifest_path.write_text(json.dumps([1, 2, 3]))
        result = screenspace.load_screenspace_manifest()
        assert result == {"regions": {}, "tasks": [], "events": [], "stashes": []}


# ---------------------------------------------------------------------------
# Events
# ---------------------------------------------------------------------------


class TestCreateEvent:
    def _make_task(self, **overrides):
        base = {
            "id": "ss_abcd1234",
            "type": "change",
            "source_video": "study_P01.mp4",
            "participant": "P01",
            "region": "healthbar",
            "parameters": {},
        }
        base.update(overrides)
        return base

    def test_basic_event_structure(self):
        task = self._make_task()
        ev = screenspace.create_event(task, 50.0, 0.85, {"magnitude": 0.12})
        assert ev["id"].startswith("ev_")
        assert ev["source_video"] == "study_P01.mp4"
        assert ev["participant"] == "P01"
        assert ev["detector"] == "change"
        assert ev["time_in"] == 50.0
        assert ev["time_out"] == 50.0
        assert ev["confidence"] == 0.85
        assert ev["metadata"] == {"magnitude": 0.12}
        assert ev["excluded"] is False
        assert ev["task_id"] == "ss_abcd1234"
        assert ev["region"] == "healthbar"

    def test_default_event_type_fallback(self):
        task = self._make_task()
        ev = screenspace.create_event(task, 10.0, 0.5)
        assert ev["event_type"] == "change: healthbar"

    def test_custom_event_label(self):
        task = self._make_task(parameters={"event_label": "low_health"})
        ev = screenspace.create_event(task, 10.0, 0.5)
        assert ev["event_type"] == "low_health"

    def test_confidence_clamping(self):
        task = self._make_task()
        ev_high = screenspace.create_event(task, 10.0, 1.5)
        assert ev_high["confidence"] == 1.0
        ev_low = screenspace.create_event(task, 10.0, -0.5)
        assert ev_low["confidence"] == 0.0

    def test_empty_metadata_default(self):
        task = self._make_task()
        ev = screenspace.create_event(task, 10.0, 0.5)
        assert ev["metadata"] == {}


class TestGenerateEventsFromResults:
    def _make_worker_and_task(self, task_type, **params):
        worker = screenspace.ScreenspaceWorker()
        task = {
            "id": "ss_test1234",
            "type": task_type,
            "source_video": "study_P01.mp4",
            "participant": "P01",
            "region": "hud",
            "parameters": params,
        }
        return worker, task

    def test_change_events(self):
        worker, task = self._make_worker_and_task("change")
        raw = [
            {"timestamp": 10.0, "magnitude": 0.15},
            {"timestamp": 20.0, "magnitude": 0.8},
        ]
        events = worker._generate_events_from_results(task, raw)
        assert len(events) == 2
        assert events[0]["confidence"] == 0.15
        assert events[0]["metadata"]["magnitude"] == 0.15
        assert events[1]["confidence"] == 0.8

    def test_similarity_events(self):
        worker, task = self._make_worker_and_task("similarity")
        raw = [{"timestamp": 5.0, "score": 0.95}]
        events = worker._generate_events_from_results(task, raw)
        assert len(events) == 1
        assert events[0]["confidence"] == 0.95
        assert events[0]["metadata"]["score"] == 0.95

    def test_text_events(self):
        worker, task = self._make_worker_and_task("text")
        raw = [{"timestamp": 30.0, "text_found": "hello", "confidence": 0.9}]
        events = worker._generate_events_from_results(task, raw)
        assert len(events) == 1
        assert events[0]["metadata"]["text_found"] == "hello"
        assert events[0]["confidence"] == 0.9

    def test_numbers_events(self):
        worker, task = self._make_worker_and_task("numbers")
        raw = [{"timestamp": 15.0, "number_found": 42}]
        events = worker._generate_events_from_results(task, raw)
        assert len(events) == 1
        assert events[0]["confidence"] == 1.0
        assert events[0]["metadata"]["value"] == 42

    def test_color_events(self):
        worker, task = self._make_worker_and_task("color")
        raw = [{"timestamp": 5.0, "_confidence": 0.7}]
        events = worker._generate_events_from_results(task, raw)
        assert len(events) == 1
        assert events[0]["confidence"] == 0.7

    def test_timelapse_no_events(self):
        worker, task = self._make_worker_and_task("timelapse")
        events = worker._generate_events_from_results(task, [{"file": "out.mp4"}])
        assert events == []

    def test_template_events(self):
        worker, task = self._make_worker_and_task("template")
        raw = [{"timestamp": 5.0, "best_score": 0.85, "match_count": 2}]
        events = worker._generate_events_from_results(task, raw)
        assert len(events) == 1
        assert events[0]["confidence"] == 0.85
        assert events[0]["metadata"]["match_count"] == 2
        assert events[0]["metadata"]["best_score"] == 0.85

    def test_flow_events(self):
        worker, task = self._make_worker_and_task("flow")
        raw = [{"timestamp": 10.0, "magnitude": 5.0, "angle": 90.0}]
        events = worker._generate_events_from_results(task, raw)
        assert len(events) == 1
        assert events[0]["confidence"] == 0.5  # 5.0 / 10.0
        assert events[0]["metadata"]["magnitude"] == 5.0
        assert events[0]["metadata"]["angle"] == 90.0

    def test_scene_events(self):
        worker, task = self._make_worker_and_task("scene")
        raw = [{"timestamp": 15.0, "scene_name": "menu", "score": 0.92}]
        events = worker._generate_events_from_results(task, raw)
        assert len(events) == 1
        assert events[0]["confidence"] == 0.92
        assert events[0]["metadata"]["scene_name"] == "menu"
        assert events[0]["metadata"]["score"] == 0.92


class TestManifestWithEvents:
    def test_roundtrip_with_events(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        events = [
            {
                "id": "ev_test1234",
                "source_video": "study_P01.mp4",
                "participant": "P01",
                "detector": "change",
                "event_type": "change: hud",
                "time_in": 10.0,
                "time_out": 10.0,
                "confidence": 0.85,
                "metadata": {"magnitude": 0.12},
                "excluded": False,
                "task_id": "ss_abcd1234",
                "region": "hud",
            }
        ]
        path = screenspace.save_screenspace_manifest({}, [], events)
        assert path is not None

        loaded = screenspace.load_screenspace_manifest()
        assert len(loaded["events"]) == 1
        assert loaded["events"][0]["id"] == "ev_test1234"
        assert loaded["events"][0]["confidence"] == 0.85

    def test_empty_events_default(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        screenspace.save_screenspace_manifest({}, [])
        loaded = screenspace.load_screenspace_manifest()
        assert loaded["events"] == []


class TestDrainNewEvents:
    def test_drain_collects_and_clears(self):
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            task_type="change",
            participant="P01",
            source_video="study_P01.mp4",
            video_path="/path/study_P01.mp4",
            region_name="hud",
            region_coords={"x": 0, "y": 0, "w": 100, "h": 50},
        )
        worker.enqueue(task)
        # Manually inject generated events for testing
        worker._tasks[task["id"]]["_generated_events"] = [
            {"id": "ev_1", "detector": "change"},
            {"id": "ev_2", "detector": "change"},
        ]
        events = worker.drain_new_events()
        assert len(events) == 2
        # Second drain should be empty
        assert worker.drain_new_events() == []


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
            "color",
            "P01",
            "s_P01.mp4",
            "/v.mp4",
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
                "/nonexistent/nope.mp4",
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
                time.sleep(0.1)
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
                "/nonexistent/v.mp4",
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
                time.sleep(0.1)

            # Second task should still process even though first callback raised
            t2 = screenspace.create_task(
                "color",
                "P02",
                "s.mp4",
                "/nonexistent/v.mp4",
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
                time.sleep(0.1)
            t = worker.get_task(t2["id"])
            assert t is not None
            assert t["status"] in ("completed", "failed")
            assert call_count["n"] >= 2
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
        assert got is not None
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
        assert t is not None
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


# ---------------------------------------------------------------------------
# Flow grid and heatmap visualization
# ---------------------------------------------------------------------------


class TestOpticalFlowGrid:
    def test_no_grid_by_default(self):
        gray = np.full((50, 50), 128, dtype=np.uint8)
        result = screenspace.compute_optical_flow(gray, gray.copy())
        assert "flow_grid" not in result

    def test_grid_returned_when_requested(self):
        prev = np.zeros((80, 80), dtype=np.uint8)
        curr = np.zeros((80, 80), dtype=np.uint8)
        prev[20:40, 20:40] = 255
        curr[30:50, 30:50] = 255
        result = screenspace.compute_optical_flow(prev, curr, return_grid=True)
        assert "flow_grid" in result
        assert isinstance(result["flow_grid"], list)

    def test_grid_entries_have_expected_keys(self):
        prev = np.zeros((80, 80), dtype=np.uint8)
        curr = np.zeros((80, 80), dtype=np.uint8)
        prev[20:40, 20:40] = 255
        curr[30:50, 30:50] = 255
        result = screenspace.compute_optical_flow(prev, curr, return_grid=True)
        grid = result["flow_grid"]
        if grid:
            cell = grid[0]
            assert "x" in cell and "y" in cell
            assert "mag" in cell and "ang" in cell
            assert 0 <= cell["x"] <= 1
            assert 0 <= cell["y"] <= 1

    def test_static_frames_empty_grid(self):
        gray = np.full((50, 50), 128, dtype=np.uint8)
        result = screenspace.compute_optical_flow(gray, gray.copy(), return_grid=True)
        assert result["flow_grid"] == []


class TestGenerateTemplateHeatmap:
    def test_basic_heatmap(self, tmp_path):
        results = [
            {
                "timestamp": 1.0,
                "matches": [{"x": 10, "y": 10, "w": 50, "h": 50, "score": 0.9}],
            },
            {
                "timestamp": 2.0,
                "matches": [{"x": 20, "y": 20, "w": 50, "h": 50, "score": 0.8}],
            },
        ]
        out = str(tmp_path / "heatmap.png")
        path = screenspace.generate_template_heatmap(results, 200, 200, out)
        assert path == out
        assert (tmp_path / "heatmap.png").is_file()
        assert (tmp_path / "heatmap.png").stat().st_size > 0

    def test_empty_results_returns_none(self, tmp_path):
        out = str(tmp_path / "heatmap.png")
        assert screenspace.generate_template_heatmap([], 200, 200, out) is None

    def test_no_matches_returns_none(self, tmp_path):
        results = [{"timestamp": 1.0, "matches": []}]
        out = str(tmp_path / "heatmap.png")
        assert screenspace.generate_template_heatmap(results, 200, 200, out) is None


class TestGenerateFlowHeatmap:
    def test_basic_heatmap(self, tmp_path):
        results = [
            {
                "timestamp": 1.0,
                "flow_grid": [
                    {"x": 0.5, "y": 0.5, "mag": 5.0, "ang": 90.0},
                    {"x": 0.2, "y": 0.2, "mag": 3.0, "ang": 180.0},
                ],
            }
        ]
        out = str(tmp_path / "flow_heatmap.png")
        path = screenspace.generate_flow_heatmap(results, 200, 200, out)
        assert path == out
        assert (tmp_path / "flow_heatmap.png").is_file()
        assert (tmp_path / "flow_heatmap.png").stat().st_size > 0

    def test_empty_grid_returns_none(self, tmp_path):
        results = [{"timestamp": 1.0, "flow_grid": []}]
        out = str(tmp_path / "heatmap.png")
        assert screenspace.generate_flow_heatmap(results, 200, 200, out) is None


class TestFlowGridStrippedFromManifest:
    def test_flow_grid_not_persisted(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        tasks = [
            {
                "id": "ss_flow1234",
                "type": "flow",
                "participant": "P01",
                "status": "completed",
                "result": [
                    {
                        "timestamp": 1.0,
                        "magnitude": 5.0,
                        "angle": 90.0,
                        "flow_grid": [{"x": 0.5, "y": 0.5, "mag": 5.0, "ang": 90.0}],
                    }
                ],
            }
        ]
        path = screenspace.save_screenspace_manifest({}, tasks)
        assert path is not None
        loaded = screenspace.load_screenspace_manifest()
        result = loaded["tasks"][0]["result"][0]
        assert "flow_grid" not in result
        assert result["magnitude"] == 5.0
