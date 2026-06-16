"""Tests for Screenspace analysis primitives, manifest I/O, and worker."""

import io
import threading
import time
from pathlib import Path
from unittest import mock

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
        matched, conf = screenspace.color_matches(
            region, hsv, {"h": 5, "s": 10, "v": 10}
        )
        assert matched
        assert conf > 0.0

    def test_mismatch(self):
        blue = np.full((10, 10, 3), [255, 0, 0], dtype=np.uint8)
        red_target = {"h": 0.0, "s": 255.0, "v": 255.0}
        matched, _conf = screenspace.color_matches(
            blue, red_target, {"h": 5, "s": 10, "v": 10}
        )
        assert not matched

    def test_hue_wraparound(self):
        # BGR red: (0, 0, 255) -> HSV ~(0, 255, 255)
        region = np.full((10, 10, 3), [0, 0, 255], dtype=np.uint8)
        # Target near the wrap boundary
        target = {"h": 175.0, "s": 255.0, "v": 255.0}
        matched, _conf = screenspace.color_matches(
            region, target, {"h": 10, "s": 10, "v": 10}
        )
        assert matched


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


class TestPrepareTemplateMask:
    def test_binarizes_alpha_mask(self):
        """Mask should come out as strictly 0 or 255 (no soft-blurred edges)."""
        template = np.full((30, 60, 3), 128, dtype=np.uint8)
        # Alpha ramp from 0..255 to exercise the boundary
        mask = np.zeros((30, 60), dtype=np.uint8)
        for c in range(60):
            mask[:, c] = int(c * 255 / 59)
        _, gray_mask, _ = screenspace._prepare_template(template, mask)
        assert gray_mask is not None
        unique = set(np.unique(gray_mask).tolist())
        assert unique.issubset({0, 255})

    def test_none_mask_stays_none(self):
        template = np.full((30, 60, 3), 128, dtype=np.uint8)
        _, gray_mask, _ = screenspace._prepare_template(template, None)
        assert gray_mask is None


_ICON_BG = 50  # flat gray background used by the icon-frame fixtures


def _make_icon(size: int, seed: int = 7) -> np.ndarray:
    """Build a square icon with a textured center and a flat border.

    The border matches the fixture background colour so that Gaussian blur
    at the icon boundary does not distort cv2.matchTemplate correlation.
    """
    import cv2

    rng = np.random.RandomState(seed)
    base = 40
    icon = np.full((base, base, 3), _ICON_BG, dtype=np.uint8)
    # Leave a 5px flat border; fill the center with high-contrast texture
    icon[5:-5, 5:-5] = rng.randint(150, 255, (base - 10, base - 10, 3), dtype=np.uint8)
    if size == base:
        return icon
    return cv2.resize(icon, (size, size), interpolation=cv2.INTER_AREA)


def _make_icon_frame(
    frame_w: int,
    frame_h: int,
    icon_positions: list[tuple[int, int, int]],
    seed: int = 7,
) -> np.ndarray:
    """Build a frame with identical icons (possibly at different sizes) placed
    at *icon_positions* (list of ``(x, y, size)``)."""
    frame = np.full((frame_h, frame_w, 3), _ICON_BG, dtype=np.uint8)
    for x, y, s in icon_positions:
        frame[y : y + s, x : x + s] = _make_icon(s, seed=seed)
    return frame


class TestScanTemplateControls:
    """Cover template_scale addition to scan_template."""

    def _patch_single_frame(self, monkeypatch, frame: np.ndarray) -> None:
        def fake_scan(video_path, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace, "scan_video_full_frames", fake_scan)
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))

    def test_scale_fixes_size_mismatch(self, monkeypatch):
        """A 40px template should miss a 20px in-frame icon at scale 1.0
        but hit at scale 0.5."""
        frame = _make_icon_frame(400, 200, [(100, 50, 20)])
        # Template at the original (larger) size — mimics an uploaded PNG
        # captured at 2x the in-video rendering.
        template = _make_icon(40)
        self._patch_single_frame(monkeypatch, frame)

        full_size = screenspace.scan_template(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 400, "h": 200},
            template,
            threshold=0.70,
            template_scale=1.0,
        )
        assert full_size == []

        scaled = screenspace.scan_template(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 400, "h": 200},
            template,
            threshold=0.70,
            template_scale=0.5,
        )
        assert len(scaled) == 1
        match = scaled[0]["matches"][0]
        assert abs(match["x"] - 100) <= 2
        assert abs(match["y"] - 50) <= 2

    def test_transparent_mask_no_false_positives_on_blank_frame(self, monkeypatch):
        """Mostly-transparent PNG + blank frame should yield no matches at
        the default threshold now that the mask is binarized."""
        template = np.full((32, 32, 3), 220, dtype=np.uint8)
        mask = np.zeros((32, 32), dtype=np.uint8)
        # Only a small opaque cross in the center
        mask[14:18, :] = 255
        mask[:, 14:18] = 255
        blank = np.full((200, 300, 3), 30, dtype=np.uint8)
        self._patch_single_frame(monkeypatch, blank)

        results = screenspace.scan_template(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 300, "h": 200},
            template,
            threshold=0.70,
            template_mask=mask,
        )
        assert results == []

    def test_scaled_masked_template_no_explosion(self, monkeypatch):
        """Regression: masked matching at non-1.0 scale must not report
        thousands of matches. TM_CCOEFF_NORMED with sparse masks at
        reduced scale previously produced near-1.0 scores at every
        position."""
        # Mostly-transparent 50x50 PNG with a small opaque central cross.
        template = np.full((50, 50, 3), 220, dtype=np.uint8)
        mask = np.zeros((50, 50), dtype=np.uint8)
        mask[22:28, :] = 255
        mask[:, 22:28] = 255
        # Random frame, not containing the template.
        rng = np.random.RandomState(3)
        frame = rng.randint(0, 255, (360, 640, 3), dtype=np.uint8)
        self._patch_single_frame(monkeypatch, frame)

        results = screenspace.scan_template(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 640, "h": 360},
            template,
            threshold=0.70,
            template_mask=mask,
            template_scale=0.75,
        )
        total = sum(r["match_count"] for r in results)
        assert total < 50, f"Expected few/no matches, got {total}"


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
        assert m == {
            "regions": {},
            "tasks": [],
            "events": [],
            "stashes": [],
            "per_participant": {},
            "pins": {},
        }

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
        assert result == {
            "regions": {},
            "tasks": [],
            "events": [],
            "stashes": [],
            "per_participant": {},
            "pins": {},
        }

    def test_load_malformed_json(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        manifest_path = tmp_path / config.SCREENSPACE_MANIFEST_FILENAME
        manifest_path.write_text("not json")
        result = screenspace.load_screenspace_manifest()
        assert result == {
            "regions": {},
            "tasks": [],
            "events": [],
            "stashes": [],
            "per_participant": {},
            "pins": {},
        }

    def test_pins_roundtrip(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        pins = {
            "P01": [
                {
                    "id": "pin_abcd1234",
                    "timestamp": 12.5,
                    "polarity": "positive",
                    "label": "health red",
                    "created_at": "2026-01-01T00:00:00+00:00",
                }
            ]
        }
        path = screenspace.save_screenspace_manifest({}, [], pins=pins)
        assert path is not None
        loaded = screenspace.load_screenspace_manifest()
        assert loaded["pins"]["P01"][0]["polarity"] == "positive"
        assert loaded["pins"]["P01"][0]["timestamp"] == 12.5


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
        raw = [{"timestamp": 15.0, "number_found": 42, "confidence": 0.42}]
        events = worker._generate_events_from_results(task, raw)
        assert len(events) == 1
        assert events[0]["confidence"] == 0.42
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

    def test_inactivity_events(self):
        worker, task = self._make_worker_and_task("inactivity")
        raw = [{"start": 5.0, "end": 15.0, "duration": 10.0, "avg_distance": 2.5}]
        events = worker._generate_events_from_results(task, raw)
        assert len(events) == 1
        assert events[0]["time_in"] == 5.0
        assert events[0]["time_out"] == 15.0
        assert events[0]["metadata"]["duration"] == 10.0
        assert events[0]["metadata"]["avg_distance"] == 2.5
        # confidence = min(10.0/30.0, 1.0) ≈ 0.3333
        assert abs(events[0]["confidence"] - 0.3333) < 0.01


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


class TestBackfillMissingEvents:
    def _completed_template_task(self):
        return {
            "id": "ss_template1",
            "type": "template",
            "status": "completed",
            "source_video": "study_P01.mp4",
            "participant": "P01",
            "region": "icon",
            "parameters": {},
            "result": [
                {"timestamp": 5.0, "best_score": 0.9, "match_count": 1},
                {"timestamp": 10.0, "best_score": 0.7, "match_count": 2},
            ],
        }

    def test_backfills_when_no_events(self, tmp_path, monkeypatch):
        import screenspace_server

        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        manifest = {
            "regions": {},
            "tasks": [self._completed_template_task()],
            "events": [],
            "stashes": [],
        }
        screenspace_server._backfill_missing_events(manifest)
        assert len(manifest["events"]) == 2
        assert all(e["task_id"] == "ss_template1" for e in manifest["events"])
        assert manifest["events"][0]["detector"] == "template"

    def test_skips_tasks_with_existing_events(self, tmp_path, monkeypatch):
        import screenspace_server

        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        manifest = {
            "regions": {},
            "tasks": [self._completed_template_task()],
            "events": [{"id": "ev_x", "task_id": "ss_template1"}],
            "stashes": [],
        }
        screenspace_server._backfill_missing_events(manifest)
        assert len(manifest["events"]) == 1

    def test_skips_non_completed_tasks(self, tmp_path, monkeypatch):
        import screenspace_server

        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        task = self._completed_template_task()
        task["status"] = "running"
        manifest = {"regions": {}, "tasks": [task], "events": [], "stashes": []}
        screenspace_server._backfill_missing_events(manifest)
        assert manifest["events"] == []


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
            # on_task_complete fires in the _run loop after the future is
            # collected, which may lag behind the task status change.
            # Wait for the callback to actually fire for task 2.
            for _ in range(20):
                if call_count["n"] >= 2:
                    break
                time.sleep(0.1)
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


class TestOcrPreprocess:
    def test_small_roi_upscaled(self):
        """Crops shorter than the min height are upscaled, preserving aspect."""
        small = np.full((20, 120, 3), 128, dtype=np.uint8)
        out = screenspace._preprocess_for_ocr(small)
        assert out.shape[0] >= 60
        # Aspect ratio preserved (6:1 → width scales with height).
        assert out.shape[1] >= 360

    def test_large_roi_size_unchanged(self):
        """Crops already tall enough are not resized."""
        large = np.full((200, 500, 3), 128, dtype=np.uint8)
        out = screenspace._preprocess_for_ocr(large)
        assert out.shape[0] == 200
        assert out.shape[1] == 500

    def test_clahe_increases_contrast(self):
        """CLAHE stretches a low-contrast crop, raising pixel variance."""
        base = np.random.RandomState(0).randint(100, 116, (60, 120)).astype(np.uint8)
        lowc = np.stack([base, base, base], axis=-1)  # equal channels → true gray
        out = screenspace._preprocess_for_ocr(lowc)
        in_var = float(np.var(lowc[:, :, 0]))
        out_var = float(np.var(out[:, :, 0]))
        assert out_var > in_var

    def test_returns_three_channels(self):
        small = np.full((20, 120, 3), 128, dtype=np.uint8)
        out = screenspace._preprocess_for_ocr(small)
        assert out.ndim == 3 and out.shape[2] == 3

    def test_empty_region_returned_unchanged(self):
        empty = np.zeros((0, 10, 3), dtype=np.uint8)
        out = screenspace._preprocess_for_ocr(empty)
        assert out.shape[0] == 0

    def test_scan_text_preprocess_enlarges_ocr_input(self, monkeypatch):
        """ocr_preprocess=True feeds an upscaled crop to the OCR reader."""
        frame = np.full((20, 120, 3), 128, dtype=np.uint8)
        seen: dict[str, int] = {}

        class _FakeReader:
            def readtext(self, pixels, **_kwargs):
                seen["h"] = pixels.shape[0]
                return []

        monkeypatch.setattr(screenspace, "_get_ocr_reader", lambda _l: _FakeReader())
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))
        monkeypatch.setattr(
            screenspace, "scan_video_frames", lambda v, r, i, cb, **k: cb(0.0, frame)
        )

        screenspace.scan_text(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 120, "h": 20},
            search_string="x",
            ocr_preprocess=True,
        )
        assert seen["h"] >= 60

    def test_scan_text_no_preprocess_keeps_native_size(self, monkeypatch):
        """Default (ocr_preprocess=False) passes the raw crop to the reader."""
        frame = np.full((20, 120, 3), 128, dtype=np.uint8)
        seen: dict[str, int] = {}

        class _FakeReader:
            def readtext(self, pixels, **_kwargs):
                seen["h"] = pixels.shape[0]
                return []

        monkeypatch.setattr(screenspace, "_get_ocr_reader", lambda _l: _FakeReader())
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))
        monkeypatch.setattr(
            screenspace, "scan_video_frames", lambda v, r, i, cb, **k: cb(0.0, frame)
        )

        screenspace.scan_text(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 120, "h": 20},
            search_string="x",
            ocr_preprocess=False,
        )
        assert seen["h"] == 20


class TestScanText:
    def test_low_confidence_rejected(self, monkeypatch):
        """OCR readings below ocr_confidence_threshold should not match."""
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)

        class _FakeReader:
            def readtext(self, _pixels, **_kwargs):
                return [([(0, 0), (10, 0), (10, 10), (0, 10)], "hello", 0.2)]

        monkeypatch.setattr(
            screenspace, "_get_ocr_reader", lambda _langs: _FakeReader()
        )
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)

        rejected = screenspace.scan_text(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 20},
            search_string="hello",
            ocr_confidence_threshold=0.5,
        )
        assert rejected == []

        accepted = screenspace.scan_text(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 20},
            search_string="hello",
            ocr_confidence_threshold=0.1,
        )
        assert len(accepted) == 1
        assert accepted[0]["text_found"] == "hello"
        assert accepted[0]["confidence"] == 0.2

    def test_normalize_letters_to_digits(self, monkeypatch):
        """ocr_normalize="digits" folds l→1 so "l00" matches a search for "100"."""
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)

        class _FakeReader:
            def readtext(self, _pixels, **_kwargs):
                # Misread: the glyphs of "100" came back as "l00".
                return [([(0, 0), (10, 0), (10, 10), (0, 10)], "l00", 0.9)]

        monkeypatch.setattr(
            screenspace, "_get_ocr_reader", lambda _langs: _FakeReader()
        )
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        # Default (ocr_normalize="off"): "l00" vs "100" stays below threshold.
        rejected = screenspace.scan_text(
            "/fake.mp4", region, search_string="100", fuzzy_threshold=0.9
        )
        assert rejected == []

        # ocr_normalize="digits": l→1 makes it an exact match.
        accepted = screenspace.scan_text(
            "/fake.mp4",
            region,
            search_string="100",
            fuzzy_threshold=0.9,
            ocr_normalize="digits",
        )
        assert len(accepted) == 1
        assert accepted[0]["text_found"] == "l00"

    def test_normalize_digits_to_letters(self, monkeypatch):
        """ocr_normalize="letters" folds 5→s so "5top" matches a search for "stop"."""
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)

        class _FakeReader:
            def readtext(self, _pixels, **_kwargs):
                # Misread: the glyphs of "stop" came back as "5top".
                return [([(0, 0), (10, 0), (10, 10), (0, 10)], "5top", 0.9)]

        monkeypatch.setattr(
            screenspace, "_get_ocr_reader", lambda _langs: _FakeReader()
        )
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        # Default ("off"): "5top" vs "stop" stays below threshold.
        assert (
            screenspace.scan_text(
                "/fake.mp4", region, search_string="stop", fuzzy_threshold=0.9
            )
            == []
        )

        # ocr_normalize="letters": 5→s makes it an exact match.
        accepted = screenspace.scan_text(
            "/fake.mp4",
            region,
            search_string="stop",
            fuzzy_threshold=0.9,
            ocr_normalize="letters",
        )
        assert len(accepted) == 1
        assert accepted[0]["text_found"] == "5top"

    def test_normalize_direction_is_distinct(self, monkeypatch):
        """The two fold directions differ: i→1 is digits-only (no 1→i inverse)."""
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)

        class _FakeReader:
            def readtext(self, _pixels, **_kwargs):
                # Want "in"; OCR rendered the i as a 1.
                return [([(0, 0), (10, 0), (10, 10), (0, 10)], "1n", 0.9)]

        monkeypatch.setattr(
            screenspace, "_get_ocr_reader", lambda _langs: _FakeReader()
        )
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        # "digits" folds search i→1, matching OCR "1n"; "letters" folds OCR 1→l
        # ("ln") which stays below threshold against "in".
        assert (
            len(
                screenspace.scan_text(
                    "/fake.mp4",
                    region,
                    search_string="in",
                    fuzzy_threshold=0.9,
                    ocr_normalize="digits",
                )
            )
            == 1
        )
        assert (
            screenspace.scan_text(
                "/fake.mp4",
                region,
                search_string="in",
                fuzzy_threshold=0.9,
                ocr_normalize="letters",
            )
            == []
        )

    def test_require_consecutive(self, monkeypatch):
        """require_consecutive=N coalesces N consecutive matches into one median event."""
        # Distinct fills so the static-frame-skip never fires between frames.
        frames = [np.full((20, 60, 3), v, dtype=np.uint8) for v in (40, 90, 140)]

        class _FakeReader:
            def readtext(self, _pixels, **_kwargs):
                return [([(0, 0), (10, 0), (10, 10), (0, 10)], "hello", 0.9)]

        monkeypatch.setattr(
            screenspace, "_get_ocr_reader", lambda _langs: _FakeReader()
        )
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for i, frame in enumerate(frames):
                callback(float(i), frame)

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        # Default (require_consecutive=1): one event per matching frame.
        default = screenspace.scan_text(
            "/fake.mp4", region, search_string="hello", fuzzy_threshold=0.8
        )
        assert len(default) == 3

        # require_consecutive=3: a single event stamped with the median timestamp.
        coalesced = screenspace.scan_text(
            "/fake.mp4",
            region,
            search_string="hello",
            fuzzy_threshold=0.8,
            require_consecutive=3,
        )
        assert len(coalesced) == 1
        assert coalesced[0]["timestamp"] == 1.0  # median([0.0, 1.0, 2.0])

    def test_require_consecutive_survives_static_frames(self, monkeypatch):
        """Static (skipped) frames carry an active run forward instead of starving it.

        The first frame matches and is OCR'd; the next two are identical to it,
        so the static-frame-skip drops them before OCR. Carry-over keeps the
        require_consecutive=3 run alive, so one event still emits (the old
        behavior would have stalled at one push and emitted nothing)."""
        base = np.full((20, 60, 3), 40, dtype=np.uint8)
        frames = [base, base.copy(), base.copy()]  # identical → static-skip fires

        reads = {"n": 0}

        class _FakeReader:
            def readtext(self_inner, _pixels, **_kwargs):
                reads["n"] += 1
                return [([(0, 0), (10, 0), (10, 10), (0, 10)], "hello", 0.9)]

        monkeypatch.setattr(
            screenspace, "_get_ocr_reader", lambda _langs: _FakeReader()
        )
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for i, frame in enumerate(frames):
                callback(float(i), frame)

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}
        out = screenspace.scan_text(
            "/fake.mp4",
            region,
            search_string="hello",
            fuzzy_threshold=0.8,
            require_consecutive=3,
        )
        assert reads["n"] == 1  # frames 2 & 3 were skipped as static (no OCR)
        assert len(out) == 1
        assert out[0]["timestamp"] == 1.0  # median([0, 1, 2]) across carried frames


class TestConsecutiveBuffer:
    def test_n1_emits_immediately(self):
        buf = screenspace._ConsecutiveBuffer(1)
        out = buf.push(5.0, {"timestamp": 5.0, "magnitude": 0.4})
        assert out is not None
        assert out["timestamp"] == 5.0
        assert out["magnitude"] == 0.4

    def test_emits_after_n_with_median_ts(self):
        buf = screenspace._ConsecutiveBuffer(3)
        assert buf.push(10.0, {"timestamp": 10.0, "v": "a"}) is None
        assert buf.push(12.0, {"timestamp": 12.0, "v": "b"}) is None
        out = buf.push(20.0, {"timestamp": 20.0, "v": "c"})
        assert out is not None
        assert out["timestamp"] == 12.0  # median([10, 12, 20])
        assert out["v"] == "b"  # middle frame's payload

    def test_miss_clears(self):
        buf = screenspace._ConsecutiveBuffer(3)
        assert buf.push(0.0, {"timestamp": 0.0}) is None
        assert buf.push(1.0, {"timestamp": 1.0}) is None
        buf.reset()
        # Only two matches accumulate after the reset, so nothing emits.
        assert buf.push(2.0, {"timestamp": 2.0}) is None
        assert buf.push(3.0, {"timestamp": 3.0}) is None

    def test_size_floor_of_one(self):
        # 0 / negative sizes clamp to 1 (passthrough behavior).
        buf = screenspace._ConsecutiveBuffer(0)
        assert buf.push(7.0, {"timestamp": 7.0}) is not None

    def test_carry_continues_active_run(self):
        # A static (skipped) frame carries the last match forward so the run
        # still reaches the threshold on stable content.
        buf = screenspace._ConsecutiveBuffer(3)
        assert buf.push(0.0, {"timestamp": 0.0, "text_found": "Save"}) is None
        assert buf.carry(1.0) is None  # static frame #1
        out = buf.carry(2.0)  # static frame #2 completes the run
        assert out is not None
        assert out["text_found"] == "Save"
        assert out["timestamp"] == 1.0  # median([0, 1, 2])

    def test_carry_noop_when_no_active_run(self):
        # Nothing to carry before any match has been pushed.
        buf = screenspace._ConsecutiveBuffer(3)
        assert buf.carry(5.0) is None

    def test_carry_noop_for_size_one(self):
        # size==1 emits and resets on every push, so no run is ever active to
        # carry — keeps the legacy passthrough path unchanged.
        buf = screenspace._ConsecutiveBuffer(1)
        assert buf.push(1.0, {"timestamp": 1.0}) is not None
        assert buf.carry(2.0) is None

    def test_even_size_pairs_median_with_nearest_frame(self):
        # Even runs interpolate the median between two frames; the payload comes
        # from the nearer real frame, not an arbitrary upper-middle one.
        buf = screenspace._ConsecutiveBuffer(2)
        assert buf.push(0.0, {"timestamp": 0.0, "v": "a"}) is None
        out = buf.push(4.0, {"timestamp": 4.0, "v": "b"})
        assert out is not None
        assert out["timestamp"] == 2.0  # median([0, 4])
        assert out["v"] in ("a", "b")  # nearest real frame to the median


def test_static_skip_uses_config():
    """The static-frame-skip sites reference the config constant, not a 2.0 literal."""
    src = Path(screenspace.__file__).read_text(encoding="utf-8")
    assert src.count("config.SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD") >= 4
    assert "< 2.0" not in src


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

    def test_low_confidence_rejected(self, monkeypatch):
        """OCR readings below ocr_confidence_threshold should not match."""
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)

        class _FakeReader:
            def readtext(self, _pixels, **_kwargs):
                # Number "5" detected at confidence 0.2 — well below 0.5 cutoff.
                return [([(0, 0), (10, 0), (10, 10), (0, 10)], "5", 0.2)]

        monkeypatch.setattr(
            screenspace, "_get_ocr_reader", lambda _langs: _FakeReader()
        )
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)

        default_rejected = screenspace.scan_numbers(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 20},
            operator="gte",
            target_value=5,
        )
        assert default_rejected == []

        rejected = screenspace.scan_numbers(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 20},
            operator="gte",
            target_value=5,
            ocr_confidence_threshold=0.5,
        )
        assert rejected == []

        accepted = screenspace.scan_numbers(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 20},
            operator="gte",
            target_value=5,
            ocr_confidence_threshold=0.1,
        )
        assert len(accepted) == 1
        assert accepted[0]["number_found"] == 5.0
        assert accepted[0]["confidence"] == 0.2

        accepted_zero = screenspace.scan_numbers(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 20},
            operator="gte",
            target_value=5,
            ocr_confidence_threshold=0.0,
        )
        assert len(accepted_zero) == 1

    @pytest.mark.parametrize("threshold", [-0.1, 1.1])
    def test_invalid_ocr_confidence_threshold_raises(self, monkeypatch, threshold):
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))
        with pytest.raises(ValueError, match="ocr_confidence_threshold"):
            screenspace.scan_numbers(
                "/fake.mp4",
                {"x": 0, "y": 0, "w": 60, "h": 20},
                operator="gte",
                target_value=5,
                ocr_confidence_threshold=threshold,
            )

    def test_allowlist_passed(self, monkeypatch):
        """Digit allowlist is forwarded to EasyOCR for English, omitted otherwise."""
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)
        seen: dict[str, object] = {}

        class _FakeReader:
            def readtext(self, _pixels, **kwargs):
                seen.clear()
                seen.update(kwargs)
                return []

        monkeypatch.setattr(
            screenspace, "_get_ocr_reader", lambda _langs: _FakeReader()
        )
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        # Default languages → ["en"]: allowlist forwarded.
        screenspace.scan_numbers("/fake.mp4", region, operator="gt", target_value=0)
        assert seen.get("allowlist") == "0123456789.,-"

        # Non-English combo: allowlist omitted (some combos reject it).
        screenspace.scan_numbers(
            "/fake.mp4",
            region,
            operator="gt",
            target_value=0,
            languages=["en", "ch_sim"],
        )
        assert "allowlist" not in seen

    def test_integers_only_narrows_allowlist(self, monkeypatch):
        """integers_only drops .,- from the English allowlist; off keeps them."""
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)
        seen: dict[str, object] = {}

        class _FakeReader:
            def readtext(self, _pixels, **kwargs):
                seen.clear()
                seen.update(kwargs)
                return []

        monkeypatch.setattr(
            screenspace, "_get_ocr_reader", lambda _langs: _FakeReader()
        )
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 1.0))

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        # integers_only=True (default English): digits-only allowlist.
        screenspace.scan_numbers(
            "/fake.mp4", region, operator="gt", target_value=0, integers_only=True
        )
        assert seen.get("allowlist") == "0123456789"

        # integers_only=False: separators/sign retained.
        screenspace.scan_numbers(
            "/fake.mp4", region, operator="gt", target_value=0, integers_only=False
        )
        assert seen.get("allowlist") == "0123456789.,-"

        # Non-English combo: allowlist omitted regardless of integers_only.
        screenspace.scan_numbers(
            "/fake.mp4",
            region,
            operator="gt",
            target_value=0,
            integers_only=True,
            languages=["en", "ch_sim"],
        )
        assert "allowlist" not in seen


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


# ---------------------------------------------------------------------------
# check_frame_for_tool
# ---------------------------------------------------------------------------


class TestCheckFrameForTool:
    def test_color_pass(self):
        # Pure blue in BGR: (255, 0, 0) -> HSV ~(120, 255, 255)
        frame = np.full((100, 100, 3), [255, 0, 0], dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        params = {
            "target_color": {"h": 120, "s": 255, "v": 255},
            "tolerance": {"h": 10, "s": 50, "v": 50},
        }
        passed, result = screenspace.check_frame_for_tool(
            frame, None, region, "color", params
        )
        assert passed is True
        assert result is not None
        assert "_confidence" in result

    def test_color_fail(self):
        frame = np.full((100, 100, 3), [255, 0, 0], dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        # Red target will not match blue frame
        params = {
            "target_color": {"h": 0, "s": 255, "v": 255},
            "tolerance": {"h": 5, "s": 10, "v": 10},
        }
        passed, result = screenspace.check_frame_for_tool(
            frame, None, region, "color", params
        )
        assert passed is False
        assert result is None

    def test_change_needs_prev_frame(self):
        frame = np.zeros((100, 100, 3), dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame, None, region, "change", {"threshold": 0.03}
        )
        assert passed is False
        assert result is None

    def test_change_pass(self):
        frame_a = np.zeros((100, 100, 3), dtype=np.uint8)
        frame_b = np.full((100, 100, 3), 200, dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame_b, frame_a, region, "change", {"threshold": 0.03}
        )
        assert passed is True
        assert result is not None
        assert "magnitude" in result

    def test_similarity_pass(self):
        frame = np.full((100, 100, 3), 128, dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        ref = np.full((100, 100, 3), 128, dtype=np.uint8)
        passed, result = screenspace.check_frame_for_tool(
            frame,
            None,
            region,
            "similarity",
            {"reference_frame": ref, "threshold": 0.5},
        )
        assert passed is True
        assert result is not None
        assert "score" in result

    def test_numbers_check_frame_honors_zero_ocr_threshold(self, monkeypatch):
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        class _FakeReader:
            def readtext(self, _pixels, **_kwargs):
                return [([(0, 0), (10, 0), (10, 10), (0, 10)], "5", 0.2)]

        monkeypatch.setattr(
            screenspace, "_get_ocr_reader", lambda _langs: _FakeReader()
        )

        passed, result = screenspace.check_frame_for_tool(
            frame,
            None,
            region,
            "numbers",
            {"operator": "gte", "target_value": 5},
        )
        assert passed is False
        assert result is None

        passed, result = screenspace.check_frame_for_tool(
            frame,
            None,
            region,
            "numbers",
            {
                "operator": "gte",
                "target_value": 5,
                "ocr_confidence_threshold": 0.0,
            },
        )
        assert passed is True
        assert result is not None
        assert result["number_found"] == 5.0
        assert result["confidence"] == 0.2

    def test_flow_needs_prev_frame(self):
        frame = np.zeros((100, 100, 3), dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame, None, region, "flow", {"magnitude_threshold": 2.0}
        )
        assert passed is False
        assert result is None

    def test_inactivity_needs_prev_frame(self):
        frame = np.zeros((100, 100, 3), dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame, None, region, "inactivity", {"threshold": 10}
        )
        assert passed is False
        assert result is None

    def test_inactivity_identical_frames(self):
        frame = np.full((100, 100, 3), 128, dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame, frame.copy(), region, "inactivity", {"threshold": 10}
        )
        assert passed is True
        assert result is not None
        assert "distance" in result
        assert result["distance"] == 0
        assert result["_confidence"] == 1.0
        assert screenspace._extract_confidence("inactivity", result) == 1.0

    def test_inactivity_different_frames(self):
        frame_a = np.zeros((100, 100, 3), dtype=np.uint8)
        # Random noise frame produces a very different perceptual hash
        rng = np.random.RandomState(42)
        frame_b = rng.randint(0, 256, (100, 100, 3), dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame_b, frame_a, region, "inactivity", {"threshold": 2}
        )
        assert passed is False
        assert result is None

    def test_unknown_type(self):
        frame = np.zeros((100, 100, 3), dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame, None, region, "bogus", {}
        )
        assert passed is False
        assert result is None


class TestExtractConfidence:
    def test_color(self):
        assert screenspace._extract_confidence("color", {"_confidence": 0.8}) == 0.8

    def test_change(self):
        assert screenspace._extract_confidence("change", {"magnitude": 0.5}) == 0.5

    def test_similarity(self):
        assert screenspace._extract_confidence("similarity", {"score": 0.95}) == 0.95

    def test_numbers(self):
        assert screenspace._extract_confidence("numbers", {}) == 1.0

    def test_numbers_uses_ocr_conf(self):
        assert screenspace._extract_confidence("numbers", {"confidence": 0.42}) == 0.42

    def test_multitool(self):
        assert (
            screenspace._extract_confidence("multitool", {"min_confidence": 0.7}) == 0.7
        )

    def test_inactivity_short(self):
        assert screenspace._extract_confidence("inactivity", {"duration": 3.0}) == 0.1

    def test_inactivity_per_frame_confidence(self):
        assert (
            screenspace._extract_confidence(
                "inactivity", {"distance": 2, "_confidence": 0.8}
            )
            == 0.8
        )

    def test_inactivity_capped(self):
        assert screenspace._extract_confidence("inactivity", {"duration": 60.0}) == 1.0

    def test_unknown_type(self):
        assert screenspace._extract_confidence("bogus", {}) == 1.0


class TestScanMultitool:
    def test_requires_min_2_steps(self):
        with pytest.raises(ValueError, match="at least 2"):
            screenspace.scan_multitool(
                "/fake/video.mp4",
                {"x": 0, "y": 0, "w": 100, "h": 100},
                steps=[{"type": "color"}],
            )

    @staticmethod
    def _setup_stubs(monkeypatch, check_fn):
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda _p: (30.0, 10.0))

        def fake_scan(
            video_path,
            interval_seconds,
            callback,
            *,
            start_seconds=0.0,
            end_seconds=None,
            fps=0.0,
            duration=0.0,
            fast_opts=None,
        ):
            frame = np.zeros((10, 10, 3), dtype=np.uint8)
            callback(1.0, frame)

        monkeypatch.setattr(screenspace, "scan_video_full_frames", fake_scan)
        monkeypatch.setattr(screenspace, "check_frame_for_tool", check_fn)

    def test_not_operator_rejects_when_negated_match(self, monkeypatch):
        def check(frame, prev, region, ttype, step):
            return True, {"_confidence": 0.9}

        self._setup_stubs(monkeypatch, check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[{"type": "color"}, {"type": "change", "logic": "NOT"}],
        )
        assert results == []

    def test_not_operator_passes_when_negated_misses(self, monkeypatch):
        def check(frame, prev, region, ttype, step):
            if ttype == "color":
                return True, {"_confidence": 0.8}
            return False, None

        self._setup_stubs(monkeypatch, check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[{"type": "color"}, {"type": "change", "logic": "NOT"}],
        )
        assert len(results) == 1
        assert results[0]["steps"][1] == {"negated": True, "type": "change"}
        assert results[0]["min_confidence"] == 0.8

    def test_and_default_when_logic_missing(self, monkeypatch):
        def check(frame, prev, region, ttype, step):
            if ttype == "color":
                return True, {"_confidence": 0.7}
            return True, {"magnitude": 0.5}

        self._setup_stubs(monkeypatch, check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[{"type": "color"}, {"type": "change"}],
        )
        assert len(results) == 1
        assert results[0]["min_confidence"] == 0.5

    def test_inactivity_step_uses_per_frame_confidence(self, monkeypatch):
        def check(frame, prev, region, ttype, step):
            if ttype == "inactivity":
                return True, {"distance": 0, "_confidence": 0.75}
            return True, {"_confidence": 0.9}

        self._setup_stubs(monkeypatch, check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[{"type": "color"}, {"type": "inactivity"}],
        )
        assert len(results) == 1
        assert results[0]["min_confidence"] == 0.75


# ---------------------------------------------------------------------------
# Fast Scan mode
# ---------------------------------------------------------------------------


class TestFastScanDispatchIntervalMultiplier:
    """Verify _dispatch() applies interval multiplier in fast scan mode."""

    def test_interval_multiplied_for_fast_scan(self, monkeypatch):
        captured = {}

        def fake_scan_color(
            video_path,
            region,
            *,
            target_color,
            tolerance,
            interval_seconds=0,
            start_seconds=0.0,
            end_seconds=None,
            on_progress=None,
            cancel_flag=None,
            on_result=None,
            fast_opts=None,
        ):
            captured["interval"] = interval_seconds
            captured["fast_opts"] = fast_opts
            return []

        monkeypatch.setattr(screenspace, "scan_color", fake_scan_color)

        worker = screenspace.ScreenspaceWorker()
        task = {
            "id": "ss_test1",
            "type": "color",
            "video_path": "/fake.mp4",
            "region_coords": {"x": 0, "y": 0, "w": 100, "h": 100},
            "parameters": {
                "scan_mode": "fast",
                "interval": 1.0,
                "target_color": {"h": 0, "s": 0, "v": 0},
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
        }
        worker._dispatch(task, lambda p: None, lambda: False, None)

        expected = 1.0 * config.SCREENSPACE_FAST_SCAN_INTERVAL_MULTIPLIER
        assert captured["interval"] == expected
        assert captured["fast_opts"] is not None
        assert captured["fast_opts"]["phash_skip"] is True
        assert captured["fast_opts"]["max_region_dim"] == 32

    def test_fast_scan_interval_not_persisted_on_task(self, monkeypatch):
        """Pause/resume re-dispatches the same task; interval must not compound."""
        captured_intervals: list[float] = []

        def fake_scan_color(
            video_path,
            region,
            *,
            target_color,
            tolerance,
            interval_seconds=0,
            start_seconds=0.0,
            end_seconds=None,
            on_progress=None,
            cancel_flag=None,
            on_result=None,
            fast_opts=None,
        ):
            captured_intervals.append(interval_seconds)
            return []

        monkeypatch.setattr(screenspace, "scan_color", fake_scan_color)

        worker = screenspace.ScreenspaceWorker()
        task = {
            "id": "ss_resume",
            "type": "color",
            "video_path": "/fake.mp4",
            "region_coords": {"x": 0, "y": 0, "w": 100, "h": 100},
            "parameters": {
                "scan_mode": "fast",
                "interval": 1.0,
                "target_color": {"h": 0, "s": 0, "v": 0},
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
        }

        def noop(_progress: float) -> None:
            return None

        worker._dispatch(task, noop, lambda: False, None)
        worker._dispatch(task, noop, lambda: False, None)

        expected = 1.0 * config.SCREENSPACE_FAST_SCAN_INTERVAL_MULTIPLIER
        assert task["parameters"]["interval"] == 1.0
        assert captured_intervals == [expected, expected]

    def test_normal_scan_no_fast_opts(self, monkeypatch):
        captured = {}

        def fake_scan_color(
            video_path,
            region,
            *,
            target_color,
            tolerance,
            interval_seconds=0,
            start_seconds=0.0,
            end_seconds=None,
            on_progress=None,
            cancel_flag=None,
            on_result=None,
            fast_opts=None,
        ):
            captured["interval"] = interval_seconds
            captured["fast_opts"] = fast_opts
            return []

        monkeypatch.setattr(screenspace, "scan_color", fake_scan_color)

        worker = screenspace.ScreenspaceWorker()
        task = {
            "id": "ss_test2",
            "type": "color",
            "video_path": "/fake.mp4",
            "region_coords": {"x": 0, "y": 0, "w": 100, "h": 100},
            "parameters": {
                "interval": 1.0,
                "target_color": {"h": 0, "s": 0, "v": 0},
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
        }
        worker._dispatch(task, lambda p: None, lambda: False, None)

        assert captured["interval"] == 1.0
        assert captured["fast_opts"] is None

    def test_template_dispatch_gets_downscale_flag(self, monkeypatch):
        captured = {}

        def fake_scan_template(
            video_path,
            region,
            *,
            template_image,
            threshold=0,
            interval_seconds=0,
            template_mask=None,
            template_scale=1.0,
            start_seconds=0.0,
            end_seconds=None,
            on_progress=None,
            cancel_flag=None,
            on_result=None,
            fast_opts=None,
        ):
            captured["fast_opts"] = fast_opts
            captured["template_shape"] = template_image.shape
            return []

        monkeypatch.setattr(screenspace, "scan_template", fake_scan_template)

        worker = screenspace.ScreenspaceWorker()
        tmpl = np.zeros((100, 200, 3), dtype=np.uint8)
        task = {
            "id": "ss_test3",
            "type": "template",
            "video_path": "/fake.mp4",
            "region_coords": {"x": 0, "y": 0, "w": 100, "h": 100},
            "parameters": {
                "scan_mode": "fast",
                "interval": 1.0,
                "template_image": tmpl,
            },
        }
        worker._dispatch(task, lambda p: None, lambda: False, None)

        assert captured["fast_opts"]["template_downscale"] is True
        # Template should be downscaled by 2x in _dispatch
        assert captured["template_shape"] == (50, 100, 3)


# ---------------------------------------------------------------------------
# 2E: ffmpeg pipe extraction
# ---------------------------------------------------------------------------


class TestFfmpegPipeFrames:
    def test_yields_frames_with_pts_from_stderr(self, monkeypatch):
        """Generator pairs each raw BGR frame with the pts_time read from stderr."""
        w, h = 4, 2
        frame1 = np.full((h, w, 3), 100, dtype=np.uint8)
        frame2 = np.full((h, w, 3), 200, dtype=np.uint8)
        raw = frame1.tobytes() + frame2.tobytes()
        # Two showinfo lines with PTS 0 and 1 (relative to the seek point of 5.0).
        stderr_lines = (
            b"[Parsed_showinfo_1 @ 0x0] n: 0 pts: 0 pts_time:0 ...\n"
            b"[Parsed_showinfo_1 @ 0x0] n: 1 pts: 1 pts_time:1 ...\n"
        )

        fake_proc = mock.MagicMock()
        fake_proc.stdout = io.BytesIO(raw)
        fake_proc.stderr = io.BytesIO(stderr_lines)
        fake_proc.terminate = mock.MagicMock()
        fake_proc.wait = mock.MagicMock()

        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("subprocess.Popen", lambda *a, **kw: fake_proc)

        frames = list(
            screenspace._ffmpeg_pipe_frames(
                "/fake.mp4",
                1.0,
                start_seconds=5.0,
                end_seconds=10.0,
                frame_width=w,
                frame_height=h,
            )
        )
        assert len(frames) == 2
        # Yielded ts = start_seconds + pts_time (relative).
        assert frames[0][0] == 5.0
        assert frames[1][0] == 6.0
        assert frames[0][1].shape == (h, w, 3)
        assert np.all(frames[0][1] == 100)
        assert np.all(frames[1][1] == 200)

    def test_empty_when_no_ffmpeg(self, monkeypatch):
        monkeypatch.setattr("shutil.which", lambda _: None)
        frames = list(
            screenspace._ffmpeg_pipe_frames(
                "/fake.mp4", 1.0, frame_width=10, frame_height=10
            )
        )
        assert frames == []

    def test_uses_single_preinput_seek_with_select_filter(self, monkeypatch):
        """Analysis pipe uses a single pre-input -ss plus the select filter +
        fps_mode vfr. The previous two-stage seek silently dropped its post-input
        -ss when paired with -vf in modern ffmpeg, and the fps filter chose a
        different source frame than the preview's accurate seek did."""
        captured: dict = {}

        def fake_popen(cmd, *a, **kw):
            captured["cmd"] = list(cmd)
            proc = mock.MagicMock()
            proc.stdout = io.BytesIO(b"")
            proc.stderr = io.BytesIO(b"")
            proc.terminate = mock.MagicMock()
            proc.wait = mock.MagicMock()
            return proc

        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("subprocess.Popen", fake_popen)

        list(
            screenspace._ffmpeg_pipe_frames(
                "/fake.mp4",
                1.0,
                start_seconds=15.0,
                end_seconds=20.0,
                frame_width=4,
                frame_height=2,
            )
        )
        cmd = captured["cmd"]
        i_idx = cmd.index("-i")
        # One pre-input -ss, nothing after -i.
        assert cmd[:i_idx].count("-ss") == 1
        assert "-ss" not in cmd[i_idx + 2 :]
        vf = cmd[cmd.index("-vf") + 1]
        assert "select=" in vf
        assert "showinfo" in vf
        # vfr is what stops ffmpeg from duplicating the kept frame to fill the
        # source frame rate.
        assert "-fps_mode" in cmd and cmd[cmd.index("-fps_mode") + 1] == "vfr"


class TestScanViaFfmpegPipe:
    def test_returns_false_when_no_ffmpeg(self, monkeypatch):
        monkeypatch.setattr("shutil.which", lambda _: None)
        result = screenspace._scan_via_ffmpeg_pipe(
            "/fake.mp4", None, 1.0, lambda ts, f: None, duration=10.0
        )
        assert result is False

    def test_returns_false_when_probe_fails(self, monkeypatch):
        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("video.probe_video_properties", lambda _: None)
        result = screenspace._scan_via_ffmpeg_pipe(
            "/fake.mp4", None, 1.0, lambda ts, f: None, duration=10.0
        )
        assert result is False

    def test_calls_callback_with_frames(self, monkeypatch):
        w, h = 4, 2
        frame_data = np.full((h, w, 3), 42, dtype=np.uint8)
        raw = frame_data.tobytes()
        # showinfo emits one line per yielded frame; pts_time is relative to
        # the seek point (here 0).
        stderr_lines = b"[Parsed_showinfo_1 @ 0x0] n: 0 pts: 0 pts_time:0 ...\n"

        fake_proc = mock.MagicMock()
        fake_proc.stdout = io.BytesIO(raw)
        fake_proc.stderr = io.BytesIO(stderr_lines)
        fake_proc.terminate = mock.MagicMock()
        fake_proc.wait = mock.MagicMock()

        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("subprocess.Popen", lambda *a, **kw: fake_proc)
        monkeypatch.setattr(
            "video.probe_video_properties",
            lambda _: {"width": w, "height": h, "video_codec": "h264"},
        )

        received = []

        def cb(ts, frame):
            received.append((ts, frame.copy()))

        result = screenspace._scan_via_ffmpeg_pipe(
            "/fake.mp4", None, 1.0, cb, duration=10.0, full_frame=True
        )
        assert result is True
        assert len(received) == 1
        assert received[0][0] == 0.0
        assert np.all(received[0][1] == 42)

    def test_builds_crop_filter_for_region(self, monkeypatch):
        w, h = 8, 6
        region = {"x": 1, "y": 2, "w": 4, "h": 3}

        captured_cmd = {}

        def fake_popen(cmd, **kw):
            captured_cmd["args"] = cmd
            proc = mock.MagicMock()
            # Return empty bytes so the generator exits immediately
            proc.stdout = io.BytesIO(b"")
            proc.terminate = mock.MagicMock()
            proc.wait = mock.MagicMock()
            return proc

        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("subprocess.Popen", fake_popen)
        monkeypatch.setattr(
            "video.probe_video_properties",
            lambda _: {"width": w, "height": h, "video_codec": "h264"},
        )

        screenspace._scan_via_ffmpeg_pipe(
            "/fake.mp4", region, 1.0, lambda ts, f: None, duration=10.0
        )

        cmd_str = " ".join(captured_cmd["args"])
        assert "crop=4:3:1:2" in cmd_str

    def test_stops_on_callback_false(self, monkeypatch):
        w, h = 4, 2
        frame = np.full((h, w, 3), 1, dtype=np.uint8)
        raw = frame.tobytes() * 5  # 5 frames
        stderr_lines = b"".join(
            b"[Parsed_showinfo_1 @ 0x0] n: %d pts_time:%d ...\n" % (i, i)
            for i in range(5)
        )

        fake_proc = mock.MagicMock()
        fake_proc.stdout = io.BytesIO(raw)
        fake_proc.stderr = io.BytesIO(stderr_lines)
        fake_proc.terminate = mock.MagicMock()
        fake_proc.wait = mock.MagicMock()

        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("subprocess.Popen", lambda *a, **kw: fake_proc)
        monkeypatch.setattr(
            "video.probe_video_properties",
            lambda _: {"width": w, "height": h, "video_codec": "h264"},
        )

        call_count = [0]

        def cb(ts, frame):
            call_count[0] += 1
            if call_count[0] >= 2:
                return False

        result = screenspace._scan_via_ffmpeg_pipe(
            "/fake.mp4", None, 1.0, cb, duration=10.0, full_frame=True
        )
        assert result is True
        assert call_count[0] == 2


class TestAnalysisPreviewAlignment:
    """End-to-end: each ts yielded by `_ffmpeg_pipe_frames` must point at the
    exact same source frame that `video.extract_frame_at_timestamp(ts)` returns,
    so clicking a result in Screenspace shows the analysed frame instead of a
    drifted neighbour.

    Requires a real ffmpeg binary on PATH; skipped otherwise.
    """

    @staticmethod
    def _have_ffmpeg() -> bool:
        import shutil as _shutil

        return (
            _shutil.which("ffmpeg") is not None and _shutil.which("ffprobe") is not None
        )

    @staticmethod
    def _synthesize(path: str, vf_extra: str = "") -> None:
        import subprocess as _sp

        vf = "testsrc=duration=12:size=320x240:rate=30"
        if vf_extra:
            vf += f",{vf_extra}"
        _sp.run(
            [
                "ffmpeg",
                "-y",
                "-f",
                "lavfi",
                "-i",
                vf,
                "-fps_mode",
                "vfr",
                "-c:v",
                "libx264",
                "-g",
                "30",
                "-pix_fmt",
                "yuv420p",
                path,
            ],
            check=True,
            stdout=_sp.DEVNULL,
            stderr=_sp.DEVNULL,
        )

    @staticmethod
    def _assert_aligned(video_path: str) -> None:
        import hashlib

        import video as video_mod

        frames = list(
            screenspace._ffmpeg_pipe_frames(
                video_path,
                interval_seconds=1.0,
                start_seconds=0.0,
                end_seconds=10.0,
                frame_width=320,
                frame_height=240,
            )
        )
        assert frames, "expected at least one analysed frame"
        for ts, analysis_frame in frames:
            preview = video_mod.extract_frame_at_timestamp(video_path, ts)
            assert preview is not None, f"preview extraction failed at ts={ts}"
            h_a = hashlib.md5(analysis_frame.tobytes()).hexdigest()
            h_p = hashlib.md5(preview.tobytes()).hexdigest()
            assert h_a == h_p, (
                f"analysed frame and preview differ at ts={ts:.4f} "
                f"(analysis={h_a[:8]} preview={h_p[:8]})"
            )

    def test_cfr_source_alignment(self, tmp_path):
        if not self._have_ffmpeg():
            pytest.skip("ffmpeg/ffprobe required for end-to-end alignment test")
        video_path = str(tmp_path / "cfr.mp4")
        self._synthesize(video_path)
        self._assert_aligned(video_path)

    def test_vfr_source_alignment(self, tmp_path):
        if not self._have_ffmpeg():
            pytest.skip("ffmpeg/ffprobe required for end-to-end alignment test")
        video_path = str(tmp_path / "vfr.mp4")
        # Drop ~2/7 of frames at non-uniform positions so the kept frames don't
        # land on integer second boundaries — the case where the old fps filter
        # picked a different source frame than the preview's accurate seek.
        self._synthesize(
            video_path,
            vf_extra="select='not(eq(mod(n,7),3))*not(eq(mod(n,11),5))'",
        )
        self._assert_aligned(video_path)


class TestScanVideoFramesFfmpegIntegration:
    def test_ffmpeg_pipe_succeeds(self, monkeypatch):
        """When ffmpeg pipe succeeds, scan completes without error."""
        monkeypatch.setattr(
            screenspace,
            "_scan_via_ffmpeg_pipe",
            lambda *a, **kw: True,
        )

        screenspace.scan_video_frames(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            1.0,
            lambda ts, f: None,
        )

    def test_warns_when_ffmpeg_pipe_fails(self, monkeypatch):
        """When ffmpeg pipe fails, a warning is emitted."""
        monkeypatch.setattr(
            screenspace,
            "_scan_via_ffmpeg_pipe",
            lambda *a, **kw: False,
        )
        warnings = []
        monkeypatch.setattr(
            screenspace.utils,
            "warning_print",
            lambda msg, *a, **kw: warnings.append(msg),
        )

        screenspace.scan_video_frames(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            1.0,
            lambda ts, f: None,
        )
        assert len(warnings) == 1
        assert "ffmpeg" in warnings[0].lower()


# ---------------------------------------------------------------------------
# 2F: Parallel worker
# ---------------------------------------------------------------------------


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
            "/v.mp4",
            "r1",
            {"x": 0, "y": 0, "w": 1, "h": 1},
        )
        t2 = screenspace.create_task(
            "color",
            "P02",
            "s.mp4",
            "/v2.mp4",
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
                f"/v{i}.mp4",
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
            "/v.mp4",
            "r1",
            {"x": 0, "y": 0, "w": 1, "h": 1},
        )
        t2 = screenspace.create_task(
            "color",
            "P02",
            "s.mp4",
            "/v2.mp4",
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


class TestBuildTimelapseCommandMarkers:
    def test_start_seconds_adds_ss_flag(self):
        cmd = screenspace.build_timelapse_command(
            "/video.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            10.0,
            "/out.mp4",
            start_seconds=30.0,
        )
        ss_idx = cmd.index("-ss")
        assert cmd[ss_idx + 1] == "30.0"
        # -ss must appear before -i for fast seeking
        i_idx = cmd.index("-i")
        assert ss_idx < i_idx

    def test_end_seconds_adds_t_flag(self):
        cmd = screenspace.build_timelapse_command(
            "/video.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            10.0,
            "/out.mp4",
            start_seconds=10.0,
            end_seconds=40.0,
        )
        t_idx = cmd.index("-t")
        assert cmd[t_idx + 1] == "30.0"  # 40 - 10

    def test_no_markers_omits_flags(self):
        cmd = screenspace.build_timelapse_command(
            "/video.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            10.0,
            "/out.mp4",
        )
        assert "-ss" not in cmd
        assert "-t" not in cmd


class TestGenerateTimelapseProgress:
    def test_reports_progress_from_ffmpeg_output(self, monkeypatch):
        """on_progress is called with values parsed from ffmpeg -progress."""
        # Simulate ffmpeg -progress output with out_time_us lines
        progress_output = (
            b"out_time_us=0\n"
            b"progress=continue\n"
            b"out_time_us=5000000\n"
            b"progress=continue\n"
            b"out_time_us=10000000\n"
            b"progress=end\n"
        )

        fake_proc = mock.MagicMock()
        fake_proc.stdout = io.BytesIO(progress_output)
        fake_proc.returncode = 0
        fake_proc.wait = mock.MagicMock()

        monkeypatch.setattr("subprocess.Popen", lambda *a, **kw: fake_proc)

        progress_values = []

        result = screenspace.generate_timelapse(
            "/video.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            10.0,
            "/out.mp4",
            start_seconds=0.0,
            end_seconds=100.0,  # 100s input / 10x speedup = 10s output = 10_000_000 us
            on_progress=lambda p: progress_values.append(round(p, 2)),
        )

        assert result == "/out.mp4"
        # Should have: 0.0 (initial), 0.0 (out_time_us=0), 0.5 (5M/10M), 0.99 (capped)
        assert 0.0 in progress_values
        assert any(0.4 <= v <= 0.6 for v in progress_values)  # ~0.5 from 5M/10M
        assert 1.0 in progress_values  # final

    def test_cancel_flag_terminates_process(self, monkeypatch):
        """cancel_flag=True stops ffmpeg and returns None."""
        # Output enough lines so the loop iterates
        progress_output = b"out_time_us=0\nprogress=continue\n" * 5

        fake_proc = mock.MagicMock()
        fake_proc.stdout = io.BytesIO(progress_output)
        fake_proc.returncode = -15  # SIGTERM
        fake_proc.wait = mock.MagicMock()
        fake_proc.terminate = mock.MagicMock()

        monkeypatch.setattr("subprocess.Popen", lambda *a, **kw: fake_proc)

        result = screenspace.generate_timelapse(
            "/video.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            10.0,
            "/out.mp4",
            start_seconds=0.0,
            end_seconds=100.0,
            cancel_flag=lambda: True,
        )

        assert result is None
        fake_proc.terminate.assert_called_once()


class TestTimelapseDispatchPassesMarkers:
    def test_dispatch_forwards_start_end_to_timelapse(self, monkeypatch):
        """_dispatch passes start_seconds and end_seconds to generate_timelapse."""
        captured = {}

        def fake_generate(
            video_path, region, speedup_factor, output_path, output_format="mp4", **kw
        ):
            captured.update(kw)
            return output_path

        monkeypatch.setattr(screenspace, "generate_timelapse", fake_generate)

        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "timelapse",
            "P01",
            "s_P01.mp4",
            "/fake.mp4",
            "region1",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            parameters={
                "speedup_factor": 5.0,
                "start_seconds": 15.0,
                "end_seconds": 45.0,
            },
        )
        worker._dispatch(task, lambda p: None, lambda: False, None)

        assert captured["start_seconds"] == 15.0
        assert captured["end_seconds"] == 45.0
        assert captured["on_progress"] is not None
        assert captured["cancel_flag"] is not None


# ---------------------------------------------------------------------------
# Inactivity tool
# ---------------------------------------------------------------------------


class TestScanInactivity:
    """Tests for scan_inactivity() function."""

    def test_identical_frames_detected(self, monkeypatch):
        """Identical consecutive frames should produce an inactivity span."""
        frame = np.full((50, 50, 3), 128, dtype=np.uint8)
        timestamps = [0.0, 1.0, 2.0, 3.0, 4.0]
        call_idx = [0]

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for ts in timestamps:
                result = callback(ts, frame.copy())
                if result is False:
                    break
                call_idx[0] += 1

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 5.0))

        results = screenspace.scan_inactivity(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 50, "h": 50},
            threshold=15,
            min_duration=2.0,
            interval_seconds=1.0,
        )
        assert len(results) == 1
        assert results[0]["start"] == 0.0
        assert results[0]["end"] == 4.0
        assert results[0]["duration"] == 4.0
        assert results[0]["avg_distance"] == 0.0

    def test_different_frames_not_detected(self, monkeypatch):
        """Frames with very different content should not produce a span."""
        timestamps = [0.0, 1.0, 2.0, 3.0]
        # Use random noise frames with different seeds for visually distinct content
        frames = [
            np.random.RandomState(seed).randint(0, 256, (50, 50, 3)).astype(np.uint8)
            for seed in [10, 20, 30, 40]
        ]

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for i, ts in enumerate(timestamps):
                result = callback(ts, frames[i])
                if result is False:
                    break

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 4.0))

        results = screenspace.scan_inactivity(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 50, "h": 50},
            threshold=2,
            min_duration=2.0,
            interval_seconds=1.0,
        )
        assert len(results) == 0

    def test_min_duration_filtering(self, monkeypatch):
        """Spans shorter than min_duration should be discarded."""
        frame = np.full((50, 50, 3), 128, dtype=np.uint8)
        # Random noise frame to ensure phash is very different from the solid frame
        diff_frame = (
            np.random.RandomState(99).randint(0, 256, (50, 50, 3)).astype(np.uint8)
        )
        # 2 identical frames then a different one — span is only 1s
        timestamps = [0.0, 1.0, 2.0]
        frame_seq = [frame, frame.copy(), diff_frame]

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for i, ts in enumerate(timestamps):
                result = callback(ts, frame_seq[i])
                if result is False:
                    break

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 3.0))

        results = screenspace.scan_inactivity(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 50, "h": 50},
            threshold=15,
            min_duration=5.0,
            interval_seconds=1.0,
        )
        assert len(results) == 0

    def test_on_result_callback(self, monkeypatch):
        """on_result should fire once per completed span."""
        frame = np.full((50, 50, 3), 128, dtype=np.uint8)
        timestamps = [0.0, 1.0, 2.0, 3.0, 4.0]
        emitted = []

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for ts in timestamps:
                result = callback(ts, frame.copy())
                if result is False:
                    break

        monkeypatch.setattr(screenspace, "scan_video_frames", fake_scan)
        monkeypatch.setattr(screenspace, "_probe_video_meta", lambda p: (30.0, 5.0))

        screenspace.scan_inactivity(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 50, "h": 50},
            threshold=15,
            min_duration=2.0,
            interval_seconds=1.0,
            on_result=lambda r: emitted.append(r),
        )
        assert len(emitted) == 1
        assert emitted[0]["duration"] == 4.0

    def test_dispatch_routes_to_scan_inactivity(self, monkeypatch):
        """ScreenspaceWorker._dispatch() should route inactivity tasks."""
        captured = {}

        def fake_scan_inactivity(video_path, region, **kwargs):
            captured.update(kwargs)
            return []

        monkeypatch.setattr(screenspace, "scan_inactivity", fake_scan_inactivity)

        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "inactivity",
            "P01",
            "s_P01.mp4",
            "/fake.mp4",
            "region1",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            parameters={
                "threshold": 8,
                "min_duration": 5.0,
                "interval": 2.0,
            },
        )
        worker._dispatch(task, lambda p: None, lambda: False, None)

        assert captured["threshold"] == 8
        assert captured["min_duration"] == 5.0
        assert captured["interval_seconds"] == 2.0


class TestMorphKernel:
    def test_shape_and_dtype(self):
        kernel = screenspace._morph_kernel(3)
        assert kernel.shape == (3, 3)
        assert kernel.dtype == np.uint8
        assert np.all(kernel == 1)

    def test_same_size_returns_cached_array(self):
        # cv2 morphology treats the kernel read-only, so callers share one array.
        assert screenspace._morph_kernel(5) is screenspace._morph_kernel(5)
