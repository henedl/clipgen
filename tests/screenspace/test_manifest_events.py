"""Tests for manifest I/O, timestamp spans, and event generation."""

import config
import screenspace


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

    def test_omitting_pins_preserves_existing_pins(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        pins = {
            "P01": [
                {
                    "id": "pin_keep",
                    "timestamp": 12.5,
                    "polarity": "positive",
                    "label": "",
                    "created_at": "2026-01-01T00:00:00+00:00",
                }
            ]
        }
        screenspace.save_screenspace_manifest({}, [], pins=pins)
        screenspace.save_screenspace_manifest({"new": {"x": 0}}, [])
        loaded = screenspace.load_screenspace_manifest()
        assert loaded["pins"] == pins

    def test_explicit_empty_pins_clears_existing_pins(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        screenspace.save_screenspace_manifest(
            {},
            [],
            pins={
                "P01": [
                    {
                        "id": "pin_clear",
                        "timestamp": 1.0,
                        "polarity": "positive",
                        "label": "",
                        "created_at": "2026-01-01T00:00:00+00:00",
                    }
                ]
            },
        )
        screenspace.save_screenspace_manifest({}, [], pins={})
        assert screenspace.load_screenspace_manifest()["pins"] == {}


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

    def test_boundary_events(self):
        worker, task = self._make_worker_and_task("boundary")
        raw = [
            {
                "timestamp": 12.0,
                "distance": 22,
                "_confidence": 0.5714,
                "period_start": 12.0,
                "period_end": 30.0,
                "scene_label": "Scene B",
            }
        ]
        events = worker._generate_events_from_results(task, raw)
        assert len(events) == 1
        assert events[0]["time_in"] == 12.0
        assert events[0]["time_out"] == 12.0
        assert events[0]["metadata"]["distance"] == 22
        # Scene/hybrid period spans and the recurrence-aware label reach metadata.
        assert events[0]["metadata"]["period_start"] == 12.0
        assert events[0]["metadata"]["period_end"] == 30.0
        assert events[0]["metadata"]["scene_label"] == "Scene B"
        # Boundaries are orientation markers; Studio intake hides them by default.
        assert events[0]["navigational"] is True
        # confidence is carried through from the per-frame _confidence scalar.
        assert abs(events[0]["confidence"] - 0.5714) < 0.001


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

    def test_roundtrip_preserves_navigational(self, tmp_path, monkeypatch):
        # Boundary events carry navigational so Studio intake can hide them and
        # timelines can render them distinctly — it must survive save → load.
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        events = [
            {
                "id": "ev_boundary1",
                "source_video": "study_P01.mp4",
                "participant": "P01",
                "detector": "boundary",
                "event_type": "boundary",
                "time_in": 12.0,
                "time_out": 12.0,
                "confidence": 0.57,
                "metadata": {"distance": 22},
                "excluded": False,
                "navigational": True,
                "task_id": "ss_bnd00001",
                "region": "full_frame",
            }
        ]
        screenspace.save_screenspace_manifest({}, [], events)
        loaded = screenspace.load_screenspace_manifest()
        assert loaded["events"][0]["navigational"] is True
        assert loaded["events"][0]["metadata"]["distance"] == 22


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
            video_paths=["/path/study_P01.mp4"],
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
