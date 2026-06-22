"""Tests for frame diff, region similarity, phash, and scene fingerprint."""

import numpy as np

import config
import screenspace
import screenspace_scans


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


class TestScanChangesEmitsGrid:
    def _run_scan(self, monkeypatch):
        frame_a = np.full((100, 100, 3), 128, dtype=np.uint8)
        frame_b = frame_a.copy()
        frame_b[20:80, 20:80] = 255  # large bright change → high magnitude

        def fake_probe(_path):
            return 30.0, 10.0

        def fake_scan_video_frames(video_path, region, interval, cb, **kwargs):
            cb(0.0, frame_a)
            cb(1.0, frame_b)
            return []

        monkeypatch.setattr(screenspace_scans, "_probe_video_meta", fake_probe)
        monkeypatch.setattr(
            screenspace_scans, "scan_video_frames", fake_scan_video_frames
        )
        return screenspace_scans.scan_changes(
            "/fake.mp4", {"x": 0, "y": 0, "w": 100, "h": 100}
        )

    def test_change_grid_emitted_on_change_frame(self, monkeypatch):
        monkeypatch.setattr(config, "SCREENSPACE_GENERATE_CHANGE_HEATMAP", True)
        results = self._run_scan(monkeypatch)
        assert len(results) == 1
        grid = results[0]["change_grid"]
        assert isinstance(grid, list) and len(grid) > 0
        # Stays small/sparse: bounded by grid^2 and thresholded (no full-res blow-up).
        assert len(grid) <= config.SCREENSPACE_CHANGE_HEATMAP_GRID**2
        for cell in grid:
            assert set(cell) == {"x", "y", "mag"}
            assert 0.0 <= cell["x"] <= 1.0 and 0.0 <= cell["y"] <= 1.0
            assert config.SCREENSPACE_CHANGE_HEATMAP_MIN_FRAC <= cell["mag"] <= 1.0

    def test_no_change_grid_when_heatmaps_disabled(self, monkeypatch):
        monkeypatch.setattr(config, "SCREENSPACE_GENERATE_CHANGE_HEATMAP", False)
        results = self._run_scan(monkeypatch)
        assert len(results) == 1
        assert "change_grid" not in results[0]  # not computed when disabled
        assert results[0]["magnitude"] > 0  # detection itself still works


class TestGenerateChangeHeatmap:
    def test_basic_heatmap(self, tmp_path):
        results = [
            {
                "timestamp": 1.0,
                "change_grid": [
                    {"x": 0.5, "y": 0.5, "mag": 0.8},
                    {"x": 0.2, "y": 0.3, "mag": 0.4},
                ],
            }
        ]
        out = str(tmp_path / "change_heatmap.png")
        path = screenspace.generate_change_heatmap(results, 200, 150, out)
        assert path == out
        assert (tmp_path / "change_heatmap.png").is_file()
        assert (tmp_path / "change_heatmap.png").stat().st_size > 0

    def test_empty_grid_returns_none(self, tmp_path):
        results = [{"timestamp": 1.0, "change_grid": []}]
        out = str(tmp_path / "change_heatmap.png")
        assert screenspace.generate_change_heatmap(results, 200, 150, out) is None


class TestWorkerOmitsChangeGrid:
    """Server-only change_grid must never be deep-copied for external reads."""

    def test_copy_for_read_drops_change_grid_keeps_flow_grid(self):
        import screenspace_worker

        task = {
            "id": "ss_x",
            "type": "change",
            "result": [
                {
                    "timestamp": 1.0,
                    "magnitude": 0.5,
                    "change_grid": [{"x": 0.5, "y": 0.5, "mag": 0.8}],
                },
                {
                    "timestamp": 2.0,
                    "flow_grid": [{"x": 0.1, "y": 0.1, "mag": 2.0, "ang": 90.0}],
                },
            ],
            "_raw_results": [
                {"timestamp": 1.0, "change_grid": [{"x": 0.5, "y": 0.5, "mag": 0.8}]},
            ],
        }
        out = screenspace_worker._copy_task_for_read(task)
        assert "change_grid" not in out["result"][0]
        assert out["result"][0]["magnitude"] == 0.5
        assert "flow_grid" in out["result"][1]  # flow overlay still needs this
        assert "change_grid" not in out["_raw_results"][0]
        # The original in-memory task is untouched (independent copy).
        assert "change_grid" in task["result"][0]

    def test_get_task_omits_change_grid(self):
        worker = screenspace.ScreenspaceWorker()
        worker._tasks["ss_y"] = {
            "id": "ss_y",
            "type": "change",
            "result": [
                {
                    "timestamp": 1.0,
                    "magnitude": 0.5,
                    "change_grid": [{"x": 0.5, "y": 0.5, "mag": 0.8}],
                }
            ],
        }
        out = worker.get_task("ss_y")
        assert out is not None
        assert "change_grid" not in out["result"][0]


class TestChangeGridStrippedFromManifest:
    def test_change_grid_not_persisted(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        tasks = [
            {
                "id": "ss_change123",
                "type": "change",
                "participant": "P01",
                "status": "completed",
                "result": [
                    {
                        "timestamp": 1.0,
                        "magnitude": 0.5,
                        "change_grid": [{"x": 0.5, "y": 0.5, "mag": 0.8}],
                    }
                ],
            }
        ]
        path = screenspace.save_screenspace_manifest({}, tasks)
        assert path is not None
        loaded = screenspace.load_screenspace_manifest()
        result = loaded["tasks"][0]["result"][0]
        assert "change_grid" not in result
        assert result["magnitude"] == 0.5


class TestHeatmapDisableSetting:
    def _change_task(self):
        return {
            "id": "ss_chg_setting",
            "type": "change",
            "region_coords": {"w": 80, "h": 60},
            "video_paths": ["/fake.mp4"],
        }

    def _change_results(self, n=8):
        return [
            {"timestamp": float(i), "change_grid": [{"x": 0.5, "y": 0.5, "mag": 0.7}]}
            for i in range(n)
        ]

    def _run_generate(self, worker, task):
        return worker._generate_heatmap(
            task["type"],
            task["id"],
            task["video_paths"],
            task["region_coords"],
            self._change_results(),
        )

    def test_skipped_when_disabled(self, monkeypatch):
        monkeypatch.setattr(config, "SCREENSPACE_GENERATE_CHANGE_HEATMAP", False)
        worker = screenspace.ScreenspaceWorker()
        attachments = self._run_generate(worker, self._change_task())
        assert attachments == {}

    def test_generated_when_enabled(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        monkeypatch.setattr(config, "SCREENSPACE_GENERATE_CHANGE_HEATMAP", True)
        worker = screenspace.ScreenspaceWorker()
        attachments = self._run_generate(worker, self._change_task())
        assert "heatmap" in attachments

    def test_settings_present_and_default_true(self):
        for name in (
            "SCREENSPACE_GENERATE_TEMPLATE_HEATMAP",
            "SCREENSPACE_GENERATE_FLOW_HEATMAP",
            "SCREENSPACE_GENERATE_CHANGE_HEATMAP",
        ):
            assert getattr(config, name) is True
            meta = config.STUDIO_SETTINGS[name]
            assert meta["tab"] == "Screenspace"
            assert meta["group"] == "Heatmaps"
            assert meta["type"] == "bool"
            assert config.SETTINGS_DESCRIPTIONS.get(name)
