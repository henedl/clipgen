"""Tests for optical flow primitives, grid, and heatmap."""

import numpy as np

import config
import screenspace


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
