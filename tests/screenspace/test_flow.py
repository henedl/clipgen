"""Tests for optical flow primitives, grid, and heatmap."""

import cv2
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

    def test_predownscaled_inputs_match_full_size(self):
        """scan_flow's carried-forward flow_downscale path must be bit-identical.

        flow_downscale replicates exactly the resize compute_optical_flow used
        to apply internally; pre-downscaled inputs (with the mask downscaled
        alongside) must therefore produce the same result dict.
        """
        rng = np.random.default_rng(5)
        prev = rng.integers(0, 256, (480, 640), dtype=np.uint8)
        curr = rng.integers(0, 256, (480, 640), dtype=np.uint8)
        mask = np.zeros((480, 640), dtype=np.uint8)
        mask[100:300, 200:500] = 255

        full = screenspace.compute_optical_flow(prev, curr, return_grid=True, mask=mask)
        small_prev, _ = screenspace.flow_downscale(prev)
        small_curr, small_mask = screenspace.flow_downscale(curr, mask)
        pre = screenspace.compute_optical_flow(
            small_prev, small_curr, return_grid=True, mask=small_mask
        )
        assert full == pre

    def test_small_inputs_are_not_resized(self):
        gray = np.full((50, 50), 128, dtype=np.uint8)
        out, mask = screenspace.flow_downscale(gray)
        assert out is gray
        assert mask is None

    def test_magnitude_and_phase_match_carttopolar(self):
        """Guards the cartToPolar -> magnitude/phase split in compute_optical_flow.

        The grid path defers angles to cv2.phase and always takes magnitude
        from cv2.magnitude. Neither is bit-identical to cartToPolar (different
        SIMD paths, last-ulp float32 drift), so this pins the bound instead:
        far below the 4-decimal mag / 1-decimal angle rounding of emitted rows.
        """
        rng = np.random.default_rng(7)
        dx = rng.standard_normal((256, 256)).astype(np.float32)
        dy = rng.standard_normal((256, 256)).astype(np.float32)
        mag_ref, ang_ref = cv2.cartToPolar(dx, dy, angleInDegrees=True)
        assert float(np.abs(cv2.magnitude(dx, dy) - mag_ref).max()) < 1e-5
        ang = cv2.phase(dx, dy, angleInDegrees=True)
        assert float(np.abs(ang - ang_ref).max()) < 1e-3

    def test_grid_min_magnitude_gates_grid(self):
        prev = np.zeros((80, 80), dtype=np.uint8)
        curr = np.zeros((80, 80), dtype=np.uint8)
        prev[20:40, 20:40] = 255
        curr[30:50, 30:50] = 255
        full = screenspace.compute_optical_flow(prev, curr, return_grid=True)
        assert full["magnitude"] > 0
        # At or above the gate: identical result, grid included.
        gated_on = screenspace.compute_optical_flow(
            prev, curr, return_grid=True, grid_min_magnitude=full["magnitude"]
        )
        assert gated_on == full
        # Below the gate: same scalars, no flow_grid key at all.
        gated_off = screenspace.compute_optical_flow(
            prev, curr, return_grid=True, grid_min_magnitude=full["magnitude"] + 1
        )
        assert "flow_grid" not in gated_off
        assert gated_off["magnitude"] == full["magnitude"]
        assert gated_off["angle"] == full["angle"]


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
