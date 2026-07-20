"""Tests for the Attention tool: saliency primitives, scan, and events."""

import numpy as np

import screenspace
import screenspace_frames
import screenspace_scans


def _bright_patch_frame(w=160, h=120, px=120, py=80, size=20):
    """Flat dark frame with one bright high-contrast square patch."""
    frame = np.full((h, w, 3), 30, dtype=np.uint8)
    frame[py : py + size, px : px + size] = 245
    return frame


class TestSpectralResidual:
    def test_shape_range_and_determinism(self):
        gray = np.random.RandomState(3).randint(0, 255, (90, 120), dtype=np.uint8)
        sal_a = screenspace.compute_spectral_residual(gray)
        sal_b = screenspace.compute_spectral_residual(gray.copy())
        assert sal_a.shape == gray.shape
        assert sal_a.dtype == np.float32
        assert float(sal_a.min()) >= 0.0 and float(sal_a.max()) <= 1.0
        assert np.array_equal(sal_a, sal_b)

    def test_flat_frame_does_not_crash(self):
        gray = np.full((60, 80), 128, dtype=np.uint8)
        sal = screenspace.compute_spectral_residual(gray)
        assert sal.shape == gray.shape
        assert np.all(np.isfinite(sal))


class TestChannels:
    def test_color_contrast_highlights_patch(self):
        frame = _bright_patch_frame()
        contrast = screenspace.compute_color_contrast(frame)
        assert contrast.shape == frame.shape[:2]
        patch_mean = float(contrast[80:100, 120:140].mean())
        background_mean = float(contrast[0:40, 0:40].mean())
        assert patch_mean > background_mean

    def test_motion_zero_without_prev(self):
        gray = np.full((60, 80), 128, dtype=np.uint8)
        motion = screenspace.compute_motion_saliency(None, gray)
        assert motion.shape == gray.shape
        assert float(motion.max()) == 0.0

    def test_motion_zero_on_shape_change(self):
        prev = np.full((50, 70), 128, dtype=np.uint8)
        curr = np.full((60, 80), 128, dtype=np.uint8)
        motion = screenspace.compute_motion_saliency(prev, curr)
        assert motion.shape == curr.shape
        assert float(motion.max()) == 0.0

    def test_motion_highlights_changed_area(self):
        prev = np.full((60, 80), 30, dtype=np.uint8)
        curr = prev.copy()
        curr[20:40, 40:60] = 240
        motion = screenspace.compute_motion_saliency(prev, curr)
        assert float(motion[25:35, 45:55].mean()) > float(motion[0:10, 0:10].mean())

    def test_face_saliency_zeros_without_faces(self):
        gray = np.full((60, 80), 128, dtype=np.uint8)
        blobs = screenspace.compute_face_saliency(gray)
        assert blobs.shape == gray.shape
        assert float(blobs.max()) == 0.0


class TestSaliencyMap:
    def test_peak_lands_on_bright_patch(self):
        frame = _bright_patch_frame()
        sal, curr_gray = screenspace.compute_saliency_map(
            frame, None, center_bias=0.0, include_face=False
        )
        assert sal.shape == frame.shape[:2]
        assert curr_gray.shape == frame.shape[:2]
        px, py, pv = screenspace.saliency_peak(sal)
        # Patch spans x 120-140 of 160 (0.75-0.875), y 80-100 of 120 (0.67-0.83)
        assert 0.70 <= px <= 0.92
        assert 0.60 <= py <= 0.90
        assert pv > 0.0

    def test_center_bias_pulls_toward_center(self):
        frame = np.full((120, 160, 3), 30, dtype=np.uint8)
        # Two identical patches: one centered, one cornered
        frame[50:70, 70:90] = 245
        frame[0:20, 0:20] = 245
        sal, _ = screenspace.compute_saliency_map(
            frame, None, center_bias=0.9, include_face=False
        )
        px, py, _ = screenspace.saliency_peak(sal)
        assert 0.3 <= px <= 0.7
        assert 0.3 <= py <= 0.7

    def test_map_in_unit_range(self):
        frame = np.random.RandomState(5).randint(0, 255, (90, 120, 3), dtype=np.uint8)
        sal, _ = screenspace.compute_saliency_map(frame, None, include_face=False)
        assert float(sal.min()) >= 0.0
        assert float(sal.max()) <= 1.0


class TestSaliencyGrid:
    def test_grid_cells_shape_and_threshold(self):
        frame = _bright_patch_frame()
        sal, _ = screenspace.compute_saliency_map(
            frame, None, center_bias=0.0, include_face=False
        )
        grid = screenspace.saliency_grid_from_map(sal, 16, 0.15)
        assert grid
        assert len(grid) <= 16 * 16
        for cell in grid:
            assert set(cell.keys()) == {"x", "y", "mag"}
            assert 0.0 <= cell["x"] <= 1.0
            assert 0.0 <= cell["y"] <= 1.0
            assert 0.15 <= cell["mag"] <= 1.0

    def test_grid_empty_on_zero_map(self):
        sal = np.zeros((60, 80), dtype=np.float32)
        assert screenspace.saliency_grid_from_map(sal, 16, 0.15) == []

    def test_grid_max_cell_near_patch(self):
        frame = _bright_patch_frame()
        sal, _ = screenspace.compute_saliency_map(
            frame, None, center_bias=0.0, include_face=False
        )
        grid = screenspace.saliency_grid_from_map(sal, 16, 0.15)
        best = max(grid, key=lambda c: c["mag"])
        assert best["mag"] == 1.0
        assert 0.70 <= best["x"] <= 0.92
        assert 0.60 <= best["y"] <= 0.90


def _run_scan(monkeypatch, frames, **kwargs):
    """Drive scan_attention over a synthetic (ts, frame) sequence."""

    def fake_scan(video_path, interval, callback, **_kw):
        for ts, frame in frames:
            if callback(ts, frame) is False:
                break

    monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", fake_scan)
    monkeypatch.setattr(
        screenspace_frames, "_probe_video_meta", lambda p: (30.0, float(len(frames)))
    )
    streamed = []
    results = screenspace.scan_attention(
        "/fake.mp4",
        None,
        interval_seconds=1.0,
        on_result=streamed.append,
        **kwargs,
    )
    return results, streamed


class TestScanAttention:
    def _patch_at(self, px, py):
        return _bright_patch_frame(w=160, h=120, px=px, py=py, size=20)

    def test_jumping_patch_emits_one_confirmed_shift(self, monkeypatch):
        frames = [(float(i), self._patch_at(10, 10)) for i in range(3)]
        frames += [(float(i), self._patch_at(130, 90)) for i in range(3, 7)]
        results, streamed = _run_scan(monkeypatch, frames, ema_alpha=1.0)

        assert len(results) == len(frames)
        for r in results:
            assert "saliency_grid" in r
            assert 0.0 <= r["peak_x"] <= 1.0
            assert 0.0 <= r["peak_y"] <= 1.0

        shifts = [r for r in results if r.get("shift")]
        assert len(shifts) == 1
        assert len(streamed) == 1
        shift = shifts[0]
        # Emitted on the frame where the jump began (confirmation back-stamps)
        assert shift["timestamp"] == 3.0
        # From the top-left patch to the bottom-right one
        assert shift["from_x"] < 0.35 and shift["from_y"] < 0.45
        assert shift["to_x"] > 0.65 and shift["to_y"] > 0.55
        assert shift["shift_distance"] > 0.5
        assert 0.05 <= shift["_confidence"] <= 1.0

    def test_streamed_payload_omits_grid(self, monkeypatch):
        frames = [(float(i), self._patch_at(10, 10)) for i in range(3)]
        frames += [(float(i), self._patch_at(130, 90)) for i in range(3, 7)]
        results, streamed = _run_scan(monkeypatch, frames, ema_alpha=1.0)
        assert streamed
        assert "saliency_grid" not in streamed[0]
        assert streamed[0]["shift"] is True
        # The returned result for the same frame keeps its grid
        match = [r for r in results if r["timestamp"] == streamed[0]["timestamp"]]
        assert match and "saliency_grid" in match[0]

    def test_static_sequence_emits_no_shifts(self, monkeypatch):
        frames = [(float(i), self._patch_at(70, 50)) for i in range(6)]
        results, streamed = _run_scan(monkeypatch, frames)
        assert len(results) == len(frames)
        assert streamed == []
        assert not any(r.get("shift") for r in results)

    def test_single_sample_jump_is_not_confirmed(self, monkeypatch):
        # Peak visits the far corner for one sample only, then returns: the
        # confirm counter (default 2) must reject the blip.
        frames = [(float(i), self._patch_at(10, 10)) for i in range(3)]
        frames.append((3.0, self._patch_at(130, 90)))
        frames += [(float(i), self._patch_at(10, 10)) for i in range(4, 7)]
        _, streamed = _run_scan(monkeypatch, frames, ema_alpha=1.0)
        assert streamed == []

    def test_cancel_stops_scan(self, monkeypatch):
        frames = [(float(i), self._patch_at(10, 10)) for i in range(10)]

        def fake_scan(video_path, interval, callback, **_kw):
            for ts, frame in frames:
                if callback(ts, frame) is False:
                    break

        monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", fake_scan)
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 10.0)
        )
        calls = [0]

        def cancel():
            calls[0] += 1
            return calls[0] > 3

        results = screenspace.scan_attention("/fake.mp4", cancel_flag=cancel)
        assert len(results) < len(frames)
