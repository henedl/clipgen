"""Tests for template matching primitives and heatmap."""

import numpy as np

import config
import screenspace
import screenspace_heatmap
import screenspace_scans
from _ss_helpers import _make_icon, _make_icon_frame


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


class TestScanTemplateControls:
    """Cover template_scale addition to scan_template."""

    def _patch_single_frame(self, monkeypatch, frame: np.ndarray) -> None:
        def fake_scan(video_path, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", fake_scan)
        monkeypatch.setattr(
            screenspace_scans, "_probe_video_meta", lambda p: (30.0, 1.0)
        )

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


class TestGenerateHeatmapGif:
    def _matches(self, n):
        return [
            {
                "timestamp": float(i),
                "matches": [{"x": 10 + i, "y": 10, "w": 30, "h": 30, "score": 0.9}],
            }
            for i in range(n)
        ]

    def test_basic_gif(self, tmp_path):
        out = str(tmp_path / "heatmap.gif")
        path = screenspace.generate_heatmap_gif(self._matches(8), 200, 200, out)
        assert path == out
        assert (tmp_path / "heatmap.gif").stat().st_size > 0

    def test_empty_returns_none(self, tmp_path):
        out = str(tmp_path / "heatmap.gif")
        assert screenspace.generate_heatmap_gif([], 200, 200, out) is None

    def test_single_result_returns_none(self, tmp_path):
        out = str(tmp_path / "heatmap.gif")
        assert screenspace.generate_heatmap_gif(self._matches(1), 200, 200, out) is None

    def test_change_type_gif(self, tmp_path):
        results = [
            {"timestamp": float(i), "change_grid": [{"x": 0.5, "y": 0.5, "mag": 0.7}]}
            for i in range(8)
        ]
        out = str(tmp_path / "heatmap_change.gif")
        path = screenspace.generate_heatmap_gif(
            results, 200, 150, out, heatmap_type="change"
        )
        assert path == out
        assert (tmp_path / "heatmap_change.gif").stat().st_size > 0


class TestGenerateRollingHeatmapGif:
    def _matches(self, n):
        return [
            {
                "timestamp": float(i),
                "matches": [{"x": 10 + i, "y": 10, "w": 30, "h": 30, "score": 0.9}],
            }
            for i in range(n)
        ]

    def test_basic_rolling_gif(self, tmp_path):
        out = str(tmp_path / "rolling.gif")
        path = screenspace.generate_rolling_heatmap_gif(self._matches(8), 200, 200, out)
        assert path == out
        assert (tmp_path / "rolling.gif").is_file()
        assert (tmp_path / "rolling.gif").stat().st_size > 0

    def test_empty_returns_none(self, tmp_path):
        out = str(tmp_path / "rolling.gif")
        assert screenspace.generate_rolling_heatmap_gif([], 200, 200, out) is None

    def test_single_result_returns_none(self, tmp_path):
        out = str(tmp_path / "rolling.gif")
        assert (
            screenspace.generate_rolling_heatmap_gif(self._matches(1), 200, 200, out)
            is None
        )

    def test_change_type_rolling_gif(self, tmp_path):
        results = [
            {"timestamp": float(i), "change_grid": [{"x": 0.5, "y": 0.5, "mag": 0.7}]}
            for i in range(8)
        ]
        out = str(tmp_path / "rolling_change.gif")
        path = screenspace.generate_rolling_heatmap_gif(
            results, 200, 150, out, heatmap_type="change"
        )
        assert path == out
        assert (tmp_path / "rolling_change.gif").stat().st_size > 0

    def test_large_count_rolling_gif(self, tmp_path):
        # 47 results / 24 frames: the old floor-division bucketing dumped the
        # remainder onto the final frame; this just guards it still renders.
        out = str(tmp_path / "rolling_large.gif")
        path = screenspace.generate_rolling_heatmap_gif(
            self._matches(47), 200, 200, out
        )
        assert path == out
        assert (tmp_path / "rolling_large.gif").stat().st_size > 0


class TestFrameBucketBounds:
    """_frame_bucket_bounds spreads results evenly, no remainder on the last frame."""

    def _buckets(self, total, num_frames):
        return [
            screenspace_heatmap._frame_bucket_bounds(i, total, num_frames)
            for i in range(num_frames)
        ]

    def test_exact_multiple(self):
        buckets = self._buckets(48, 24)
        assert all(end - start == 2 for start, end in buckets)

    def test_non_multiple_counts_are_even(self):
        for total in (25, 47, 100):
            num_frames = 24
            buckets = self._buckets(total, num_frames)
            sizes = [end - start for start, end in buckets]
            # Contiguous, fully covering [0, total), adjacent sizes differ by ≤ 1.
            assert buckets[0][0] == 0
            assert buckets[-1][1] == total
            for i in range(num_frames - 1):
                assert buckets[i][1] == buckets[i + 1][0]
            assert max(sizes) - min(sizes) <= 1
            assert sum(sizes) == total


class TestHeatmapConfigConstants:
    def test_rolling_window_and_change_grid_present(self):
        assert isinstance(config.SCREENSPACE_HEATMAP_ROLLING_WINDOW, int)
        assert config.SCREENSPACE_HEATMAP_ROLLING_WINDOW >= 1
        assert isinstance(config.SCREENSPACE_CHANGE_HEATMAP_GRID, int)
        assert config.SCREENSPACE_CHANGE_HEATMAP_GRID >= 2
