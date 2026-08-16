"""Tests for template matching primitives and heatmap."""

import numpy as np

import config
import screenspace
import screenspace_frames
import screenspace_heatmap
import screenspace_primitives
import screenspace_scans
import screenspace_tools
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


class TestCorrelationMapReuse:
    """The correlation map is the most expensive op here; compute it once."""

    @staticmethod
    def _frame_and_template():
        rng = np.random.RandomState(42)
        frame = rng.randint(0, 255, (100, 200, 3), dtype=np.uint8)
        return frame, frame[30:60, 80:140].copy()

    def test_supplied_corr_gives_the_same_matches(self):
        frame, template = self._frame_and_template()
        prepared = screenspace_primitives._prepare_template(template, None)
        corr = screenspace_primitives._template_correlation_map(frame, prepared)
        assert screenspace.match_template(
            frame, template, threshold=0.9, prepared=prepared, corr=corr
        ) == screenspace.match_template(frame, template, threshold=0.9)

    def test_tool_check_frame_computes_the_map_once(self, monkeypatch):
        frame, template = self._frame_and_template()
        calls = []
        real = screenspace_primitives._template_correlation_map

        def counting(frame_arg, prepared):
            calls.append(1)
            return real(frame_arg, prepared)

        monkeypatch.setattr(
            screenspace_primitives, "_template_correlation_map", counting
        )
        monkeypatch.setattr(screenspace_tools, "_template_correlation_map", counting)
        params = {"template_image": template, "threshold": 0.9}
        matched, details = screenspace_tools.TemplateTool().check_frame(
            frame, None, {"x": 0, "y": 0, "w": 200, "h": 100}, params
        )
        assert matched
        assert details["match_count"] >= 1
        assert len(calls) == 1


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
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
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
        info = screenspace.generate_heatmap_gif(self._matches(8), 200, 200, out)
        assert info is not None
        assert info["path"] == out
        assert info["frames"] == 8
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
        info = screenspace.generate_heatmap_gif(
            results, 200, 150, out, heatmap_type="change"
        )
        assert info is not None
        assert info["path"] == out
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
        info = screenspace.generate_rolling_heatmap_gif(self._matches(8), 200, 200, out)
        assert info is not None
        assert info["path"] == out
        assert info["frames"] == 8
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
        info = screenspace.generate_rolling_heatmap_gif(
            results, 200, 150, out, heatmap_type="change"
        )
        assert info is not None
        assert info["path"] == out
        assert (tmp_path / "rolling_change.gif").stat().st_size > 0

    def test_large_count_rolling_gif(self, tmp_path):
        # 47 results / 24 frames: the old floor-division bucketing dumped the
        # remainder onto the final frame; this just guards it still renders.
        out = str(tmp_path / "rolling_large.gif")
        info = screenspace.generate_rolling_heatmap_gif(
            self._matches(47), 200, 200, out
        )
        assert info is not None
        assert info["path"] == out
        assert info["frames"] == 24  # capped at num_frames, not one per result
        assert (tmp_path / "rolling_large.gif").stat().st_size > 0


class TestHeatmapSprite:
    """The sprite sheet the browser hover-scrubs in place of the unseekable GIF.

    It is rendered on demand from the GIF and never written to disk, so these
    cover the descriptor the task carries and the tiling the route performs.
    """

    def _matches(self, n):
        return [
            {
                "timestamp": float(i),
                "matches": [{"x": 10 + i, "y": 10, "w": 30, "h": 30, "score": 0.9}],
            }
            for i in range(n)
        ]

    def test_descriptor_is_intrinsic_to_the_gif(self, tmp_path):
        info = screenspace.generate_heatmap_gif(
            self._matches(8), 200, 100, str(tmp_path / "heatmap.gif")
        )
        assert info is not None
        # Frame count and cell shape come from the animation itself, never from
        # config, so a stored descriptor can't drift from a later-rendered sheet.
        assert info["frames"] == 8
        assert (info["w"], info["h"]) == (200, 100)
        # The GIF is the only file written.
        assert [p.name for p in tmp_path.iterdir()] == ["heatmap.gif"]

    def test_sprite_grid_wraps_at_the_configured_columns(self, monkeypatch):
        monkeypatch.setattr(config, "SCREENSPACE_HEATMAP_SPRITE_COLS", 6)
        assert screenspace.sprite_grid(8) == (6, 2)
        assert screenspace.sprite_grid(24) == (6, 4)
        # Fewer frames than columns: no empty columns.
        assert screenspace.sprite_grid(3) == (3, 1)

    def test_sprite_bytes_tile_the_gif_frames(self, tmp_path, monkeypatch):
        import io

        from PIL import Image

        monkeypatch.setattr(config, "SCREENSPACE_HEATMAP_SPRITE_FRAME_WIDTH", 80)
        gif = str(tmp_path / "heatmap.gif")
        info = screenspace.generate_heatmap_gif(self._matches(8), 200, 100, gif)
        assert info is not None

        data = screenspace.build_gif_sprite_bytes(gif, 6)
        assert data is not None
        with Image.open(io.BytesIO(data)) as sheet:
            # 8 frames at 6 columns → 2 rows; frames downscaled 200x100 → 80x40.
            assert sheet.size == (6 * 80, 2 * 40)

    def test_sprite_bytes_never_upscale(self, tmp_path, monkeypatch):
        import io

        from PIL import Image

        monkeypatch.setattr(config, "SCREENSPACE_HEATMAP_SPRITE_FRAME_WIDTH", 320)
        gif = str(tmp_path / "rolling.gif")
        screenspace.generate_rolling_heatmap_gif(self._matches(4), 48, 24, gif)
        data = screenspace.build_gif_sprite_bytes(gif, 6)
        assert data is not None
        with Image.open(io.BytesIO(data)) as sheet:
            assert sheet.size == (4 * 48, 1 * 24)

    def test_frame_count_reflects_what_pil_actually_wrote(self, tmp_path):
        """PIL collapses a run of identical frames, so the descriptor must be
        re-read from the file — otherwise the frontend scrubs to cells the
        sprite sheet doesn't have."""
        # Every result paints the same cell, so every rendered frame is identical.
        results = [
            {"timestamp": float(i), "change_grid": [{"x": 0.5, "y": 0.5, "mag": 0.7}]}
            for i in range(8)
        ]
        gif = str(tmp_path / "heatmap.gif")
        info = screenspace.generate_heatmap_gif(
            results, 100, 50, gif, heatmap_type="change"
        )
        assert info is not None
        assert info["frames"] == 1  # not the 8 buckets we handed it

        data = screenspace.build_gif_sprite_bytes(gif, 6)
        assert data is not None
        import io

        from PIL import Image

        with Image.open(io.BytesIO(data)) as sheet:
            assert sheet.size[0] // 100 == 1  # one column of cells, matching

    def test_single_frame_gif_gets_no_sprite_descriptor(self, tmp_path, monkeypatch):
        """A collapsed one-frame animation is a still: no scrub geometry, so the
        thumb degrades to plain playback instead of scrubbing a 1-cell sheet."""
        import screenspace_worker

        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        gif = str(tmp_path / "heatmap_ss_x.gif")
        info = screenspace.generate_heatmap_gif(
            [
                {
                    "timestamp": float(i),
                    "change_grid": [{"x": 0.5, "y": 0.5, "mag": 0.7}],
                }
                for i in range(8)
            ],
            100,
            50,
            gif,
            heatmap_type="change",
        )
        attachments = screenspace_worker._heatmap_gif_attachments(info, "heatmap_gif")
        assert attachments == {"heatmap_gif": "heatmap_ss_x.gif"}

    def test_unreadable_gif_yields_no_sprite(self, tmp_path):
        broken = tmp_path / "heatmap_broken.gif"
        broken.write_bytes(b"not a gif")
        assert screenspace.build_gif_sprite_bytes(str(broken), 6) is None
        assert screenspace.build_gif_sprite_bytes(str(tmp_path / "gone.gif"), 6) is None


class TestHeatmapPngWriteFailure:
    """A PNG that never landed must not be reported as a generated artifact."""

    def _results(self):
        return [{"timestamp": 0.0, "change_grid": [{"x": 0.5, "y": 0.5, "mag": 0.7}]}]

    def test_missing_directory_returns_none(self, tmp_path):
        out = str(tmp_path / "nope" / "heatmap.png")
        assert screenspace.generate_change_heatmap(self._results(), 64, 64, out) is None

    def test_encode_failure_returns_none(self, tmp_path, monkeypatch):
        monkeypatch.setattr(
            screenspace_heatmap.cv2, "imencode", lambda *a, **kw: (False, None)
        )
        out = str(tmp_path / "heatmap.png")
        assert screenspace.generate_change_heatmap(self._results(), 64, 64, out) is None
        assert not (tmp_path / "heatmap.png").exists()


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
