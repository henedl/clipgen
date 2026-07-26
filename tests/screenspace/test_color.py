"""Tests for region extraction and colour primitives."""

import numpy as np
import pytest

import screenspace
from _ss_helpers import _gray_with_red_patch


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


class TestResolveRegionRequest:
    MANIFEST = {
        "regions": {"hud": {"x": 0.0, "y": 0.0, "w": 0.5, "h": 0.5}},
        "stashes": [
            {
                "id": "stash_a",
                "regions": {"hud": {"x": 0.1, "y": 0.1, "w": 0.1, "h": 0.1}},
            },
            {
                "id": "stash_b",
                "regions": {"hud": {"x": 0.2, "y": 0.2, "w": 0.2, "h": 0.2}},
            },
        ],
    }

    def test_bare_name_prefers_active_over_stash(self):
        name, region = screenspace.resolve_region_request("hud", None, self.MANIFEST)
        assert name == "hud"
        assert region["w"] == 0.5  # active "hud", not a stashed one

    def test_bare_name_falls_back_to_stash(self):
        manifest = {"regions": {}, "stashes": self.MANIFEST["stashes"]}
        name, region = screenspace.resolve_region_request("hud", None, manifest)
        assert name == "hud"
        # No active match: first stash wins (no flattening, but order is preserved).
        assert region["w"] == 0.1

    def test_full_frame_by_bare_name(self):
        name, region = screenspace.resolve_region_request(
            "full_frame", None, self.MANIFEST
        )
        assert name == screenspace.FULL_FRAME_REGION_NAME
        assert region == screenspace.FULL_FRAME_REGION

    def test_full_frame_by_ref(self):
        name, region = screenspace.resolve_region_request(
            "", {"source": "full_frame"}, self.MANIFEST
        )
        assert name == screenspace.FULL_FRAME_REGION_NAME
        assert region["w"] == 1.0

    def test_active_ref(self):
        _name, region = screenspace.resolve_region_request(
            "", {"source": "active", "name": "hud"}, self.MANIFEST
        )
        assert region["w"] == 0.5

    def test_stash_ref_honors_stash_id(self):
        _, region = screenspace.resolve_region_request(
            "", {"source": "stash", "stash_id": "stash_a", "name": "hud"}, self.MANIFEST
        )
        assert region["w"] == 0.1  # stash_a, not last-write-wins stash_b

    def test_unknown_name_raises(self):
        with pytest.raises(ValueError, match="Region 'ghost' not found"):
            screenspace.resolve_region_request("ghost", None, self.MANIFEST)

    def test_stash_ref_missing_stash_id_raises(self):
        with pytest.raises(ValueError, match="region_ref.stash_id is required"):
            screenspace.resolve_region_request(
                "", {"source": "stash", "name": "hud"}, self.MANIFEST
            )

    def test_unknown_stash_raises(self):
        with pytest.raises(ValueError, match="Stash 'nope' not found"):
            screenspace.resolve_region_request(
                "",
                {"source": "stash", "stash_id": "nope", "name": "hud"},
                self.MANIFEST,
            )

    def test_invalid_source_raises(self):
        with pytest.raises(ValueError, match="region_ref.source must be"):
            screenspace.resolve_region_request(
                "", {"source": "bogus", "name": "hud"}, self.MANIFEST
            )


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


class TestColorPresent:
    target = {"h": 0.0, "s": 255.0, "v": 139.0}  # dark red in OpenCV HSV
    tol = {"h": 10.0, "s": 60.0, "v": 60.0}

    def test_small_patch_detected(self):
        # 1% red patch: averaged away by color_matches, but present per-pixel.
        region = _gray_with_red_patch(10)
        avg_matched, _ = screenspace.color_matches(region, self.target, self.tol)
        assert not avg_matched
        matched, coverage = screenspace.color_present(region, self.target, self.tol)
        assert matched
        assert coverage == pytest.approx(0.01, abs=1e-3)

    def test_min_coverage_gate(self):
        region = _gray_with_red_patch(10)  # ~1% coverage
        matched_hi, cov = screenspace.color_present(
            region, self.target, self.tol, min_coverage=0.05
        )
        assert not matched_hi
        assert cov < 0.05
        matched_lo, _ = screenspace.color_present(
            region, self.target, self.tol, min_coverage=0.005
        )
        assert matched_lo

    def test_absent_color_misses(self):
        region = np.full((100, 100, 3), 128, dtype=np.uint8)  # all gray, no red
        matched, coverage = screenspace.color_present(region, self.target, self.tol)
        assert not matched
        assert coverage == 0.0

    def test_empty_region(self):
        region = np.zeros((0, 0, 3), dtype=np.uint8)
        matched, coverage = screenspace.color_present(region, self.target, self.tol)
        assert not matched
        assert coverage == 0.0
