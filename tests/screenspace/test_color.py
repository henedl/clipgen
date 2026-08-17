"""Tests for region extraction and colour primitives."""

import cv2
import numpy as np
import pytest

import screenspace
import screenspace_primitives
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

    def test_downsample_matches_one_step_mean(self):
        """Color downsample must be a single INTER_AREA resize.

        The two-pass integer-ratio split used by pHash shifts the HSV mean
        enough to flip color_matches at default tolerances.
        """
        rng = np.random.default_rng(11)
        for h, w in ((720, 1280), (1280, 720), (45, 61)):
            frame = rng.integers(0, 256, (h, w, 3), dtype=np.uint8)
            got = screenspace.average_color_hsv(frame)
            new_w, new_h = min(w, 64), min(h, 64)
            if h > 64 or w > 64:
                one = cv2.resize(frame, (new_w, new_h), interpolation=cv2.INTER_AREA)
            else:
                one = frame
            hsv = cv2.cvtColor(one, cv2.COLOR_BGR2HSV)
            mean = hsv.mean(axis=(0, 1))
            assert abs(got["h"] - float(mean[0])) < 1e-6, (h, w, got, mean)
            assert abs(got["s"] - float(mean[1])) < 1e-6, (h, w, got, mean)
            assert abs(got["v"] - float(mean[2])) < 1e-6, (h, w, got, mean)


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


def _float32_match_mask(region_pixels, target_color, tolerance):
    """The per-pixel float32 test ``color_present`` used before the band rewrite.

    Kept verbatim as the oracle: the integer-band path is a speedup, so it must
    agree *exactly*, not approximately.
    """
    hsv = cv2.cvtColor(region_pixels, cv2.COLOR_BGR2HSV)
    h = hsv[..., 0].astype(np.float32)
    s = hsv[..., 1].astype(np.float32)
    v = hsv[..., 2].astype(np.float32)
    hue_diff = np.abs(h - float(target_color["h"]))
    hue_dist = np.minimum(hue_diff, 180.0 - hue_diff)
    return (
        (hue_dist <= tolerance["h"])
        & (np.abs(s - float(target_color["s"])) <= tolerance["s"])
        & (np.abs(v - float(target_color["v"])) <= tolerance["v"])
    )


def _reference_color_present(
    region_pixels, target_color, tolerance, min_coverage=0.0, mask=None
):
    """``color_present`` semantics on top of the float32 oracle mask."""
    if region_pixels.size == 0:
        return False, 0.0
    match = _float32_match_mask(region_pixels, target_color, tolerance)
    if mask is not None and not np.any(mask):
        mask = None
    denom = match.size
    if mask is not None:
        match &= mask > 0
        denom = int(np.count_nonzero(mask))
    count = int(np.count_nonzero(match))
    coverage = count / denom if denom else 0.0
    matched = count > 0 if min_coverage <= 0 else coverage >= min_coverage
    return matched, float(coverage)


class TestColorPresentBands:
    """The uint8 ``cv2.inRange`` bands must reproduce the float32 test exactly."""

    @pytest.mark.parametrize("hue", [0.0, 5.0, 90.0, 175.5, 179.0])
    @pytest.mark.parametrize("tol_hue", [0.0, 10.0, 89.0, 90.0, 95.0])
    @pytest.mark.parametrize("sat_val", [0.0, 7.5, 128.0, 255.0])
    @pytest.mark.parametrize("tol_sat_val", [0.0, 0.5, 60.0, 255.0])
    def test_matches_float32_oracle(self, hue, tol_hue, sat_val, tol_sat_val):
        region = np.random.RandomState(11).randint(0, 256, (24, 32, 3), dtype=np.uint8)
        target = {"h": hue, "s": sat_val, "v": sat_val}
        tol = {"h": tol_hue, "s": tol_sat_val, "v": tol_sat_val}
        assert screenspace.color_present(
            region, target, tol
        ) == _reference_color_present(region, target, tol)

    def test_matches_oracle_with_mask_and_min_coverage(self):
        region = np.random.RandomState(5).randint(0, 256, (40, 40, 3), dtype=np.uint8)
        mask = np.zeros((40, 40), dtype=np.uint8)
        mask[5:20, 5:20] = 255
        target = {"h": 90.0, "s": 128.0, "v": 128.0}
        tol = {"h": 20.0, "s": 80.0, "v": 80.0}
        for min_coverage in (0.0, 0.05, 0.9):
            assert screenspace.color_present(
                region, target, tol, min_coverage, mask
            ) == _reference_color_present(region, target, tol, min_coverage, mask)

    def test_empty_mask_falls_back_to_full_rect(self):
        region = _gray_with_red_patch(10)
        empty = np.zeros(region.shape[:2], dtype=np.uint8)
        target = {"h": 0.0, "s": 255.0, "v": 139.0}
        tol = {"h": 10.0, "s": 60.0, "v": 60.0}
        assert screenspace.color_present(
            region, target, tol, 0.0, empty
        ) == screenspace.color_present(region, target, tol)

    def test_uniform_frames_agree_with_oracle(self):
        target = {"h": 60.0, "s": 200.0, "v": 200.0}
        tol = {"h": 15.0, "s": 40.0, "v": 40.0}
        for fill in (0, 128, 255):
            region = np.full((16, 16, 3), fill, dtype=np.uint8)
            assert screenspace.color_present(
                region, target, tol
            ) == _reference_color_present(region, target, tol)

    def test_impossible_tolerance_yields_no_bands(self):
        # Nothing can match a saturation 300 apart from any uint8 value.
        assert (
            screenspace_primitives._hsv_match_bands(90.0, 300.0, 128.0, 10.0, 5.0, 60.0)
            == ()
        )
        region = np.random.RandomState(2).randint(0, 256, (8, 8, 3), dtype=np.uint8)
        matched, coverage = screenspace.color_present(
            region,
            {"h": 90.0, "s": 300.0, "v": 128.0},
            {"h": 10.0, "s": 5.0, "v": 60.0},
        )
        assert not matched
        assert coverage == 0.0

    def test_hue_wraparound_splits_into_two_bands(self):
        bands = screenspace_primitives._hsv_match_bands(
            2.0, 128.0, 128.0, 5.0, 255.0, 255.0
        )
        assert len(bands) == 2
        hue_ranges = sorted((lower[0], upper[0]) for lower, upper in bands)
        assert hue_ranges[0][0] == 0
        assert hue_ranges[-1][1] == 255

    def test_wide_hue_tolerance_is_a_single_band(self):
        bands = screenspace_primitives._hsv_match_bands(
            90.0, 128.0, 128.0, 90.0, 255.0, 255.0
        )
        assert len(bands) == 1
        assert bands[0][0][0] == 0 and bands[0][1][0] == 255
