"""Shaped-region mask primitives and per-tool masked-statistics behavior."""

import numpy as np

import screenspace_ocr
from screenspace_ocr import _ocr_region_readings
from screenspace_primitives import (
    average_color_hsv,
    color_present,
    compare_scene_fingerprints,
    compute_frame_diff,
    compute_optical_flow,
    compute_scene_fingerprint,
    denormalize_region,
    filter_matches_by_region_mask,
    mask_points_key,
    point_in_mask_points,
    region_mask_for,
    region_masker,
)

# A right triangle covering the lower-left half of the bbox (u, v pairs).
TRIANGLE = [[0.0, 0.0], [0.0, 1.0], [1.0, 1.0]]


class TestRegionMaskFor:
    def test_rect_region_returns_none(self):
        assert region_mask_for({"x": 0, "y": 0, "w": 10, "h": 10}, 10, 10) is None

    def test_empty_points_returns_none(self):
        region = {"x": 0, "y": 0, "w": 10, "h": 10, "mask_points": []}
        assert region_mask_for(region, 10, 10) is None

    def test_zero_dims_return_none(self):
        region = {"mask_points": [TRIANGLE]}
        assert region_mask_for(region, 0, 10) is None
        assert region_mask_for(region, 10, 0) is None

    def test_triangle_mask_shape_and_coverage(self):
        region = {"mask_points": [TRIANGLE]}
        mask = region_mask_for(region, 100, 200)
        assert mask is not None
        assert mask.shape == (100, 200)
        assert mask.dtype == np.uint8
        assert set(np.unique(mask)) <= {0, 255}
        # A half-bbox triangle fills ~50% of the mask.
        coverage = np.count_nonzero(mask) / mask.size
        assert 0.45 <= coverage <= 0.55
        # Lower-left corner inside, upper-right corner outside.
        assert mask[95, 5] == 255
        assert mask[5, 195] == 0

    def test_full_bbox_polygon_fills_everything(self):
        region = {"mask_points": [[[0, 0], [1, 0], [1, 1], [0, 1]]]}
        mask = region_mask_for(region, 20, 30)
        assert mask is not None
        assert np.count_nonzero(mask) == mask.size

    def test_mask_scales_with_raster_size(self):
        """Same points rasterized at two sizes cover the same fraction."""
        region = {"mask_points": [TRIANGLE]}
        small = region_mask_for(region, 32, 32)
        large = region_mask_for(region, 256, 256)
        assert small is not None and large is not None
        frac_small = np.count_nonzero(small) / small.size
        frac_large = np.count_nonzero(large) / large.size
        assert abs(frac_small - frac_large) < 0.05

    def test_degenerate_polygon_rasterizes_near_empty(self):
        """A collinear 'polygon' has (almost) no filled area."""
        region = {"mask_points": [[[0.1, 0.1], [0.5, 0.5], [0.9, 0.9]]]}
        mask = region_mask_for(region, 100, 100)
        assert mask is not None
        assert np.count_nonzero(mask) / mask.size < 0.05

    def test_disjoint_contours_union(self):
        """A multi-part shape (merge/add results) masks every part."""
        region = {
            "mask_points": [
                [[0.0, 0.0], [0.25, 0.0], [0.25, 1.0], [0.0, 1.0]],
                [[0.75, 0.0], [1.0, 0.0], [1.0, 1.0], [0.75, 1.0]],
            ]
        }
        mask = region_mask_for(region, 100, 100)
        assert mask is not None
        assert mask[50, 10] == 255  # left part
        assert mask[50, 50] == 0  # gap between parts
        assert mask[50, 90] == 255  # right part

    def test_overlapping_contours_stay_filled(self):
        """Overlaps fill per contour (union), never XOR out."""
        region = {
            "mask_points": [
                [[0.0, 0.0], [0.6, 0.0], [0.6, 1.0], [0.0, 1.0]],
                [[0.4, 0.0], [1.0, 0.0], [1.0, 1.0], [0.4, 1.0]],
            ]
        }
        mask = region_mask_for(region, 100, 100)
        assert mask is not None
        assert mask[50, 50] == 255  # inside both contours
        assert np.count_nonzero(mask) == mask.size

    def test_short_contours_are_skipped(self):
        region = {"mask_points": [[[0.0, 0.0], [1.0, 1.0]]]}
        mask = region_mask_for(region, 50, 50)
        assert mask is not None
        assert np.count_nonzero(mask) == 0


class TestPointInMaskPoints:
    def test_inside_and_outside_triangle(self):
        assert point_in_mask_points(0.2, 0.8, [TRIANGLE]) is True
        assert point_in_mask_points(0.8, 0.2, [TRIANGLE]) is False

    def test_too_few_points(self):
        assert point_in_mask_points(0.5, 0.5, []) is False
        assert point_in_mask_points(0.5, 0.5, [[[0, 0], [1, 1]]]) is False

    def test_square_contains_center(self):
        square = [[0.25, 0.25], [0.75, 0.25], [0.75, 0.75], [0.25, 0.75]]
        assert point_in_mask_points(0.5, 0.5, [square]) is True
        assert point_in_mask_points(0.1, 0.5, [square]) is False

    def test_concave_polygon(self):
        # A "C" shape: the notch on the right side is outside.
        c_shape = [
            [0.0, 0.0],
            [1.0, 0.0],
            [1.0, 0.25],
            [0.4, 0.25],
            [0.4, 0.75],
            [1.0, 0.75],
            [1.0, 1.0],
            [0.0, 1.0],
        ]
        assert point_in_mask_points(0.2, 0.5, [c_shape]) is True
        assert point_in_mask_points(0.8, 0.5, [c_shape]) is False

    def test_multi_contour_any_hit(self):
        contours = [
            [[0.0, 0.0], [0.25, 0.0], [0.25, 1.0], [0.0, 1.0]],
            [[0.75, 0.0], [1.0, 0.0], [1.0, 1.0], [0.75, 1.0]],
        ]
        assert point_in_mask_points(0.1, 0.5, contours) is True
        assert point_in_mask_points(0.9, 0.5, contours) is True
        assert point_in_mask_points(0.5, 0.5, contours) is False

    def test_none_contours(self):
        assert point_in_mask_points(0.5, 0.5, None) is False


class TestDenormalizePassthrough:
    def test_rect_region_unchanged(self):
        out = denormalize_region({"x": 0.1, "y": 0.2, "w": 0.5, "h": 0.25}, 1000, 800)
        assert out == {"x": 100, "y": 160, "w": 500, "h": 200}
        assert "mask_points" not in out
        assert "shape" not in out

    def test_points_and_shape_pass_through_verbatim(self):
        region = {
            "x": 0.1,
            "y": 0.2,
            "w": 0.5,
            "h": 0.25,
            "points": [TRIANGLE],
            "shape": "lasso",
        }
        out = denormalize_region(region, 1000, 800)
        assert out["mask_points"] == [TRIANGLE]
        assert out["shape"] == "lasso"
        assert out["x"] == 100 and out["w"] == 500


# Left half of the bbox: a rectangle polygon covering u in [0, 0.5].
LEFT_HALF = [[0.0, 0.0], [0.5, 0.0], [0.5, 1.0], [0.0, 1.0]]


def _left_red_right_blue(h=64, w=64):
    """BGR crop: pure red left half, pure blue right half."""
    crop = np.zeros((h, w, 3), dtype=np.uint8)
    crop[:, : w // 2, 2] = 255  # red (BGR)
    crop[:, w // 2 :, 0] = 255  # blue
    return crop


class TestMaskedColor:
    def test_average_restricted_to_mask(self):
        crop = _left_red_right_blue()
        mask = region_mask_for({"mask_points": [LEFT_HALF]}, 64, 64)
        avg = average_color_hsv(crop, mask=mask)
        # Pure red in OpenCV HSV is hue 0, s=v=255; the unmasked mean would be
        # a red/blue blend with far lower saturation coherence.
        assert avg["h"] < 5 or avg["h"] > 175
        assert avg["s"] > 250 and avg["v"] > 250

    def test_average_masked_downscale_path(self):
        crop = _left_red_right_blue(200, 200)  # forces the <=64 INTER_AREA path
        mask = region_mask_for({"mask_points": [LEFT_HALF]}, 200, 200)
        avg = average_color_hsv(crop, mask=mask)
        assert avg["s"] > 240 and avg["v"] > 240

    def test_presence_coverage_uses_mask_denominator(self):
        crop = _left_red_right_blue()
        red: dict[str, float] = {"h": 0, "s": 255, "v": 255}
        tol: dict[str, float] = {"h": 5, "s": 20, "v": 20}
        _, coverage_full = color_present(crop, red, tol)
        mask = region_mask_for({"mask_points": [LEFT_HALF]}, 64, 64)
        _, coverage_masked = color_present(crop, red, tol, mask=mask)
        # Red fills ~half the rect but ~all of the left-half polygon.
        assert 0.45 <= coverage_full <= 0.55
        assert coverage_masked > 0.95

    def test_presence_excludes_matches_outside_mask(self):
        crop = _left_red_right_blue()
        blue: dict[str, float] = {"h": 120, "s": 255, "v": 255}
        tol: dict[str, float] = {"h": 5, "s": 20, "v": 20}
        matched_full, _ = color_present(crop, blue, tol)
        mask = region_mask_for({"mask_points": [LEFT_HALF]}, 64, 64)
        # fillPoly edges are inclusive, so the boundary column can leak a few
        # blue pixels into the mask — gate on coverage rather than any-pixel.
        matched_masked, coverage = color_present(
            crop, blue, tol, min_coverage=0.05, mask=mask
        )
        assert matched_full is True
        assert matched_masked is False and coverage < 0.05


class TestMaskedChange:
    def test_change_outside_mask_ignored(self):
        a = np.zeros((64, 64, 3), dtype=np.uint8)
        b = a.copy()
        b[:, 40:, :] = 255  # change only in the right half
        mask = region_mask_for({"mask_points": [LEFT_HALF]}, 64, 64)
        assert compute_frame_diff(a, b) > 0.2
        assert compute_frame_diff(a, b, mask=mask) < 0.05

    def test_change_inside_mask_amplified_by_mask_denominator(self):
        a = np.zeros((64, 64, 3), dtype=np.uint8)
        b = a.copy()
        b[:, :24, :] = 255  # change inside the left-half polygon
        mask = region_mask_for({"mask_points": [LEFT_HALF]}, 64, 64)
        full = compute_frame_diff(a, b)
        masked = compute_frame_diff(a, b, mask=mask)
        # Same changed pixels over half the denominator ≈ double the ratio.
        assert masked > full * 1.5


class TestMaskedFlow:
    def test_flow_stats_restricted_to_mask(self):
        rng = np.random.RandomState(7)
        prev = rng.randint(0, 255, (64, 64), dtype=np.uint8)
        curr = prev.copy()
        # Shift only the right half rightward to create motion outside the mask.
        curr[:, 33:] = prev[:, 32:63]
        full = compute_optical_flow(prev, curr)
        mask = region_mask_for({"mask_points": [LEFT_HALF]}, 64, 64)
        masked = compute_optical_flow(prev, curr, mask=mask)
        assert masked["magnitude"] < full["magnitude"]

    def test_flow_grid_cells_outside_mask_dropped(self):
        rng = np.random.RandomState(7)
        prev = rng.randint(0, 255, (128, 128), dtype=np.uint8)
        curr = np.roll(prev, 3, axis=1)  # uniform motion everywhere
        mask = region_mask_for({"mask_points": [LEFT_HALF]}, 128, 128)
        full = compute_optical_flow(prev, curr, return_grid=True)
        masked = compute_optical_flow(prev, curr, return_grid=True, mask=mask)
        assert masked["flow_grid"]
        assert len(masked["flow_grid"]) < len(full["flow_grid"])
        assert all(cell["x"] <= 0.55 for cell in masked["flow_grid"])


class TestMaskedScene:
    def test_masked_fingerprints_ignore_outside_content(self):
        base = _left_red_right_blue(128, 128)
        variant = base.copy()
        variant[:, 64:, :] = 0  # right half differs wildly
        mask = region_mask_for({"mask_points": [LEFT_HALF]}, 128, 128)
        fp_base_m = compute_scene_fingerprint(base, mask=mask)
        fp_var_m = compute_scene_fingerprint(variant, mask=mask)
        score_masked = compare_scene_fingerprints(fp_base_m, fp_var_m)
        score_full = compare_scene_fingerprints(
            compute_scene_fingerprint(base), compute_scene_fingerprint(variant)
        )
        # Inside the polygon the two frames are identical (up to Canny edge
        # bleed at the inclusive polygon boundary).
        assert score_masked > 0.95
        assert score_masked > score_full


class TestTemplateMatchFilter:
    REGION = {
        "x": 100,
        "y": 100,
        "w": 200,
        "h": 100,
        "mask_points": [LEFT_HALF],
    }

    def test_matches_outside_polygon_dropped(self):
        inside = {"x": 120, "y": 130, "w": 20, "h": 20, "score": 0.9}
        outside = {"x": 260, "y": 130, "w": 20, "h": 20, "score": 0.95}
        kept = filter_matches_by_region_mask([inside, outside], self.REGION)
        assert kept == [inside]

    def test_rect_region_passes_through(self):
        rect = {"x": 100, "y": 100, "w": 200, "h": 100}
        matches = [{"x": 999, "y": 999, "w": 5, "h": 5, "score": 0.8}]
        assert filter_matches_by_region_mask(matches, rect) == matches


class TestMaskedOcrReadings:
    def test_readings_outside_polygon_dropped(self, monkeypatch):
        # Two fake readings: one centered in the left half, one in the right.
        left = ([[4, 4], [24, 4], [24, 14], [4, 14]], "inside", 0.9)
        right = ([[40, 4], [60, 4], [60, 14], [40, 14]], "outside", 0.9)
        monkeypatch.setattr(
            screenspace_ocr, "_ocr_readtext", lambda langs, img, **kw: [left, right]
        )
        crop = np.zeros((20, 64, 3), dtype=np.uint8)
        unfiltered = _ocr_region_readings(crop)
        filtered = _ocr_region_readings(crop, mask_points=[LEFT_HALF])
        assert len(unfiltered) == 2
        assert [r[1] for r in filtered] == ["inside"]

    def test_filter_uses_post_preprocess_shape(self, monkeypatch):
        """Centers are normalized by the upscaled image when preprocess runs."""
        captured = {}

        def fake_readtext(langs, img, **kw):
            captured["shape"] = img.shape[:2]
            h, w = img.shape[:2]
            # A reading centered at 25% width — inside LEFT_HALF regardless
            # of the upscale, but only if normalized by the *upscaled* size.
            return [([[0, 0], [w // 2, 0], [w // 2, h], [0, h]], "ok", 0.9)]

        monkeypatch.setattr(screenspace_ocr, "_ocr_readtext", fake_readtext)
        crop = np.zeros((20, 200, 3), dtype=np.uint8)  # short → preprocess upscales
        filtered = _ocr_region_readings(crop, preprocess=True, mask_points=[LEFT_HALF])
        assert captured["shape"][0] > 20  # preprocess actually upscaled
        assert [r[1] for r in filtered] == ["ok"]


class TestMaskPointsKey:
    def test_hashable_and_distinguishes_shapes(self):
        a = mask_points_key([TRIANGLE])
        b = mask_points_key([TRIANGLE, LEFT_HALF])
        assert hash(a) != hash(b) or a != b
        assert a != b
        assert mask_points_key(None) == ()
        assert mask_points_key([]) == ()
        {a: 1, b: 2}  # usable as dict keys


class TestRegionMasker:
    def test_caches_per_shape_and_handles_rects(self):
        masker = region_masker({"mask_points": [TRIANGLE]})
        a = masker(np.zeros((40, 40, 3), dtype=np.uint8))
        b = masker(np.zeros((40, 40, 3), dtype=np.uint8))
        assert a is b  # cached
        c = masker(np.zeros((20, 20, 3), dtype=np.uint8))
        assert c is not None and c.shape == (20, 20)
        rect_masker = region_masker({"x": 0, "y": 0, "w": 10, "h": 10})
        assert rect_masker(np.zeros((40, 40, 3), dtype=np.uint8)) is None
