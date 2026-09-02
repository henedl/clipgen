"""Tests for the Shape tool: scale-swept edge matching primitives, scan, tool."""

from concurrent.futures import ThreadPoolExecutor
from unittest.mock import Mock

import cv2
import numpy as np
import pytest

import config
import screenspace
import screenspace_frames
import screenspace_primitives
import screenspace_scans

_BG = 40  # flat dark background with strong luma contrast to white outlines


def _make_outline(size: int, color=(255, 255, 255)) -> np.ndarray:
    """Square reference: rectangle outline + circle, drawn (never resized)."""
    ref = np.full((size, size, 3), _BG, dtype=np.uint8)
    m = max(2, size // 8)
    t = max(2, size // 20)
    cv2.rectangle(ref, (m, m), (size - m, size - m), color, t)
    cv2.circle(ref, (size // 2, size // 2), size // 4, color, t)
    return ref


def _make_outline_frame(
    frame_w: int, frame_h: int, x: int, y: int, size: int, color=(0, 220, 220)
) -> np.ndarray:
    frame = np.full((frame_h, frame_w, 3), _BG, dtype=np.uint8)
    frame[y : y + size, x : x + size] = _make_outline(size, color=color)
    return frame


def _full_shape(frame_edges, prepared, threshold, overlap, window):
    """Previous full-map Shape matcher used as an exact reference."""
    candidates = []
    best_peak = -1.0
    fh, fw = frame_edges.shape[:2]
    for entry in prepared:
        tw, th = entry["w"], entry["h"]
        if th > fh or tw > fw:
            continue
        result = cv2.matchTemplate(frame_edges, entry["edges"], cv2.TM_CCOEFF_NORMED)
        if not np.all(np.isfinite(result)):
            result = np.where(np.isfinite(result), result, -1.0)
        np.clip(result, -1.0, 1.0, out=result)
        if window is not None:
            result = screenspace_primitives._mask_corr_outside_window(
                result, tw, th, window
            )
            if result is None:
                continue
        if result.size:
            best_peak = max(best_peak, float(result.max()))
        locs = np.where(result >= threshold)
        ys, xs = locs[0], locs[1]
        scores = result[locs]
        if scores.size > 5000:
            top_idx = np.argpartition(scores, -5000)[-5000:]
            ys, xs, scores = ys[top_idx], xs[top_idx], scores[top_idx]
        candidates.extend(
            {
                "x": int(x),
                "y": int(y),
                "w": tw,
                "h": th,
                "score": float(score),
                "scale": entry["scale"],
                "scale_y": entry["scale_y"],
            }
            for y, x, score in zip(ys, xs, scores, strict=True)
        )
    if len(candidates) > 5000:
        candidates.sort(key=lambda candidate: -candidate["score"])
        del candidates[5000:]
    return screenspace.nms_boxes_iou(candidates, overlap), best_peak


class TestPrepareShapeReference:
    def test_ladder_is_geometric(self):
        prepared = screenspace._prepare_shape_reference(
            _make_outline(60), scale_min=0.5, scale_max=2.0, scale_steps=7
        )
        scales = [e["scale"] for e in prepared]
        assert scales[0] == 0.5
        assert scales[-1] == 2.0
        ratios = [scales[i + 1] / scales[i] for i in range(len(scales) - 1)]
        assert max(ratios) - min(ratios) < 0.01

    def test_flat_reference_degenerate(self):
        flat = np.full((40, 40, 3), 128, dtype=np.uint8)
        assert screenspace._prepare_shape_reference(flat) == []

    def test_near_edgeless_reference_degenerate(self):
        # A single 2px dot yields fewer raw Canny pixels than the floor.
        ref = np.full((40, 40, 3), _BG, dtype=np.uint8)
        ref[20:22, 20:22] = 255
        assert screenspace._prepare_shape_reference(ref) == []

    def test_empty_mask_degenerate(self):
        ref = _make_outline(40)
        mask = np.zeros((40, 40), dtype=np.uint8)
        assert screenspace._prepare_shape_reference(ref, mask) == []

    def test_unlinked_ladder_is_cross_product(self):
        prepared = screenspace._prepare_shape_reference(
            _make_outline(60),
            scale_min=0.8,
            scale_max=1.2,
            scale_steps=3,
            scale_y_min=0.9,
            scale_y_max=1.1,
            scale_y_steps=2,
        )
        assert len(prepared) == 6
        pairs = {(e["scale"], e["scale_y"]) for e in prepared}
        assert (0.8, 0.9) in pairs and (1.2, 1.1) in pairs
        assert all(e["scale_y"] in (0.9, 1.1) for e in prepared)

    def test_linked_ladder_keeps_uniform_pairs(self):
        prepared = screenspace._prepare_shape_reference(_make_outline(60))
        assert all(e["scale"] == e["scale_y"] for e in prepared)

    def test_mask_drops_alpha_ring_edges(self):
        # Mask out the rectangle outline; only the circle's edges remain.
        ref = _make_outline(60)
        mask = np.zeros((60, 60), dtype=np.uint8)
        cv2.circle(mask, (30, 30), 22, 255, -1)
        prepared = screenspace._prepare_shape_reference(
            ref, mask, scale_min=1.0, scale_max=1.0, scale_steps=1
        )
        unmasked = screenspace._prepare_shape_reference(
            ref, None, scale_min=1.0, scale_max=1.0, scale_steps=1
        )
        assert prepared and unmasked
        masked_px = np.count_nonzero(prepared[0]["edges"])
        assert masked_px < np.count_nonzero(unmasked[0]["edges"])


class TestMatchShape:
    def test_recolored_rescaled_outline_found(self):
        # Reference is white at 60px; the frame holds a recolored 1.5x copy.
        ref = _make_outline(60, color=(255, 255, 255))
        big = cv2.resize(
            _make_outline(60, color=(0, 220, 220)),
            (90, 90),
            interpolation=cv2.INTER_CUBIC,
        )
        frame = np.full((240, 320, 3), _BG, dtype=np.uint8)
        frame[90:180, 110:200] = big
        prepared = screenspace._prepare_shape_reference(ref)
        matches, peak = screenspace.match_shape(
            screenspace._frame_edge_map(frame), prepared, 0.5
        )
        assert peak > 0.5
        # Some surviving box must cover the shape center at a ~1.5x scale
        # (smaller rungs may also fire on sub-structure; that is acceptable).
        cx, cy = 155, 135
        hits = [
            m
            for m in matches
            if m["x"] <= cx <= m["x"] + m["w"] and m["y"] <= cy <= m["y"] + m["h"]
        ]
        assert hits
        assert any(1.2 <= m["scale"] <= 2.0 for m in hits)

    def test_template_misses_where_shape_hits(self):
        # Intensity inversion flips grayscale correlation negative (template
        # misses) but leaves the edge geometry intact (shape hits).
        ref = _make_outline(60, color=(255, 255, 255))
        inverted = 255 - ref
        frame = np.full((240, 320, 3), 255 - _BG, dtype=np.uint8)
        frame[90:150, 110:170] = inverted
        tmpl_matches = screenspace.match_template(frame, ref, threshold=0.7)
        prepared = screenspace._prepare_shape_reference(
            ref, scale_min=1.0, scale_max=1.0, scale_steps=1
        )
        shape_matches, _peak = screenspace.match_shape(
            screenspace._frame_edge_map(frame), prepared, 0.5
        )
        assert tmpl_matches == []
        assert len(shape_matches) >= 1
        m = shape_matches[0]
        assert abs(m["x"] - 110) <= 6 and abs(m["y"] - 90) <= 6

    def test_blank_frame_no_matches(self):
        prepared = screenspace._prepare_shape_reference(_make_outline(60))
        blank = np.full((240, 320, 3), _BG, dtype=np.uint8)
        matches, peak = screenspace.match_shape(
            screenspace._frame_edge_map(blank), prepared, 0.5
        )
        assert matches == []
        assert peak <= 1.0

    def test_peak_clipped_to_one(self):
        ref = _make_outline(60)
        frame = _make_outline_frame(320, 240, 100, 80, 60, color=(255, 255, 255))
        prepared = screenspace._prepare_shape_reference(ref)
        _matches, peak = screenspace.match_shape(
            screenspace._frame_edge_map(frame), prepared, 0.3
        )
        assert peak <= 1.0

    def test_stretched_copy_needs_unlinked_axes(self):
        # A 2x-wide, same-height copy (a content-stretched button): the
        # uniform ladder scores poorly; unlinking the axes finds it at
        # scale 2 / scale_y 1.
        ref = _make_outline(60, color=(255, 255, 255))
        stretched = cv2.resize(
            _make_outline(60, color=(0, 220, 220)),
            (120, 60),
            interpolation=cv2.INTER_CUBIC,
        )
        frame = np.full((240, 320, 3), _BG, dtype=np.uint8)
        frame[90:150, 100:220] = stretched
        fe = screenspace._frame_edge_map(frame)
        uniform = screenspace._prepare_shape_reference(ref)
        aniso = screenspace._prepare_shape_reference(
            ref,
            scale_min=1.0,
            scale_max=2.0,
            scale_steps=3,
            scale_y_min=1.0,
            scale_y_max=1.0,
            scale_y_steps=1,
        )
        _um, uniform_peak = screenspace.match_shape(fe, uniform, 0.99)
        matches, aniso_peak = screenspace.match_shape(fe, aniso, 0.5)
        assert aniso_peak > uniform_peak
        hits = [
            m for m in matches if abs(m["scale"] - 2.0) < 0.01 and m["scale_y"] == 1.0
        ]
        assert hits
        assert abs(hits[0]["x"] - 100) <= 6 and abs(hits[0]["y"] - 90) <= 6

    def test_one_instance_one_box(self):
        # Adjacent ladder scales fire on the same instance; NMS keeps one box.
        ref = _make_outline(60)
        frame = _make_outline_frame(320, 240, 100, 80, 60, color=(0, 220, 220))
        prepared = screenspace._prepare_shape_reference(ref)
        matches, _peak = screenspace.match_shape(
            screenspace._frame_edge_map(frame), prepared, 0.45
        )
        assert len(matches) == 1


class TestSearchWindow:
    def test_roi_matches_full_map(self):
        rng = np.random.RandomState(84)
        frame = rng.randint(0, 255, (96, 144, 3), dtype=np.uint8)
        reference = frame[20:52, 37:81].copy()
        prepared = screenspace._prepare_shape_reference(
            reference, scale_min=0.75, scale_max=1.4, scale_steps=4
        )
        edges = screenspace._frame_edge_map(frame)
        for window in ((0.0, 0.0, 33.0, 26.0), (31.0, 17.0, 93.0, 72.0)):
            want = _full_shape(edges, prepared, 0.05, 0.3, window)
            got = screenspace.match_shape(edges, prepared, 0.05, 0.3, window)
            got_matches, got_peak = got
            want_matches, want_peak = want
            assert [
                {key: value for key, value in row.items() if key != "score"}
                for row in got_matches
            ] == [
                {key: value for key, value in row.items() if key != "score"}
                for row in want_matches
            ]
            assert all(
                abs(got_row["score"] - want_row["score"]) < 5e-5
                for got_row, want_row in zip(got_matches, want_matches, strict=True)
            )
            assert abs(got_peak - want_peak) < 5e-5

    def test_unmatchable_scale_is_identical(self):
        edges = screenspace._frame_edge_map(_make_outline_frame(80, 60, 10, 8, 30))
        prepared = screenspace._prepare_shape_reference(
            _make_outline(120), scale_min=1.0, scale_max=1.0, scale_steps=1
        )
        window = (0.0, 0.0, 20.0, 20.0)
        assert screenspace.match_shape(edges, prepared, 0.1, 0.3, window) == (
            [],
            -1.0,
        )

    def test_executor_is_exact_and_ordered(self):
        frame = _make_outline_frame(180, 120, 72, 36, 42)
        edges = screenspace._frame_edge_map(frame)
        prepared = screenspace._prepare_shape_reference(
            _make_outline(42), scale_min=0.7, scale_max=1.4, scale_steps=5
        )
        sequential = screenspace_primitives._match_shape_scales(
            edges, prepared, 0.1, 0.3
        )
        with ThreadPoolExecutor(max_workers=4) as executor:
            parallel = screenspace_primitives._match_shape_scales(
                edges, prepared, 0.1, 0.3, executor=executor
            )
        assert parallel == sequential

        pool = Mock()
        seen = []

        def ordered(fn, entries):
            entries = list(entries)
            seen.extend(entry["scale"] for entry in entries)
            return map(fn, entries)

        pool.map.side_effect = ordered
        got = screenspace_primitives._match_shape_scales(
            edges, prepared, 0.1, 0.3, executor=pool
        )
        assert got == sequential
        assert seen == [entry["scale"] for entry in prepared]

    def test_window_scopes_matches_and_peak(self):
        # Same frame, two run regions: over the shape → hit; elsewhere → the
        # peak drops too (region-local calibration honesty).
        ref = _make_outline(60)
        frame = _make_outline_frame(320, 240, 110, 90, 60, color=(0, 220, 220))
        fe = screenspace._frame_edge_map(frame)
        prepared = screenspace._prepare_shape_reference(ref)
        over = screenspace.region_search_window({"x": 100, "y": 80, "w": 80, "h": 80})
        away = screenspace.region_search_window({"x": 0, "y": 0, "w": 60, "h": 60})
        m_over, p_over = screenspace.match_shape(fe, prepared, 0.5, window=over)
        m_away, p_away = screenspace.match_shape(fe, prepared, 0.5, window=away)
        assert len(m_over) == 1
        assert m_away == []
        assert p_over > 0.5 > p_away

    def test_zero_size_region_means_full_frame(self):
        assert screenspace.region_search_window({"w": 0, "h": 0}) is None
        assert screenspace.region_search_window({}) is None

    def test_scan_region_scopes_rows(self, monkeypatch):
        frame = _make_outline_frame(320, 240, 110, 90, 60)

        def fake_scan(video_path, interval, callback, **kwargs):
            callback(0.0, frame)

        import screenspace_frames
        import screenspace_scans

        monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", fake_scan)
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )
        hit = screenspace.scan_shape(
            "/fake.mp4",
            {"x": 100, "y": 80, "w": 80, "h": 80},
            _make_outline(60),
            threshold=0.5,
        )
        miss = screenspace.scan_shape(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 60},
            _make_outline(60),
            threshold=0.5,
        )
        assert len(hit) == 1
        assert miss == []

    def test_check_frame_region_scopes(self):
        tool = screenspace.TOOLS["shape"]
        frame = _make_outline_frame(320, 240, 110, 90, 60)
        base = {"shape_image": _make_outline(60), "threshold": 0.5}
        passed, _detail = tool.check_frame(
            frame, None, {"x": 100, "y": 80, "w": 80, "h": 80}, dict(base)
        )
        assert passed is True
        away, away_detail = tool.check_frame(
            frame, None, {"x": 0, "y": 0, "w": 60, "h": 60}, dict(base)
        )
        assert away is False
        assert away_detail is not None and away_detail["best_score"] < 0.5


class TestNmsBoxesIou:
    def test_overlapping_suppressed_highest_wins(self):
        boxes = [
            {"x": 10, "y": 10, "w": 40, "h": 40, "score": 0.8},
            {"x": 12, "y": 12, "w": 50, "h": 50, "score": 0.9},
            {"x": 200, "y": 100, "w": 40, "h": 40, "score": 0.7},
        ]
        kept = screenspace.nms_boxes_iou(boxes, 0.5)
        assert len(kept) == 2
        assert kept[0]["score"] == 0.9
        assert kept[1]["x"] == 200

    def test_disjoint_all_kept(self):
        boxes = [
            {"x": 0, "y": 0, "w": 20, "h": 20, "score": 0.6},
            {"x": 100, "y": 0, "w": 30, "h": 30, "score": 0.5},
        ]
        assert len(screenspace.nms_boxes_iou(boxes, 0.5)) == 2


class TestScanShape:
    def _patch_single_frame(self, monkeypatch, frame: np.ndarray) -> None:
        def fake_scan(video_path, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", fake_scan)
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )

    def test_basic_hit_row_shape(self, monkeypatch):
        frame = _make_outline_frame(320, 240, 110, 90, 60)
        self._patch_single_frame(monkeypatch, frame)
        results = screenspace.scan_shape(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 320, "h": 240},
            _make_outline(60),
            threshold=0.5,
        )
        assert len(results) == 1
        row = results[0]
        assert row["timestamp"] == 0.0
        assert row["match_count"] == len(row["matches"]) == 1
        assert 0.5 <= row["best_score"] <= 1.0
        assert "scale" in row["matches"][0]

    def test_degenerate_reference_skips_pipe(self, monkeypatch):
        def boom(*args, **kwargs):
            raise AssertionError("pipe opened for a degenerate reference")

        monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", boom)
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )
        flat = np.full((40, 40, 3), 128, dtype=np.uint8)
        results = screenspace.scan_shape(
            "/fake.mp4", {"x": 0, "y": 0, "w": 320, "h": 240}, flat
        )
        assert results == []

    def test_static_frame_carries_row(self, monkeypatch):
        frame = _make_outline_frame(320, 240, 110, 90, 60)

        def fake_scan(video_path, interval, callback, **kwargs):
            callback(0.0, frame)
            callback(1.0, frame.copy())

        monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", fake_scan)
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 2.0)
        )
        results = screenspace.scan_shape(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 320, "h": 240},
            _make_outline(60),
            threshold=0.5,
        )
        assert [r["timestamp"] for r in results] == [0.0, 1.0]
        assert results[0]["best_score"] == results[1]["best_score"]

    def test_executor_shutdown(self, monkeypatch):
        frame = _make_outline_frame(160, 120, 50, 30, 40)
        pool = Mock()
        created = {}
        pool.map.side_effect = lambda fn, entries: map(fn, entries)

        def make_pool(**kwargs):
            created.update(kwargs)
            return pool

        monkeypatch.setattr(screenspace_scans, "ThreadPoolExecutor", make_pool)
        monkeypatch.setattr(screenspace_scans.os, "cpu_count", lambda: 8)
        monkeypatch.setattr(config, "SCREENSPACE_PARALLEL_WORKERS", 1)
        self._patch_single_frame(monkeypatch, frame)
        screenspace.scan_shape(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 160, "h": 120},
            _make_outline(40),
            threshold=0.5,
        )
        assert created["max_workers"] == 4
        pool.shutdown.assert_called_once_with()

    def test_executor_shutdown_on_cancel(self, monkeypatch):
        pool = Mock()
        monkeypatch.setattr(screenspace_scans, "ThreadPoolExecutor", lambda **_k: pool)
        monkeypatch.setattr(screenspace_scans.os, "cpu_count", lambda: 8)
        monkeypatch.setattr(config, "SCREENSPACE_PARALLEL_WORKERS", 1)
        self._patch_single_frame(monkeypatch, _make_outline_frame(160, 120, 50, 30, 40))
        assert (
            screenspace.scan_shape(
                "/fake.mp4",
                {"x": 0, "y": 0, "w": 160, "h": 120},
                _make_outline(40),
                cancel_flag=lambda: True,
            )
            == []
        )
        pool.shutdown.assert_called_once_with()

    def test_executor_shutdown_on_error(self, monkeypatch):
        pool = Mock()
        monkeypatch.setattr(screenspace_scans, "ThreadPoolExecutor", lambda **_k: pool)
        monkeypatch.setattr(screenspace_scans.os, "cpu_count", lambda: 8)
        monkeypatch.setattr(config, "SCREENSPACE_PARALLEL_WORKERS", 1)
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda _path: (30.0, 1.0)
        )

        def fail(*_args, **_kwargs):
            raise RuntimeError("scan failed")

        monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", fail)
        with pytest.raises(RuntimeError, match="scan failed"):
            screenspace.scan_shape(
                "/fake.mp4",
                {"x": 0, "y": 0, "w": 160, "h": 120},
                _make_outline(40),
            )
        pool.shutdown.assert_called_once_with()


class TestShapeTool:
    def test_check_frame_miss_keeps_best_score(self):
        # The calibration scalar must be populated on both branches.
        tool = screenspace.TOOLS["shape"]
        blank = np.full((240, 320, 3), _BG, dtype=np.uint8)
        params = {"shape_image": _make_outline(60)}
        passed, detail = tool.check_frame(blank, None, {"w": 0, "h": 0}, params)
        assert passed is False
        assert detail is not None and "best_score" in detail
        assert "_prepared_shape" in params

    def test_check_frame_hit(self):
        tool = screenspace.TOOLS["shape"]
        frame = _make_outline_frame(320, 240, 110, 90, 60)
        params = {"shape_image": _make_outline(60), "threshold": 0.5}
        passed, detail = tool.check_frame(frame, None, {"w": 0, "h": 0}, params)
        assert passed is True
        assert detail is not None
        assert detail["match_count"] == 1

    def test_degenerate_reference_not_evaluable(self):
        tool = screenspace.TOOLS["shape"]
        blank = np.full((240, 320, 3), _BG, dtype=np.uint8)
        flat = np.full((40, 40, 3), 128, dtype=np.uint8)
        passed, detail = tool.check_frame(
            blank, None, {"w": 0, "h": 0}, {"shape_image": flat}
        )
        assert passed is False
        assert detail is None

    def test_zero_size_region_dispatchable(self):
        # Uploaded reference with no region: shape shares template's exemption.
        frame = _make_outline_frame(320, 240, 110, 90, 60)
        passed, detail = screenspace.check_frame_for_tool(
            frame,
            None,
            {"w": 0, "h": 0},
            "shape",
            {"shape_image": _make_outline(60), "threshold": 0.5},
        )
        assert passed is True
        assert detail is not None
        assert detail["match_count"] == 1

    def test_threshold_default_from_config(self, monkeypatch):
        captured = {}

        def fake_scan_shape(video_path, region, **kwargs):
            captured.update(kwargs)
            return []

        import screenspace_tools

        monkeypatch.setattr(screenspace_tools, "scan_shape", fake_scan_shape)
        tool = screenspace.TOOLS["shape"]
        tool.scan(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 320, "h": 240},
            {"shape_image": _make_outline(60)},
            task_id="t1",
            scan_mode="full",
            on_progress=lambda _p: None,
            cancel_flag=lambda: False,
            on_result=None,
            fast_opts=None,
        )
        # scan_shape resolves threshold<=0 to the config default itself.
        assert captured["threshold"] == 0
        assert config.SCREENSPACE_SHAPE_MATCH_THRESHOLD > 0

    def test_fast_mode_halves_reference(self, monkeypatch):
        captured = {}

        def fake_scan_shape(video_path, region, **kwargs):
            captured.update(kwargs)
            return []

        import screenspace_tools

        monkeypatch.setattr(screenspace_tools, "scan_shape", fake_scan_shape)
        tool = screenspace.TOOLS["shape"]
        tool.scan(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 320, "h": 240},
            {"shape_image": _make_outline(60)},
            task_id="t1",
            scan_mode="fast",
            on_progress=lambda _p: None,
            cancel_flag=lambda: False,
            on_result=None,
            fast_opts={"template_downscale": True},
        )
        assert captured["shape_image"].shape[:2] == (30, 30)
