"""Tests for per-frame tool scoring, boundaries, and multitool chaining."""

import copy
import math

import numpy as np
import pytest

import config
import screenspace
import screenspace_frames
import screenspace_multitool
import screenspace_ocr
import screenspace_scans
from _ss_helpers import _gray_with_red_patch, _make_icon, _make_icon_frame


@pytest.fixture(autouse=True)
def _reset_ocr_pool(monkeypatch):
    """Pin the OCR reader pool to one slot and clear it between tests."""
    monkeypatch.setattr(config, "SCREENSPACE_OCR_POOL_SIZE", 1)
    screenspace_ocr._ocr_pools.clear()
    yield
    screenspace_ocr._ocr_pools.clear()


class TestCheckFrameForTool:
    def test_color_pass(self):
        # Pure blue in BGR: (255, 0, 0) -> HSV ~(120, 255, 255)
        frame = np.full((100, 100, 3), [255, 0, 0], dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        params = {
            "target_color": {"h": 120, "s": 255, "v": 255},
            "tolerance": {"h": 10, "s": 50, "v": 50},
        }
        passed, result = screenspace.check_frame_for_tool(
            frame, None, region, "color", params
        )
        assert passed is True
        assert result is not None
        assert "_confidence" in result

    def test_color_fail(self):
        frame = np.full((100, 100, 3), [255, 0, 0], dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        # Red target will not match blue frame
        params = {
            "target_color": {"h": 0, "s": 255, "v": 255},
            "tolerance": {"h": 5, "s": 10, "v": 10},
        }
        passed, result = screenspace.check_frame_for_tool(
            frame, None, region, "color", params
        )
        assert passed is False
        # The scalar is now exposed on the fail branch too (calibration reads it).
        assert result is not None
        assert "_confidence" in result

    def test_color_presence_mode(self):
        # Mostly-gray frame with a small dark-red patch: average mode misses it,
        # presence mode detects it.
        frame = _gray_with_red_patch(10)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        target = {"h": 0, "s": 255, "v": 139}
        tol = {"h": 10, "s": 60, "v": 60}
        avg_passed, _ = screenspace.check_frame_for_tool(
            frame, None, region, "color", {"target_color": target, "tolerance": tol}
        )
        assert avg_passed is False
        passed, result = screenspace.check_frame_for_tool(
            frame,
            None,
            region,
            "color",
            {"target_color": target, "tolerance": tol, "color_mode": "presence"},
        )
        assert passed is True
        assert result is not None
        assert "_confidence" in result

    def test_color_presence_min_coverage(self):
        # The same 1% patch fails a 5% min-area gate.
        frame = _gray_with_red_patch(10)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        params = {
            "target_color": {"h": 0, "s": 255, "v": 139},
            "tolerance": {"h": 10, "s": 60, "v": 60},
            "color_mode": "presence",
            "min_coverage": 0.05,
        }
        passed, _ = screenspace.check_frame_for_tool(
            frame, None, region, "color", params
        )
        assert passed is False

    def test_change_needs_prev_frame(self):
        frame = np.zeros((100, 100, 3), dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame, None, region, "change", {"threshold": 0.03}
        )
        assert passed is False
        assert result is None

    def test_change_pass(self):
        frame_a = np.zeros((100, 100, 3), dtype=np.uint8)
        frame_b = np.full((100, 100, 3), 200, dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame_b, frame_a, region, "change", {"threshold": 0.03}
        )
        assert passed is True
        assert result is not None
        assert "magnitude" in result

    def test_similarity_pass(self):
        frame = np.full((100, 100, 3), 128, dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        ref = np.full((100, 100, 3), 128, dtype=np.uint8)
        passed, result = screenspace.check_frame_for_tool(
            frame,
            None,
            region,
            "similarity",
            {"reference_frame": ref, "threshold": 0.5},
        )
        assert passed is True
        assert result is not None
        assert "score" in result

    def test_template_check_frame_applies_scale(self):
        """The multitool per-frame matcher honors template_scale: a 40px
        template misses a 20px in-frame icon at scale 1.0 but hits at 0.5."""
        frame = _make_icon_frame(400, 200, [(100, 50, 20)])
        template = _make_icon(40)
        region = {"x": 0, "y": 0, "w": 400, "h": 200}

        passed_full, _ = screenspace.check_frame_for_tool(
            frame,
            None,
            region,
            "template",
            {"template_image": template.copy(), "threshold": 0.70},
        )
        assert passed_full is False

        passed_scaled, detail = screenspace.check_frame_for_tool(
            frame,
            None,
            region,
            "template",
            {
                "template_image": template.copy(),
                "threshold": 0.70,
                "template_scale": 0.5,
            },
        )
        assert passed_scaled is True
        assert detail is not None
        assert detail["match_count"] >= 1

    def test_template_zero_region_matches_full_frame(self):
        # A multitool uploaded-template step with no region gets zero-size
        # region_coords; check_frame_for_tool must NOT reject it as degenerate
        # (template ignores the region and scans the whole frame). Without the
        # exemption the AND chain could never pass for an upload-only step.
        frame = _make_icon_frame(400, 200, [(100, 50, 40)])
        template = _make_icon(40)
        passed, detail = screenspace.check_frame_for_tool(
            frame,
            None,
            {"x": 0, "y": 0, "w": 0, "h": 0},
            "template",
            {"template_image": template.copy(), "threshold": 0.70},
        )
        assert passed is True
        assert detail is not None
        assert detail["match_count"] >= 1

    def test_numbers_check_frame_honors_zero_ocr_threshold(self, monkeypatch):
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        from types import SimpleNamespace

        # Callable RapidOCR engine stub emitting one "5" reading at conf 0.2.
        monkeypatch.setattr(
            screenspace_ocr,
            "_build_ocr_reader",
            lambda _m: (
                lambda _px: SimpleNamespace(
                    boxes=[[(0, 0), (10, 0), (10, 10), (0, 10)]],
                    txts=("5",),
                    scores=(0.2,),
                )
            ),
        )

        passed, result = screenspace.check_frame_for_tool(
            frame,
            None,
            region,
            "numbers",
            {"operator": "gte", "target_value": 5},
        )
        assert passed is False
        # Best OCR confidence is exposed even when no reading clears the floor.
        assert result is not None
        assert "confidence" in result

        passed, result = screenspace.check_frame_for_tool(
            frame,
            None,
            region,
            "numbers",
            {
                "operator": "gte",
                "target_value": 5,
                "ocr_confidence_threshold": 0.0,
            },
        )
        assert passed is True
        assert result is not None
        assert result["number_found"] == 5.0
        assert result["confidence"] == 0.2

    def test_flow_needs_prev_frame(self):
        frame = np.zeros((100, 100, 3), dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame, None, region, "flow", {"magnitude_threshold": 2.0}
        )
        assert passed is False
        assert result is None

    def test_inactivity_needs_prev_frame(self):
        frame = np.zeros((100, 100, 3), dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame, None, region, "inactivity", {"threshold": 10}
        )
        assert passed is False
        assert result is None

    def test_inactivity_identical_frames(self):
        frame = np.full((100, 100, 3), 128, dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame, frame.copy(), region, "inactivity", {"threshold": 10}
        )
        assert passed is True
        assert result is not None
        assert "distance" in result
        assert result["distance"] == 0
        assert result["_confidence"] == 1.0
        assert screenspace._extract_confidence("inactivity", result) == 1.0

    def test_inactivity_different_frames(self):
        frame_a = np.zeros((100, 100, 3), dtype=np.uint8)
        # Random noise frame produces a very different perceptual hash
        rng = np.random.RandomState(42)
        frame_b = rng.randint(0, 256, (100, 100, 3), dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame_b, frame_a, region, "inactivity", {"threshold": 2}
        )
        assert passed is False
        # phash distance is exposed on the fail branch (the calibration scalar).
        assert result is not None
        assert "distance" in result

    def test_boundary_not_a_multitool_step(self):
        # Boundary is scan-only: it defines no check_frame, so the generic
        # dispatch falls back to the base "not a multitool step" result.
        frame_a = np.zeros((100, 100, 3), dtype=np.uint8)
        rng = np.random.RandomState(42)
        frame_b = rng.randint(0, 256, (100, 100, 3), dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame_b, frame_a, region, "boundary", {"threshold": 2}
        )
        assert passed is False
        assert result is None

    def test_unknown_type(self):
        frame = np.zeros((100, 100, 3), dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 100, "h": 100}
        passed, result = screenspace.check_frame_for_tool(
            frame, None, region, "bogus", {}
        )
        assert passed is False
        assert result is None


class _FakeHash:
    """Stand-in for imagehash.ImageHash whose subtraction is a plain |Δ|.

    Lets the boundary scan tests drive exact phash distances by encoding the
    "scene id" in a frame's pixel value, isolating the threshold + debounce
    logic from real perceptual-hash behaviour.
    """

    def __init__(self, val: int):
        self.val = val

    def __sub__(self, other: "_FakeHash") -> int:
        return abs(self.val - other.val)


class TestScanBoundaries:
    @staticmethod
    def _setup(monkeypatch, frames):
        # frames: list of (timestamp, tag). Same tag → distance 0; different
        # tags → distance == |Δtag|. Each frame is filled with its tag value.
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda _p: (30.0, 100.0)
        )
        monkeypatch.setattr(
            screenspace_scans,
            "compute_phash",
            lambda f, gray=None: _FakeHash(int(f.reshape(-1)[0])),
        )

        def fake_scan(video_path, interval_seconds, callback, **kwargs):
            for ts, tag in frames:
                frame = np.full((8, 8, 3), tag, dtype=np.uint8)
                if callback(ts, frame) is False:
                    break

        monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", fake_scan)

    def test_hard_cuts_fire_at_cut_frames_only(self, monkeypatch):
        frames = [
            (0.0, 0),
            (1.0, 0),
            (2.0, 0),
            (3.0, 40),
            (4.0, 40),
            (5.0, 40),
            (6.0, 0),
            (7.0, 0),
        ]
        self._setup(monkeypatch, frames)
        results = screenspace.scan_boundaries(
            "/fake.mp4", threshold=14, min_gap=3.0, interval_seconds=1.0
        )
        assert [r["timestamp"] for r in results] == [3.0, 6.0]
        assert all(r["distance"] == 40 for r in results)
        assert all(0.0 < r["_confidence"] <= 1.0 for r in results)

    def test_min_gap_suppresses_storms(self, monkeypatch):
        # Every other frame is a big jump; only the first spike per min_gap
        # window survives (the run's first frame is the boundary).
        frames = [(float(i), 0 if i % 2 == 0 else 50) for i in range(8)]
        self._setup(monkeypatch, frames)
        results = screenspace.scan_boundaries(
            "/fake.mp4", threshold=14, min_gap=3.0, interval_seconds=1.0
        )
        assert [r["timestamp"] for r in results] == [1.0, 4.0, 7.0]

    def test_gradual_drift_below_threshold_never_fires(self, monkeypatch):
        # Documented behaviour: scan_boundaries compares CONSECUTIVE samples, not
        # cumulative drift. A slow fade where each step stays under threshold
        # produces no boundary at all, even though the first and last frames
        # differ wildly (0 → 45). A boundary needs a per-sample jump.
        frames = [(float(i), i * 5) for i in range(10)]  # +5 per step, threshold 14
        self._setup(monkeypatch, frames)
        results = screenspace.scan_boundaries(
            "/fake.mp4", threshold=14, min_gap=3.0, interval_seconds=1.0
        )
        assert results == []

    def test_cancel_stops_scan_early(self, monkeypatch):
        # 20 alternating frames would otherwise yield many boundaries; a cancel
        # flag that trips after a couple frames must short-circuit the sweep.
        frames = [(float(i), 0 if i % 2 == 0 else 50) for i in range(20)]
        self._setup(monkeypatch, frames)
        calls = {"n": 0}

        def cancel():
            calls["n"] += 1
            return calls["n"] > 2

        results = screenspace.scan_boundaries(
            "/fake.mp4",
            threshold=14,
            min_gap=0.5,
            interval_seconds=1.0,
            cancel_flag=cancel,
        )
        assert len(results) <= 1

    def test_scene_labels_are_sequential_for_phash(self, monkeypatch):
        # phash has no fingerprints, so labels are sequential: B, C, ...
        frames = [(0.0, 0), (1.0, 0), (2.0, 40), (3.0, 40), (4.0, 0)]
        self._setup(monkeypatch, frames)
        results = screenspace.scan_boundaries(
            "/fake.mp4", threshold=14, min_gap=1.0, interval_seconds=1.0
        )
        assert [r["scene_label"] for r in results] == ["Scene B", "Scene C"]


def _tag_frame(tag: int, phash: int = 0) -> np.ndarray:
    """8×8 frame encoding a scene tag at [0,0,0] and a phash value at [0,1,0].

    Lets the scene/hybrid boundary tests drive the content fingerprint and the
    phash spike independently, deterministically, without real CV.
    """
    f = np.zeros((8, 8, 3), dtype=np.uint8)
    f[0, 0, 0] = tag
    f[0, 1, 0] = phash
    return f


class TestConsolidateBoundaryPeriods:
    """The post-run sanity pass over the scene/hybrid period list."""

    @staticmethod
    def _patch(monkeypatch):
        # Identical tags = same scene (distance 0); different tags fully differ
        # (distance 1). Isolates the merge/prune logic from real fingerprints.
        monkeypatch.setattr(
            screenspace_scans,
            "compute_scene_fingerprint",
            lambda f: {"tag": int(f[0, 0, 0])},
        )
        monkeypatch.setattr(
            screenspace_scans,
            "compare_scene_fingerprints",
            lambda a, b: 1.0 if a["tag"] == b["tag"] else 0.0,
        )

    @staticmethod
    def _period(start_ts, tag, entry_dist):
        return {
            "start_ts": float(start_ts),
            "pixels": _tag_frame(tag),
            "entry_dist": entry_dist,
            "entry_conf": 0.9,
        }

    def _consolidate(
        self,
        periods,
        *,
        merge_threshold=0.15,
        short_period_seconds=3.0,
        relative_prune_enabled=False,
        relative_prune_factor=0.5,
        type_threshold=0.35,
    ):
        return screenspace_scans._consolidate_boundary_periods(
            periods,
            end_seconds=1000.0,
            merge_threshold=merge_threshold,
            short_period_seconds=short_period_seconds,
            relative_prune_enabled=relative_prune_enabled,
            relative_prune_factor=relative_prune_factor,
            type_threshold=type_threshold,
        )

    def test_adjacent_merge_drops_spurious_boundary(self, monkeypatch):
        self._patch(monkeypatch)
        # Two periods, same scene → the boundary between them is spurious.
        periods = [self._period(0, 1, 0), self._period(10, 1, 80)]
        assert self._consolidate(periods) == []

    def test_real_change_kept_with_span(self, monkeypatch):
        self._patch(monkeypatch)
        periods = [self._period(0, 1, 0), self._period(10, 2, 80)]
        results = self._consolidate(periods)
        assert [r["timestamp"] for r in results] == [10.0]
        assert results[0]["period_start"] == 10.0
        assert results[0]["period_end"] == 1000.0  # last period runs to end

    def test_roundtrip_transient_dissolved(self, monkeypatch):
        self._patch(monkeypatch)
        # Short period (tag 2) bracketed by the same scene (tag 1) → dissolve both.
        periods = [
            self._period(0, 1, 0),
            self._period(10, 2, 80),
            self._period(11, 1, 80),
        ]
        assert self._consolidate(periods) == []

    def test_distinct_short_period_kept(self, monkeypatch):
        self._patch(monkeypatch)
        # Short period (tag 3) differs from both neighbors → a real loading screen.
        periods = [
            self._period(0, 1, 0),
            self._period(10, 3, 80),
            self._period(11, 2, 80),
        ]
        results = self._consolidate(periods)
        assert [r["timestamp"] for r in results] == [10.0, 11.0]
        # period_end chains to the next boundary, then to end_seconds.
        assert results[0]["period_end"] == 11.0
        assert results[1]["period_end"] == 1000.0

    def test_relative_prune_drops_weak_boundary(self, monkeypatch):
        self._patch(monkeypatch)
        # Four distinct, well-spaced scenes; one boundary far below the median.
        periods = [
            self._period(0, 0, 0),
            self._period(10, 1, 50),
            self._period(20, 2, 48),
            self._period(30, 3, 52),
            self._period(40, 4, 5),  # weak: median≈49, cutoff 0.5×49=24.5
        ]
        kept = self._consolidate(periods, relative_prune_enabled=True)
        assert [r["timestamp"] for r in kept] == [10.0, 20.0, 30.0]
        # With the prune off, the weak boundary survives.
        all_b = self._consolidate(periods, relative_prune_enabled=False)
        assert [r["timestamp"] for r in all_b] == [10.0, 20.0, 30.0, 40.0]

    def test_relative_prune_guard_skips_when_few_boundaries(self, monkeypatch):
        self._patch(monkeypatch)
        # Only three boundaries (< min count) → prune is skipped even when enabled.
        periods = [
            self._period(0, 0, 0),
            self._period(10, 1, 50),
            self._period(20, 2, 48),
            self._period(30, 3, 5),  # weak, but guard protects it
        ]
        kept = self._consolidate(periods, relative_prune_enabled=True)
        assert [r["timestamp"] for r in kept] == [10.0, 20.0, 30.0]

    def test_recurrence_labels_reuse_for_same_scene(self, monkeypatch):
        self._patch(monkeypatch)
        # Scenes 1 → 2 → back to 1 → 3: the revisit to scene 1 reuses its label.
        # With the binary metric every distinct scene is its own type, so labels
        # are <type letter>1 and the revisit reuses "Scene A1".
        periods = [
            self._period(0, 1, 0),
            self._period(10, 2, 80),
            self._period(20, 1, 80),
            self._period(30, 3, 80),
        ]
        results = self._consolidate(periods)
        assert [r["scene_label"] for r in results] == [
            "Scene B1",
            "Scene A1",
            "Scene C1",
        ]

    def test_similar_scenes_share_type_letter(self, monkeypatch):
        # Graded metric: distance grows with |Δtag|*0.2. Tags 1 & 2 are similar
        # (distinct scenes, same type → A1/A2); tag 5 is a different type (B1);
        # revisiting tag 1 reuses A1.
        monkeypatch.setattr(
            screenspace_scans,
            "compute_scene_fingerprint",
            lambda f: {"tag": int(f[0, 0, 0])},
        )
        monkeypatch.setattr(
            screenspace_scans,
            "compare_scene_fingerprints",
            lambda a, b: max(0.0, 1.0 - min(1.0, abs(a["tag"] - b["tag"]) * 0.2)),
        )
        periods = [
            self._period(0, 1, 0),
            self._period(10, 2, 80),
            self._period(20, 5, 80),
            self._period(30, 1, 80),
        ]
        results = self._consolidate(periods)
        assert [r["scene_label"] for r in results] == [
            "Scene A2",
            "Scene B1",
            "Scene A1",
        ]

    def test_prune_then_remerge_collapses_duplicate_boundary(self, monkeypatch):
        self._patch(monkeypatch)
        # A B A B C with a weak revisit to A. Pruning the weak A@20 leaves two
        # adjacent B periods (same scene); the post-prune merge must collapse
        # them into a single boundary rather than emitting a duplicate.
        periods = [
            self._period(0, 1, 0),
            self._period(10, 2, 50),
            self._period(20, 1, 5),  # weak revisit → pruned
            self._period(30, 2, 50),
            self._period(40, 3, 50),
        ]
        results = self._consolidate(periods, relative_prune_enabled=True)
        assert [r["timestamp"] for r in results] == [10.0, 40.0]
        assert [r["scene_label"] for r in results] == ["Scene B1", "Scene C1"]


class TestScanBoundariesSceneHybrid:
    """The scene/hybrid period-reference detector (metric != phash)."""

    @staticmethod
    def _setup(monkeypatch, frames):
        # frames: list of (ts, scene_tag, phash_val)
        monkeypatch.setattr(
            screenspace_frames,
            "_probe_video_meta",
            lambda _p: (30.0, float(len(frames) + 1)),
        )
        monkeypatch.setattr(
            screenspace_scans,
            "compute_scene_fingerprint",
            lambda f: {"tag": int(f[0, 0, 0])},
        )
        monkeypatch.setattr(
            screenspace_scans,
            "compare_scene_fingerprints",
            lambda a, b: 1.0 if a["tag"] == b["tag"] else 0.0,
        )
        monkeypatch.setattr(
            screenspace_scans,
            "compute_phash",
            lambda f, gray=None: _FakeHash(int(f[0, 1, 0])),
        )

        def fake_scan(video_path, interval_seconds, callback, **kwargs):
            for ts, tag, ph in frames:
                if callback(ts, _tag_frame(tag, ph)) is False:
                    break

        monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", fake_scan)

    def test_motion_within_period_no_boundary(self, monkeypatch):
        # Same scene throughout; phash jumps around (motion) — scene metric ignores it.
        frames = [(float(i), 5, (i * 70) % 256) for i in range(8)]
        self._setup(monkeypatch, frames)
        results = screenspace.scan_boundaries(
            "/fake.mp4", metric="scene", interval_seconds=1.0
        )
        assert results == []

    def test_scene_change_fires_once_at_transition_start(self, monkeypatch):
        # tag 1 → tag 2 sustained: one boundary at the first tag-2 frame.
        frames = [
            (0.0, 1, 0),
            (1.0, 1, 0),
            (2.0, 1, 0),
            (3.0, 2, 0),
            (4.0, 2, 0),
            (5.0, 2, 0),
        ]
        self._setup(monkeypatch, frames)
        results = screenspace.scan_boundaries(
            "/fake.mp4", metric="scene", interval_seconds=1.0
        )
        assert [r["timestamp"] for r in results] == [3.0]
        assert results[0]["period_start"] == 3.0

    def test_confirm_window_suppresses_single_sample_blip(self, monkeypatch):
        # One-sample blip (tag 9) then back to tag 1 → never confirmed.
        frames = [(0.0, 1, 0), (1.0, 1, 0), (2.0, 9, 0), (3.0, 1, 0), (4.0, 1, 0)]
        self._setup(monkeypatch, frames)
        results = screenspace.scan_boundaries(
            "/fake.mp4", metric="scene", interval_seconds=1.0
        )
        assert results == []

    def test_hybrid_requires_phash_corroboration(self, monkeypatch):
        # Sustained scene shift but NO phash spike → hybrid does not fire.
        frames = [(0.0, 1, 10), (1.0, 1, 10), (2.0, 2, 10), (3.0, 2, 10), (4.0, 2, 10)]
        self._setup(monkeypatch, frames)
        results = screenspace.scan_boundaries(
            "/fake.mp4", metric="hybrid", threshold=14, interval_seconds=1.0
        )
        assert results == []

    def test_hybrid_fires_when_phash_spike_coincides(self, monkeypatch):
        # Sustained scene shift WITH a phash spike at the transition → fires once.
        frames = [
            (0.0, 1, 10),
            (1.0, 1, 10),
            (2.0, 2, 200),
            (3.0, 2, 200),
            (4.0, 2, 200),
        ]
        self._setup(monkeypatch, frames)
        results = screenspace.scan_boundaries(
            "/fake.mp4", metric="hybrid", threshold=14, interval_seconds=1.0
        )
        assert [r["timestamp"] for r in results] == [2.0]

    def test_min_gap_does_not_stick_period_reference(self, monkeypatch):
        # Regression: a transition within min_gap of the previous boundary has its
        # boundary suppressed, but the period reference must still advance — else
        # ref_fp sticks on the old scene and every later transition is lost.
        # Boundary at t2 (1→2); 2→3 at t4 is within min_gap (suppressed); the
        # genuine 3→4 at t8 (beyond the gap) must still be detected.
        frames = [
            (0.0, 1, 0),
            (1.0, 1, 0),
            (2.0, 2, 0),
            (3.0, 2, 0),
            (4.0, 3, 0),
            (5.0, 3, 0),
            (6.0, 3, 0),
            (7.0, 3, 0),
            (8.0, 4, 0),
            (9.0, 4, 0),
        ]
        self._setup(monkeypatch, frames)
        results = screenspace.scan_boundaries(
            "/fake.mp4", metric="scene", min_gap=3.0, interval_seconds=1.0
        )
        assert [r["timestamp"] for r in results] == [2.0, 8.0]


class TestScanMultitool:
    def test_requires_min_2_steps(self):
        with pytest.raises(ValueError, match="at least 2"):
            screenspace.scan_multitool(
                "/fake/video.mp4",
                {"x": 0, "y": 0, "w": 100, "h": 100},
                steps=[{"type": "color"}],
            )

    @staticmethod
    def _setup_stubs(monkeypatch, check_fn):
        monkeypatch.setattr(
            screenspace_multitool, "_probe_video_meta", lambda _p: (30.0, 10.0)
        )

        def fake_scan(
            video_path,
            interval_seconds,
            callback,
            *,
            start_seconds=0.0,
            end_seconds=None,
            fps=0.0,
            duration=0.0,
            fast_opts=None,
            profile_kind="",
        ):
            frame = np.zeros((10, 10, 3), dtype=np.uint8)
            callback(1.0, frame)

        monkeypatch.setattr(screenspace_multitool, "scan_video_full_frames", fake_scan)
        monkeypatch.setattr(screenspace_multitool, "check_frame_for_tool", check_fn)

    def test_not_operator_rejects_when_negated_match(self, monkeypatch):
        def check(frame, prev, region, ttype, step, cache=None, prev_cache=None):
            return True, {"_confidence": 0.9}

        self._setup_stubs(monkeypatch, check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[{"type": "color"}, {"type": "change", "logic": "NOT"}],
        )
        assert results == []

    def test_not_operator_passes_when_negated_misses(self, monkeypatch):
        def check(frame, prev, region, ttype, step, cache=None, prev_cache=None):
            if ttype == "color":
                return True, {"_confidence": 0.8}
            return False, None

        self._setup_stubs(monkeypatch, check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[{"type": "color"}, {"type": "change", "logic": "NOT"}],
        )
        assert len(results) == 1
        assert results[0]["steps"][1] == {"negated": True, "type": "change"}
        assert results[0]["min_confidence"] == 0.8

    def test_and_default_when_logic_missing(self, monkeypatch):
        def check(frame, prev, region, ttype, step, cache=None, prev_cache=None):
            if ttype == "color":
                return True, {"_confidence": 0.7}
            return True, {"magnitude": 0.5}

        self._setup_stubs(monkeypatch, check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[{"type": "color"}, {"type": "change"}],
        )
        assert len(results) == 1
        assert results[0]["min_confidence"] == 0.5

    def test_inactivity_step_uses_per_frame_confidence(self, monkeypatch):
        def check(frame, prev, region, ttype, step, cache=None, prev_cache=None):
            if ttype == "inactivity":
                return True, {"distance": 0, "_confidence": 0.75}
            return True, {"_confidence": 0.9}

        self._setup_stubs(monkeypatch, check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[{"type": "color"}, {"type": "inactivity"}],
        )
        assert len(results) == 1
        assert results[0]["min_confidence"] == 0.75

    def test_failing_and_step_excludes_frame_with_nonnull_detail(self, monkeypatch):
        # Regression: after Phase 2, check_frame returns a detail dict on the
        # fail branch too. The AND chain must still exclude the frame — it
        # short-circuits on `not passed`, never on `rd is None`.
        def check(frame, prev, region, ttype, step, cache=None, prev_cache=None):
            if ttype == "color":
                return True, {"_confidence": 0.9}
            return False, {"magnitude": 0.01}  # miss, but detail is non-None now

        self._setup_stubs(monkeypatch, check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[{"type": "color"}, {"type": "change"}],
        )
        assert results == []

    # ---- Offset (two-phase) path -----------------------------------------

    @staticmethod
    def _setup_multiframe_stubs(monkeypatch, frames_ts, check_fn):
        """Stub the decode to emit one frame per ts in *frames_ts*.

        ``check_fn(ts, ttype, step)`` returns ``(passed, detail)`` and may vary
        its answer by the current timestamp — the stubbed ``check_frame_for_tool``
        injects the live ts that ``scan_video_full_frames`` is replaying.
        """
        monkeypatch.setattr(
            screenspace_multitool,
            "_probe_video_meta",
            lambda _p: (30.0, max(frames_ts) + 1.0),
        )
        state = {"ts": 0.0}

        def fake_scan(
            video_path,
            interval_seconds,
            callback,
            *,
            start_seconds=0.0,
            end_seconds=None,
            fps=0.0,
            duration=0.0,
            fast_opts=None,
            profile_kind="",
        ):
            frame = np.zeros((10, 10, 3), dtype=np.uint8)
            for ts in frames_ts:
                state["ts"] = ts
                if callback(ts, frame) is False:
                    break

        def check(frame, prev, region, ttype, step, cache=None, prev_cache=None):
            return check_fn(state["ts"], ttype, step)

        monkeypatch.setattr(screenspace_multitool, "scan_video_full_frames", fake_scan)
        monkeypatch.setattr(screenspace_multitool, "check_frame_for_tool", check)

    def test_offset_inverted_window_warns(self, monkeypatch):
        # min > max is an empty window: warn, but don't raise.
        from unittest.mock import Mock

        warn = Mock()
        monkeypatch.setattr(screenspace_multitool.utils, "warning_print", warn)

        def check(ts, ttype, step):
            return True, {"_confidence": 0.9}

        self._setup_multiframe_stubs(monkeypatch, [0, 1, 2], check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color"},
                {"type": "change", "offset": {"min": 5, "max": 1}},
            ],
        )
        assert results == []
        warn.assert_called_once()

    def test_offset_out_of_range_window_warns(self, monkeypatch):
        # Window starts beyond the whole scan span → can never match: warn.
        from unittest.mock import Mock

        warn = Mock()
        monkeypatch.setattr(screenspace_multitool.utils, "warning_print", warn)

        def check(ts, ttype, step):
            return True, {"_confidence": 0.9}

        # _probe_video_meta -> duration = max(ts)+1 = 4, so total_range ~= 4.
        self._setup_multiframe_stubs(monkeypatch, [0, 1, 2, 3], check)
        screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color"},
                {"type": "change", "offset": {"min": 100, "max": 200}},
            ],
        )
        warn.assert_called_once()

    def test_offset_and_hit(self, monkeypatch):
        # color @2; change @4 within the [+0,+3] window from the anchor.
        def check(ts, ttype, step):
            if ttype == "color":
                return (ts == 2.0), {"_confidence": 0.9}
            return (ts == 4.0), {"magnitude": 0.5}

        self._setup_multiframe_stubs(monkeypatch, [0, 1, 2, 3, 4, 5], check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color"},
                {"type": "change", "offset": {"min": 0, "max": 3}},
            ],
        )
        assert len(results) == 1
        assert results[0]["timestamp"] == 2.0  # anchored on the trigger frame
        assert results[0]["min_confidence"] == 0.5

    def test_offset_and_miss_outside_window(self, monkeypatch):
        # change only matches at t=6, outside the [2, 5] window from anchor @2.
        def check(ts, ttype, step):
            if ttype == "color":
                return (ts == 2.0), {"_confidence": 0.9}
            return (ts == 6.0), {"magnitude": 0.5}

        self._setup_multiframe_stubs(monkeypatch, [0, 1, 2, 3, 4, 5, 6], check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color"},
                {"type": "change", "offset": {"min": 0, "max": 3}},
            ],
        )
        assert results == []

    def test_offset_not_passes_when_absent_in_window(self, monkeypatch):
        # NOT change never matches inside [2, 5] → the chain passes.
        def check(ts, ttype, step):
            if ttype == "color":
                return (ts == 2.0), {"_confidence": 0.8}
            return (ts == 6.0), {"magnitude": 0.5}  # only outside the window

        self._setup_multiframe_stubs(monkeypatch, [0, 1, 2, 3, 4, 5, 6], check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color"},
                {"type": "change", "logic": "NOT", "offset": {"min": 0, "max": 3}},
            ],
        )
        assert len(results) == 1
        assert results[0]["timestamp"] == 2.0
        assert results[0]["steps"][1] == {"negated": True, "type": "change"}
        assert results[0]["min_confidence"] == 0.8

    def test_offset_not_rejects_when_present_in_window(self, monkeypatch):
        # NOT change matches at t=4 inside [2, 5] → the chain fails.
        def check(ts, ttype, step):
            if ttype == "color":
                return (ts == 2.0), {"_confidence": 0.8}
            return (ts == 4.0), {"magnitude": 0.5}

        self._setup_multiframe_stubs(monkeypatch, [0, 1, 2, 3, 4, 5], check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color"},
                {"type": "change", "logic": "NOT", "offset": {"min": 0, "max": 3}},
            ],
        )
        assert results == []

    def test_cumulative_three_step_chain(self, monkeypatch):
        # color @1; change @2 (window [1,3] from anchor, ref→2); flow @4
        # (window [2,4] from step 1's match). Measured from the anchor (t=1)
        # step 2's window would be [1,3] and miss @4 — so a hit proves the
        # window is cumulative (relative to the previous step).
        def check(ts, ttype, step):
            if ttype == "color":
                return (ts == 1.0), {"_confidence": 0.9}
            if ttype == "change":
                return (ts == 2.0), {"magnitude": 0.6}
            return (ts == 4.0), {"magnitude": 0.5}  # flow

        self._setup_multiframe_stubs(monkeypatch, [1, 2, 3, 4], check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color"},
                {"type": "change", "offset": {"min": 0, "max": 2}},
                {"type": "flow", "offset": {"min": 0, "max": 2}},
            ],
        )
        assert len(results) == 1
        assert results[0]["timestamp"] == 1.0

    def test_offset_then_exact_frame_step(self, monkeypatch):
        # Mixed chain: step1 has an offset (advances ref to its match @2),
        # step2 has NO offset so it matches the same frame as step1's match.
        # flow matches only @2 (not the anchor @1) — a hit proves the
        # exact-frame step uses the *advanced* ref, not the anchor.
        def check(ts, ttype, step):
            if ttype == "color":
                return (ts == 1.0), {"_confidence": 0.9}
            if ttype == "change":
                return (ts == 2.0), {"magnitude": 0.6}
            return (ts == 2.0), {"magnitude": 0.5}  # flow, exact-frame on ref

        self._setup_multiframe_stubs(monkeypatch, [1, 2, 3], check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color"},
                {"type": "change", "offset": {"min": 0, "max": 2}},
                {"type": "flow"},  # no offset → same frame as step 1's match
            ],
        )
        assert len(results) == 1
        assert results[0]["timestamp"] == 1.0

    def test_offset_then_exact_frame_miss(self, monkeypatch):
        # Same mixed chain, but flow matches @3 (not on step1's matched frame
        # @2), so the exact-frame step misses and the chain fails.
        def check(ts, ttype, step):
            if ttype == "color":
                return (ts == 1.0), {"_confidence": 0.9}
            if ttype == "change":
                return (ts == 2.0), {"magnitude": 0.6}
            return (ts == 3.0), {"magnitude": 0.5}  # flow off the ref frame

        self._setup_multiframe_stubs(monkeypatch, [1, 2, 3], check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color"},
                {"type": "change", "offset": {"min": 0, "max": 2}},
                {"type": "flow"},
            ],
        )
        assert results == []

    def test_offset_negative_window(self, monkeypatch):
        # change @4 lies *before* the anchor @5, inside the [-2, 0] window.
        def check(ts, ttype, step):
            if ttype == "color":
                return (ts == 5.0), {"_confidence": 0.9}
            return (ts == 4.0), {"magnitude": 0.5}

        self._setup_multiframe_stubs(monkeypatch, [3, 4, 5], check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color"},
                {"type": "change", "offset": {"min": -2, "max": 0}},
            ],
        )
        assert len(results) == 1
        assert results[0]["timestamp"] == 5.0

    def test_offset_path_coalesces_adjacent_anchors(self, monkeypatch):
        # color matches at consecutive frames 2,3,4 (confidence peaks @3);
        # change matches everywhere. The three resolving anchors collapse into
        # one event represented by the highest-confidence anchor.
        def check(ts, ttype, step):
            if ttype == "color":
                conf = {2.0: 0.5, 3.0: 0.9, 4.0: 0.5}.get(ts, 0.0)
                return (ts in (2.0, 3.0, 4.0)), {"_confidence": conf}
            return True, {"magnitude": 1.0}

        self._setup_multiframe_stubs(monkeypatch, [2, 3, 4], check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color", "interval": 1.0},
                {"type": "change", "offset": {"min": 0, "max": 3}},
            ],
        )
        assert len(results) == 1
        # Representative = highest-confidence anchor (@3): min(color 0.9, change 1.0).
        assert results[0]["timestamp"] == 3.0
        assert results[0]["min_confidence"] == 0.9

    def test_offset_path_anchor_is_step0_timestamp(self, monkeypatch):
        # The hit lands on the step-0 frame, not the later matched frame.
        def check(ts, ttype, step):
            if ttype == "color":
                return (ts == 1.0), {"_confidence": 0.9}
            return (ts == 5.0), {"magnitude": 0.5}

        self._setup_multiframe_stubs(monkeypatch, [1, 2, 3, 4, 5], check)
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color"},
                {"type": "change", "offset": {"min": 0, "max": 5}},
            ],
        )
        assert len(results) == 1
        assert results[0]["timestamp"] == 1.0

    def test_offset_path_detect_first_stops_after_first(self, monkeypatch):
        # Three well-separated anchors all resolve; a cancel_flag that trips
        # once the first result is emitted (mirrors detect_first's on_result
        # hook) must stop after the first event.
        def check(ts, ttype, step):
            if ttype == "color":
                return (ts in (1.0, 5.0, 9.0)), {"_confidence": 0.9}
            return True, {"magnitude": 0.5}

        self._setup_multiframe_stubs(monkeypatch, [1, 5, 9], check)
        emitted: list = []
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[
                {"type": "color", "interval": 1.0},
                {"type": "change", "offset": {"min": 0, "max": 3}},
            ],
            on_result=emitted.append,
            cancel_flag=lambda: len(emitted) >= 1,
        )
        assert len(results) == 1
        assert len(emitted) == 1

    def test_no_offset_path_streams_per_frame(self, monkeypatch):
        # No step has an offset → the original short-circuit path runs and
        # emits per matching frame (streaming preserved).
        def check(ts, ttype, step):
            if ttype == "color":
                return (ts in (1.0, 2.0)), {"_confidence": 0.7}
            return (ts in (1.0, 2.0, 3.0)), {"magnitude": 0.5}

        self._setup_multiframe_stubs(monkeypatch, [1, 2, 3], check)
        emitted: list = []
        results = screenspace.scan_multitool(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            steps=[{"type": "color"}, {"type": "change"}],
            on_result=emitted.append,
        )
        assert [r["timestamp"] for r in results] == [1.0, 2.0]
        assert emitted == results  # streamed, one call per result


# ---------------------------------------------------------------------------
# Pin calibration scoring (Phase 2)
# ---------------------------------------------------------------------------


class TestScoreFrameForTool:
    region = {"x": 0, "y": 0, "w": 100, "h": 100}

    def test_color_miss_still_scored(self):
        frame = np.full((100, 100, 3), [255, 0, 0], dtype=np.uint8)  # blue
        params = {
            "target_color": {"h": 0, "s": 255, "v": 255},  # red target -> miss
            "tolerance": {"h": 5, "s": 10, "v": 10},
        }
        res = screenspace.score_frame_for_tool(
            "color", frame, None, self.region, params
        )
        assert res["status"] == "ok"
        assert res["passed"] is False
        assert math.isfinite(res["score"])

    def test_similarity_miss_still_scored(self):
        frame = np.zeros((100, 100, 3), dtype=np.uint8)
        ref = np.full((100, 100, 3), 255, dtype=np.uint8)
        params = {"reference_frame": ref, "threshold": 0.99}
        res = screenspace.score_frame_for_tool(
            "similarity", frame, None, self.region, params
        )
        assert res["status"] == "ok"
        assert res["passed"] is False
        assert "score" in res["detail"]

    def test_change_below_threshold_scored(self):
        a = np.zeros((100, 100, 3), dtype=np.uint8)
        b = np.zeros((100, 100, 3), dtype=np.uint8)  # identical -> magnitude 0
        res = screenspace.score_frame_for_tool(
            "change", b, a, self.region, {"threshold": 0.5}
        )
        assert res["status"] == "ok"
        assert res["passed"] is False
        assert res["score"] == 0.0

    def test_change_missing_companion_not_evaluable(self):
        frame = np.zeros((100, 100, 3), dtype=np.uint8)
        res = screenspace.score_frame_for_tool(
            "change", frame, None, self.region, {"threshold": 0.5}
        )
        assert res["status"] == "not_evaluable"

    def test_flow_below_threshold_scored(self):
        a = np.zeros((100, 100, 3), dtype=np.uint8)
        b = np.zeros((100, 100, 3), dtype=np.uint8)
        res = screenspace.score_frame_for_tool(
            "flow", b, a, self.region, {"magnitude_threshold": 1.0}
        )
        assert res["status"] == "ok"
        assert res["passed"] is False
        assert "magnitude" in res["detail"]

    def test_inactivity_miss_scored(self):
        a = np.zeros((100, 100, 3), dtype=np.uint8)
        rng = np.random.RandomState(7)
        b = rng.randint(0, 256, (100, 100, 3), dtype=np.uint8)
        res = screenspace.score_frame_for_tool(
            "inactivity", b, a, self.region, {"threshold": 2}
        )
        assert res["status"] == "ok"
        assert res["passed"] is False
        assert res["score"] >= 0  # raw phash distance

    def test_template_miss_scored(self):
        tmpl = np.zeros((20, 40, 3), dtype=np.uint8)
        tmpl[:, 20:] = 255  # half/half -> non-degenerate template
        rng = np.random.RandomState(3)
        frame = rng.randint(0, 256, (100, 100, 3), dtype=np.uint8)
        params = {"template_image": tmpl, "threshold": 0.99}
        res = screenspace.score_frame_for_tool(
            "template", frame, None, self.region, params
        )
        assert res["status"] == "ok"
        assert res["passed"] is False
        assert "best_score" in res["detail"]

    def test_scene_miss_scored(self):
        ref = np.full((100, 100, 3), [255, 0, 0], dtype=np.uint8)
        frame = np.full((100, 100, 3), [0, 255, 0], dtype=np.uint8)
        params = {
            "reference_scenes": [{"name": "menu", "frame": ref}],
            "threshold": 0.99,
        }
        res = screenspace.score_frame_for_tool(
            "scene", frame, None, self.region, params
        )
        assert res["status"] == "ok"
        assert "score" in res["detail"]

    def test_degenerate_region_not_evaluable(self):
        frame = np.zeros((100, 100, 3), dtype=np.uint8)
        res = screenspace.score_frame_for_tool(
            "color", frame, None, {"x": 0, "y": 0, "w": 0, "h": 0}, {}
        )
        assert res["status"] == "not_evaluable"

    def test_template_zero_region_scored_full_frame(self):
        # Uploaded template with no region_ref scans the full frame, so a
        # zero-size region must NOT be rejected as degenerate (template ignores
        # the region entirely).
        tmpl = np.zeros((20, 40, 3), dtype=np.uint8)
        tmpl[:, 20:] = 255  # non-degenerate template
        rng = np.random.RandomState(5)
        frame = rng.randint(0, 256, (100, 100, 3), dtype=np.uint8)
        res = screenspace.score_frame_for_tool(
            "template",
            frame,
            None,
            {"x": 0, "y": 0, "w": 0, "h": 0},
            {"template_image": tmpl, "threshold": 0.99},
        )
        assert res["status"] == "ok"
        assert "best_score" in res["detail"]

    def test_timelapse_not_calibratable(self):
        frame = np.zeros((100, 100, 3), dtype=np.uint8)
        res = screenspace.score_frame_for_tool(
            "timelapse", frame, None, self.region, {}
        )
        assert res["status"] == "not_evaluable"


class TestScoreMultitoolFrame:
    region = {"x": 0, "y": 0, "w": 100, "h": 100}

    def test_scores_every_step_even_when_one_fails(self):
        red = np.full((100, 100, 3), [0, 0, 255], dtype=np.uint8)
        steps = [
            {
                "type": "color",
                "region_coords": self.region,
                "target_color": {"h": 60, "s": 255, "v": 255},  # green -> fails on red
                "tolerance": {"h": 5, "s": 10, "v": 10},
            },
            {
                "type": "color",
                "logic": "AND",
                "region_coords": self.region,
                "target_color": screenspace.average_color_hsv(red),  # matches red
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
        ]
        res = screenspace.score_multitool_frame(red, None, steps)
        assert len(res["steps"]) == 2
        assert all(s["status"] == "ok" for s in res["steps"])
        assert res["steps"][0]["passed"] is False
        assert res["steps"][1]["passed"] is True
        assert res["passed"] is False  # AND chain: one step failed

    def test_color_presence_step_passes_on_small_patch(self):
        # A gray frame with a 1% dark-red patch: a presence-mode color step
        # fires where an average-mode step on the same target would not.
        frame = _gray_with_red_patch(10)
        target = {"h": 0, "s": 255, "v": 139}
        tol = {"h": 10, "s": 60, "v": 60}
        avg_step = [
            {
                "type": "color",
                "region_coords": self.region,
                "target_color": target,
                "tolerance": tol,
            },
        ]
        assert (
            screenspace.score_multitool_frame(frame, None, avg_step)["passed"] is False
        )
        presence_step = [
            {
                "type": "color",
                "region_coords": self.region,
                "target_color": target,
                "tolerance": tol,
                "color_mode": "presence",
            },
        ]
        assert (
            screenspace.score_multitool_frame(frame, None, presence_step)["passed"]
            is True
        )

    def test_not_evaluable_step_makes_chain_none(self):
        red = np.full((100, 100, 3), [0, 0, 255], dtype=np.uint8)
        steps = [
            {
                "type": "color",
                "region_coords": self.region,
                "target_color": screenspace.average_color_hsv(red),
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
            {
                "type": "change",
                "logic": "AND",
                "region_coords": self.region,
                "threshold": 0.5,
            },  # no companion frame -> not_evaluable
        ]
        res = screenspace.score_multitool_frame(red, None, steps)
        assert res["steps"][1]["status"] == "not_evaluable"
        assert res["passed"] is None

    def test_not_step_passes_chain_when_condition_misses(self):
        red = np.full((100, 100, 3), [0, 0, 255], dtype=np.uint8)
        steps = [
            {
                "type": "color",
                "region_coords": self.region,
                "target_color": screenspace.average_color_hsv(red),
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
            {
                "type": "color",
                "logic": "NOT",
                "region_coords": self.region,
                "target_color": {"h": 60, "s": 255, "v": 255},
                "tolerance": {"h": 5, "s": 10, "v": 10},
            },
        ]
        res = screenspace.score_multitool_frame(red, None, steps)
        assert res["steps"][1]["passed"] is False
        assert res["passed"] is True

    def test_not_step_fails_chain_when_condition_matches(self):
        red = np.full((100, 100, 3), [0, 0, 255], dtype=np.uint8)
        steps = [
            {
                "type": "color",
                "region_coords": self.region,
                "target_color": screenspace.average_color_hsv(red),
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
            {
                "type": "color",
                "logic": "NOT",
                "region_coords": self.region,
                "target_color": screenspace.average_color_hsv(red),
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
        ]
        res = screenspace.score_multitool_frame(red, None, steps)
        assert res["steps"][1]["passed"] is True
        assert res["passed"] is False

    def test_offset_is_ignored_in_calibration(self):
        # Calibration is single-frame: a step's offset window is not evaluated
        # here, so the score is identical with or without it (documented limit).
        red = np.full((100, 100, 3), [0, 0, 255], dtype=np.uint8)
        base = [
            {
                "type": "color",
                "region_coords": self.region,
                "target_color": screenspace.average_color_hsv(red),
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
            {
                "type": "color",
                "logic": "AND",
                "region_coords": self.region,
                "target_color": screenspace.average_color_hsv(red),
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
        ]
        without = screenspace.score_multitool_frame(red, None, base)
        with_offset = copy.deepcopy(base)
        with_offset[1]["offset"] = {"min": 2.0, "max": 5.0}
        scored = screenspace.score_multitool_frame(red, None, with_offset)
        assert scored["passed"] == without["passed"]
        assert [s["passed"] for s in scored["steps"]] == [
            s["passed"] for s in without["steps"]
        ]


# ---------------------------------------------------------------------------
# Per-frame memoization (crop / gray / phash / OCR reuse within a chain)
# ---------------------------------------------------------------------------


class TestRegionKeyMaskIdentity:
    """Same bbox + different polygon must not collide in the per-frame cache."""

    def test_mask_points_distinguish_cache_keys(self):
        import screenspace_tools

        rect = {"x": 0, "y": 0, "w": 40, "h": 40}
        tri = dict(rect, mask_points=[[[0, 0], [1, 0], [1, 1]]])
        tri2 = dict(rect, mask_points=[[[0, 0], [0, 1], [1, 1]]])
        keys = {screenspace_tools._region_key("crop", r) for r in (rect, tri, tri2)}
        assert len(keys) == 3

    def test_cached_mask_memoizes_and_distinguishes(self):
        import screenspace_tools

        frame = np.zeros((80, 80, 3), dtype=np.uint8)
        cache: dict = {}
        rect = {"x": 0, "y": 0, "w": 40, "h": 40}
        tri = dict(rect, mask_points=[[[0, 0], [1, 0], [1, 1]]])
        assert screenspace_tools._cached_mask(cache, frame, rect) is None
        m1 = screenspace_tools._cached_mask(cache, frame, tri)
        m2 = screenspace_tools._cached_mask(cache, frame, tri)
        assert m1 is m2
        assert m1 is not None and m1.shape == (40, 40)


class TestMultitoolMemoization:
    """The per-frame cache must not change detections vs the uncached path."""

    region = {"x": 4, "y": 4, "w": 48, "h": 36}

    @staticmethod
    def _frames():
        # Two deterministic frames with a bright block that moves between them,
        # so change / flow / inactivity all see real motion (non-trivial results).
        f0 = np.zeros((80, 80, 3), dtype=np.uint8)
        f0[10:40, 10:40] = (200, 120, 60)
        f1 = np.zeros((80, 80, 3), dtype=np.uint8)
        f1[16:46, 18:48] = (200, 120, 60)
        return f0, f1

    def _steps(self):
        # Spatial (color) + temporal (change/flow/inactivity) tools sharing one
        # region: exercises crop reuse across steps and prev-frame roll-forward.
        return [
            {
                "type": "color",
                "target_color": {"h": 15, "s": 180, "v": 200},
                "tolerance": {"h": 40, "s": 120, "v": 120},
            },
            {"type": "change"},
            {"type": "flow"},
            {"type": "inactivity"},
        ]

    def test_shared_cache_and_rollforward_match_uncached(self):
        f0, f1 = self._frames()
        steps = self._steps()
        region = self.region

        def uncached(frame, prev):
            return [
                screenspace.check_frame_for_tool(frame, prev, region, s["type"], s)
                for s in steps
            ]

        def cached(frame, prev, cache, prev_cache):
            # One shared cache per frame, passed to every step — mirrors _cb.
            return [
                screenspace.check_frame_for_tool(
                    frame, prev, region, s["type"], s, cache, prev_cache
                )
                for s in steps
            ]

        ref0, ref1 = uncached(f0, None), uncached(f1, f0)

        c0: dict = {}
        got0 = cached(f0, None, c0, None)
        c1: dict = {}
        got1 = cached(f1, f0, c1, c0)  # prev_cache=c0 exercises roll-forward

        assert got0 == ref0
        assert got1 == ref1

    def test_ocr_memoized_across_text_and_numbers_steps(self, monkeypatch):
        import screenspace_tools

        calls = {"n": 0}

        def fake_readings(
            pixels,
            *,
            languages=None,
            preprocess=False,
            mask_points=None,
        ):
            calls["n"] += 1
            return [([[0, 0], [10, 0], [10, 10], [0, 10]], "hello", 0.99)]

        monkeypatch.setattr(screenspace_tools, "_ocr_region_readings", fake_readings)

        frame = np.zeros((40, 40, 3), dtype=np.uint8)
        region = {"x": 0, "y": 0, "w": 40, "h": 40}
        text = {"type": "text", "search_string": "hello"}
        cache: dict = {}

        # Two identical Text steps sharing region + cache -> OCR runs once.
        screenspace.check_frame_for_tool(frame, None, region, "text", text, cache, None)
        screenspace.check_frame_for_tool(frame, None, region, "text", text, cache, None)
        assert calls["n"] == 1

        # Numbers on the same region now shares the memo: the engine has no
        # per-tool recognition mode, and integers_only/operators are applied at
        # scoring time. (Under EasyOCR the digit allowlist forced a re-run.)
        numbers = {"type": "numbers", "operator": "gt", "target_value": 0}
        screenspace.check_frame_for_tool(
            frame, None, region, "numbers", numbers, cache, None
        )
        assert calls["n"] == 1


# ---------------------------------------------------------------------------
# Fast Scan mode
# ---------------------------------------------------------------------------
