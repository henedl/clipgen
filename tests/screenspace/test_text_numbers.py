"""Tests for OCR text/number scanning and confidence scoring."""

from pathlib import Path
from types import SimpleNamespace

import numpy as np
import pytest

import config
import screenspace
import screenspace_frames
import screenspace_ocr
import screenspace_primitives
import screenspace_scans


@pytest.fixture(autouse=True)
def _reset_ocr_pool(monkeypatch):
    """Isolate the process-wide OCR engine pool between tests.

    Engines are cached in ``screenspace_ocr._ocr_pools``; pin the pool to a
    single slot so each test builds exactly one (patched) fake engine, and clear
    it so a fake from a prior test never leaks in.
    """
    monkeypatch.setattr(config, "SCREENSPACE_OCR_POOL_SIZE", 1)
    screenspace_ocr._ocr_pools.clear()
    # scan_text/scan_numbers call utils.require_optional("rapidocr") before the
    # fake engine is ever used; stub it so this file stays runnable in an
    # environment without the OCR stack. Production still fail-fasts; these
    # tests patch the engine constructor.
    import utils as _utils

    monkeypatch.setattr(_utils, "require_optional", lambda *_a, **_k: None)
    yield
    screenspace_ocr._ocr_pools.clear()


# The bbox every fake detection carries; scorers ignore it, the mask filter
# consumes it as four (x, y) points.
_BOX = [(0, 0), (10, 0), (10, 10), (0, 10)]


class _FakeEngine:
    """RapidOCR engine stub: a callable returning a RapidOCROutput-shaped result.

    Built from ``(bbox, text, conf)`` reading tuples (empty by default — detects
    nothing) so tests state detections in the same shape the adapter emits.
    ``on_call`` lets input-inspecting tests capture the pixels handed to the
    engine.
    """

    def __init__(self, readings=(), on_call=None):
        self._readings = list(readings)
        self._on_call = on_call

    def __call__(self, pixels):
        if self._on_call is not None:
            self._on_call(pixels)
        if not self._readings:
            return SimpleNamespace(boxes=None, txts=None, scores=None)
        return SimpleNamespace(
            boxes=[r[0] for r in self._readings],
            txts=tuple(r[1] for r in self._readings),
            scores=tuple(r[2] for r in self._readings),
        )


class TestOcrPreprocess:
    def test_small_roi_upscaled(self):
        """Crops shorter than the min height are upscaled, preserving aspect."""
        small = np.full((20, 120, 3), 128, dtype=np.uint8)
        out = screenspace._preprocess_for_ocr(small)
        assert out.shape[0] >= 60
        # Aspect ratio preserved (6:1 → width scales with height).
        assert out.shape[1] >= 360

    def test_large_roi_size_unchanged(self):
        """Crops already tall enough are not resized."""
        large = np.full((200, 500, 3), 128, dtype=np.uint8)
        out = screenspace._preprocess_for_ocr(large)
        assert out.shape[0] == 200
        assert out.shape[1] == 500

    def test_clahe_increases_contrast(self):
        """CLAHE stretches a low-contrast crop, raising pixel variance."""
        base = np.random.RandomState(0).randint(100, 116, (60, 120)).astype(np.uint8)
        lowc = np.stack([base, base, base], axis=-1)  # equal channels → true gray
        out = screenspace._preprocess_for_ocr(lowc)
        in_var = float(np.var(lowc[:, :, 0]))
        out_var = float(np.var(out[:, :, 0]))
        assert out_var > in_var

    def test_returns_three_channels(self):
        small = np.full((20, 120, 3), 128, dtype=np.uint8)
        out = screenspace._preprocess_for_ocr(small)
        assert out.ndim == 3 and out.shape[2] == 3

    def test_empty_region_returned_unchanged(self):
        empty = np.zeros((0, 10, 3), dtype=np.uint8)
        out = screenspace._preprocess_for_ocr(empty)
        assert out.shape[0] == 0

    def test_scan_text_preprocess_enlarges_ocr_input(self, monkeypatch):
        """ocr_preprocess=True feeds an upscaled crop to the OCR reader."""
        frame = np.full((20, 120, 3), 128, dtype=np.uint8)
        seen: dict[str, int] = {}

        monkeypatch.setattr(
            screenspace_ocr,
            "_build_ocr_reader",
            lambda _m: _FakeEngine(on_call=lambda px: seen.update(h=px.shape[0])),
        )
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )
        monkeypatch.setattr(
            screenspace_scans,
            "scan_video_frames",
            lambda v, r, i, cb, **k: cb(0.0, frame),
        )

        screenspace.scan_text(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 120, "h": 20},
            search_string="x",
            ocr_preprocess=True,
        )
        assert seen["h"] >= 60

    def test_scan_text_no_preprocess_keeps_native_size(self, monkeypatch):
        """Default (ocr_preprocess=False) passes the raw crop to the reader."""
        frame = np.full((20, 120, 3), 128, dtype=np.uint8)
        seen: dict[str, int] = {}

        monkeypatch.setattr(
            screenspace_ocr,
            "_build_ocr_reader",
            lambda _m: _FakeEngine(on_call=lambda px: seen.update(h=px.shape[0])),
        )
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )
        monkeypatch.setattr(
            screenspace_scans,
            "scan_video_frames",
            lambda v, r, i, cb, **k: cb(0.0, frame),
        )

        screenspace.scan_text(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 120, "h": 20},
            search_string="x",
            ocr_preprocess=False,
        )
        assert seen["h"] == 20


class TestScanText:
    def test_low_confidence_rejected(self, monkeypatch):
        """OCR readings below ocr_confidence_threshold should not match."""
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)

        monkeypatch.setattr(
            screenspace_ocr,
            "_build_ocr_reader",
            lambda _m: _FakeEngine([(_BOX, "hello", 0.2)]),
        )
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)

        rejected = screenspace.scan_text(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 20},
            search_string="hello",
            ocr_confidence_threshold=0.5,
        )
        assert rejected == []

        accepted = screenspace.scan_text(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 20},
            search_string="hello",
            ocr_confidence_threshold=0.1,
        )
        assert len(accepted) == 1
        assert accepted[0]["text_found"] == "hello"
        assert accepted[0]["confidence"] == 0.2

    def test_normalize_letters_to_digits(self, monkeypatch):
        """ocr_normalize="digits" folds l→1 so "l00" matches a search for "100"."""
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)

        monkeypatch.setattr(
            screenspace_ocr,
            "_build_ocr_reader",
            lambda _m: _FakeEngine([(_BOX, "l00", 0.9)]),
        )
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        # Default (ocr_normalize="off"): "l00" vs "100" stays below threshold.
        rejected = screenspace.scan_text(
            "/fake.mp4", region, search_string="100", fuzzy_threshold=0.9
        )
        assert rejected == []

        # ocr_normalize="digits": l→1 makes it an exact match.
        accepted = screenspace.scan_text(
            "/fake.mp4",
            region,
            search_string="100",
            fuzzy_threshold=0.9,
            ocr_normalize="digits",
        )
        assert len(accepted) == 1
        assert accepted[0]["text_found"] == "l00"

    def test_normalize_digits_to_letters(self, monkeypatch):
        """ocr_normalize="letters" folds 5→s so "5top" matches a search for "stop"."""
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)

        monkeypatch.setattr(
            screenspace_ocr,
            "_build_ocr_reader",
            lambda _m: _FakeEngine([(_BOX, "5top", 0.9)]),
        )
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        # Default ("off"): "5top" vs "stop" stays below threshold.
        assert (
            screenspace.scan_text(
                "/fake.mp4", region, search_string="stop", fuzzy_threshold=0.9
            )
            == []
        )

        # ocr_normalize="letters": 5→s makes it an exact match.
        accepted = screenspace.scan_text(
            "/fake.mp4",
            region,
            search_string="stop",
            fuzzy_threshold=0.9,
            ocr_normalize="letters",
        )
        assert len(accepted) == 1
        assert accepted[0]["text_found"] == "5top"

    def test_normalize_direction_is_distinct(self, monkeypatch):
        """The two fold directions differ: i→1 is digits-only (no 1→i inverse)."""
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)

        monkeypatch.setattr(
            screenspace_ocr,
            "_build_ocr_reader",
            lambda _m: _FakeEngine([(_BOX, "1n", 0.9)]),
        )
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        # "digits" folds search i→1, matching OCR "1n"; "letters" folds OCR 1→l
        # ("ln") which stays below threshold against "in".
        assert (
            len(
                screenspace.scan_text(
                    "/fake.mp4",
                    region,
                    search_string="in",
                    fuzzy_threshold=0.9,
                    ocr_normalize="digits",
                )
            )
            == 1
        )
        assert (
            screenspace.scan_text(
                "/fake.mp4",
                region,
                search_string="in",
                fuzzy_threshold=0.9,
                ocr_normalize="letters",
            )
            == []
        )

    def test_require_consecutive(self, monkeypatch):
        """require_consecutive=N coalesces N consecutive matches into one median event."""
        # Distinct fills so the static-frame-skip never fires between frames.
        frames = [np.full((20, 60, 3), v, dtype=np.uint8) for v in (40, 90, 140)]

        monkeypatch.setattr(
            screenspace_ocr,
            "_build_ocr_reader",
            lambda _m: _FakeEngine([(_BOX, "hello", 0.9)]),
        )
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for i, frame in enumerate(frames):
                callback(float(i), frame)

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        # Default (require_consecutive=1): one event per matching frame.
        default = screenspace.scan_text(
            "/fake.mp4", region, search_string="hello", fuzzy_threshold=0.8
        )
        assert len(default) == 3

        # require_consecutive=3: a single event stamped with the median timestamp.
        coalesced = screenspace.scan_text(
            "/fake.mp4",
            region,
            search_string="hello",
            fuzzy_threshold=0.8,
            require_consecutive=3,
        )
        assert len(coalesced) == 1
        assert coalesced[0]["timestamp"] == 1.0  # median([0.0, 1.0, 2.0])

    def test_require_consecutive_survives_static_frames(self, monkeypatch):
        """Static (skipped) frames carry an active run forward instead of starving it.

        The first frame matches and is OCR'd; the next two are identical to it,
        so the static-frame-skip drops them before OCR. Carry-over keeps the
        require_consecutive=3 run alive, so one event still emits (the old
        behavior would have stalled at one push and emitted nothing)."""
        base = np.full((20, 60, 3), 40, dtype=np.uint8)
        frames = [base, base.copy(), base.copy()]  # identical → static-skip fires

        reads = {"n": 0}

        monkeypatch.setattr(
            screenspace_ocr,
            "_build_ocr_reader",
            lambda _m: _FakeEngine(
                [(_BOX, "hello", 0.9)],
                on_call=lambda _px: reads.update(n=reads["n"] + 1),
            ),
        )
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for i, frame in enumerate(frames):
                callback(float(i), frame)

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}
        out = screenspace.scan_text(
            "/fake.mp4",
            region,
            search_string="hello",
            fuzzy_threshold=0.8,
            require_consecutive=3,
        )
        assert reads["n"] == 1  # frames 2 & 3 were skipped as static (no OCR)
        assert len(out) == 1
        assert out[0]["timestamp"] == 1.0  # median([0, 1, 2]) across carried frames


class TestConsecutiveBuffer:
    def test_n1_emits_immediately(self):
        buf = screenspace._ConsecutiveBuffer(1)
        out = buf.push(5.0, {"timestamp": 5.0, "magnitude": 0.4})
        assert out is not None
        assert out["timestamp"] == 5.0
        assert out["magnitude"] == 0.4

    def test_emits_after_n_with_median_ts(self):
        buf = screenspace._ConsecutiveBuffer(3)
        assert buf.push(10.0, {"timestamp": 10.0, "v": "a"}) is None
        assert buf.push(12.0, {"timestamp": 12.0, "v": "b"}) is None
        out = buf.push(20.0, {"timestamp": 20.0, "v": "c"})
        assert out is not None
        assert out["timestamp"] == 12.0  # median([10, 12, 20])
        assert out["v"] == "b"  # middle frame's payload

    def test_miss_clears(self):
        buf = screenspace._ConsecutiveBuffer(3)
        assert buf.push(0.0, {"timestamp": 0.0}) is None
        assert buf.push(1.0, {"timestamp": 1.0}) is None
        buf.reset()
        # Only two matches accumulate after the reset, so nothing emits.
        assert buf.push(2.0, {"timestamp": 2.0}) is None
        assert buf.push(3.0, {"timestamp": 3.0}) is None

    def test_size_floor_of_one(self):
        # 0 / negative sizes clamp to 1 (passthrough behavior).
        buf = screenspace._ConsecutiveBuffer(0)
        assert buf.push(7.0, {"timestamp": 7.0}) is not None

    def test_carry_continues_active_run(self):
        # A static (skipped) frame carries the last match forward so the run
        # still reaches the threshold on stable content.
        buf = screenspace._ConsecutiveBuffer(3)
        assert buf.push(0.0, {"timestamp": 0.0, "text_found": "Save"}) is None
        assert buf.carry(1.0) is None  # static frame #1
        out = buf.carry(2.0)  # static frame #2 completes the run
        assert out is not None
        assert out["text_found"] == "Save"
        assert out["timestamp"] == 1.0  # median([0, 1, 2])

    def test_carry_noop_when_no_active_run(self):
        # Nothing to carry before any match has been pushed.
        buf = screenspace._ConsecutiveBuffer(3)
        assert buf.carry(5.0) is None

    def test_carry_noop_for_size_one(self):
        # size==1 emits and resets on every push, so no run is ever active to
        # carry — keeps the legacy passthrough path unchanged.
        buf = screenspace._ConsecutiveBuffer(1)
        assert buf.push(1.0, {"timestamp": 1.0}) is not None
        assert buf.carry(2.0) is None

    def test_even_size_pairs_median_with_nearest_frame(self):
        # Even runs interpolate the median between two frames; the payload comes
        # from the nearer real frame, not an arbitrary upper-middle one.
        buf = screenspace._ConsecutiveBuffer(2)
        assert buf.push(0.0, {"timestamp": 0.0, "v": "a"}) is None
        out = buf.push(4.0, {"timestamp": 4.0, "v": "b"})
        assert out is not None
        assert out["timestamp"] == 2.0  # median([0, 4])
        assert out["v"] in ("a", "b")  # nearest real frame to the median


def test_static_skip_uses_config():
    """The static-frame-skip sites reference the config constant, not a 2.0
    literal. scan_text/scan_numbers share the _is_static_skip helper and
    similarity/flow/inactivity/boundaries/scene/template share the
    _frame_is_static predicate (both in screenspace_primitives)."""
    src = Path(screenspace_primitives.__file__).read_text(encoding="utf-8")
    src += Path(screenspace_scans.__file__).read_text(encoding="utf-8")
    # Both shared helpers (_is_static_skip, _frame_is_static) live in
    # screenspace_primitives and reference the constant; every scan routes
    # through one of them rather than inlining the threshold.
    assert src.count("config.SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD") >= 2
    assert "< 2.0" not in src


class TestScanNumbers:
    def test_unknown_operator_raises(self):
        with pytest.raises(ValueError, match="Unknown.*operator"):
            screenspace.scan_numbers(
                "/nonexistent.mp4",
                {"x": 0, "y": 0, "w": 10, "h": 10},
                operator="invalid",
                target_value=100,
            )

    def test_valid_operators_accepted(self, monkeypatch):
        # scan_numbers builds the OCR engine before opening the video, so a
        # missing-video test would otherwise construct the real RapidOCR
        # engine. Stub it out to stay fast and offline.
        monkeypatch.setattr(
            screenspace_ocr, "_build_ocr_reader", lambda _m: _FakeEngine()
        )
        for op in ("eq", "gt", "lt", "gte", "lte", "range"):
            # Should not raise ValueError -- returns [] because video doesn't exist
            result = screenspace.scan_numbers(
                "/nonexistent.mp4",
                {"x": 0, "y": 0, "w": 10, "h": 10},
                operator=op,
                target_value=100,
                range_min=0,
                range_max=200,
            )
            assert result == []

    def test_dispatch_routes_numbers(self, monkeypatch):
        # Same reason as test_valid_operators_accepted: avoid loading real OCR.
        monkeypatch.setattr(
            screenspace_ocr, "_build_ocr_reader", lambda _m: _FakeEngine()
        )
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "numbers",
            "P01",
            "s_P01.mp4",
            ["/nonexistent.mp4"],
            "r",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            parameters={"operator": "gt", "target_value": 50},
        )
        # Should not raise -- dispatches to scan_numbers, returns [] for bad video
        result = worker._dispatch(task, lambda p: None, lambda: False)
        assert result == []

    def test_dispatch_unknown_type_raises(self):
        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "bogus",
            "P01",
            "s.mp4",
            ["/v.mp4"],
            "r",
            {"x": 0, "y": 0, "w": 1, "h": 1},
        )
        with pytest.raises(ValueError, match="Unknown task type"):
            worker._dispatch(task, lambda p: None, lambda: False)

    def test_low_confidence_rejected(self, monkeypatch):
        """OCR readings below ocr_confidence_threshold should not match."""
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)

        monkeypatch.setattr(
            screenspace_ocr,
            "_build_ocr_reader",
            lambda _m: _FakeEngine([(_BOX, "5", 0.2)]),
        )
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)

        default_rejected = screenspace.scan_numbers(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 20},
            operator="gte",
            target_value=5,
        )
        assert default_rejected == []

        rejected = screenspace.scan_numbers(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 20},
            operator="gte",
            target_value=5,
            ocr_confidence_threshold=0.5,
        )
        assert rejected == []

        accepted = screenspace.scan_numbers(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 20},
            operator="gte",
            target_value=5,
            ocr_confidence_threshold=0.1,
        )
        assert len(accepted) == 1
        assert accepted[0]["number_found"] == 5.0
        assert accepted[0]["confidence"] == 0.2

        accepted_zero = screenspace.scan_numbers(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 60, "h": 20},
            operator="gte",
            target_value=5,
            ocr_confidence_threshold=0.0,
        )
        assert len(accepted_zero) == 1

    @pytest.mark.parametrize("threshold", [-0.1, 1.1])
    def test_invalid_ocr_confidence_threshold_raises(self, monkeypatch, threshold):
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )
        with pytest.raises(ValueError, match="ocr_confidence_threshold"):
            screenspace.scan_numbers(
                "/fake.mp4",
                {"x": 0, "y": 0, "w": 60, "h": 20},
                operator="gte",
                target_value=5,
                ocr_confidence_threshold=threshold,
            )

    def test_integers_only_rejects_decimals_and_signs(self, monkeypatch):
        """integers_only post-filters extracted values: no decimals, no signs.

        (Replaces the old English-only EasyOCR recognition allowlist — the
        filter now applies to every language.)
        """
        frame = np.full((20, 60, 3), 128, dtype=np.uint8)

        monkeypatch.setattr(
            screenspace_ocr,
            "_build_ocr_reader",
            lambda _m: _FakeEngine(
                [(_BOX, "3.5", 0.9), (_BOX, "-12", 0.9), (_BOX, "1,234", 0.9)]
            ),
        )
        monkeypatch.setattr(
            screenspace_frames, "_probe_video_meta", lambda p: (30.0, 1.0)
        )

        def fake_scan(video_path, region, interval, callback, **kwargs):
            callback(0.0, frame)

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)

        region = {"x": 0, "y": 0, "w": 60, "h": 20}

        # integers_only=True: "3.5" and "-12" are rejected outright; the
        # comma-separated "1,234" is a whole number and survives.
        results = screenspace.scan_numbers(
            "/fake.mp4",
            region,
            operator="gt",
            target_value=0,
            integers_only=True,
            ocr_confidence_threshold=0.5,
        )
        assert len(results) == 1
        assert results[0]["number_found"] == 1234.0

        # integers_only=False: the decimal reading matches too (best-satisfying
        # reading wins; all confidences are equal so any is acceptable).
        loose = screenspace.scan_numbers(
            "/fake.mp4",
            region,
            operator="gt",
            target_value=0,
            ocr_confidence_threshold=0.5,
        )
        assert len(loose) == 1

        # The filter is language-independent now.
        non_english = screenspace.scan_numbers(
            "/fake.mp4",
            region,
            operator="lt",
            target_value=0,
            integers_only=True,
            languages=["ja"],
            ocr_confidence_threshold=0.5,
        )
        assert non_english == []  # only "-12" satisfies lt 0, and it's rejected


# ---------------------------------------------------------------------------
# Flow grid and heatmap visualization
# ---------------------------------------------------------------------------


class TestExtractConfidence:
    def test_color(self):
        assert screenspace._extract_confidence("color", {"_confidence": 0.8}) == 0.8

    def test_change(self):
        assert screenspace._extract_confidence("change", {"magnitude": 0.5}) == 0.5

    def test_similarity(self):
        assert screenspace._extract_confidence("similarity", {"score": 0.95}) == 0.95

    def test_numbers(self):
        assert screenspace._extract_confidence("numbers", {}) == 1.0

    def test_numbers_uses_ocr_conf(self):
        assert screenspace._extract_confidence("numbers", {"confidence": 0.42}) == 0.42

    def test_multitool(self):
        assert (
            screenspace._extract_confidence("multitool", {"min_confidence": 0.7}) == 0.7
        )

    def test_inactivity_short(self):
        assert screenspace._extract_confidence("inactivity", {"duration": 3.0}) == 0.1

    def test_inactivity_per_frame_confidence(self):
        assert (
            screenspace._extract_confidence(
                "inactivity", {"distance": 2, "_confidence": 0.8}
            )
            == 0.8
        )

    def test_inactivity_capped(self):
        assert screenspace._extract_confidence("inactivity", {"duration": 60.0}) == 1.0

    def test_boundary_per_frame_confidence(self):
        assert (
            screenspace._extract_confidence(
                "boundary", {"distance": 20, "_confidence": 0.43}
            )
            == 0.43
        )

    def test_boundary_from_distance(self):
        # Fallback when _confidence is absent: (distance - threshold) / threshold,
        # clamped to [0, 1] against the default threshold (14).
        assert screenspace._extract_confidence("boundary", {"distance": 28}) == 1.0
        assert screenspace._extract_confidence("boundary", {"distance": 21}) == 0.5
        assert screenspace._extract_confidence("boundary", {"distance": 10}) == 0.0

    def test_unknown_type(self):
        assert screenspace._extract_confidence("bogus", {}) == 1.0


class TestScoreOcrReadings:
    def test_text_best_ratio_among_clearing_readings(self):
        readings = [
            (None, "hello", 0.9),
            (None, "xxxxx", 0.95),
            (None, "hello", 0.2),  # below the conf floor -> ignored
        ]
        params = {
            "search_string": "hello",
            "fuzzy_threshold": 0.8,
            "ocr_confidence_threshold": 0.5,
        }
        passed, detail = screenspace._score_text_readings(readings, params)
        assert passed is True
        assert detail["fuzzy_ratio"] == 1.0
        assert detail["text_found"] == "hello"

    def test_text_no_clearing_reading_scores_zero(self):
        readings = [(None, "hello", 0.2)]
        params = {
            "search_string": "hello",
            "fuzzy_threshold": 0.8,
            "ocr_confidence_threshold": 0.5,
        }
        passed, detail = screenspace._score_text_readings(readings, params)
        assert passed is False
        assert detail["fuzzy_ratio"] == 0.0
        assert detail["text_found"] == ""

    def test_numbers_matching_conf_is_scalar_operator_is_polarity(self):
        readings = [(None, "5", 0.9), (None, "99", 0.6)]
        params = {
            "operator": "gte",
            "target_value": 50,
            "ocr_confidence_threshold": 0.5,
        }
        passed, detail = screenspace._score_numbers_readings(readings, params)
        # Scalar = best OCR confidence among readings satisfying the operator.
        assert detail["confidence"] == 0.6
        assert passed is True
        assert detail["number_found"] == 99.0

    def test_numbers_no_operator_match(self):
        readings = [(None, "5", 0.9)]
        params = {
            "operator": "gte",
            "target_value": 50,
            "ocr_confidence_threshold": 0.5,
        }
        passed, detail = screenspace._score_numbers_readings(readings, params)
        assert passed is False
        assert detail["confidence"] == 0.0
        assert "number_found" not in detail

    def test_numbers_low_confidence_operator_match_still_scores(self):
        readings = [(None, "99", 0.4)]
        params = {
            "operator": "gte",
            "target_value": 50,
            "ocr_confidence_threshold": 0.5,
        }
        passed, detail = screenspace._score_numbers_readings(readings, params)
        assert passed is False
        assert detail["confidence"] == 0.4
        assert detail["number_found"] == 99.0

    def test_numbers_integers_only_is_a_scoring_filter(self):
        readings = [(None, "3.5", 0.9), (None, "-12", 0.9), (None, "1,234", 0.8)]
        params = {
            "operator": "gt",
            "target_value": 0,
            "ocr_confidence_threshold": 0.5,
            "integers_only": True,
        }
        passed, detail = screenspace._score_numbers_readings(readings, params)
        assert passed is True
        assert detail["number_found"] == 1234.0  # 3.5 and -12 rejected outright


class TestReadingsFromResult:
    def test_empty_result_maps_to_no_readings(self):
        empty = SimpleNamespace(boxes=None, txts=None, scores=None)
        assert screenspace_ocr._readings_from_result(empty) == []
        assert screenspace_ocr._readings_from_result(object()) == []

    def test_ndarray_boxes_become_plain_float_points(self):
        # RapidOCR emits boxes as a (N, 4, 2) ndarray; readings cross the
        # server's pin OCR cache boundary, so nothing numpy may survive.
        result = SimpleNamespace(
            boxes=np.array([[[0, 0], [10, 0], [10, 10], [0, 10]]], dtype=np.float32),
            txts=("hello",),
            scores=(0.9,),
        )
        readings = screenspace_ocr._readings_from_result(result)
        assert readings == [
            ([(0.0, 0.0), (10.0, 0.0), (10.0, 10.0), (0.0, 10.0)], "hello", 0.9)
        ]
        bbox, text, conf = readings[0]
        assert all(isinstance(v, float) for point in bbox for v in point)
        assert isinstance(text, str) and isinstance(conf, float)

    def test_missing_scores_default_to_zero(self):
        result = SimpleNamespace(boxes=[_BOX], txts=("hi",), scores=None)
        readings = screenspace_ocr._readings_from_result(result)
        assert readings[0][2] == 0.0


class TestResolveOcrModel:
    def test_language_table(self):
        assert screenspace._resolve_ocr_model(None) == "default"
        assert screenspace._resolve_ocr_model(["en"]) == "default"
        assert screenspace._resolve_ocr_model(["zh"]) == "default"
        for lang in ("es", "fr", "de"):
            assert screenspace._resolve_ocr_model([lang]) == "latin"
        assert screenspace._resolve_ocr_model(["ja"]) == "japan"
        assert screenspace._resolve_ocr_model(["ko"]) == "korean"

    def test_english_plus_latin_resolves_to_latin(self):
        # The latin model covers English glyphs, so the mix is servable.
        assert screenspace._resolve_ocr_model(["en", "de"]) == "latin"

    def test_incompatible_mix_raises(self):
        with pytest.raises(ValueError, match="incompatible"):
            screenspace._resolve_ocr_model(["ja", "ko"])

    def test_unknown_language_raises(self):
        with pytest.raises(ValueError, match="Unsupported"):
            screenspace._resolve_ocr_model(["ch_sim"])
