"""Tests for inactivity scanning and morphology kernel."""

import numpy as np

import screenspace
import screenspace_scans
import screenspace_tools


class TestScanInactivity:
    """Tests for scan_inactivity() function."""

    def test_identical_frames_detected(self, monkeypatch):
        """Identical consecutive frames should produce an inactivity span."""
        frame = np.full((50, 50, 3), 128, dtype=np.uint8)
        timestamps = [0.0, 1.0, 2.0, 3.0, 4.0]
        call_idx = [0]

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for ts in timestamps:
                result = callback(ts, frame.copy())
                if result is False:
                    break
                call_idx[0] += 1

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)
        monkeypatch.setattr(
            screenspace_scans, "_probe_video_meta", lambda p: (30.0, 5.0)
        )

        results = screenspace.scan_inactivity(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 50, "h": 50},
            threshold=15,
            min_duration=2.0,
            interval_seconds=1.0,
        )
        assert len(results) == 1
        assert results[0]["start"] == 0.0
        assert results[0]["end"] == 4.0
        assert results[0]["duration"] == 4.0
        assert results[0]["avg_distance"] == 0.0

    def test_span_start_not_negative_early_match(self, monkeypatch):
        """First similar pair with ts < interval must not start before 0:00."""
        frame = np.full((50, 50, 3), 128, dtype=np.uint8)
        # First similar pair lands at ts=0.5 with a 1.0s interval, so the
        # naive ``ts - interval_seconds`` would be -0.5.
        timestamps = [0.0, 0.5, 1.5, 2.5, 3.5]

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for ts in timestamps:
                if callback(ts, frame.copy()) is False:
                    break

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)
        monkeypatch.setattr(
            screenspace_scans, "_probe_video_meta", lambda p: (30.0, 5.0)
        )

        results = screenspace.scan_inactivity(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 50, "h": 50},
            threshold=15,
            min_duration=1.0,
            interval_seconds=1.0,
        )
        assert len(results) == 1
        assert results[0]["start"] >= 0.0

    def test_span_start_clamped_to_scan_start(self, monkeypatch):
        """With start_seconds > 0 the span can't begin before the requested start."""
        frame = np.full((50, 50, 3), 128, dtype=np.uint8)
        timestamps = [10.0, 11.0, 12.0, 13.0]

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for ts in timestamps:
                if callback(ts, frame.copy()) is False:
                    break

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)
        monkeypatch.setattr(
            screenspace_scans, "_probe_video_meta", lambda p: (30.0, 20.0)
        )

        results = screenspace.scan_inactivity(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 50, "h": 50},
            threshold=15,
            min_duration=1.0,
            interval_seconds=1.0,
            start_seconds=10.0,
        )
        assert len(results) == 1
        assert results[0]["start"] >= 10.0

    def test_different_frames_not_detected(self, monkeypatch):
        """Frames with very different content should not produce a span."""
        timestamps = [0.0, 1.0, 2.0, 3.0]
        # Use random noise frames with different seeds for visually distinct content
        frames = [
            np.random.RandomState(seed).randint(0, 256, (50, 50, 3)).astype(np.uint8)
            for seed in [10, 20, 30, 40]
        ]

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for i, ts in enumerate(timestamps):
                result = callback(ts, frames[i])
                if result is False:
                    break

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)
        monkeypatch.setattr(
            screenspace_scans, "_probe_video_meta", lambda p: (30.0, 4.0)
        )

        results = screenspace.scan_inactivity(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 50, "h": 50},
            threshold=2,
            min_duration=2.0,
            interval_seconds=1.0,
        )
        assert len(results) == 0

    def test_min_duration_filtering(self, monkeypatch):
        """Spans shorter than min_duration should be discarded."""
        frame = np.full((50, 50, 3), 128, dtype=np.uint8)
        # Random noise frame to ensure phash is very different from the solid frame
        diff_frame = (
            np.random.RandomState(99).randint(0, 256, (50, 50, 3)).astype(np.uint8)
        )
        # 2 identical frames then a different one — span is only 1s
        timestamps = [0.0, 1.0, 2.0]
        frame_seq = [frame, frame.copy(), diff_frame]

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for i, ts in enumerate(timestamps):
                result = callback(ts, frame_seq[i])
                if result is False:
                    break

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)
        monkeypatch.setattr(
            screenspace_scans, "_probe_video_meta", lambda p: (30.0, 3.0)
        )

        results = screenspace.scan_inactivity(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 50, "h": 50},
            threshold=15,
            min_duration=5.0,
            interval_seconds=1.0,
        )
        assert len(results) == 0

    def test_on_result_callback(self, monkeypatch):
        """on_result should fire once per completed span."""
        frame = np.full((50, 50, 3), 128, dtype=np.uint8)
        timestamps = [0.0, 1.0, 2.0, 3.0, 4.0]
        emitted = []

        def fake_scan(video_path, region, interval, callback, **kwargs):
            for ts in timestamps:
                result = callback(ts, frame.copy())
                if result is False:
                    break

        monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)
        monkeypatch.setattr(
            screenspace_scans, "_probe_video_meta", lambda p: (30.0, 5.0)
        )

        screenspace.scan_inactivity(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 50, "h": 50},
            threshold=15,
            min_duration=2.0,
            interval_seconds=1.0,
            on_result=lambda r: emitted.append(r),
        )
        assert len(emitted) == 1
        assert emitted[0]["duration"] == 4.0

    def test_dispatch_routes_to_scan_inactivity(self, monkeypatch):
        """ScreenspaceWorker._dispatch() should route inactivity tasks."""
        captured = {}

        def fake_scan_inactivity(video_path, region, **kwargs):
            captured.update(kwargs)
            return []

        monkeypatch.setattr(screenspace_tools, "scan_inactivity", fake_scan_inactivity)

        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "inactivity",
            "P01",
            "s_P01.mp4",
            ["/fake.mp4"],
            "region1",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            parameters={
                "threshold": 8,
                "min_duration": 5.0,
                "interval": 2.0,
            },
        )
        worker._dispatch(task, lambda p: None, lambda: False, None)

        assert captured["threshold"] == 8
        assert captured["min_duration"] == 5.0
        assert captured["interval_seconds"] == 2.0


class TestMorphKernel:
    def test_shape_and_dtype(self):
        kernel = screenspace._morph_kernel(3)
        assert kernel.shape == (3, 3)
        assert kernel.dtype == np.uint8
        assert np.all(kernel == 1)

    def test_same_size_returns_cached_array(self):
        # cv2 morphology treats the kernel read-only, so callers share one array.
        assert screenspace._morph_kernel(5) is screenspace._morph_kernel(5)
