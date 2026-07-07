"""Static-frame gating parity + savings for the expensive scans.

Each scan that gained a cheap gray-diff pre-gate (flow, inactivity, boundaries
phash path, template) is exercised two ways on the same frozen-then-changing
frame sequence:

* **gate on** — the real code path.
* **gate off** — ``_frame_is_static`` monkeypatched to always return False,
  reproducing the pre-gate behaviour.

The gate must be *result-equivalent* (same spans/events/rows) while invoking the
expensive per-frame op (Farneback / phash / matchTemplate) strictly fewer times.
"""

import cv2
import numpy as np

import screenspace
import screenspace_primitives
import screenspace_scans
from _ss_helpers import _make_icon, _make_icon_frame

_REGION = {"x": 0, "y": 0, "w": 60, "h": 60}


def _disable_gate(monkeypatch):
    monkeypatch.setattr(screenspace_scans, "_frame_is_static", lambda prev, curr: False)


def _feed_region(monkeypatch, sequence):
    """Patch the region-frame scan driver to replay ``[(ts, frame), ...]``."""

    def fake_scan(video_path, region, interval, callback, **kwargs):
        for ts, frame in sequence:
            if callback(ts, frame.copy()) is False:
                break

    monkeypatch.setattr(screenspace_scans, "scan_video_frames", fake_scan)
    monkeypatch.setattr(screenspace_scans, "_probe_video_meta", lambda p: (30.0, 10.0))


def _feed_full(monkeypatch, sequence):
    """Patch the full-frame scan driver to replay ``[(ts, frame), ...]``."""

    def fake_scan(video_path, interval, callback, **kwargs):
        for ts, frame in sequence:
            if callback(ts, frame.copy()) is False:
                break

    monkeypatch.setattr(screenspace_scans, "scan_video_full_frames", fake_scan)
    monkeypatch.setattr(screenspace_scans, "_probe_video_meta", lambda p: (30.0, 10.0))


class TestFlowStaticGate:
    def _run(self, monkeypatch, gate_on):
        # g0/g1 differ maximally: g0->g1 is real motion, repeats are static.
        g0 = np.zeros((60, 60, 3), dtype=np.uint8)
        g1 = np.full((60, 60, 3), 255, dtype=np.uint8)
        seq = [(0.0, g0), (1.0, g0), (2.0, g1), (3.0, g1)]

        calls = [0]

        def counting_flow(prev, curr, return_grid=False):
            calls[0] += 1
            return {
                "magnitude": float(np.mean(cv2.absdiff(prev, curr))),
                "angle": 0.0,
                "flow_grid": [],
            }

        _feed_region(monkeypatch, seq)
        monkeypatch.setattr(screenspace_scans, "compute_optical_flow", counting_flow)
        if not gate_on:
            _disable_gate(monkeypatch)

        results = screenspace.scan_flow(
            "/fake.mp4", _REGION, magnitude_threshold=5.0, interval_seconds=1.0
        )
        return results, calls[0]

    def test_parity_and_fewer_farneback_calls(self, monkeypatch):
        on_results, on_calls = self._run(monkeypatch, gate_on=True)
        off_results, off_calls = self._run(monkeypatch, gate_on=False)

        assert on_results == off_results
        assert len(on_results) == 1  # single motion event at the g0->g1 transition
        assert on_results[0]["timestamp"] == 2.0
        assert on_calls < off_calls
        assert on_calls == 1 and off_calls == 3


class TestInactivityStaticGate:
    def _run(self, monkeypatch, gate_on):
        frame = np.full((60, 60, 3), 128, dtype=np.uint8)
        seq = [(float(i), frame) for i in range(5)]

        calls = [0]
        real_phash = screenspace_primitives.compute_phash

        def counting_phash(px):
            calls[0] += 1
            return real_phash(px)

        _feed_region(monkeypatch, seq)
        monkeypatch.setattr(screenspace_scans, "compute_phash", counting_phash)
        if not gate_on:
            _disable_gate(monkeypatch)

        results = screenspace.scan_inactivity(
            "/fake.mp4", _REGION, threshold=15, min_duration=2.0, interval_seconds=1.0
        )
        return results, calls[0]

    def test_parity_and_fewer_phash_calls(self, monkeypatch):
        on_results, on_calls = self._run(monkeypatch, gate_on=True)
        off_results, off_calls = self._run(monkeypatch, gate_on=False)

        assert on_results == off_results
        assert len(on_results) == 1
        assert on_results[0]["start"] == 0.0 and on_results[0]["end"] == 4.0
        assert on_results[0]["avg_distance"] == 0.0
        assert on_calls < off_calls
        assert on_calls == 1 and off_calls == 5  # gate skips phash on 4 frozen frames


class TestBoundariesStaticGate:
    def _run(self, monkeypatch, gate_on):
        a = np.random.RandomState(1).randint(0, 256, (60, 60, 3)).astype(np.uint8)
        b = np.random.RandomState(2).randint(0, 256, (60, 60, 3)).astype(np.uint8)
        seq = [(0.0, a), (1.0, a.copy()), (2.0, b)]

        calls = [0]
        real_phash = screenspace_primitives.compute_phash

        def counting_phash(px):
            calls[0] += 1
            return real_phash(px)

        _feed_full(monkeypatch, seq)
        monkeypatch.setattr(screenspace_scans, "compute_phash", counting_phash)
        if not gate_on:
            _disable_gate(monkeypatch)

        results = screenspace.scan_boundaries(
            "/fake.mp4",
            _REGION,
            threshold=10,
            min_gap=0.1,
            interval_seconds=1.0,
            metric="phash",
        )
        return results, calls[0]

    def test_parity_and_fewer_phash_calls(self, monkeypatch):
        on_results, on_calls = self._run(monkeypatch, gate_on=True)
        off_results, off_calls = self._run(monkeypatch, gate_on=False)

        assert on_results == off_results
        assert len(on_results) == 1  # one boundary at the a->b cut
        assert on_results[0]["timestamp"] == 2.0
        assert on_calls < off_calls
        assert on_calls == 2 and off_calls == 3  # gate skips phash on the frozen frame


class TestTemplateStaticGate:
    def _run(self, monkeypatch, gate_on):
        matched = _make_icon_frame(400, 200, [(100, 50, 40)])
        blank = np.full((200, 400, 3), 200, dtype=np.uint8)
        seq = [
            (0.0, matched),
            (1.0, matched.copy()),
            (2.0, matched.copy()),
            (3.0, blank),
        ]
        template = _make_icon(40)

        calls = [0]
        real_match = screenspace_primitives._match_template_prepared

        def counting_match(*args, **kwargs):
            calls[0] += 1
            return real_match(*args, **kwargs)

        _feed_full(monkeypatch, seq)
        monkeypatch.setattr(
            screenspace_scans, "_match_template_prepared", counting_match
        )
        if not gate_on:
            _disable_gate(monkeypatch)

        results = screenspace.scan_template(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 400, "h": 200},
            template,
            threshold=0.70,
        )
        return results, calls[0]

    def test_carry_preserves_rows_and_fewer_match_calls(self, monkeypatch):
        on_results, on_calls = self._run(monkeypatch, gate_on=True)
        off_results, off_calls = self._run(monkeypatch, gate_on=False)

        # Carried rows must be identical to freshly-matched rows, so the
        # per-frame detection never flickers across a frozen run.
        assert on_results == off_results
        assert [r["timestamp"] for r in on_results] == [0.0, 1.0, 2.0]
        assert all(r["match_count"] == 1 for r in on_results)
        assert on_calls < off_calls
        assert on_calls == 2 and off_calls == 4  # 2 frozen frames carried, not matched
