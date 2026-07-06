"""Tests for scan dispatch, ffmpeg pipe extraction, and preview alignment."""

import io
from unittest import mock

import numpy as np
import pytest

import config
import screenspace
import screenspace_tools


class TestFastScanDispatchIntervalMultiplier:
    """Verify _dispatch() applies interval multiplier in fast scan mode."""

    def test_interval_multiplied_for_fast_scan(self, monkeypatch):
        captured = {}

        def fake_scan_color(
            video_path,
            region,
            *,
            target_color,
            tolerance,
            interval_seconds=0,
            start_seconds=0.0,
            end_seconds=None,
            color_mode="average",
            min_coverage=0.0,
            on_progress=None,
            cancel_flag=None,
            on_result=None,
            fast_opts=None,
        ):
            captured["interval"] = interval_seconds
            captured["fast_opts"] = fast_opts
            return []

        monkeypatch.setattr(screenspace_tools, "scan_color", fake_scan_color)

        worker = screenspace.ScreenspaceWorker()
        task = {
            "id": "ss_test1",
            "type": "color",
            "video_paths": ["/fake.mp4"],
            "region_coords": {"x": 0, "y": 0, "w": 100, "h": 100},
            "parameters": {
                "scan_mode": "fast",
                "interval": 1.0,
                "target_color": {"h": 0, "s": 0, "v": 0},
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
        }
        worker._dispatch(task, lambda p: None, lambda: False, None)

        expected = 1.0 * config.SCREENSPACE_FAST_SCAN_INTERVAL_MULTIPLIER
        assert captured["interval"] == expected
        assert captured["fast_opts"] is not None
        assert captured["fast_opts"]["phash_skip"] is True
        assert captured["fast_opts"]["max_region_dim"] == 32

    def test_fast_scan_interval_not_persisted_on_task(self, monkeypatch):
        """Pause/resume re-dispatches the same task; interval must not compound."""
        captured_intervals: list[float] = []

        def fake_scan_color(
            video_path,
            region,
            *,
            target_color,
            tolerance,
            interval_seconds=0,
            start_seconds=0.0,
            end_seconds=None,
            color_mode="average",
            min_coverage=0.0,
            on_progress=None,
            cancel_flag=None,
            on_result=None,
            fast_opts=None,
        ):
            captured_intervals.append(interval_seconds)
            return []

        monkeypatch.setattr(screenspace_tools, "scan_color", fake_scan_color)

        worker = screenspace.ScreenspaceWorker()
        task = {
            "id": "ss_resume",
            "type": "color",
            "video_paths": ["/fake.mp4"],
            "region_coords": {"x": 0, "y": 0, "w": 100, "h": 100},
            "parameters": {
                "scan_mode": "fast",
                "interval": 1.0,
                "target_color": {"h": 0, "s": 0, "v": 0},
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
        }

        def noop(_progress: float) -> None:
            return None

        worker._dispatch(task, noop, lambda: False, None)
        worker._dispatch(task, noop, lambda: False, None)

        expected = 1.0 * config.SCREENSPACE_FAST_SCAN_INTERVAL_MULTIPLIER
        assert task["parameters"]["interval"] == 1.0
        assert captured_intervals == [expected, expected]

    def test_normal_scan_no_fast_opts(self, monkeypatch):
        captured = {}

        def fake_scan_color(
            video_path,
            region,
            *,
            target_color,
            tolerance,
            interval_seconds=0,
            start_seconds=0.0,
            end_seconds=None,
            color_mode="average",
            min_coverage=0.0,
            on_progress=None,
            cancel_flag=None,
            on_result=None,
            fast_opts=None,
        ):
            captured["interval"] = interval_seconds
            captured["fast_opts"] = fast_opts
            return []

        monkeypatch.setattr(screenspace_tools, "scan_color", fake_scan_color)

        worker = screenspace.ScreenspaceWorker()
        task = {
            "id": "ss_test2",
            "type": "color",
            "video_paths": ["/fake.mp4"],
            "region_coords": {"x": 0, "y": 0, "w": 100, "h": 100},
            "parameters": {
                "interval": 1.0,
                "target_color": {"h": 0, "s": 0, "v": 0},
                "tolerance": {"h": 10, "s": 50, "v": 50},
            },
        }
        worker._dispatch(task, lambda p: None, lambda: False, None)

        assert captured["interval"] == 1.0
        assert captured["fast_opts"] is None

    def test_presence_color_drops_fast_opts(self, monkeypatch):
        """Presence mode must scan full-res: ColorTool.scan nulls fast_opts so
        the max_region_dim INTER_AREA downscale can't erase small patches."""
        captured = {}

        def fake_scan_color(
            video_path,
            region,
            *,
            target_color,
            tolerance,
            interval_seconds=0,
            start_seconds=0.0,
            end_seconds=None,
            color_mode="average",
            min_coverage=0.0,
            on_progress=None,
            cancel_flag=None,
            on_result=None,
            fast_opts=None,
        ):
            captured["color_mode"] = color_mode
            captured["min_coverage"] = min_coverage
            captured["fast_opts"] = fast_opts
            return []

        monkeypatch.setattr(screenspace_tools, "scan_color", fake_scan_color)

        worker = screenspace.ScreenspaceWorker()
        task = {
            "id": "ss_presence",
            "type": "color",
            "video_paths": ["/fake.mp4"],
            "region_coords": {"x": 0, "y": 0, "w": 100, "h": 100},
            "parameters": {
                "scan_mode": "fast",
                "interval": 1.0,
                "target_color": {"h": 0, "s": 255, "v": 139},
                "tolerance": {"h": 10, "s": 60, "v": 60},
                "color_mode": "presence",
                "min_coverage": 0.02,
            },
        }
        worker._dispatch(task, lambda p: None, lambda: False, None)

        assert captured["color_mode"] == "presence"
        assert captured["min_coverage"] == 0.02
        assert captured["fast_opts"] is None  # downscale skipped for presence

    def test_template_dispatch_gets_downscale_flag(self, monkeypatch):
        captured = {}

        def fake_scan_template(
            video_path,
            region,
            *,
            template_image,
            threshold=0,
            interval_seconds=0,
            template_mask=None,
            template_scale=1.0,
            start_seconds=0.0,
            end_seconds=None,
            on_progress=None,
            cancel_flag=None,
            on_result=None,
            fast_opts=None,
        ):
            captured["fast_opts"] = fast_opts
            captured["template_shape"] = template_image.shape
            return []

        monkeypatch.setattr(screenspace_tools, "scan_template", fake_scan_template)

        worker = screenspace.ScreenspaceWorker()
        tmpl = np.zeros((100, 200, 3), dtype=np.uint8)
        task = {
            "id": "ss_test3",
            "type": "template",
            "video_paths": ["/fake.mp4"],
            "region_coords": {"x": 0, "y": 0, "w": 100, "h": 100},
            "parameters": {
                "scan_mode": "fast",
                "interval": 1.0,
                "template_image": tmpl,
            },
        }
        worker._dispatch(task, lambda p: None, lambda: False, None)

        assert captured["fast_opts"]["template_downscale"] is True
        # Template should be downscaled by 2x in _dispatch
        assert captured["template_shape"] == (50, 100, 3)


# ---------------------------------------------------------------------------
# 2E: ffmpeg pipe extraction
# ---------------------------------------------------------------------------


class TestFfmpegPipeFrames:
    def test_yields_frames_with_pts_from_stderr(self, monkeypatch):
        """Generator pairs each raw BGR frame with the pts_time read from stderr."""
        w, h = 4, 2
        frame1 = np.full((h, w, 3), 100, dtype=np.uint8)
        frame2 = np.full((h, w, 3), 200, dtype=np.uint8)
        raw = frame1.tobytes() + frame2.tobytes()
        # Two showinfo lines with PTS 0 and 1 (relative to the seek point of 5.0).
        stderr_lines = (
            b"[Parsed_showinfo_1 @ 0x0] n: 0 pts: 0 pts_time:0 ...\n"
            b"[Parsed_showinfo_1 @ 0x0] n: 1 pts: 1 pts_time:1 ...\n"
        )

        fake_proc = mock.MagicMock()
        fake_proc.stdout = io.BytesIO(raw)
        fake_proc.stderr = io.BytesIO(stderr_lines)
        fake_proc.terminate = mock.MagicMock()
        fake_proc.wait = mock.MagicMock()

        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("subprocess.Popen", lambda *a, **kw: fake_proc)

        frames = list(
            screenspace._ffmpeg_pipe_frames(
                "/fake.mp4",
                1.0,
                start_seconds=5.0,
                end_seconds=10.0,
                frame_width=w,
                frame_height=h,
            )
        )
        assert len(frames) == 2
        # Yielded ts = start_seconds + pts_time (relative).
        assert frames[0][0] == 5.0
        assert frames[1][0] == 6.0
        assert frames[0][1].shape == (h, w, 3)
        assert np.all(frames[0][1] == 100)
        assert np.all(frames[1][1] == 200)

    def test_empty_when_no_ffmpeg(self, monkeypatch):
        monkeypatch.setattr("shutil.which", lambda _: None)
        frames = list(
            screenspace._ffmpeg_pipe_frames(
                "/fake.mp4", 1.0, frame_width=10, frame_height=10
            )
        )
        assert frames == []

    def test_uses_single_preinput_seek_with_select_filter(self, monkeypatch):
        """Analysis pipe uses a single pre-input -ss plus the select filter +
        fps_mode vfr. The previous two-stage seek silently dropped its post-input
        -ss when paired with -vf in modern ffmpeg, and the fps filter chose a
        different source frame than the preview's accurate seek did."""
        captured: dict = {}

        def fake_popen(cmd, *a, **kw):
            captured["cmd"] = list(cmd)
            proc = mock.MagicMock()
            proc.stdout = io.BytesIO(b"")
            proc.stderr = io.BytesIO(b"")
            proc.terminate = mock.MagicMock()
            proc.wait = mock.MagicMock()
            return proc

        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("subprocess.Popen", fake_popen)

        list(
            screenspace._ffmpeg_pipe_frames(
                "/fake.mp4",
                1.0,
                start_seconds=15.0,
                end_seconds=20.0,
                frame_width=4,
                frame_height=2,
            )
        )
        cmd = captured["cmd"]
        i_idx = cmd.index("-i")
        # One pre-input -ss, nothing after -i.
        assert cmd[:i_idx].count("-ss") == 1
        assert "-ss" not in cmd[i_idx + 2 :]
        vf = cmd[cmd.index("-vf") + 1]
        assert "select=" in vf
        assert "showinfo" in vf
        # vfr is what stops ffmpeg from duplicating the kept frame to fill the
        # source frame rate.
        assert "-fps_mode" in cmd and cmd[cmd.index("-fps_mode") + 1] == "vfr"


class TestScanViaFfmpegPipe:
    def test_returns_false_when_no_ffmpeg(self, monkeypatch):
        monkeypatch.setattr("shutil.which", lambda _: None)
        result = screenspace._scan_via_ffmpeg_pipe(
            "/fake.mp4", None, 1.0, lambda ts, f: None, duration=10.0
        )
        assert result is False

    def test_returns_false_when_probe_fails(self, monkeypatch):
        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("video.probe_video_properties", lambda _: None)
        result = screenspace._scan_via_ffmpeg_pipe(
            "/fake.mp4", None, 1.0, lambda ts, f: None, duration=10.0
        )
        assert result is False

    def test_calls_callback_with_frames(self, monkeypatch):
        w, h = 4, 2
        frame_data = np.full((h, w, 3), 42, dtype=np.uint8)
        raw = frame_data.tobytes()
        # showinfo emits one line per yielded frame; pts_time is relative to
        # the seek point (here 0).
        stderr_lines = b"[Parsed_showinfo_1 @ 0x0] n: 0 pts: 0 pts_time:0 ...\n"

        fake_proc = mock.MagicMock()
        fake_proc.stdout = io.BytesIO(raw)
        fake_proc.stderr = io.BytesIO(stderr_lines)
        fake_proc.terminate = mock.MagicMock()
        fake_proc.wait = mock.MagicMock()

        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("subprocess.Popen", lambda *a, **kw: fake_proc)
        monkeypatch.setattr(
            "video.probe_video_properties",
            lambda _: {"width": w, "height": h, "video_codec": "h264"},
        )

        received = []

        def cb(ts, frame):
            received.append((ts, frame.copy()))

        result = screenspace._scan_via_ffmpeg_pipe(
            "/fake.mp4", None, 1.0, cb, duration=10.0, full_frame=True
        )
        assert result is True
        assert len(received) == 1
        assert received[0][0] == 0.0
        assert np.all(received[0][1] == 42)

    def test_builds_crop_filter_for_region(self, monkeypatch):
        w, h = 8, 6
        region = {"x": 1, "y": 2, "w": 4, "h": 3}

        captured_cmd = {}

        def fake_popen(cmd, **kw):
            captured_cmd["args"] = cmd
            proc = mock.MagicMock()
            # Return empty bytes so the generator exits immediately
            proc.stdout = io.BytesIO(b"")
            proc.terminate = mock.MagicMock()
            proc.wait = mock.MagicMock()
            return proc

        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("subprocess.Popen", fake_popen)
        monkeypatch.setattr(
            "video.probe_video_properties",
            lambda _: {"width": w, "height": h, "video_codec": "h264"},
        )

        screenspace._scan_via_ffmpeg_pipe(
            "/fake.mp4", region, 1.0, lambda ts, f: None, duration=10.0
        )

        cmd_str = " ".join(captured_cmd["args"])
        assert "crop=4:3:1:2" in cmd_str

    def test_stops_on_callback_false(self, monkeypatch):
        w, h = 4, 2
        frame = np.full((h, w, 3), 1, dtype=np.uint8)
        raw = frame.tobytes() * 5  # 5 frames
        stderr_lines = b"".join(
            b"[Parsed_showinfo_1 @ 0x0] n: %d pts_time:%d ...\n" % (i, i)
            for i in range(5)
        )

        fake_proc = mock.MagicMock()
        fake_proc.stdout = io.BytesIO(raw)
        fake_proc.stderr = io.BytesIO(stderr_lines)
        fake_proc.terminate = mock.MagicMock()
        fake_proc.wait = mock.MagicMock()

        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("subprocess.Popen", lambda *a, **kw: fake_proc)
        monkeypatch.setattr(
            "video.probe_video_properties",
            lambda _: {"width": w, "height": h, "video_codec": "h264"},
        )

        call_count = [0]

        def cb(ts, frame):
            call_count[0] += 1
            if call_count[0] >= 2:
                return False

        result = screenspace._scan_via_ffmpeg_pipe(
            "/fake.mp4", None, 1.0, cb, duration=10.0, full_frame=True
        )
        assert result is True
        assert call_count[0] == 2


class TestKeyframeSkipGating:
    """`-skip_frame nokey` is emitted only in fast-scan mode when the probed
    keyframe interval confirms short-enough GOP; the precise path never gets it."""

    def _capture_cmd(self, monkeypatch):
        captured: dict = {}

        def fake_popen(cmd, *a, **kw):
            captured["cmd"] = list(cmd)
            proc = mock.MagicMock()
            proc.stdout = io.BytesIO(b"")  # empty → generator exits immediately
            proc.stderr = io.BytesIO(b"")
            proc.terminate = mock.MagicMock()
            proc.wait = mock.MagicMock()
            return proc

        monkeypatch.setattr("shutil.which", lambda _: "/usr/bin/ffmpeg")
        monkeypatch.setattr("subprocess.Popen", fake_popen)
        return captured

    def test_skip_frame_present_when_enabled(self, monkeypatch):
        captured = self._capture_cmd(monkeypatch)
        list(
            screenspace._ffmpeg_pipe_frames(
                "/fake.mp4",
                1.0,
                start_seconds=15.0,
                end_seconds=20.0,
                frame_width=4,
                frame_height=2,
                skip_non_keyframes=True,
            )
        )
        cmd = captured["cmd"]
        assert "-skip_frame" in cmd
        sf_idx = cmd.index("-skip_frame")
        assert cmd[sf_idx + 1] == "nokey"
        # Input-level option: after any pre-input -ss, before -i.
        assert sf_idx < cmd.index("-i")
        assert sf_idx > cmd.index("-ss")
        # Additive: the accurate-PTS machinery is unchanged.
        vf = cmd[cmd.index("-vf") + 1]
        assert "select=" in vf and "showinfo" in vf
        assert cmd[cmd.index("-fps_mode") + 1] == "vfr"

    def test_skip_frame_absent_by_default(self, monkeypatch):
        captured = self._capture_cmd(monkeypatch)
        list(
            screenspace._ffmpeg_pipe_frames(
                "/fake.mp4", 1.0, frame_width=4, frame_height=2
            )
        )
        assert "-skip_frame" not in captured["cmd"]

    def _run_scan(self, monkeypatch, *, codec, kf_gap, fast_opts, interval=3.0):
        captured = self._capture_cmd(monkeypatch)
        monkeypatch.setattr(
            "video.probe_video_properties",
            lambda _: {"width": 4, "height": 2, "video_codec": codec},
        )
        monkeypatch.setattr("video.probe_max_keyframe_gap", lambda _: kf_gap)
        monkeypatch.setattr(config, "SCREENSPACE_FAST_SCAN_SKIP_NONKEY", True)
        screenspace._scan_via_ffmpeg_pipe(
            "/fake.mp4",
            None,
            interval,
            lambda ts, f: None,
            duration=30.0,
            full_frame=True,
            fast_opts=fast_opts,
        )
        return captured["cmd"]

    def test_enables_skip_when_gop_short(self, monkeypatch):
        cmd = self._run_scan(
            monkeypatch, codec="h264", kf_gap=1.0, fast_opts={"phash_skip": True}
        )
        assert "-skip_frame" in cmd and cmd[cmd.index("-skip_frame") + 1] == "nokey"

    def test_select_interval_tightened_by_keyframe_gap(self, monkeypatch):
        # #1 fix: with a 1s worst-case gap and a 3s interval, the select grid is
        # tightened to 3-1=2s so keyframe snapping never overshoots past 3s.
        cmd = self._run_scan(
            monkeypatch,
            codec="h264",
            kf_gap=1.0,
            interval=3.0,
            fast_opts={"phash_skip": True},
        )
        vf = cmd[cmd.index("-vf") + 1]
        assert "gte(t-prev_selected_t,2.0)" in vf
        assert "gte(t-prev_selected_t,3.0)" not in vf

    def test_disables_skip_when_gop_long(self, monkeypatch):
        cmd = self._run_scan(
            monkeypatch, codec="h264", kf_gap=10.0, fast_opts={"phash_skip": True}
        )
        assert "-skip_frame" not in cmd
        # Grid unchanged when skip is off: select still uses the full interval.
        vf = cmd[cmd.index("-vf") + 1]
        assert "gte(t-prev_selected_t,3.0)" in vf

    def test_disables_skip_when_probe_none(self, monkeypatch):
        cmd = self._run_scan(
            monkeypatch, codec="h264", kf_gap=None, fast_opts={"phash_skip": True}
        )
        assert "-skip_frame" not in cmd

    def test_no_skip_without_fast_opts(self, monkeypatch):
        # Precise path: fast_opts is None, so the probe is never consulted and
        # the flag never appears regardless of GOP.
        probe_calls = []
        captured = self._capture_cmd(monkeypatch)
        monkeypatch.setattr(
            "video.probe_video_properties",
            lambda _: {"width": 4, "height": 2, "video_codec": "h264"},
        )

        def _probe(_):
            probe_calls.append(1)
            return 1.0

        monkeypatch.setattr("video.probe_max_keyframe_gap", _probe)
        monkeypatch.setattr(config, "SCREENSPACE_FAST_SCAN_SKIP_NONKEY", True)
        screenspace._scan_via_ffmpeg_pipe(
            "/fake.mp4", None, 3.0, lambda ts, f: None, duration=30.0, full_frame=True
        )
        assert "-skip_frame" not in captured["cmd"]
        assert probe_calls == []

    def test_no_skip_for_non_h264_codec(self, monkeypatch):
        cmd = self._run_scan(
            monkeypatch, codec="vp9", kf_gap=1.0, fast_opts={"phash_skip": True}
        )
        assert "-skip_frame" not in cmd


class TestAnalysisPreviewAlignment:
    """End-to-end: each ts yielded by `_ffmpeg_pipe_frames` must point at the
    exact same source frame that `video.extract_frame_at_timestamp(ts)` returns,
    so clicking a result in Screenspace shows the analysed frame instead of a
    drifted neighbour.

    Requires a real ffmpeg binary on PATH; skipped otherwise.
    """

    @staticmethod
    def _have_ffmpeg() -> bool:
        import shutil as _shutil

        return (
            _shutil.which("ffmpeg") is not None and _shutil.which("ffprobe") is not None
        )

    @staticmethod
    def _synthesize(path: str, vf_extra: str = "") -> None:
        import subprocess as _sp

        vf = "testsrc=duration=12:size=320x240:rate=30"
        if vf_extra:
            vf += f",{vf_extra}"
        _sp.run(
            [
                "ffmpeg",
                "-y",
                "-f",
                "lavfi",
                "-i",
                vf,
                "-fps_mode",
                "vfr",
                "-c:v",
                "libx264",
                "-g",
                "30",
                "-pix_fmt",
                "yuv420p",
                path,
            ],
            check=True,
            stdout=_sp.DEVNULL,
            stderr=_sp.DEVNULL,
        )

    @staticmethod
    def _assert_aligned(video_path: str) -> None:
        import hashlib

        import video as video_mod

        frames = list(
            screenspace._ffmpeg_pipe_frames(
                video_path,
                interval_seconds=1.0,
                start_seconds=0.0,
                end_seconds=10.0,
                frame_width=320,
                frame_height=240,
            )
        )
        assert frames, "expected at least one analysed frame"
        for ts, analysis_frame in frames:
            preview = video_mod.extract_frame_at_timestamp(video_path, ts)
            assert preview is not None, f"preview extraction failed at ts={ts}"
            h_a = hashlib.md5(analysis_frame.tobytes()).hexdigest()
            h_p = hashlib.md5(preview.tobytes()).hexdigest()
            assert h_a == h_p, (
                f"analysed frame and preview differ at ts={ts:.4f} "
                f"(analysis={h_a[:8]} preview={h_p[:8]})"
            )

    def test_cfr_source_alignment(self, tmp_path):
        if not self._have_ffmpeg():
            pytest.skip("ffmpeg/ffprobe required for end-to-end alignment test")
        video_path = str(tmp_path / "cfr.mp4")
        self._synthesize(video_path)
        self._assert_aligned(video_path)

    def test_vfr_source_alignment(self, tmp_path):
        if not self._have_ffmpeg():
            pytest.skip("ffmpeg/ffprobe required for end-to-end alignment test")
        video_path = str(tmp_path / "vfr.mp4")
        # Drop ~2/7 of frames at non-uniform positions so the kept frames don't
        # land on integer second boundaries — the case where the old fps filter
        # picked a different source frame than the preview's accurate seek.
        self._synthesize(
            video_path,
            vf_extra="select='not(eq(mod(n,7),3))*not(eq(mod(n,11),5))'",
        )
        self._assert_aligned(video_path)


class TestScanVideoFramesFfmpegIntegration:
    def test_ffmpeg_pipe_succeeds(self, monkeypatch):
        """When ffmpeg pipe succeeds, scan completes without error."""
        monkeypatch.setattr(
            screenspace,
            "_scan_via_ffmpeg_pipe",
            lambda *a, **kw: True,
        )

        screenspace.scan_video_frames(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            1.0,
            lambda ts, f: None,
        )

    def test_warns_when_ffmpeg_pipe_fails(self, monkeypatch):
        """When ffmpeg pipe fails, a warning is emitted."""
        monkeypatch.setattr(
            screenspace,
            "_scan_via_ffmpeg_pipe",
            lambda *a, **kw: False,
        )
        warnings = []
        monkeypatch.setattr(
            screenspace.utils,
            "warning_print",
            lambda msg, *a, **kw: warnings.append(msg),
        )

        screenspace.scan_video_frames(
            "/fake.mp4",
            {"x": 0, "y": 0, "w": 10, "h": 10},
            1.0,
            lambda ts, f: None,
        )
        assert len(warnings) == 1
        assert "ffmpeg" in warnings[0].lower()


class TestFacadeReExports:
    """The screenspace facade must keep exposing former god-file seams."""

    def test_probe_video_meta_reexported(self):
        import screenspace_frames

        assert hasattr(screenspace, "_probe_video_meta")
        assert screenspace._probe_video_meta is screenspace_frames._probe_video_meta


# ---------------------------------------------------------------------------
# 2F: Parallel worker
# ---------------------------------------------------------------------------
