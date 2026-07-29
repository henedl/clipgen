"""Tests for timelapse command building and progress."""

import io
from unittest import mock


import screenspace
import screenspace_tools


class TestBuildTimelapseCommand:
    def test_mp4_output(self):
        cmd = screenspace.build_timelapse_command(
            "/video.mp4", {"x": 10, "y": 20, "w": 300, "h": 100}, 10.0, "/out.mp4"
        )
        assert "ffmpeg" in cmd[0]
        assert "/video.mp4" in cmd
        assert "/out.mp4" in cmd
        vf_idx = cmd.index("-vf")
        assert "crop=300:100:10:20" in cmd[vf_idx + 1]
        assert "setpts=PTS/10.0" in cmd[vf_idx + 1]

    def test_gif_output(self):
        cmd = screenspace.build_timelapse_command(
            "/v.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            5.0,
            "/out.gif",
            "gif",
        )
        assert "-loop" in cmd

    def test_mp4_defaults_to_libx264(self):
        cmd = screenspace.build_timelapse_command(
            "/video.mp4", {"x": 0, "y": 0, "w": 100, "h": 100}, 10.0, "/out.mp4"
        )
        assert cmd[cmd.index("-c:v") + 1] == "libx264"
        assert cmd[cmd.index("-preset") + 1] == "fast"
        assert cmd[cmd.index("-crf") + 1] == "23"

    def test_mp4_honors_hardware_encoder(self):
        cmd = screenspace.build_timelapse_command(
            "/video.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            10.0,
            "/out.mp4",
            encoder="h264_videotoolbox",
        )
        assert cmd[cmd.index("-c:v") + 1] == "h264_videotoolbox"
        assert "-q:v" in cmd
        assert "libx264" not in cmd
        assert "-crf" not in cmd


class TestBuildTimelapseCommandMarkers:
    def test_start_seconds_adds_ss_flag(self):
        cmd = screenspace.build_timelapse_command(
            "/video.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            10.0,
            "/out.mp4",
            start_seconds=30.0,
        )
        ss_idx = cmd.index("-ss")
        assert cmd[ss_idx + 1] == "30.0"
        # -ss must appear before -i for fast seeking
        i_idx = cmd.index("-i")
        assert ss_idx < i_idx

    def test_end_seconds_adds_t_flag(self):
        cmd = screenspace.build_timelapse_command(
            "/video.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            10.0,
            "/out.mp4",
            start_seconds=10.0,
            end_seconds=40.0,
        )
        t_idx = cmd.index("-t")
        assert cmd[t_idx + 1] == "30.0"  # 40 - 10

    def test_no_markers_omits_flags(self):
        cmd = screenspace.build_timelapse_command(
            "/video.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            10.0,
            "/out.mp4",
        )
        assert "-ss" not in cmd
        assert "-t" not in cmd


class TestGenerateTimelapseProgress:
    def test_reports_progress_from_ffmpeg_output(self, monkeypatch):
        """on_progress is called with values parsed from ffmpeg -progress."""
        # Simulate ffmpeg -progress output with out_time_us lines
        progress_output = (
            b"out_time_us=0\n"
            b"progress=continue\n"
            b"out_time_us=5000000\n"
            b"progress=continue\n"
            b"out_time_us=10000000\n"
            b"progress=end\n"
        )

        fake_proc = mock.MagicMock()
        fake_proc.stdout = io.BytesIO(progress_output)
        fake_proc.returncode = 0
        fake_proc.wait = mock.MagicMock()

        monkeypatch.setattr("subprocess.Popen", lambda *a, **kw: fake_proc)

        progress_values = []

        result = screenspace.generate_timelapse(
            "/video.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            10.0,
            "/out.mp4",
            start_seconds=0.0,
            end_seconds=100.0,  # 100s input / 10x speedup = 10s output = 10_000_000 us
            on_progress=lambda p: progress_values.append(round(p, 2)),
        )

        assert result == "/out.mp4"
        # Should have: 0.0 (initial), 0.0 (out_time_us=0), 0.5 (5M/10M), 0.99 (capped)
        assert 0.0 in progress_values
        assert any(0.4 <= v <= 0.6 for v in progress_values)  # ~0.5 from 5M/10M
        assert 1.0 in progress_values  # final

    def test_cancel_flag_terminates_process(self, monkeypatch):
        """cancel_flag=True stops ffmpeg and returns None."""
        # Output enough lines so the loop iterates
        progress_output = b"out_time_us=0\nprogress=continue\n" * 5

        fake_proc = mock.MagicMock()
        fake_proc.stdout = io.BytesIO(progress_output)
        fake_proc.returncode = -15  # SIGTERM
        fake_proc.wait = mock.MagicMock()
        fake_proc.terminate = mock.MagicMock()

        monkeypatch.setattr("subprocess.Popen", lambda *a, **kw: fake_proc)

        result = screenspace.generate_timelapse(
            "/video.mp4",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            10.0,
            "/out.mp4",
            start_seconds=0.0,
            end_seconds=100.0,
            cancel_flag=lambda: True,
        )

        assert result is None
        fake_proc.terminate.assert_called_once()


class TestTimelapseDispatchPassesMarkers:
    def test_dispatch_forwards_start_end_to_timelapse(self, monkeypatch):
        """_dispatch passes start_seconds and end_seconds to generate_timelapse."""
        captured = {}

        def fake_generate(
            video_path, region, speedup_factor, output_path, output_format="mp4", **kw
        ):
            captured.update(kw)
            return output_path

        monkeypatch.setattr(screenspace_tools, "generate_timelapse", fake_generate)

        worker = screenspace.ScreenspaceWorker()
        task = screenspace.create_task(
            "timelapse",
            "P01",
            "s_P01.mp4",
            ["/fake.mp4"],
            "region1",
            {"x": 0, "y": 0, "w": 100, "h": 100},
            parameters={
                "speedup_factor": 5.0,
                "start_seconds": 15.0,
                "end_seconds": 45.0,
            },
        )
        worker._dispatch(task, lambda p: None, lambda: False, None)

        assert captured["start_seconds"] == 15.0
        assert captured["end_seconds"] == 45.0
        assert captured["on_progress"] is not None
        assert captured["cancel_flag"] is not None


# ---------------------------------------------------------------------------
# Inactivity tool
# ---------------------------------------------------------------------------
