import json
import subprocess

import video

_MATCHING_PROPS = {
    "width": 1920,
    "height": 1080,
    "video_codec": "h264",
    "audio_codec": "aac",
}


def test_build_ffmpeg_cut_command_includes_expected_flags():
    cmd_copy = video.build_ffmpeg_cut_command(
        input_file="in.mp4",
        output_file="out.mp4",
        start_pos="00:10",
        duration_seconds=5,
        reencode=False,
        audio_normalize=True,
    )
    assert "-c:v" in cmd_copy
    assert "copy" in cmd_copy
    assert "-af" in cmd_copy
    assert "loudnorm=I=-16:TP=-1.5:LRA=11" in cmd_copy

    cmd_reencode = video.build_ffmpeg_cut_command(
        input_file="in.mp4",
        output_file="out.mp4",
        start_pos="00:10",
        duration_seconds=5,
        reencode=True,
        audio_normalize=False,
    )
    assert "-c:v" not in cmd_reencode
    assert "-af" not in cmd_reencode
    assert "out.mp4" in cmd_reencode


def test_concatenate_clips_reencode_fallback(monkeypatch):
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_args, **_kwargs: True)
    monkeypatch.setattr(
        video, "probe_video_properties", lambda _path: dict(_MATCHING_PROPS)
    )

    captured_commands = []
    results = [
        subprocess.CompletedProcess(
            args=["ffmpeg"], returncode=1, stderr="copy failed"
        ),
        subprocess.CompletedProcess(args=["ffmpeg"], returncode=0, stderr=""),
    ]

    def fake_run_ffmpeg_process(command, **_kwargs):
        captured_commands.append(command)
        return results.pop(0)

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run_ffmpeg_process)

    ok = video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4", reencode_on_fail=True)
    assert ok is True
    assert len(captured_commands) == 2
    assert "-c" in captured_commands[0] and "copy" in captured_commands[0]
    assert "-c:v" in captured_commands[1] and "libx264" in captured_commands[1]
    assert "-c:a" in captured_commands[1] and "aac" in captured_commands[1]


# -- probe_video_properties tests --


def test_probe_video_properties_parses_output(monkeypatch):
    fake_json = json.dumps(
        {
            "streams": [
                {
                    "codec_type": "video",
                    "codec_name": "h264",
                    "width": 1920,
                    "height": 1080,
                },
                {"codec_type": "audio", "codec_name": "aac"},
            ]
        }
    )
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video.subprocess, "check_output", lambda _cmd, **_kw: fake_json
    )

    result = video.probe_video_properties("clip.mp4")
    assert result == {
        "width": 1920,
        "height": 1080,
        "video_codec": "h264",
        "audio_codec": "aac",
    }


def test_probe_video_properties_no_audio(monkeypatch):
    fake_json = json.dumps(
        {
            "streams": [
                {
                    "codec_type": "video",
                    "codec_name": "hevc",
                    "width": 1280,
                    "height": 720,
                },
            ]
        }
    )
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video.subprocess, "check_output", lambda _cmd, **_kw: fake_json
    )

    result = video.probe_video_properties("clip.mp4")
    assert result is not None
    assert result["audio_codec"] is None
    assert result["video_codec"] == "hevc"
    assert result["width"] == 1280


def test_probe_video_properties_failure(monkeypatch):
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)

    def raise_cpe(_cmd, **_kw):
        raise subprocess.CalledProcessError(returncode=1, cmd="ffprobe")

    monkeypatch.setattr(video.subprocess, "check_output", raise_cpe)
    assert video.probe_video_properties("clip.mp4") is None


def test_probe_video_properties_file_not_found(monkeypatch):
    monkeypatch.setattr(video.Path, "is_file", lambda self: False)
    assert video.probe_video_properties("missing.mp4") is None


# -- concatenate mismatch detection tests --


def test_concatenate_clips_matching_uses_demuxer(monkeypatch):
    """Identical properties → concat demuxer path (fast, stream copy)."""
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)
    monkeypatch.setattr(
        video, "probe_video_properties", lambda _p: dict(_MATCHING_PROPS)
    )

    captured = []

    def fake_run(command, **_kw):
        captured.append(command)
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4")
    assert ok is True
    assert len(captured) == 1
    assert "-f" in captured[0] and "concat" in captured[0]
    assert "-c" in captured[0] and "copy" in captured[0]
    assert "-filter_complex" not in captured[0]


def test_concatenate_clips_resolution_mismatch_uses_filter_complex(monkeypatch):
    """Different resolutions → filter_complex path with scale+pad."""
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)

    props_by_path = {
        "a.mp4": {"width": 1920, "height": 1080, "video_codec": "h264", "audio_codec": "aac"},
        "b.mp4": {"width": 1280, "height": 720, "video_codec": "h264", "audio_codec": "aac"},
    }
    monkeypatch.setattr(
        video, "probe_video_properties", lambda p: props_by_path.get(p)
    )

    captured = []

    def fake_run(command, **_kw):
        captured.append(command)
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4")
    assert ok is True
    assert len(captured) == 1
    cmd = captured[0]
    assert "-filter_complex" in cmd
    fc_idx = cmd.index("-filter_complex")
    fc_str = cmd[fc_idx + 1]
    assert "scale=1920:1080" in fc_str
    assert "pad=1920:1080" in fc_str
    assert "concat=n=2:v=1:a=1" in fc_str
    assert "-f" not in cmd  # no concat demuxer


def test_concatenate_clips_warns_on_mismatch(monkeypatch):
    """Mismatch detection prints warnings."""
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)

    props_by_path = {
        "a.mp4": {"width": 1920, "height": 1080, "video_codec": "h264", "audio_codec": "aac"},
        "b.mp4": {"width": 1280, "height": 720, "video_codec": "hevc", "audio_codec": "aac"},
    }
    monkeypatch.setattr(
        video, "probe_video_properties", lambda p: props_by_path.get(p)
    )
    monkeypatch.setattr(
        video, "run_ffmpeg_process",
        lambda cmd, **_kw: subprocess.CompletedProcess(args=cmd, returncode=0, stderr=""),
    )

    warnings = []
    original_warning = video.utils.warning_print

    def capture_warning(msg, details=None):
        warnings.append(msg)
        original_warning(msg, details)

    monkeypatch.setattr(video.utils, "warning_print", capture_warning)

    video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4")
    warning_text = " ".join(warnings)
    assert "Resolution mismatch" in warning_text
    assert "1920x1080" in warning_text
    assert "1280x720" in warning_text
    assert "Video codec mismatch" in warning_text


def test_concatenate_clips_mixed_audio_presence(monkeypatch):
    """One clip with audio, one without → anullsrc for the silent clip."""
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)
    monkeypatch.setattr(video, "get_file_duration", lambda _p: 10)

    props_by_path = {
        "a.mp4": {"width": 1920, "height": 1080, "video_codec": "h264", "audio_codec": "aac"},
        "b.mp4": {"width": 1920, "height": 1080, "video_codec": "h264", "audio_codec": None},
    }
    monkeypatch.setattr(
        video, "probe_video_properties", lambda p: props_by_path.get(p)
    )

    captured = []

    def fake_run(command, **_kw):
        captured.append(command)
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4")
    assert ok is True
    cmd = captured[0]
    assert "-filter_complex" in cmd
    fc_str = cmd[cmd.index("-filter_complex") + 1]
    assert "anullsrc" in fc_str
    assert "concat=n=2:v=1:a=1" in fc_str


def test_concatenate_clips_all_no_audio(monkeypatch):
    """All clips lack audio → concat with a=0, no audio mapping."""
    monkeypatch.setattr(video.Path, "is_file", lambda self: True)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_kw: True)

    no_audio_props = {
        "width": 1280,
        "height": 720,
        "video_codec": "h264",
        "audio_codec": None,
    }
    monkeypatch.setattr(
        video, "probe_video_properties", lambda _p: dict(no_audio_props)
    )

    # Same resolution, but audio mismatch is about presence—need at least one
    # mismatch to trigger filter_complex. Use resolution mismatch instead.
    call_count = [0]
    def props_alternating(_p):
        call_count[0] += 1
        if call_count[0] % 2 == 1:
            return {"width": 1920, "height": 1080, "video_codec": "h264", "audio_codec": None}
        return {"width": 1280, "height": 720, "video_codec": "h264", "audio_codec": None}

    monkeypatch.setattr(video, "probe_video_properties", props_alternating)

    captured = []

    def fake_run(command, **_kw):
        captured.append(command)
        return subprocess.CompletedProcess(args=command, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run)

    ok = video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4")
    assert ok is True
    cmd = captured[0]
    fc_str = cmd[cmd.index("-filter_complex") + 1]
    assert "concat=n=2:v=1:a=0" in fc_str
    assert "[outa]" not in fc_str
    assert '"-map", "[outa]"' not in str(cmd)


def test_get_duration_valid_and_invalid_values(monkeypatch):
    # Valid MM:SS
    assert video.get_duration("00:10", "00:20") == 10
    # Valid HH:MM:SS
    assert video.get_duration("01:00:00", "01:10:00") == 600
    # INVALID_END_TIMESTAMP sentinel returns None
    assert video.get_duration("00:10", video.INVALID_END_TIMESTAMP) is None
    # Completely invalid format returns None
    assert video.get_duration("not-a-time", "also-bad") is None


def test_calculate_target_bitrate_typical_and_min_floor():
    kbps = video.calculate_target_bitrate(target_size_mb=50, duration_seconds=600)
    assert kbps > 100
    small = video.calculate_target_bitrate(target_size_mb=1, duration_seconds=5)
    assert small >= 100
    zero_duration = video.calculate_target_bitrate(
        target_size_mb=10, duration_seconds=0
    )
    assert zero_duration == 100


def test_get_file_duration_error_paths(monkeypatch):
    # Missing file
    monkeypatch.setattr(video.os.path, "isfile", lambda _path: False)
    assert video.get_file_duration("missing.mp4") is None

    # ffprobe not found
    monkeypatch.setattr(video.os.path, "isfile", lambda _path: True)

    def raise_fnf(_cmd, **_kwargs):
        raise FileNotFoundError()

    monkeypatch.setattr(video.subprocess, "check_output", raise_fnf)
    assert video.get_file_duration("video.mp4") is None

    # ffprobe CalledProcessError
    def raise_cpe(_cmd, **_kwargs):
        raise video.subprocess.CalledProcessError(returncode=1, cmd="ffprobe")

    monkeypatch.setattr(video.subprocess, "check_output", raise_cpe)
    assert video.get_file_duration("video.mp4") is None


def test_check_ffmpeg_tools_available_missing(monkeypatch):
    monkeypatch.setattr(video.shutil, "which", lambda _tool: None)
    ok = video.check_ffmpeg_tools_available()
    assert ok is False
