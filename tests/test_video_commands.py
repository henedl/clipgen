import subprocess

import video


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
    monkeypatch.setattr(video, "_verify_output_file", lambda *_args, **_kwargs: True)

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

    monkeypatch.setattr(video, "_run_ffmpeg_process", fake_run_ffmpeg_process)

    ok = video.concatenate_clips(["a.mp4", "b.mp4"], "reel.mp4", reencode_on_fail=True)
    assert ok is True
    assert len(captured_commands) == 2
    assert "-c" in captured_commands[0] and "copy" in captured_commands[0]
    assert "-c:v" in captured_commands[1] and "libx264" in captured_commands[1]
    assert "-c:a" in captured_commands[1] and "aac" in captured_commands[1]


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
