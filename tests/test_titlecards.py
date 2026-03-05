import subprocess
from types import SimpleNamespace

import pytest

import titlecards
import video


def test_build_titlecard_frame_uses_png_background_and_scaling(monkeypatch, make_clip):
    clip = make_clip(desc="My description")

    commands = []

    def fake_run_ffmpeg_process(cmd, **_kwargs):
        commands.append(cmd)
        return subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")

    monkeypatch.setattr(video, "_run_ffmpeg_process", fake_run_ffmpeg_process)
    monkeypatch.setattr(video, "_verify_output_file", lambda *_args, **_kwargs: True)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)

    path = titlecards.build_titlecard_frame(clip, "1280x720")
    assert path is not None
    assert commands, "ffmpeg was not invoked for titlecard generation"

    cmd = commands[0]
    joined = " ".join(cmd)
    assert "assets/titlecard.png" in joined
    assert "scale=1280x720:force_original_aspect_ratio=decrease" in joined
    assert "pad=1280:720:(ow-iw)/2:(oh-ih)/2" in joined
    assert "drawtext=text='My description'" in joined


def test_build_titlecard_frame_uses_black_background_when_no_asset(monkeypatch, make_clip):
    clip = make_clip(desc="No asset")

    commands = []

    def fake_run_ffmpeg_process(cmd, **_kwargs):
        commands.append(cmd)
        return subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")

    monkeypatch.setattr(video, "_run_ffmpeg_process", fake_run_ffmpeg_process)
    monkeypatch.setattr(video, "_verify_output_file", lambda *_args, **_kwargs: True)

    def fake_is_file(self):
        return False

    monkeypatch.setattr(titlecards.Path, "is_file", fake_is_file)

    path = titlecards.build_titlecard_frame(clip, "1920x1080")
    assert path is not None
    cmd = commands[0]
    joined = " ".join(cmd)
    assert "assets/titlecard.png" not in joined
    assert "color=c=black:s=1920x1080" in joined
    assert "drawtext=text='No asset'" in joined


def test_prepend_titlecard_to_clip_respects_disabled_flag(monkeypatch, make_clip):
    clip = make_clip()

    calls = {"build": 0, "ffmpeg": 0}

    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", False)
    monkeypatch.setattr(
        titlecards, "build_titlecard_frame", lambda *_args, **_kwargs: calls.__setitem__("build", calls["build"] + 1)
    )
    monkeypatch.setattr(video, "_run_ffmpeg_process", lambda *_args, **_kwargs: calls.__setitem__("ffmpeg", calls["ffmpeg"] + 1))

    ok = titlecards.prepend_titlecard_to_clip(clip, "clip.mp4")
    assert ok is True
    assert calls["build"] == 0
    assert calls["ffmpeg"] == 0


def test_prepend_titlecard_to_clip_uses_filter_concat(monkeypatch, make_clip):
    clip = make_clip()

    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(titlecards, "_get_video_resolution", lambda _path: "1280x720")
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)
    monkeypatch.setattr(titlecards, "build_titlecard_frame", lambda *_args, **_kwargs: "titlecard.mp4")
    monkeypatch.setattr(video, "_verify_output_file", lambda *_args, **_kwargs: True)

    commands = []

    def fake_run_ffmpeg_process(cmd, **_kwargs):
        commands.append(cmd)
        return subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")

    monkeypatch.setattr(video, "_run_ffmpeg_process", fake_run_ffmpeg_process)
    replaced = {}
    monkeypatch.setattr(titlecards.os, "replace", lambda src, dst: replaced.update({"src": src, "dst": dst}))

    ok = titlecards.prepend_titlecard_to_clip(clip, "clip.mp4")
    assert ok is True
    assert commands, "ffmpeg was not invoked for prepend"
    joined = " ".join(commands[0])
    assert "[0:v][1:v]concat=n=2:v=1:a=0[v]" in joined
    assert "-map [v]" in joined.replace(",", " ")
    assert "clip.mp4" in joined
    assert replaced.get("dst") == "clip.mp4"

