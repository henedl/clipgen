import subprocess


import config
import titlecards
import utils
import video


def test_build_titlecard_frame_uses_png_background_and_scaling(monkeypatch, make_clip):
    clip = make_clip(desc="My description")

    commands = []

    def fake_run_ffmpeg_process(cmd, **_kwargs):
        commands.append(cmd)
        return subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run_ffmpeg_process)
    monkeypatch.setattr(video, "verify_output_file", lambda *_args, **_kwargs: True)
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


def test_build_titlecard_frame_uses_black_background_when_no_asset(
    monkeypatch, make_clip
):
    clip = make_clip(desc="No asset")

    commands = []

    def fake_run_ffmpeg_process(cmd, **_kwargs):
        commands.append(cmd)
        return subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run_ffmpeg_process)
    monkeypatch.setattr(video, "verify_output_file", lambda *_args, **_kwargs: True)

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


def test_wrap_clip_with_cards_respects_disabled_flag(monkeypatch, make_clip):
    clip = make_clip()

    calls = {"build_title": 0, "build_end": 0, "ffmpeg": 0}

    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", False)
    monkeypatch.setattr(
        titlecards,
        "build_titlecard_frame",
        lambda *_a, **_k: calls.__setitem__("build_title", calls["build_title"] + 1),
    )
    monkeypatch.setattr(
        titlecards,
        "get_or_build_endcard",
        lambda *_a, **_k: calls.__setitem__("build_end", calls["build_end"] + 1),
    )
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda *_a, **_k: calls.__setitem__("ffmpeg", calls["ffmpeg"] + 1),
    )

    ok = titlecards.wrap_clip_with_cards(clip, "clip.mp4")
    assert ok is True
    assert calls == {"build_title": 0, "build_end": 0, "ffmpeg": 0}


def test_wrap_clip_with_cards_single_encode_happy_path(monkeypatch, make_clip):
    clip = make_clip()

    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _p: {
            "width": 1280,
            "height": 720,
            "video_codec": "h264",
            "audio_codec": "aac",
            "fps": 30.0,
            "duration": 12.0,
            "nb_frames": 360,
        },
    )
    monkeypatch.setattr(
        titlecards, "build_titlecard_frame", lambda *_a, **_k: "titlecard.mp4"
    )
    monkeypatch.setattr(
        titlecards, "get_or_build_endcard", lambda *_a, **_k: "endcard.mp4"
    )
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_k: True)

    commands = []

    def fake_run_ffmpeg_process(cmd, **_kwargs):
        commands.append(cmd)
        return subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")

    monkeypatch.setattr(video, "run_ffmpeg_process", fake_run_ffmpeg_process)
    replaced = {}
    monkeypatch.setattr(
        titlecards.os,
        "replace",
        lambda src, dst: replaced.update({"src": src, "dst": dst}),
    )

    ok = titlecards.wrap_clip_with_cards(clip, "clip.mp4")
    assert ok is True
    assert len(commands) == 1, "expected exactly one ffmpeg invocation"

    cmd = commands[0]
    joined = " ".join(cmd)
    # All three video inputs (title, clip, end) should be present.
    assert cmd.count("-i") == 5  # 3 video inputs + 2 silent anullsrc audio inputs
    assert "titlecard.mp4" in cmd
    assert "endcard.mp4" in cmd
    assert "clip.mp4" in cmd
    # Single concat covers all three segments with audio.
    assert "concat=n=3:v=1:a=1" in joined
    # Clip audio is normalized before concat.
    assert "aresample=48000" in joined
    # Silent audio pads the two cards.
    assert joined.count("anullsrc=channel_layout=stereo:sample_rate=48000") == 2
    # Output replaces the original clip.
    assert replaced.get("dst") == "clip.mp4"


def test_wrap_clip_with_cards_title_only_fallback(monkeypatch, make_clip):
    clip = make_clip()

    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _p: {
            "width": 1920,
            "height": 1080,
            "video_codec": "h264",
            "audio_codec": "aac",
            "fps": 30.0,
            "duration": 8.0,
            "nb_frames": 240,
        },
    )
    monkeypatch.setattr(
        titlecards, "build_titlecard_frame", lambda *_a, **_k: "titlecard.mp4"
    )
    # Endcard build fails.
    monkeypatch.setattr(titlecards, "get_or_build_endcard", lambda *_a, **_k: None)
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_k: True)

    commands = []
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda cmd, **_k: (
            commands.append(cmd)
            or subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")
        ),
    )
    monkeypatch.setattr(titlecards.os, "replace", lambda src, dst: None)

    ok = titlecards.wrap_clip_with_cards(clip, "clip.mp4")
    assert ok is True
    assert len(commands) == 1
    joined = " ".join(commands[0])
    assert "concat=n=2:v=1:a=1" in joined
    assert "titlecard.mp4" in commands[0]
    assert "endcard.mp4" not in commands[0]


def test_wrap_clip_with_cards_no_audio_drops_audio_stream(monkeypatch, make_clip):
    clip = make_clip()

    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _p: {
            "width": 1280,
            "height": 720,
            "video_codec": "h264",
            "audio_codec": None,
            "fps": 30.0,
            "duration": 4.0,
            "nb_frames": 120,
        },
    )
    monkeypatch.setattr(
        titlecards, "build_titlecard_frame", lambda *_a, **_k: "titlecard.mp4"
    )
    monkeypatch.setattr(
        titlecards, "get_or_build_endcard", lambda *_a, **_k: "endcard.mp4"
    )
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_k: True)

    commands = []
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda cmd, **_k: (
            commands.append(cmd)
            or subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")
        ),
    )
    monkeypatch.setattr(titlecards.os, "replace", lambda src, dst: None)

    ok = titlecards.wrap_clip_with_cards(clip, "clip.mp4")
    assert ok is True
    cmd = commands[0]
    joined = " ".join(cmd)
    # No audio stream in output: concat a=0 and no anullsrc inputs.
    assert "concat=n=3:v=1:a=0" in joined
    assert "anullsrc" not in joined
    assert "-c:a" not in cmd


def test_wrap_clip_with_cards_no_cards_is_noop(monkeypatch, make_clip):
    clip = make_clip()

    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _p: {
            "width": 1280,
            "height": 720,
            "video_codec": "h264",
            "audio_codec": "aac",
            "fps": 30.0,
            "duration": 4.0,
            "nb_frames": 120,
        },
    )
    monkeypatch.setattr(titlecards, "build_titlecard_frame", lambda *_a, **_k: None)
    monkeypatch.setattr(titlecards, "get_or_build_endcard", lambda *_a, **_k: None)

    ffmpeg_calls = []
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda cmd, **_k: ffmpeg_calls.append(cmd) or None,
    )

    ok = titlecards.wrap_clip_with_cards(clip, "clip.mp4")
    assert ok is True
    assert ffmpeg_calls == []


def test_resolve_card_background_title_default(monkeypatch):
    monkeypatch.setattr(config, "TITLECARD_IMAGE", "")
    path, allow_color, skip, fill_color = titlecards.resolve_card_background("title")
    assert path == utils.get_bundled_assets_root() / "assets" / "titlecard.png"
    assert allow_color is True
    assert skip is False
    assert fill_color == "black"  # missing-image fallback stays black


def test_resolve_card_background_title_color(monkeypatch):
    monkeypatch.setattr(config, "TITLECARD_IMAGE", config.CARD_IMAGE_COLOR)
    monkeypatch.setattr(config, "TITLECARD_COLOR", "#ff8800")
    path, allow_color, skip, fill_color = titlecards.resolve_card_background("title")
    assert path is None
    assert allow_color is True
    assert skip is False
    assert fill_color == "#ff8800"


def test_resolve_card_background_end_default_no_color(monkeypatch):
    monkeypatch.setattr(config, "ENDCARD_IMAGE", "")
    path, allow_color, skip, _fill = titlecards.resolve_card_background("end")
    assert path == utils.get_bundled_assets_root() / "assets" / "endcard.png"
    # Endcards keep historical behavior: render only when an image is present.
    assert allow_color is False
    assert skip is False


def test_resolve_card_background_end_none_skips(monkeypatch):
    monkeypatch.setattr(config, "ENDCARD_IMAGE", config.CARD_IMAGE_NONE)
    path, _allow_color, skip, _fill = titlecards.resolve_card_background("end")
    assert path is None
    assert skip is True


def test_resolve_card_background_end_color(monkeypatch):
    monkeypatch.setattr(config, "ENDCARD_IMAGE", config.CARD_IMAGE_COLOR)
    monkeypatch.setattr(config, "ENDCARD_COLOR", "#123456")
    path, allow_color, skip, fill_color = titlecards.resolve_card_background("end")
    assert path is None
    assert allow_color is True
    assert skip is False
    assert fill_color == "#123456"


def test_resolve_card_background_upload_existing(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    images = tmp_path / config.TITLECARD_IMAGES_DIRNAME
    images.mkdir()
    (images / "mycard.png").write_bytes(b"x")
    monkeypatch.setattr(config, "TITLECARD_IMAGE", "mycard.png")
    path, allow_color, skip, _fill = titlecards.resolve_card_background("title")
    assert path == images / "mycard.png"
    assert allow_color is True
    assert skip is False


def test_resolve_card_background_upload_missing_falls_back(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(config, "TITLECARD_IMAGE", "ghost.png")
    path, _allow_color, skip, _fill = titlecards.resolve_card_background("title")
    assert path == utils.get_bundled_assets_root() / "assets" / "titlecard.png"
    assert skip is False


def test_resolve_card_background_rejects_traversal(monkeypatch, tmp_path):
    """A '../' value must fall back to the default, not resolve outside the pool."""
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    # Plant a real file one level above the pool that a traversal value targets.
    (tmp_path / "outside.png").write_bytes(b"x")
    monkeypatch.setattr(config, "TITLECARD_IMAGE", "../outside.png")
    path, _allow_color, skip, _fill = titlecards.resolve_card_background("title")
    assert path == utils.get_bundled_assets_root() / "assets" / "titlecard.png"
    assert skip is False


def test_build_titlecard_frame_uses_configured_solid_color(monkeypatch, make_clip):
    """An explicit solid-color titlecard fills with the configured color."""
    clip = make_clip(desc="Colored")
    monkeypatch.setattr(config, "TITLECARD_IMAGE", config.CARD_IMAGE_COLOR)
    monkeypatch.setattr(config, "TITLECARD_COLOR", "#ff8800")

    commands = []
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda cmd, **_k: (
            commands.append(cmd)
            or subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")
        ),
    )
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_k: True)

    path = titlecards.build_titlecard_frame(clip, "1280x720")
    assert path is not None
    joined = " ".join(commands[0])
    # #rrggbb is converted to ffmpeg's 0xRRGGBB.
    assert "color=c=0xff8800:s=1280x720" in joined
    assert "drawtext=text='Colored'" in joined


def test_resolve_titlecard_images_bakes_color_into_identity(monkeypatch):
    import pipeline

    monkeypatch.setattr(
        pipeline.config, "TITLECARD_IMAGE", pipeline.config.CARD_IMAGE_COLOR
    )
    monkeypatch.setattr(pipeline.config, "TITLECARD_COLOR", "#ff0000")
    monkeypatch.setattr(pipeline.config, "ENDCARD_IMAGE", "")
    title, end = pipeline._resolve_titlecard_images(True)
    assert title == pipeline.config.CARD_IMAGE_COLOR + "#ff0000"
    assert end == ""


def test_get_or_build_endcard_cache_keyed_by_selection(monkeypatch):
    titlecards.clear_endcard_cache()
    builds = []

    def fake_build(_resolution, *, cancel_flag=None, card_duration_seconds=None):
        path = f"endcard_{config.ENDCARD_IMAGE or 'default'}.mp4"
        builds.append(path)
        return path

    monkeypatch.setattr(titlecards, "build_endcard_frame", fake_build)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)

    monkeypatch.setattr(config, "ENDCARD_IMAGE", "")
    first = titlecards.get_or_build_endcard("1280x720")
    again = titlecards.get_or_build_endcard("1280x720")  # cached, no rebuild
    monkeypatch.setattr(config, "ENDCARD_IMAGE", "custom.png")
    second = titlecards.get_or_build_endcard("1280x720")  # new key, rebuild

    assert first == "endcard_default.mp4"
    assert again == first
    assert second == "endcard_custom.png.mp4"
    assert builds == ["endcard_default.mp4", "endcard_custom.png.mp4"]
    titlecards.clear_endcard_cache()


def test_resolve_titlecard_images(monkeypatch, tmp_path):
    import pipeline

    # Uploads must exist on disk to be recorded by their filename.
    monkeypatch.setattr(pipeline.config, "OUTPUT_DIR", str(tmp_path))
    images = tmp_path / pipeline.config.TITLECARD_IMAGES_DIRNAME
    images.mkdir()
    (images / "a.png").write_bytes(b"x")
    (images / "b.png").write_bytes(b"x")
    monkeypatch.setattr(pipeline.config, "TITLECARD_IMAGE", "a.png")
    monkeypatch.setattr(pipeline.config, "ENDCARD_IMAGE", "b.png")
    assert pipeline._resolve_titlecard_images(True) == ("a.png", "b.png")
    assert pipeline._resolve_titlecard_images(False) == ("", "")


def test_resolve_titlecard_images_missing_upload_collapses_to_default(
    monkeypatch, tmp_path
):
    """A selected upload missing on disk records the default identity (""), so a
    clip rendered with the silent fallback regenerates once the file appears —
    rather than cache-matching on a filename it was never rendered with."""
    import pipeline

    monkeypatch.setattr(pipeline.config, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(pipeline.config, "TITLECARD_IMAGE", "ghost.png")
    monkeypatch.setattr(pipeline.config, "ENDCARD_IMAGE", "")
    title, _end = pipeline._resolve_titlecard_images(True)
    assert title == ""  # not "ghost.png"

    # Once the file appears, the identity changes -> cache miss -> regenerate.
    images = tmp_path / pipeline.config.TITLECARD_IMAGES_DIRNAME
    images.mkdir()
    (images / "ghost.png").write_bytes(b"x")
    title2, _ = pipeline._resolve_titlecard_images(True)
    assert title2 == "ghost.png"


def test_get_or_build_endcard_cache_keyed_by_color(monkeypatch):
    titlecards.clear_endcard_cache()
    builds = []

    def fake_build(_resolution, *, cancel_flag=None, card_duration_seconds=None):
        path = f"endcard_{config.ENDCARD_COLOR}.mp4"
        builds.append(path)
        return path

    monkeypatch.setattr(titlecards, "build_endcard_frame", fake_build)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)
    monkeypatch.setattr(config, "ENDCARD_IMAGE", config.CARD_IMAGE_COLOR)

    monkeypatch.setattr(config, "ENDCARD_COLOR", "#111111")
    first = titlecards.get_or_build_endcard("1280x720")
    monkeypatch.setattr(config, "ENDCARD_COLOR", "#222222")
    second = titlecards.get_or_build_endcard("1280x720")

    assert first == "endcard_#111111.mp4"
    assert second == "endcard_#222222.mp4"
    assert len(builds) == 2
    titlecards.clear_endcard_cache()


def test_card_encode_uses_fast_preset(monkeypatch, make_clip):
    """Card generation encodes with the configured fast preset + CRF."""
    clip = make_clip(desc="Fast")
    monkeypatch.setattr(config, "TITLECARD_IMAGE", "")
    monkeypatch.setattr(config, "TITLECARD_ENCODE_PRESET", "veryfast")
    monkeypatch.setattr(config, "TITLECARD_ENCODE_CRF", 20)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)

    commands = []
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda cmd, **_k: (
            commands.append(cmd)
            or subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")
        ),
    )
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_k: True)

    path = titlecards.build_titlecard_frame(clip, "1280x720")
    assert path is not None
    joined = " ".join(commands[0])
    assert "-preset veryfast" in joined
    assert "-crf 20" in joined


def test_wrap_reencode_uses_fast_preset(monkeypatch, make_clip):
    """The clip-body wrap re-encode honors the fast preset + CRF."""
    clip = make_clip()
    monkeypatch.setattr(titlecards.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(titlecards.config, "TITLECARD_ENCODE_PRESET", "veryfast")
    monkeypatch.setattr(titlecards.config, "TITLECARD_ENCODE_CRF", 20)
    monkeypatch.setattr(titlecards.Path, "is_file", lambda self: True)
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _p: {
            "width": 1280,
            "height": 720,
            "video_codec": "h264",
            "audio_codec": "aac",
            "fps": 30.0,
            "duration": 12.0,
            "nb_frames": 360,
        },
    )
    monkeypatch.setattr(
        titlecards, "build_titlecard_frame", lambda *_a, **_k: "titlecard.mp4"
    )
    monkeypatch.setattr(
        titlecards, "get_or_build_endcard", lambda *_a, **_k: "endcard.mp4"
    )
    monkeypatch.setattr(video, "verify_output_file", lambda *_a, **_k: True)

    commands = []
    monkeypatch.setattr(
        video,
        "run_ffmpeg_process",
        lambda cmd, **_k: (
            commands.append(cmd)
            or subprocess.CompletedProcess(args=cmd, returncode=0, stderr="")
        ),
    )
    monkeypatch.setattr(titlecards.os, "replace", lambda src, dst: None)

    ok = titlecards.wrap_clip_with_cards(clip, "clip.mp4")
    assert ok is True
    joined = " ".join(commands[0])
    assert "-preset veryfast" in joined
    assert "-crf 20" in joined
    # Still a single encode that preserves audio.
    assert "-c:a aac" in joined


def test_pipeline_skips_per_output_probe(monkeypatch, make_clip):
    """Regression: pipeline should probe the source video, not each generated output."""
    import pipeline

    clip = make_clip(value="00:10-00:20")
    clip["times"] = [("00:10", "00:20"), ("00:30", "00:40")]
    clip["severity"] = ""

    monkeypatch.setattr(pipeline.config, "TITLECARDS_ENABLED", True)
    monkeypatch.setattr(pipeline.config, "REENCODING", False)
    monkeypatch.setattr(pipeline.config, "DEBUGGING", False)

    probe_calls = []

    def fake_probe(path):
        probe_calls.append(path)
        return {
            "width": 1280,
            "height": 720,
            "video_codec": "h264",
            "audio_codec": "aac",
            "fps": 30.0,
            "duration": 60.0,
            "nb_frames": 1800,
        }

    monkeypatch.setattr(pipeline.video, "probe_video_properties", fake_probe)
    monkeypatch.setattr(pipeline.video, "run_ffmpeg", lambda **_k: True)

    unique_names = iter(["out1.mp4", "out2.mp4"])
    monkeypatch.setattr(
        pipeline.files, "get_unique_filename", lambda *_a, **_k: next(unique_names)
    )

    wrap_calls = []

    def fake_wrap(_clip, out_name, resolution=None, **_kwargs):
        wrap_calls.append((out_name, resolution))
        return True

    monkeypatch.setattr(pipeline.titlecards, "wrap_clip_with_cards", fake_wrap)

    generated, _paths = pipeline._process_single_clip_segments(
        clip, "source.mp4", set(), output_format="clip"
    )

    assert generated == 2
    # Source is probed once for two segments; outputs are not probed separately.
    assert probe_calls == ["source.mp4"]
    assert wrap_calls == [
        ("out1.mp4", "1280x720"),
        ("out2.mp4", "1280x720"),
    ]
