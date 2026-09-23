"""Real-ffmpeg I/O tests: the pipeline must write usable media, not just argv.

Everything else in the default suite mocks ``run_ffmpeg`` and asserts arguments;
these prove the end of the pipe — a cut clip a player can open, a readable
screenshot, a reel that is the sum of its parts (or, on a failed part, no reel
at all — the silent-short-reel class, locked here on the *real* concat path;
the mock-only version lives in ``test_clip_pipeline.py``), and the
frame-accuracy of the single pre-input seek (an assumption about ffmpeg itself,
unmockable by construction).

Deliberately capped at four (see ``agents/skills/test/SKILL.md``): titlecards
(machine-dependent ``drawtext``), OCR, Whisper, remux (already covered by
``test_container_seekability.py``), and padding/format/worker combinatorics all
stay mocked elsewhere.
"""

import shutil
import subprocess

import pytest
from PIL import Image

import config
import pipeline
import video


def _have_ffmpeg() -> bool:
    return shutil.which("ffmpeg") is not None and shutil.which("ffprobe") is not None


requires_ffmpeg = pytest.mark.skipif(
    not _have_ffmpeg(), reason="needs real ffmpeg/ffprobe to cut media"
)


def _encode(path) -> None:
    """Encode a 2 s testsrc clip; keyframe every second so short cuts stay tight."""
    command = ["ffmpeg", "-y", "-v", "error"]
    command += ["-f", "lavfi", "-i", "testsrc=duration=2:size=160x120:rate=15"]
    command += ["-c:v", "libx264", "-pix_fmt", "yuv420p", "-g", "15"]
    command += ["-f", "mp4", str(path)]
    subprocess.run(command, check=True, capture_output=True)


@pytest.fixture(scope="module")
def source_dir(tmp_path_factory):
    """One shared 2 s source at the ``{study}_{participant}.mp4`` contract path."""
    if not _have_ffmpeg():
        pytest.skip("needs real ffmpeg/ffprobe")
    directory = tmp_path_factory.mktemp("io_input")
    _encode(directory / "study_P01.mp4")
    return directory


@pytest.fixture
def output_dir(monkeypatch, source_dir, tmp_path):
    """Point the pipeline at the shared source and a fresh output dir."""
    out = tmp_path / "out"
    out.mkdir()
    monkeypatch.setattr(config, "INPUT_DIR", str(source_dir), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(out), raising=False)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    # Keep Rich progress/spinner rendering out of the test output.
    monkeypatch.setattr(pipeline.utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(pipeline.utils, "use_progress", lambda: False)
    return out


@requires_ffmpeg
def test_cut_writes_playable_clip(make_clip, output_dir):
    clip = make_clip(value="0:00-0:01")

    generated, records = pipeline.process_clips(
        [clip], output_format="clip", titlecards_enabled=False
    )

    assert generated == 1
    assert len(records) == 1
    outputs = list(output_dir.glob("*.mp4"))
    assert len(outputs) == 1
    assert outputs[0].stat().st_size > 0
    # A ~1 s window from a 2 s source; stream copy may run to the next keyframe,
    # and get_file_duration rounds to whole seconds. None would mean unprobeable.
    duration = video.get_file_duration(str(outputs[0]))
    assert duration is not None and 1 <= duration <= 2


@requires_ffmpeg
def test_screenshot_is_a_readable_image(make_clip, output_dir):
    clip = make_clip(value="0:01-0:02")  # screenshots key off the start time only

    generated, _records = pipeline.process_clips(
        [clip], output_format="screen", titlecards_enabled=False
    )

    assert generated == 1
    outputs = list(output_dir.glob("*.png"))
    assert len(outputs) == 1
    with Image.open(outputs[0]) as image:
        assert image.width > 0 and image.height > 0


@requires_ffmpeg
def test_reel_concatenates_both_windows_or_writes_nothing(
    monkeypatch, make_clip, output_dir
):
    def _clips():
        return [
            make_clip(row=3, value="0:00-0:01"),
            make_clip(row=4, value="0:01-0:02"),
        ]

    # Failed second cut: no reel file may exist afterwards. With one worker the
    # part cuts run in clip order, so failing the second call is deterministic.
    with monkeypatch.context() as patch:
        calls: list[str] = []
        real_run_ffmpeg = video.run_ffmpeg

        def fail_second_cut(**kwargs):
            calls.append(kwargs["output_file"])
            if len(calls) == 2:
                return False
            return real_run_ffmpeg(**kwargs)

        patch.setattr(pipeline.video, "run_ffmpeg", fail_second_cut)
        errors: list[tuple] = []
        patch.setattr(pipeline.utils, "error_print", lambda *a, **_k: errors.append(a))

        generated, records = pipeline.process_reel(_clips(), titlecards_enabled=False)

        assert generated == 0
        assert records == []
        assert any("Reel aborted" in e[0] for e in errors)
        # The successful first part was unlinked and its reservation released.
        assert list(output_dir.iterdir()) == []

    # Clean run: the reel is roughly the sum of both windows, not one of them.
    generated, records = pipeline.process_reel(_clips(), titlecards_enabled=False)

    assert generated == 1
    assert len(records) == 1
    reel_path = output_dir / "study_reel.mp4"
    assert reel_path.is_file() and reel_path.stat().st_size > 0
    duration = video.get_file_duration(str(reel_path))
    assert duration is not None and 2 <= duration <= 3


@requires_ffmpeg
def test_single_seek_matches_two_stage_seek_exactly(tmp_path):
    """The single pre-input -ss must decode the exact frame the old split did.

    accurate_seek_args emits one pre-input -ss and relies on ffmpeg
    decoding-and-discarding from the prior keyframe when the output is
    re-encoded. That is a claim about ffmpeg's behavior, not clipgen's, so it
    can only be locked against the real binary: every extracted frame must be
    bit-identical to the obsolete two-stage form (pre-input -ss to target-2s
    plus post-input -ss 2) it replaced. Needs its own 5 s fixture (not the
    shared 2 s source) so the timestamps straddle the old 2 s split threshold
    — both branches of the old idiom — as well as keyframe boundaries.
    """
    path = str(tmp_path / "seek.mp4")
    command = ["ffmpeg", "-y", "-v", "error"]
    command += ["-f", "lavfi", "-i", "testsrc=duration=5:size=160x120:rate=15"]
    command += ["-c:v", "libx264", "-pix_fmt", "yuv420p", "-g", "15", "-f", "mp4", path]
    subprocess.run(command, check=True, capture_output=True)

    def raw_frame(pre: list[str], post: list[str]) -> bytes:
        cmd = ["ffmpeg", *pre, "-i", path, *post, "-frames:v", "1"]
        cmd += ["-pix_fmt", "bgr24", "-f", "rawvideo", "-loglevel", "error", "pipe:1"]
        return subprocess.run(cmd, capture_output=True, check=True).stdout

    for ts in (0.4, 1.75, 3.0, 4.6):
        old = raw_frame(
            ["-ss", str(ts - 2.0)] if ts > 2.0 else [],
            ["-ss", "2.0"] if ts > 2.0 else ["-ss", str(ts)],
        )
        new = raw_frame(video.accurate_seek_args(ts), [])
        assert len(new) > 0
        assert new == old, f"frame at t={ts} drifted off the two-stage result"
