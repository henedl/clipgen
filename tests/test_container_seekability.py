"""Fragmented-MP4 detection, the remux that fixes it, and the shared routes.

The bug this guards is invisible server-side: ffmpeg reads a fragmented
recording perfectly (it uses the ``mfra`` tail index), so every probe, scan and
clip cut looks healthy while the browser cannot seek the file at all. Only the
container shape distinguishes the two fixtures below — same frames, same
duration, same codecs.
"""

import shutil
import struct
import subprocess
from pathlib import Path

import pytest

import files
import video


def _have_ffmpeg() -> bool:
    return shutil.which("ffmpeg") is not None and shutil.which("ffprobe") is not None


requires_ffmpeg = pytest.mark.skipif(
    not _have_ffmpeg(), reason="needs real ffmpeg/ffprobe to build container fixtures"
)


def _encode(path, fragmented: bool, audio_tracks: int = 0) -> None:
    """Encode a 2 s clip, optionally as a fragmented MP4 and/or with audio."""
    command = ["ffmpeg", "-y", "-v", "error"]
    command += ["-f", "lavfi", "-i", "testsrc=duration=2:size=160x120:rate=15"]
    for index in range(audio_tracks):
        command += [
            "-f",
            "lavfi",
            "-i",
            f"sine=frequency={440 * (index + 1)}:duration=2",
        ]
    command += ["-map", "0:v"]
    for index in range(audio_tracks):
        command += ["-map", f"{index + 1}:a"]
    command += ["-c:v", "libx264", "-pix_fmt", "yuv420p", "-g", "15"]
    if audio_tracks:
        command += ["-c:a", "aac"]
    if fragmented:
        command += ["-movflags", "frag_keyframe+empty_moov+default_base_moof"]
    command += ["-f", "mp4", str(path)]
    subprocess.run(command, check=True, capture_output=True)


def _seekability(path) -> dict:
    """probe_container_seekability, narrowed — these fixtures are always MP4s."""
    probed = video.probe_container_seekability(str(path))
    assert probed is not None, f"expected a classifiable container: {path}"
    return probed


def _properties(path) -> dict:
    """probe_video_properties, narrowed — these fixtures always probe clean."""
    props = video.probe_video_properties(str(path))
    assert props is not None, f"expected probeable media: {path}"
    return props


@pytest.fixture(scope="module")
def containers(tmp_path_factory):
    """Two otherwise-identical clips: one fragmented, one normal."""
    if not _have_ffmpeg():
        pytest.skip("needs real ffmpeg/ffprobe")
    directory = tmp_path_factory.mktemp("containers")
    fragmented = directory / "study_P90.mp4"
    normal = directory / "study_P91.mp4"
    _encode(fragmented, fragmented=True)
    _encode(normal, fragmented=False)
    return {"dir": directory, "fragmented": fragmented, "normal": normal}


# ---- Detection ----


@requires_ffmpeg
class TestProbeContainerSeekability:
    def test_flags_a_fragmented_recording(self, containers):
        probed = _seekability(containers["fragmented"])
        assert probed["fragmented"] is True
        assert probed["browser_seekable"] is False
        # The tell: a fragmented moov carries no movie duration, which is
        # exactly why the browser reports Infinity.
        assert not probed["header_duration"]

    def test_passes_a_normal_recording(self, containers):
        probed = _seekability(containers["normal"])
        assert probed["fragmented"] is False
        assert probed["browser_seekable"] is True
        assert probed["header_duration"] == pytest.approx(2.0, abs=0.5)

    def test_ffprobe_cannot_tell_the_two_apart(self, containers):
        """Guards the premise: this is not something an existing probe caught.

        ffmpeg reads the fragmented file via its mfra index, so duration and
        codec come back identical and no existing code path notices a problem.
        """
        frag = _properties(containers["fragmented"])
        norm = _properties(containers["normal"])
        assert frag["duration"] == pytest.approx(norm["duration"], abs=0.2)
        assert frag["video_codec"] == norm["video_codec"]

    @pytest.mark.parametrize(
        "content",
        [
            b"",
            b"not an mp4 whatsoever" * 32,
            b"\x00\x00\x00\x18ftypqt  " + b"\x00" * 64,
        ],
        ids=["empty", "garbage", "mp4-header-without-moov"],
    )
    def test_unknown_shapes_return_none(self, tmp_path, content):
        """None means "no claim" — warning on every non-MP4 would be noise."""
        path = tmp_path / "odd.mp4"
        path.write_bytes(content)
        assert video.probe_container_seekability(str(path)) is None

    def test_missing_file_returns_none(self, tmp_path):
        assert video.probe_container_seekability(str(tmp_path / "nope.mp4")) is None

    @requires_ffmpeg
    def test_result_is_cached_per_mtime(self, containers, monkeypatch):
        path = str(containers["normal"])
        video.probe_container_seekability(path)  # warm

        def _explode(_path):
            raise AssertionError("cache miss — the walk should not re-run")

        monkeypatch.setattr(video, "_walk_mp4_for_seekability", _explode)
        assert _seekability(path)["browser_seekable"] is True


# ---- Remux ----


@requires_ffmpeg
class TestRemuxToFaststart:
    def test_fixes_the_container_and_keeps_the_original(self, tmp_path):
        source = tmp_path / "study_P92.mp4"
        _encode(source, fragmented=True)
        before = _properties(source)

        succeeded, message = video.remux_to_faststart(str(source))

        assert succeeded, message
        assert _seekability(source)["browser_seekable"]
        backup = video.original_backup_path(str(source))
        assert backup.is_file()
        # The picture is untouched — this is a container rewrite, not a re-encode.
        after = _properties(source)
        assert after["duration"] == pytest.approx(before["duration"], abs=0.2)
        assert (after["width"], after["height"]) == (before["width"], before["height"])

    def test_keeps_every_audio_track(self, tmp_path):
        """Without an explicit -map 0, ffmpeg's default mapping drops all but one.

        These recordings routinely carry mic + system audio, and losing the
        second track silently would be worse than the bug being fixed.
        """
        source = tmp_path / "study_P93.mp4"
        _encode(source, fragmented=True, audio_tracks=2)
        assert _properties(source)["audio_track_count"] == 2

        succeeded, message = video.remux_to_faststart(str(source))

        assert succeeded, message
        assert _properties(source)["audio_track_count"] == 2

    def test_backup_is_invisible_to_participant_discovery(self, tmp_path, monkeypatch):
        """The .orig must not come back as a second, phantom participant."""
        source = tmp_path / "study_P94.mp4"
        _encode(source, fragmented=True)
        monkeypatch.setattr("utils.get_effective_input_dir", lambda: str(tmp_path))
        import utils

        utils._discover_videos_cache.clear()

        assert video.remux_to_faststart(str(source))[0]
        utils._discover_videos_cache.clear()

        assert [p["id"] for p in files.resolve_participant_videos()] == ["P94"]

    def test_refuses_while_an_original_is_still_kept(self, tmp_path):
        source = tmp_path / "study_P95.mp4"
        _encode(source, fragmented=True)
        assert video.remux_to_faststart(str(source))[0]

        succeeded, message = video.remux_to_faststart(str(source))

        assert succeeded is False
        assert "still kept" in message

    def test_a_bad_output_leaves_the_source_untouched(self, tmp_path, monkeypatch):
        """A verification failure must not swap: a wrong file beats no file only
        in the sense that it is far worse, so the source is never given up."""
        source = tmp_path / "study_P96.mp4"
        _encode(source, fragmented=True)
        original_bytes = source.read_bytes()
        monkeypatch.setattr(
            video, "_remux_output_mismatch", lambda *_a, **_k: "synthetic mismatch"
        )

        succeeded, message = video.remux_to_faststart(str(source))

        assert succeeded is False
        assert "synthetic mismatch" in message
        assert source.read_bytes() == original_bytes
        assert not video.original_backup_path(str(source)).exists()
        assert not list(tmp_path.glob(".*"))  # scratch file cleaned up

    def test_restore_and_discard_round_trip(self, tmp_path):
        source = tmp_path / "study_P97.mp4"
        _encode(source, fragmented=True)
        assert video.remux_to_faststart(str(source))[0]

        assert video.restore_remux_original(str(source))[0]
        assert _seekability(source)["fragmented"] is True
        assert not video.original_backup_path(str(source)).exists()

        assert video.remux_to_faststart(str(source))[0]
        assert video.discard_remux_original(str(source))[0]
        assert not video.original_backup_path(str(source)).exists()
        assert _seekability(source)["browser_seekable"]

    def test_discard_without_a_backup_reports_failure(self, tmp_path):
        source = tmp_path / "study_P98.mp4"
        _encode(source, fragmented=False)
        assert video.discard_remux_original(str(source))[0] is False


# ---- In-place audio normalization ----


def _encode_two_codec(path) -> None:
    """A 2 s clip with an AAC track 0 and an MP3 track 1.

    The codec split is the test's tracer: after normalizing one track, the
    other's codec proves whether it was stream-copied (mp3 survives) or
    silently re-encoded (it would come back aac).
    """
    command = ["ffmpeg", "-y", "-v", "error"]
    command += ["-f", "lavfi", "-i", "testsrc=duration=2:size=160x120:rate=15"]
    command += ["-f", "lavfi", "-i", "sine=frequency=440:duration=2"]
    command += ["-f", "lavfi", "-i", "sine=frequency=880:duration=2"]
    command += ["-map", "0:v", "-map", "1:a", "-map", "2:a"]
    command += ["-c:v", "libx264", "-pix_fmt", "yuv420p"]
    command += ["-c:a:0", "aac", "-c:a:1", "libmp3lame"]
    command += ["-f", "mp4", str(path)]
    subprocess.run(command, check=True, capture_output=True)


@requires_ffmpeg
class TestNormalizeAudioInplace:
    def test_normalizes_only_the_selected_track(self, tmp_path):
        source = tmp_path / "study_P80.mp4"
        _encode_two_codec(source)

        succeeded, message = video.normalize_audio_inplace(str(source), [0])

        assert succeeded is True, message
        assert video.original_backup_path(str(source)).is_file()
        after = _properties(source)
        assert after["audio_track_count"] == 2
        codecs = [t["codec"] for t in after["audio_tracks"]]
        # Track 0 went through the loudnorm encode; track 1 was stream-copied,
        # which only its surviving mp3 codec can prove.
        assert codecs == ["aac", "mp3"]
        assert _seekability(source)["browser_seekable"] is True

    def test_refuses_while_an_original_is_still_kept(self, tmp_path):
        """The backup slot is shared with remux on purpose: two in-place
        rewriters stacking .orig files would leave 'restore' ambiguous."""
        source = tmp_path / "study_P81.mp4"
        _encode(source, fragmented=True, audio_tracks=1)
        assert video.remux_to_faststart(str(source))[0]

        succeeded, message = video.normalize_audio_inplace(str(source), [0])

        assert succeeded is False
        assert "still kept" in message

    def test_a_bad_output_leaves_the_source_untouched(self, tmp_path, monkeypatch):
        source = tmp_path / "study_P82.mp4"
        _encode(source, fragmented=False, audio_tracks=1)
        original_bytes = source.read_bytes()
        monkeypatch.setattr(
            video, "_remux_output_mismatch", lambda *_a, **_k: "synthetic mismatch"
        )

        succeeded, message = video.normalize_audio_inplace(str(source), [0])

        assert succeeded is False
        assert "synthetic mismatch" in message
        assert source.read_bytes() == original_bytes
        assert not video.original_backup_path(str(source)).exists()
        assert not list(tmp_path.glob(".*"))  # scratch file cleaned up

    def test_refuses_a_container_it_cannot_rewrite(self, tmp_path):
        source = tmp_path / "study_P83.webm"
        source.write_bytes(b"stand-in")

        succeeded, message = video.normalize_audio_inplace(str(source), [0])

        assert succeeded is False
        assert ".webm" in message

    def test_refuses_a_file_with_no_audio(self, tmp_path):
        source = tmp_path / "study_P84.mp4"
        _encode(source, fragmented=False, audio_tracks=0)

        succeeded, message = video.normalize_audio_inplace(str(source), [0])

        assert succeeded is False
        assert "no audio" in message


# ---- Shared blueprint routes ----


@pytest.fixture(scope="module")
def remux_app():
    Flask = pytest.importorskip("flask").Flask
    import remux_server

    app = Flask(__name__)
    blueprint = pytest.importorskip("flask").Blueprint("probe", __name__)
    remux_server.register_remux_routes(blueprint, lambda: None)
    app.register_blueprint(blueprint, url_prefix="/probe")
    return app


@pytest.fixture
def remux_client(remux_app, tmp_path, monkeypatch):
    import remux_server
    import utils

    monkeypatch.setattr(utils, "get_effective_input_dir", lambda: str(tmp_path))
    utils._discover_videos_cache.clear()
    monkeypatch.setattr(
        utils,
        "discover_participant_videos",
        lambda study_name="": [
            {
                "id": "P01",
                "has_video": True,
                "video_paths": [str(tmp_path / "study_P01.mp4")],
            }
        ],
    )
    (tmp_path / "study_P01.mp4").write_bytes(b"stand-in for a real recording")
    monkeypatch.setattr(remux_server, "_jobs", {})
    return remux_app.test_client(), tmp_path


@requires_ffmpeg
class TestRemuxRoutes:
    def test_status_starts_empty(self, remux_client):
        client, _ = remux_client
        body = client.get("/probe/api/remux/status").get_json()
        # `unseekable` is empty here only because the fixture writes a stand-in
        # byte string, which probes as "unknown" rather than "fragmented".
        assert body == {"ok": True, "jobs": {}, "kept": {}, "unseekable": []}

    def test_unknown_participant_is_404(self, remux_client):
        client, _ = remux_client
        response = client.post("/probe/api/remux/P99")
        assert response.status_code == 404
        assert response.get_json()["ok"] is False

    def test_second_start_is_rejected_while_running(self, remux_client, monkeypatch):
        """The in-flight gate is a check-and-set, not a check-then-act: two
        clicks must never put two ffmpeg runs on the same file."""
        import threading

        import remux_server

        release = threading.Event()

        def _blocking(path, **_kwargs):
            release.wait(timeout=5)
            return True, "done"

        monkeypatch.setattr(remux_server.video, "remux_to_faststart", _blocking)
        client, _ = remux_client

        first = client.post("/probe/api/remux/P01")
        second = client.post("/probe/api/remux/P01")

        assert first.get_json()["ok"] is True
        assert second.status_code == 409
        assert "already running" in second.get_json()["error"]
        release.set()

    def test_completed_job_and_kept_original_are_reported(
        self, remux_client, monkeypatch
    ):
        import remux_server

        client, tmp_path = remux_client
        backup = tmp_path / "study_P01.mp4.orig"

        def _fake(path, **_kwargs):
            backup.write_bytes(b"the original")
            return True, "Remuxed. Original kept as 'study_P01.mp4.orig'."

        monkeypatch.setattr(remux_server.video, "remux_to_faststart", _fake)
        client.post("/probe/api/remux/P01")

        for _ in range(50):
            body = client.get("/probe/api/remux/status").get_json()
            if body["jobs"].get("P01", {}).get("state") != "running":
                break
            _sleep_briefly()
        assert body["jobs"]["P01"]["state"] == "done"
        assert body["kept"] == {"P01": ["study_P01.mp4.orig"]}

        assert client.post("/probe/api/remux/P01/discard-original").get_json()["ok"]
        assert not backup.exists()
        # Discarding clears the job so the banner stops reporting a stale result.
        assert client.get("/probe/api/remux/status").get_json()["jobs"] == {}

    def test_a_failing_remux_surfaces_the_reason(self, remux_client, monkeypatch):
        import remux_server

        monkeypatch.setattr(
            remux_server.video,
            "remux_to_faststart",
            lambda path, **_k: (False, "ffmpeg exploded"),
        )
        client, _ = remux_client
        client.post("/probe/api/remux/P01")

        for _ in range(50):
            body = client.get("/probe/api/remux/status").get_json()
            if body["jobs"].get("P01", {}).get("state") != "running":
                break
            _sleep_briefly()
        assert body["jobs"]["P01"]["state"] == "error"
        assert "ffmpeg exploded" in body["jobs"]["P01"]["error"]


@requires_ffmpeg
class TestMultiPartRemux:
    """A participant whose session spans several files.

    The retry path is the interesting one: a run that converts part 1 and then
    fails on part 2 must be resumable, or the user is stranded with a job they
    can only finish by hand-deleting a backup.
    """

    @pytest.fixture
    def two_parts(self, tmp_path):
        parts = [tmp_path / "study_P01-1.mp4", tmp_path / "study_P01-2.mp4"]
        for part in parts:
            _encode(part, fragmented=True)
        return parts

    def test_retry_skips_the_part_that_already_succeeded(self, two_parts, monkeypatch):
        import remux_server

        real = video.remux_to_faststart
        calls: list[str] = []

        def _fail_on_second(path, **kwargs):
            calls.append(Path(path).name)
            if Path(path).name.endswith("-2.mp4"):
                return False, "synthetic failure"
            return real(path, **kwargs)

        monkeypatch.setattr(remux_server.video, "remux_to_faststart", _fail_on_second)
        token: dict = {"state": "running", "progress": 0.0, "error": "", "message": ""}
        monkeypatch.setattr(remux_server, "_jobs", {"P01": token})

        remux_server._run_remux("P01", [str(p) for p in two_parts], token)
        assert token["state"] == "error"
        assert calls == ["study_P01-1.mp4", "study_P01-2.mp4"]

        # Retry: part 1 is done and must not be attempted again — re-running it
        # would trip the "an earlier original is still kept" guard and the job
        # could never reach 'done'.
        calls.clear()
        monkeypatch.setattr(remux_server.video, "remux_to_faststart", real)
        token2: dict = {"state": "running", "progress": 0.0, "error": "", "message": ""}
        monkeypatch.setattr(remux_server, "_jobs", {"P01": token2})

        remux_server._run_remux("P01", [str(p) for p in two_parts], token2)

        assert token2["state"] == "done", token2["error"]
        assert all(_seekability(p)["browser_seekable"] for p in two_parts)

    def test_already_remuxed_needs_both_a_backup_and_a_fixed_file(self, two_parts):
        import remux_server

        part = str(two_parts[0])
        assert remux_server._already_remuxed(part) is False  # no backup yet
        assert video.remux_to_faststart(part)[0]
        assert remux_server._already_remuxed(part) is True
        # Restoring removes the backup, so the part is fair game again.
        assert video.restore_remux_original(part)[0]
        assert remux_server._already_remuxed(part) is False

    def test_status_reports_seekability_from_disk_not_the_client_snapshot(
        self, remux_app, two_parts, tmp_path, monkeypatch
    ):
        """The banner reads `unseekable` from here, deliberately.

        A page holds its /api/participants snapshot until it reloads, so a remux
        run from one of the other two pages — or any run whose job has since been
        cleared — would otherwise leave it warning about a file that is already
        fixed. Re-probing per poll is what makes the banner correct.
        """
        import remux_server
        import utils

        monkeypatch.setattr(utils, "get_effective_input_dir", lambda: str(tmp_path))
        utils._discover_videos_cache.clear()
        monkeypatch.setattr(remux_server, "_jobs", {})
        client = remux_app.test_client()
        assert client.get("/probe/api/remux/status").get_json()["unseekable"] == ["P01"]

        for part in two_parts:
            assert video.remux_to_faststart(str(part))[0]
            assert video.discard_remux_original(str(part))[0]

        body = client.get("/probe/api/remux/status").get_json()
        # No job left in the registry and no kept original — only a fresh probe
        # can tell the banner this participant is fine now.
        assert body["jobs"] == {}
        assert body["kept"] == {}
        assert body["unseekable"] == []

    def test_partial_discard_reports_the_untouched_parts(
        self, remux_app, two_parts, tmp_path, monkeypatch
    ):
        """The client shows a success toast off this response — it must not
        claim a clean sweep when only one part had a backup."""
        import remux_server
        import utils

        monkeypatch.setattr(utils, "get_effective_input_dir", lambda: str(tmp_path))
        utils._discover_videos_cache.clear()
        monkeypatch.setattr(remux_server, "_jobs", {})
        # Only the first part gets remuxed, so only it has a kept original.
        assert video.remux_to_faststart(str(two_parts[0]))[0]

        body = (
            remux_app.test_client()
            .post("/probe/api/remux/P01/discard-original")
            .get_json()
        )

        assert body["ok"] is True
        assert body["applied"] == 1
        assert len(body["warnings"]) == 1
        assert "study_P01-2.mp4" in body["warnings"][0]


def _sleep_briefly() -> None:
    import time

    time.sleep(0.02)


# ---- Participant plumbing ----


@requires_ffmpeg
def test_resolve_participant_videos_carries_the_flag(
    containers, tmp_path, monkeypatch
) -> None:
    """The pages read this key; without it every banner stays silent."""
    import utils

    shutil.copy(containers["fragmented"], tmp_path / "study_P90.mp4")
    shutil.copy(containers["normal"], tmp_path / "study_P91.mp4")
    monkeypatch.setattr(utils, "get_effective_input_dir", lambda: str(tmp_path))
    utils._discover_videos_cache.clear()

    resolved = {
        p["id"]: p["browser_seekable"] for p in files.resolve_participant_videos()
    }

    assert resolved == {"P90": False, "P91": True}


def test_unknown_container_leaves_the_flag_unset(tmp_path, monkeypatch) -> None:
    """None, not False — a non-MP4 source must not be reported as broken."""
    import utils

    (tmp_path / "study_P80.mp4").write_bytes(b"definitely not an mp4")
    monkeypatch.setattr(utils, "get_effective_input_dir", lambda: str(tmp_path))
    utils._discover_videos_cache.clear()

    resolved = files.resolve_participant_videos()

    assert [p["browser_seekable"] for p in resolved] == [None]


# ---- mvhd parsing ----


class TestReadMvhdDuration:
    def test_version_0(self):
        body = b"\x00\x00\x00\x00" + b"\x00" * 8 + struct.pack(">II", 1000, 5695000)
        assert video._read_mvhd_duration(body) == pytest.approx(5695.0)

    def test_version_1(self):
        body = b"\x01\x00\x00\x00" + b"\x00" * 16 + struct.pack(">IQ", 600, 3_417_000)
        assert video._read_mvhd_duration(body) == pytest.approx(5695.0)

    def test_zero_timescale_is_not_a_division_error(self):
        body = b"\x00\x00\x00\x00" + b"\x00" * 8 + struct.pack(">II", 0, 123)
        assert video._read_mvhd_duration(body) is None

    def test_truncated_body(self):
        assert video._read_mvhd_duration(b"\x00\x00\x00") is None
