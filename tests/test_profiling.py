"""Tests for the opt-in profiling layer (source/profiling.py).

Covers the accumulator contract (no-op when ``config.PROFILING`` is off,
accumulation when on), the report line shape agents grep for, and the
``/api/profile`` endpoint's opt-in gating on the combined Flask app.
"""

from __future__ import annotations

import time

import pytest

import config
import profiling


@pytest.fixture(autouse=True)
def _clean_profiling(monkeypatch):
    """Every test starts with profiling off and an empty accumulator."""
    monkeypatch.setattr(config, "PROFILING", False)
    profiling.reset()
    yield
    profiling.reset()


# ---------- accumulator ------------------------------------------------------


def test_recording_is_noop_when_off():
    profiling.add("x", 1.0)
    profiling.count("y")
    with profiling.span("z"):
        pass
    assert profiling.snapshot() == {}


def test_add_and_count_accumulate_when_on(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("work", 0.5, 2)
    profiling.add("work", 0.25)
    profiling.count("hits", 3)
    snap = profiling.snapshot()
    # The n=2 batch contributes no peak (a sum has no per-item max); the n=1 add does.
    assert snap["work"] == {"seconds": 0.75, "count": 3, "max": 0.25}
    assert snap["hits"] == {"seconds": 0.0, "count": 3, "max": 0.0}


def test_span_times_the_block(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    with profiling.span("block"):
        pass
    snap = profiling.snapshot()
    assert snap["block"]["count"] == 1
    assert snap["block"]["seconds"] >= 0.0


def test_span_records_on_exception(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    with pytest.raises(ValueError, match="boom"), profiling.span("failing"):
        raise ValueError("boom")
    assert profiling.snapshot()["failing"]["count"] == 1


def test_timed_decorator(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)

    @profiling.timed("fn")
    def fn(x):
        return x + 1

    assert fn(1) == 2
    assert fn(2) == 3
    assert profiling.snapshot()["fn"]["count"] == 2


def test_snapshot_sorted_by_seconds_desc(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("small", 0.1)
    profiling.add("big", 5.0)
    assert list(profiling.snapshot()) == ["big", "small"]


def test_reset_clears(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("x", 1.0)
    profiling.reset()
    assert profiling.snapshot() == {}


def test_label_cap_drops_new_labels(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(profiling, "_MAX_LABELS", 2)
    profiling.add("a", 1.0)
    profiling.add("b", 1.0)
    profiling.add("c", 1.0)  # dropped: at cap
    profiling.add("a", 1.0)  # existing label still accumulates
    snap = profiling.snapshot()
    assert set(snap) == {"a", "b"}
    assert snap["a"]["seconds"] == 2.0


# ---------- deep profiling ---------------------------------------------------


def _deep_probe_workload():
    """A named function the deep profile must be able to attribute time to."""
    return sum(range(2000))


def test_deep_profiler_none_when_unrequested(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(config, "PROFILE_DEEP", "")
    assert profiling.deep_profiler("scan.callback.template") is None


def test_deep_profiler_none_when_profiling_off(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", False)
    monkeypatch.setattr(config, "PROFILE_DEEP", "scan.callback")
    assert profiling.deep_profiler("scan.callback.template") is None


def test_deep_profiler_matches_substring_and_reuses(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(config, "PROFILE_DEEP", "callback")
    prof = profiling.deep_profiler("scan.callback.template")
    assert prof is not None
    assert profiling.deep_profiler("scan.callback.template") is prof
    assert profiling.deep_profiler("ffmpeg.run") is None


def test_deep_report_names_the_hot_function(monkeypatch, capsys):
    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(config, "PROFILE_DEEP", "unit.deep")
    with profiling.span("unit.deep"):
        _deep_probe_workload()
    profiling.report()
    out = capsys.readouterr().out
    assert "profile-deep | unit.deep" in out
    assert "_deep_probe_workload" in out


def test_deep_report_silent_when_nothing_ran(monkeypatch, capsys):
    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(config, "PROFILE_DEEP", "")
    with profiling.span("plain"):
        pass
    profiling.report()
    assert "profile-deep" not in capsys.readouterr().out


def test_deep_nested_matching_spans_do_not_break_the_work(monkeypatch, capsys):
    """A broad match hitting an outer span AND nested work must not raise.

    cProfile raises on a second enable() in one thread; letting it propagate
    aborts the instrumented work (a --profile-deep heatmap. run lost its GIFs
    to exactly this). The outermost profiler keeps running and absorbs the
    nested work.
    """
    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(config, "PROFILE_DEEP", "unit.nest")

    ran = []

    @profiling.timed("unit.nest.inner")
    def inner():
        ran.append(True)
        return _deep_probe_workload()

    with profiling.span("unit.nest.outer"):
        inner()

    assert ran == [True]
    profiling.report()
    out = capsys.readouterr().out
    assert "profile-deep | unit.nest.outer" in out
    assert "_deep_probe_workload" in out  # inner work attributed to the outer


def test_deep_report_covers_timed_decorator(monkeypatch, capsys):
    """@timed functions (often run in executor threads) get deep-profiled too."""
    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(config, "PROFILE_DEEP", "unit.timed")

    @profiling.timed("unit.timed")
    def work():
        return _deep_probe_workload()

    work()
    profiling.report()
    out = capsys.readouterr().out
    assert "profile-deep | unit.timed" in out
    assert "_deep_probe_workload" in out


def test_reset_clears_deep_profiles(monkeypatch, capsys):
    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(config, "PROFILE_DEEP", "unit.deep")
    with profiling.span("unit.deep"):
        _deep_probe_workload()
    profiling.reset()
    profiling.report()
    assert "profile-deep" not in capsys.readouterr().out


def test_enable_flips_config_and_registers_report_once(monkeypatch):
    registered = []
    monkeypatch.setattr(profiling.atexit, "register", registered.append)
    monkeypatch.setattr(profiling, "_REPORT_REGISTERED", False)
    profiling.enable()
    profiling.enable()
    assert config.PROFILING is True
    assert registered == [profiling.report]


# ---------- report shape -----------------------------------------------------


def test_report_line_shape(monkeypatch, capsys):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("scan.callback", 12.8, 312)
    profiling.count("media_cache.hit", 42)
    profiling.report()
    lines = capsys.readouterr().out.splitlines()
    assert all(line.startswith("profile | ") for line in lines)
    assert lines[0].startswith("profile | scan.callback")  # seconds-desc order
    assert "n=312" in lines[0]
    assert "avg=" in lines[0]
    assert "n=42" in lines[1]
    assert "avg=" not in lines[1]  # pure counters carry no average
    # Neither fixture supplied a peak=, so no max token is invented for either.
    assert "max=" not in lines[0]
    assert "max=" not in lines[1]


def test_report_silent_when_empty(capsys):
    profiling.report()
    assert capsys.readouterr().out == ""


def test_scan_summary_single_line(monkeypatch, capsys):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.scan_summary("bench_P01.mp4", [("decode_wait", 8.1, 300)])
    out = capsys.readouterr().out
    assert out.startswith("profile | scan bench_P01.mp4: ")
    assert "decode_wait=8.100s/n=300" in out


def test_scan_summary_noop_when_off(capsys):
    profiling.scan_summary("x", [("a", 1.0, 1)])
    assert capsys.readouterr().out == ""


# ---------- /api/profile endpoint ---------------------------------------------


@pytest.fixture(scope="module")
def combined_app():
    pytest.importorskip("flask")
    import server

    return server.build_combined_app(worksheet=None, default_page="studio")


@pytest.fixture
def client(combined_app):
    return combined_app.test_client()


def test_api_profile_404_when_off(client):
    resp = client.get("/api/profile")
    assert resp.status_code == 404
    assert resp.get_json()["ok"] is False


def test_api_profile_snapshot_and_reset(client, monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("scan.callback", 1.5, 10)

    resp = client.get("/api/profile")
    assert resp.status_code == 200
    body = resp.get_json()
    assert body["ok"] is True
    assert body["profile"]["scan.callback"] == {
        "seconds": 1.5,
        "count": 10,
        "max": 0.0,
    }
    # Monotonic and process-global, so it rides beside the label map, not in it.
    assert "peak_rss_mb" in body
    # Request itself was timed by the route hook.
    assert any(label.startswith("route ") for label in profiling.snapshot())

    resp = client.get("/api/profile?reset=1")
    assert resp.status_code == 200
    resp = client.get("/api/profile")
    body = resp.get_json()
    assert "scan.callback" not in body["profile"]


# ---------- streaming responses -----------------------------------------------
#
# Flask's after_request hook runs on the Response *object*, before the WSGI
# server iterates the body, so `route <rule>` records generator construction
# only. These pin the fix: the body's own wall time lands under `stream <rule>`.


@pytest.fixture
def stream_app(monkeypatch):
    pytest.importorskip("flask")
    from flask import Flask, Response

    import server_utils

    app = Flask(__name__)

    @app.route("/slow")
    def slow():
        def body():
            time.sleep(0.05)
            yield "a"
            time.sleep(0.05)
            yield "b"

        return Response(
            server_utils.profiled_stream(body()), mimetype="application/x-ndjson"
        )

    return app


def test_route_label_alone_misses_streamed_body(stream_app, monkeypatch):
    """The defect itself: the route hook sees ~0 for a body that takes 0.1s."""
    import server

    monkeypatch.setattr(config, "PROFILING", True)
    stream_app.before_request(server._profile_request_start)
    stream_app.after_request(server._profile_request_end)

    resp = stream_app.test_client().get("/slow")
    route_only = profiling.snapshot().get("route /slow", {}).get("seconds", 0.0)
    assert route_only < 0.02, f"route hook unexpectedly saw the body ({route_only}s)"

    assert resp.data == b"ab"  # drains the generator
    snap = profiling.snapshot()
    assert snap["stream /slow"]["seconds"] >= 0.09
    assert snap["stream /slow"]["count"] == 1
    profiling.reset()


def test_stream_span_is_passthrough_when_off(stream_app, monkeypatch):
    monkeypatch.setattr(config, "PROFILING", False)
    resp = stream_app.test_client().get("/slow")
    assert resp.data == b"ab"
    assert profiling.snapshot() == {}


def test_stream_span_records_on_client_disconnect(monkeypatch):
    """An abandoned generation still reports what it spent (GeneratorExit)."""
    monkeypatch.setattr(config, "PROFILING", True)

    def body():
        for _ in range(100):
            time.sleep(0.01)
            yield "x"

    gen = profiling.stream_span("stream /abandoned", body())
    next(gen)
    gen.close()  # what the WSGI server does when the client drops
    assert profiling.snapshot()["stream /abandoned"]["seconds"] > 0
    profiling.reset()


# ---------- max / tail -------------------------------------------------------


def test_max_tracks_largest_single_add(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("x", 0.2)
    profiling.add("x", 0.9)
    profiling.add("x", 0.4)
    assert profiling.snapshot()["x"]["max"] == 0.9


def test_batched_add_leaves_max_alone(monkeypatch):
    """A batched flush's `seconds` is a sum, so it has no per-item max to report."""
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("scan.callback", 5.0, 100)
    assert profiling.snapshot()["scan.callback"]["max"] == 0.0


def test_batched_add_honors_explicit_peak(monkeypatch):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("scan.callback", 5.0, 100, peak=0.3)
    profiling.add("scan.callback", 4.0, 80, peak=0.1)  # smaller: must not lower it
    assert profiling.snapshot()["scan.callback"]["max"] == 0.3


def test_report_shows_max_when_supplied(monkeypatch, capsys):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("transcribe.decode", 12.0, 400, peak=0.85)
    profiling.report()
    line = capsys.readouterr().out.splitlines()[0]
    assert "max=850.0ms" in line


def test_snapshot_sort_ignores_max(monkeypatch):
    """Max must never reorder the report — seconds-desc is the contract."""
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("small_but_spiky", 0.1, 1)
    profiling.add("big", 5.0, 100, peak=0.01)
    assert list(profiling.snapshot()) == ["big", "small_but_spiky"]


# ---------- peak RSS ---------------------------------------------------------


def test_peak_rss_is_plausible():
    peak = profiling.peak_rss_mb()
    if peak is None:  # non-POSIX
        return
    # A Python process running pytest is comfortably inside these bounds; the
    # point is to catch the macOS-bytes / Linux-kilobytes unit mixup, which is
    # off by 1024x in one direction or the other.
    assert 5.0 < peak < 20000.0, f"peak_rss_mb looks unit-scaled wrong: {peak}"


def test_report_appends_peak_rss(monkeypatch, capsys):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.add("x", 1.0)
    profiling.report()
    lines = capsys.readouterr().out.splitlines()
    assert all(line.startswith("profile | ") for line in lines)
    if profiling.peak_rss_mb() is not None:
        assert lines[-1].startswith("profile | peak_rss")
        assert lines[-1].rstrip().endswith("MB")


def test_report_silent_when_empty_including_rss(capsys):
    """No labels means no output at all — peak_rss must not break that contract."""
    profiling.report()
    assert capsys.readouterr().out == ""


def test_scan_summary_kind_and_extra(monkeypatch, capsys):
    monkeypatch.setattr(config, "PROFILING", True)
    profiling.scan_summary(
        "study_P01.mp4",
        [("decode", 240.1, 412)],
        kind="whisper",
        extra="audio=3598.4s  xrt=15.0x",
    )
    out = capsys.readouterr().out
    assert out.startswith("profile | whisper study_P01.mp4: ")
    assert "decode=240.100s/n=412" in out
    assert "xrt=15.0x" in out


# ---------- profiling hooks --------------------------------------------------
#
# The split exists because faster-whisper runs audio load, feature extraction,
# VAD and language detection *eagerly* inside model.transcribe(), then returns a
# lazy generator. Measured on a 30 s file with the tiny model: enabling VAD moved
# prepare 0.185s -> 0.641s while decode collapsed 1.138s -> 0.000s. An xRT
# computed on decode alone therefore reports *infinity* with VAD on, endorsing
# the knob no matter what it did; on prepare+decode it correctly reports the
# real 2x win. These pin that.


class _FakeSeg:
    def __init__(self, start, end, text):
        self.start, self.end, self.text = start, end, text


class _FakeInfo:
    language = "en"
    duration = 30.0
    duration_after_vad = 12.0


def _fake_model(segs, *, prepare_delay=0.0, per_segment_delay=0.0):
    class _M:
        def transcribe(self, _source, **_kwargs):
            time.sleep(prepare_delay)

            def gen():
                for s in segs:
                    time.sleep(per_segment_delay)
                    yield s

            return gen(), _FakeInfo()

    return _M()


@pytest.fixture
def whisper_probe(monkeypatch, tmp_path):
    """A real-shaped transcribe_video call with the model and probes stubbed out."""
    import transcripts
    import video as video_mod

    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(
        video_mod,
        "probe_video_properties",
        lambda _p: {"audio_codec": "aac", "audio_tracks": [{"index": 0}]},
    )
    import numpy as np

    monkeypatch.setattr(
        video_mod,
        "decode_audio_pcm",
        lambda _path, _idx=0: np.zeros(16000, dtype=np.float32),
    )
    src = tmp_path / "study_P01.mp4"
    src.write_bytes(b"stub")

    def run(model, **kwargs):
        monkeypatch.setattr(transcripts, "_load_model", lambda *a, **k: model)
        return transcripts.transcribe_video(str(src), **kwargs)

    return run


def test_transcribe_splits_prepare_decode_and_callback(whisper_probe):
    segs = [_FakeSeg(0.0, 5.0, "one"), _FakeSeg(5.0, 11.0, "two")]
    calls = []
    whisper_probe(
        _fake_model(segs, prepare_delay=0.05, per_segment_delay=0.03),
        on_segment=lambda end, seg: (time.sleep(0.02), calls.append(seg)),
    )
    snap = profiling.snapshot()
    assert len(calls) == 2
    # Eager setup is attributed to prepare, not to the first generator pull.
    assert snap["transcribe.prepare"]["seconds"] >= 0.04
    assert snap["transcribe.decode"]["count"] == 2
    assert snap["transcribe.decode"]["seconds"] >= 0.05
    # The callback is charged to callback, never folded into decode.
    assert snap["transcribe.callback"]["count"] == 2
    assert snap["transcribe.callback"]["seconds"] >= 0.03
    assert snap["transcribe.decode"]["max"] > 0  # batched flush still carries a peak


def test_transcribe_flushes_on_cancel(whisper_probe):
    """A cancelled long run is exactly the one whose numbers you want."""
    import transcripts

    segs = [_FakeSeg(i, i + 1.0, f"s{i}") for i in range(10)]
    state = {"n": 0}

    def cancelled():
        state["n"] += 1
        # Four pre-loop checks (before/after model load, after PCM decode,
        # after prepare) — trip on the first between-segments check so the
        # cancel lands mid-loop, where the finally-flush is what's under test.
        return state["n"] > 4

    with pytest.raises(transcripts._TranscriptionCancelled):
        whisper_probe(_fake_model(segs, per_segment_delay=0.01), cancel_flag=cancelled)
    snap = profiling.snapshot()
    assert snap["transcribe.decode"]["count"] >= 1
    assert snap["transcribe.decode"]["seconds"] > 0


def test_whisper_summary_uses_full_wall_for_xrt(whisper_probe, capsys):
    """xRT must divide audio by prepare+decode; decode alone goes to infinity with VAD."""
    segs = [_FakeSeg(0.0, 20.0, "hello")]
    whisper_probe(_fake_model(segs, prepare_delay=0.10, per_segment_delay=0.0))
    line = next(
        ln
        for ln in capsys.readouterr().out.splitlines()
        if ln.startswith("profile | whisper ")
    )
    assert "prepare=" in line and "decode=" in line
    assert "audio=20.0s" in line and "file=30.0s" in line and "vad=12.0s" in line
    xrt = float(line.split("xrt=")[1].rstrip("x"))
    # 20s of audio over >=0.10s of prepare-dominated wall. Decode was ~0, so a
    # decode-only ratio would be astronomically higher than this bound.
    assert xrt < 205, f"xrt {xrt} looks computed on decode alone, not prepare+decode"


def test_transcribe_records_nothing_when_profiling_off(whisper_probe, monkeypatch):
    monkeypatch.setattr(config, "PROFILING", False)
    whisper_probe(_fake_model([_FakeSeg(0.0, 1.0, "x")]))
    assert profiling.snapshot() == {}


# ---------- expanded hooks (ffprobe / scan kind / excel / heatmap / ollama) --


def test_scan_callback_label_uses_profile_kind(monkeypatch):
    """A named tool must not dump analysis time into the un-suffixed bucket."""
    import numpy as np
    import screenspace_frames as sf

    monkeypatch.setattr(config, "PROFILING", True)
    frame = np.zeros((8, 8, 3), dtype=np.uint8)

    def fake_pipe(*_a, **_k):
        yield (0.0, frame)
        yield (0.1, frame)

    monkeypatch.setattr(sf, "_ffmpeg_pipe_frames", fake_pipe)
    monkeypatch.setattr(
        sf.video,
        "probe_video_properties",
        lambda _p: {"width": 8, "height": 8, "video_codec": "h264"},
    )
    monkeypatch.setattr(sf.shutil, "which", lambda _x: "/bin/ffmpeg")

    def _cb(_ts, _pixels):
        time.sleep(0.01)

    sf.scan_video_frames(
        "/x.mp4",
        {"x": 0, "y": 0, "w": 8, "h": 8},
        0.1,
        _cb,
        fps=30,
        duration=1,
        profile_kind="change",
    )
    snap = profiling.snapshot()
    assert "scan.callback.change" in snap
    assert "scan.callback" not in snap
    assert snap["scan.callback.change"]["count"] == 2
    assert snap["scan.callback.change"]["max"] >= 0.01


def test_ffprobe_run_records_on_properties_probe(monkeypatch, tmp_path):
    import video as video_mod

    monkeypatch.setattr(config, "PROFILING", True)
    video_mod._video_properties_cache.clear()
    src = tmp_path / "probe.mp4"
    src.write_bytes(b"stub")
    payload = (
        '{"streams":[{"codec_type":"video","width":1280,"height":720,'
        '"codec_name":"h264","r_frame_rate":"30/1","nb_frames":"30"}],'
        '"format":{"duration":"1.0"}}'
    )
    monkeypatch.setattr(video_mod.subprocess, "check_output", lambda *_a, **_k: payload)
    assert video_mod.probe_video_properties(str(src)) is not None
    assert profiling.snapshot()["ffprobe.run"]["count"] == 1


def test_excel_load_records(monkeypatch, tmp_path):
    import excel_io
    import openpyxl

    monkeypatch.setattr(config, "PROFILING", True)
    path = tmp_path / "sheet.xlsx"
    openpyxl.Workbook().save(path)
    assert excel_io.open_excel_workbook(str(path)) is not None
    assert "sheets.excel_load" in profiling.snapshot()


def test_heatmap_gif_records(monkeypatch, tmp_path):
    import screenspace_heatmap as hm

    monkeypatch.setattr(config, "PROFILING", True)
    results = [
        {
            "timestamp": float(i),
            "matches": [{"x": 10, "y": 10, "w": 20, "h": 20, "score": 0.9}],
        }
        for i in range(4)
    ]
    out = str(tmp_path / "heat.gif")
    assert hm.generate_heatmap_gif(results, 64, 64, out) is not None
    assert "heatmap.gif" in profiling.snapshot()


def test_heatmap_gifs_pair_wall_records(monkeypatch, tmp_path):
    """Rolling heatmap writes record the pair wall so overlap is visible."""
    import screenspace

    monkeypatch.setattr(config, "PROFILING", True)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    results = [
        {
            "timestamp": float(i),
            "matches": [{"x": 10, "y": 10, "w": 20, "h": 20, "score": 0.9}],
        }
        for i in range(4)
    ]
    worker = screenspace.ScreenspaceWorker()
    attachments = worker._write_heatmap_gifs(
        "t_pair", results, 64, 64, "template", rolling=True
    )
    assert "heatmap_gif" in attachments
    assert "heatmap_rolling_gif" in attachments
    snap = profiling.snapshot()
    assert "heatmap.gifs" in snap
    assert "heatmap.gif" in snap
    assert "heatmap.rolling" in snap


def test_ollama_generate_records_even_on_connection_error(monkeypatch):
    import urllib.error
    import urllib.request

    import ollama_client

    monkeypatch.setattr(config, "PROFILING", True)

    def _boom(*_a, **_k):
        raise urllib.error.URLError("offline")

    monkeypatch.setattr(urllib.request, "urlopen", _boom)
    with pytest.raises(urllib.error.URLError):
        ollama_client._do_generate({"model": "x", "prompt": "y"})
    assert "ollama.generate" in profiling.snapshot()
