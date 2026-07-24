"""Smoke tests for Transcripts Flask API (prewarm + model status)."""

import threading
from typing import cast

import pytest

Flask = pytest.importorskip("flask").Flask

import config  # noqa: E402
import thinking_agents  # noqa: E402
import transcripts  # noqa: E402
import transcripts_server  # noqa: E402
import video  # noqa: E402
import viewer  # noqa: E402


@pytest.fixture
def tr_client(tmp_path, monkeypatch):
    app = Flask(__name__)
    app.register_blueprint(transcripts_server.transcripts_bp, url_prefix="/transcripts")

    # Seed module globals via monkeypatch so they auto-restore on teardown —
    # otherwise a later test that reads these globals without the fixture would
    # inherit this test's state (matters under random ordering).
    monkeypatch.setattr(
        transcripts_server,
        "_manifest",
        {
            "source_transcripts": {},
            "corrections": [],
            "marks": [],
        },
    )
    monkeypatch.setattr(
        transcripts_server,
        "_participants",
        [
            {
                "id": "P01",
                "video_paths": [str(tmp_path / "study_P01.mp4")],
                "has_video": False,
            }
        ],
    )
    monkeypatch.setattr(transcripts_server, "_worker", None)
    monkeypatch.setattr(transcripts_server, "_input_dir", str(tmp_path))
    monkeypatch.setattr(transcripts_server, "_transcript_model_warming", False)
    # Fresh corrected-segments cache + merged-task set per test (auto-restored).
    monkeypatch.setattr(transcripts_server, "_corrected_cache", {})
    monkeypatch.setattr(transcripts_server, "_merged_task_ids", set())
    monkeypatch.setattr(transcripts_server, "_pending_chain_pids", [])

    monkeypatch.setattr(viewer, "load_manifest_artifacts", lambda: [])

    with app.test_client() as c:
        yield c
    # Cancel any debounced manifest write armed during the test so a stray Timer
    # doesn't fire _do_persist into torn-down state after the fixture exits.
    transcripts_server._cancel_pending_persist_timer()


def test_participants_includes_transcribe_prewarm(tr_client):
    resp = tr_client.get("/transcripts/api/participants")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["transcribe_prewarm"] in ("off", "queue_open", "page_load")


def test_audio_info_reports_tracks(tr_client, monkeypatch):
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda p: {
            "audio_tracks": [
                {"index": 0, "label": "Microphone"},
                {"index": 1, "label": "System"},
            ],
            "audio_track_count": 2,
        },
    )
    data = tr_client.get("/transcripts/api/audio-info/P01").get_json()
    assert data["ok"] is True
    assert data["audio_track_count"] == 2
    assert [t["label"] for t in data["audio_tracks"]] == ["Microphone", "System"]


def test_audio_info_unknown_participant_404(tr_client):
    assert tr_client.get("/transcripts/api/audio-info/ZZ").status_code == 404


def test_audio_track_streams(tr_client, monkeypatch, tmp_path):
    track = tmp_path / "track.m4a"
    track.write_bytes(b"audio")
    monkeypatch.setattr(video, "extract_audio_track", lambda p, idx: track)

    resp = tr_client.get("/transcripts/api/audio-track/P01/0")
    assert resp.status_code == 200
    assert resp.mimetype == "audio/mp4"
    assert tr_client.get("/transcripts/api/audio-track/ZZ/0").status_code == 404


def test_prewarm_invalid_config_normalized_in_participants(tr_client, monkeypatch):
    monkeypatch.setattr(config, "TRANSCRIBE_PREWARM", "bogus")
    resp = tr_client.get("/transcripts/api/participants")
    assert resp.status_code == 200
    assert resp.get_json()["transcribe_prewarm"] == "queue_open"


def test_transcribe_status_slim_and_segments_tail(tr_client):
    """Status poll reports partial_count (not the array); the segments endpoint
    serves the tail from a cursor."""
    worker = transcripts.TranscriptWorker()
    task = transcripts.create_transcript_task("P01", ["/v.mp4"])
    worker.enqueue(task)
    with worker._lock:
        t = worker._tasks[task["id"]]
        t["status"] = transcripts.TASK_STATUS_RUNNING
        t["partial_segments"] = [
            {"start": float(i), "end": i + 1.0, "text": str(i)} for i in range(3)
        ]
    transcripts_server._worker = worker

    status = tr_client.get("/transcripts/api/transcribe/status").get_json()
    entry = next(x for x in status["tasks"] if x["id"] == task["id"])
    assert "partial_segments" not in entry
    assert entry["partial_count"] == 3

    seg = tr_client.get(
        f"/transcripts/api/transcribe/{task['id']}/segments?since=1"
    ).get_json()
    assert [s["text"] for s in seg["segments"]] == ["1", "2"]
    assert seg["total"] == 3


def test_completion_side_effects_survive_debounced_merge_race(tr_client, monkeypatch):
    """A debounced _do_persist can win _manifest_lock and merge a completed
    transcription before _on_task_complete runs. The completion side effects —
    clearing stale agent fields and starting the thinking-agent chain — must
    still fire (they drain _pending_chain_pids, not the handler's own merge)."""
    monkeypatch.setattr(transcripts, "save_transcripts_manifest", lambda *a, **k: None)
    worker = transcripts.TranscriptWorker()
    task = transcripts.create_transcript_task("P01", ["/v.mp4"])
    worker.enqueue(task)
    new_segments = [{"start": 0.0, "end": 1.0, "text": "new words"}]
    with worker._lock:
        t = worker._tasks[task["id"]]
        t["status"] = transcripts.TASK_STATUS_COMPLETED
        t["result"] = {"segments": new_segments}
    monkeypatch.setattr(transcripts_server, "_worker", worker)

    # Stale AI output from a previous transcription of the same participant.
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [{"start": 0.0, "end": 1.0, "text": "old words"}],
        "summary": "stale summary",
    }

    chained: list[str] = []
    monkeypatch.setattr(
        transcripts_server._orchestrator,
        "run_chain",
        lambda pid: chained.append(pid),
    )

    # The debounce timer fires first and wins the merge...
    with transcripts_server._manifest_lock:
        transcripts_server._do_persist()
    # ...then the worker's completion callback runs.
    transcripts_server._on_task_complete()

    entry = transcripts_server._manifest["source_transcripts"]["P01"]
    assert entry["segments"] == new_segments
    assert "summary" not in entry  # stale output cleared
    assert chained == ["P01"]  # chain still kicked off


def test_participants_includes_video_version(tr_client, tmp_path):
    """video_version (mtime_ns) drives the frontend's media/<file>?v=… cache-bust."""
    video_file = tmp_path / "study_P09.mp4"
    video_file.write_bytes(b"\x00data")
    transcripts_server._participants = [
        {"id": "P09", "video_paths": [str(video_file)], "has_video": True},
        {
            "id": "P10",
            "video_paths": [str(tmp_path / "missing.mp4")],
            "has_video": False,
        },
    ]
    resp = tr_client.get("/transcripts/api/participants")
    assert resp.status_code == 200
    by_id = {p["id"]: p for p in resp.get_json()["participants"]}
    assert by_id["P09"]["video_version"] == video_file.stat().st_mtime_ns
    assert by_id["P10"]["video_version"] is None


def test_participants_video_version_combines_all_parts(tr_client, tmp_path):
    """Cache-bust ?v= must change when ANY part (not just the first) is replaced."""
    a = tmp_path / "study_P11-1.mp4"
    b = tmp_path / "study_P11-2.mp4"
    a.write_bytes(b"\x00a")
    b.write_bytes(b"\x00b")
    transcripts_server._participants = [
        {"id": "P11", "video_paths": [str(a), str(b)], "has_video": True},
    ]
    resp = tr_client.get("/transcripts/api/participants")
    p = resp.get_json()["participants"][0]
    assert p["video_version"] == a.stat().st_mtime_ns + b.stat().st_mtime_ns


def test_participants_stale_artifact_detection(tr_client, monkeypatch):
    transcripts_server._participants = [
        {"id": "P01", "video_paths": ["study_P01.mp4"], "has_video": True},
        {"id": "P02", "video_paths": ["study_P02.mp4"], "has_video": True},
    ]
    transcripts_server._manifest = {
        "source_transcripts": {
            "P01": {
                "segments": [{"start": 0.0, "end": 1.0, "text": "hi"}],
                "transcribed_at": "2026-05-22T12:00:00Z",
            },
            "P02": {
                "segments": [{"start": 0.0, "end": 1.0, "text": "yo"}],
                "transcribed_at": "2026-05-22T12:00:00Z",
            },
        },
        "corrections": [],
        "marks": [],
    }
    artifacts = [
        # P01's artifact was transcribed before the current source transcript.
        {
            "participant": "P01",
            "transcript": "clip_P01.srt",
            "transcript_version": "2026-05-22T09:00:00Z",
        },
        # P02's artifact is current.
        {
            "participant": "P02",
            "transcript": "clip_P02.srt",
            "transcript_version": "2026-05-22T12:00:00Z",
        },
    ]
    monkeypatch.setattr(viewer, "load_manifest_artifacts", lambda: artifacts)

    resp = tr_client.get("/transcripts/api/participants")
    assert resp.status_code == 200
    by_id = {p["id"]: p for p in resp.get_json()["participants"]}
    # The per-participant index must not let P01's stale artifact bleed into P02.
    assert by_id["P01"]["has_stale_artifacts"] is True
    assert by_id["P02"]["has_stale_artifacts"] is False


def test_warmup_skipped_when_prewarm_off(tr_client, monkeypatch):
    monkeypatch.setattr(config, "TRANSCRIBE_PREWARM", "off")
    resp = tr_client.post("/transcripts/api/transcribe/warmup", json={})
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data.get("skipped") is True
    assert data.get("reason") == "prewarm_disabled"


def test_model_status_shape(tr_client):
    resp = tr_client.get("/transcripts/api/transcribe/model-status")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert "loaded" in data
    assert "warming" in data
    assert "model" in data
    assert data["prewarm"] in ("off", "queue_open", "page_load")


def test_warmup_already_loaded_in_debugging(tr_client, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True)
    monkeypatch.setattr(config, "TRANSCRIBE_PREWARM", "queue_open")
    resp = tr_client.post("/transcripts/api/transcribe/warmup", json={})
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data.get("already_loaded") is True


def test_warmup_confirms_before_downloading_uncached_model(tr_client, monkeypatch):
    """Prewarm must never download silently — skip with model_not_cached."""
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(config, "TRANSCRIBE_PREWARM", "queue_open")
    monkeypatch.setattr(config, "TRANSCRIBE_MODEL", "medium")
    monkeypatch.setattr(transcripts, "is_transcription_model_loaded", lambda: False)
    monkeypatch.setattr(transcripts, "is_whisper_model_cached", lambda n=None: False)

    resp = tr_client.post("/transcripts/api/transcribe/warmup", json={})
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["skipped"] is True
    assert data["reason"] == "model_not_cached"
    assert data["model"] == "medium"
    assert data["size_mb"] == 1500


def test_warmup_force_downloads_uncached_model(tr_client, monkeypatch):
    """With force=true, prewarm proceeds even when the model isn't cached."""
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(config, "TRANSCRIBE_PREWARM", "queue_open")
    monkeypatch.setattr(transcripts, "is_transcription_model_loaded", lambda: False)
    monkeypatch.setattr(transcripts, "is_whisper_model_cached", lambda n=None: False)
    monkeypatch.setattr(transcripts, "warmup_transcription_model", lambda: True)

    resp = tr_client.post("/transcripts/api/transcribe/warmup", json={"force": True})
    assert resp.status_code == 200
    assert resp.get_json().get("started") is True


def test_warmup_proceeds_when_model_cached(tr_client, monkeypatch):
    """A cached model warms without any download confirmation."""
    monkeypatch.setattr(config, "DEBUGGING", False)
    monkeypatch.setattr(config, "TRANSCRIBE_PREWARM", "queue_open")
    monkeypatch.setattr(transcripts, "is_transcription_model_loaded", lambda: False)
    monkeypatch.setattr(transcripts, "is_whisper_model_cached", lambda n=None: True)
    monkeypatch.setattr(transcripts, "warmup_transcription_model", lambda: True)

    resp = tr_client.post("/transcripts/api/transcribe/warmup", json={})
    assert resp.status_code == 200
    assert resp.get_json().get("started") is True


def test_transcribe_applies_per_participant_overrides(tr_client, monkeypatch):
    """POST /api/transcribe threads {model, language} overrides onto the task."""

    captured: list[dict] = []

    class _StubWorker:
        def enqueue(self, task):
            captured.append(task)

    transcripts_server._participants = [
        {"id": "P01", "video_paths": ["/tmp/P01.mp4"], "has_video": True}
    ]
    transcripts_server._worker = cast("transcripts.TranscriptWorker", _StubWorker())
    monkeypatch.setattr(transcripts, "is_whisper_model_cached", lambda n=None: True)

    resp = tr_client.post(
        "/transcripts/api/transcribe",
        json={
            "participants": ["P01"],
            "force": True,
            "overrides": {"P01": {"model": "tiny", "language": "sv"}},
        },
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert len(captured) == 1
    assert captured[0]["model"] == "tiny"
    assert captured[0]["language"] == "sv"


def test_transcribe_without_overrides_defaults_to_none(tr_client, monkeypatch):
    """Missing overrides → task uses None (worker falls back to config defaults)."""

    captured: list[dict] = []

    class _StubWorker:
        def enqueue(self, task):
            captured.append(task)

    transcripts_server._participants = [
        {"id": "P01", "video_paths": ["/tmp/P01.mp4"], "has_video": True}
    ]
    transcripts_server._worker = cast("transcripts.TranscriptWorker", _StubWorker())
    monkeypatch.setattr(transcripts, "is_whisper_model_cached", lambda n=None: True)

    resp = tr_client.post(
        "/transcripts/api/transcribe",
        json={"participants": ["P01"], "force": True},
    )
    assert resp.status_code == 200
    assert len(captured) == 1
    assert captured[0]["model"] is None
    assert captured[0]["language"] is None


def test_transcribe_rejects_uncached_model(tr_client, monkeypatch):
    """An uncached Whisper model is gated: no enqueue, model_not_cached reason."""

    captured: list[dict] = []

    class _StubWorker:
        def enqueue(self, task):
            captured.append(task)

    transcripts_server._participants = [
        {"id": "P01", "video_paths": ["/tmp/P01.mp4"], "has_video": True}
    ]
    transcripts_server._worker = cast("transcripts.TranscriptWorker", _StubWorker())
    monkeypatch.setattr(transcripts, "is_whisper_model_cached", lambda n=None: False)

    resp = tr_client.post(
        "/transcripts/api/transcribe",
        json={
            "participants": ["P01"],
            "force": True,
            "overrides": {"P01": {"model": "large-v3"}},
        },
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is False
    assert data["reason"] == "model_not_cached"
    assert [u["model"] for u in data["uncached"]] == ["large-v3"]
    assert captured == []  # nothing enqueued


def test_transcribe_allows_uncached_with_allow_download(tr_client, monkeypatch):
    """allow_download bypasses the cache gate and enqueues normally."""

    captured: list[dict] = []

    class _StubWorker:
        def enqueue(self, task):
            captured.append(task)

    transcripts_server._participants = [
        {"id": "P01", "video_paths": ["/tmp/P01.mp4"], "has_video": True}
    ]
    transcripts_server._worker = cast("transcripts.TranscriptWorker", _StubWorker())
    monkeypatch.setattr(transcripts, "is_whisper_model_cached", lambda n=None: False)

    resp = tr_client.post(
        "/transcripts/api/transcribe",
        json={
            "participants": ["P01"],
            "force": True,
            "overrides": {"P01": {"model": "large-v3"}},
            "allow_download": True,
        },
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert len(captured) == 1
    assert captured[0]["model"] == "large-v3"


# ---- Thinking-agent stop endpoints ----


@pytest.fixture
def _agent_state_clean():
    """Clear the agent in-flight / cancel-event registries before and after.

    Also cancels any pending model-unload timers so a previous test does not
    fire an unload during the next test's setup.
    """

    def _reset() -> None:
        for key in transcripts_server._orchestrator._in_flight:
            transcripts_server._orchestrator._in_flight[key].clear()
        for key in transcripts_server._orchestrator._cancel_events:
            transcripts_server._orchestrator._cancel_events[key].clear()
        for key in transcripts_server._orchestrator._started_at:
            transcripts_server._orchestrator._started_at[key].clear()
        for key in transcripts_server._orchestrator._partial:
            transcripts_server._orchestrator._partial[key].clear()
        with transcripts_server._pending_model_unloads_lock:
            for timer in transcripts_server._pending_model_unloads.values():
                timer.cancel()
            transcripts_server._pending_model_unloads.clear()

    _reset()
    yield
    _reset()


def test_summary_stop_sets_cancel_event(tr_client, _agent_state_clean):
    """POST /api/agent/summary/<pid>/stop must set the registered event and remove
    the participant from the in-flight set so the UI flips to idle."""
    pid = "P01"
    evt = threading.Event()
    transcripts_server._orchestrator._in_flight["summary"].add(pid)
    transcripts_server._orchestrator._cancel_events["summary"][pid] = evt

    resp = tr_client.post(f"/transcripts/api/agent/summary/{pid}/stop")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["running"] is False
    assert evt.is_set()
    assert pid not in transcripts_server._orchestrator._in_flight["summary"]


def test_citations_stop_sets_cancel_event(tr_client, _agent_state_clean):
    """POST /api/agent/citations/<pid>/stop must set the registered event and
    remove the participant from the in-flight set."""
    pid = "P01"
    evt = threading.Event()
    transcripts_server._orchestrator._in_flight["citations"].add(pid)
    transcripts_server._orchestrator._cancel_events["citations"][pid] = evt

    resp = tr_client.post(f"/transcripts/api/agent/citations/{pid}/stop")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["running"] is False
    assert evt.is_set()
    assert pid not in transcripts_server._orchestrator._in_flight["citations"]


def test_summary_stop_when_not_running_is_noop(tr_client, _agent_state_clean):
    """When nothing is in flight, /stop is a no-op returning running: False."""
    resp = tr_client.post("/transcripts/api/agent/summary/P01/stop")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["running"] is False
    assert "P01" not in transcripts_server._orchestrator._cancel_events["summary"]


def test_citations_stop_when_not_running_is_noop(tr_client, _agent_state_clean):
    """When nothing is in flight, /stop is a no-op returning running: False."""
    resp = tr_client.post("/transcripts/api/agent/citations/P01/stop")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["running"] is False
    assert "P01" not in transcripts_server._orchestrator._cancel_events["citations"]


def test_orchestrator_stop_then_restart_isolates_run_state(
    _agent_state_clean, monkeypatch
):
    """Stop-then-Regenerate must not let the old daemon's ``finally`` clobber
    the successor run's slot.

    Repro: monkeypatch the summary agent's ``run`` to block on a
    ``threading.Event``. Spawn run_agent, call stop, spawn run_agent again
    (which claims the slot with a fresh cancel_event), then unblock the first
    thread. The first daemon's ``finally`` must detect that the slot is no
    longer its own and skip cleanup, leaving the successor visible and
    cancellable.
    """
    orch = transcripts_server._orchestrator
    pid = "P01"

    transcripts_server._manifest = {
        "source_transcripts": {pid: {"segments": [{"id": "s0", "text": "x"}]}},
        "corrections": [],
        "marks": [],
    }

    summary_agent = thinking_agents.get_agent("summary")
    assert summary_agent is not None

    started_first = threading.Event()
    release_first = threading.Event()

    def blocking_run_first(snapshot, cancel_event, on_token=None):
        started_first.set()
        release_first.wait(timeout=5)
        return None

    monkeypatch.setitem(summary_agent, "run", blocking_run_first)

    # T0: spawn run #1 — claims the slot, blocks inside the agent's run.
    threads_before = set(orch._threads["summary"])
    orch.run_agent("summary", pid, force=True)
    assert started_first.wait(timeout=2), "run #1 never entered blocking_run_first"
    new_threads = set(orch._threads["summary"]) - threads_before
    assert len(new_threads) == 1, "expected exactly one new daemon thread for run #1"
    thread_first = next(iter(new_threads))

    # T1: stop run #1 — releases _in_flight, fires event_1, leaves event_1
    # in _cancel_events for the daemon's finally to find.
    assert orch.stop("summary", pid) is True
    assert not orch.is_generating(pid, "summary")

    # T2: spawn run #2 — claims the slot, overwrites _cancel_events[..][pid]
    # with a fresh event_2.
    started_second = threading.Event()
    release_second = threading.Event()

    def blocking_run_second(snapshot, cancel_event, on_token=None):
        started_second.set()
        release_second.wait(timeout=5)
        return None

    monkeypatch.setitem(summary_agent, "run", blocking_run_second)
    orch.run_agent("summary", pid, force=True)
    assert started_second.wait(timeout=2), "run #2 never entered blocking_run_second"
    assert orch.is_generating(pid, "summary"), "run #2 did not claim the slot"

    # T3: unblock run #1 — its finally must NOT clobber run #2's slot. Without
    # the identity gate, the next two assertions would fail (run #2 would
    # become invisible and uncancellable).
    release_first.set()
    # Wait for thread #1 specifically to exit so its finally has definitely
    # fired (both daemons share the same thread name, so we must join by ref).
    thread_first.join(timeout=2.0)
    assert not thread_first.is_alive(), "run #1 daemon thread never exited"

    assert orch.is_generating(pid, "summary"), (
        "run #2 lost its slot to run #1's stale finally cleanup"
    )
    assert orch.stop("summary", pid) is True, (
        "run #2 became uncancellable after run #1's finally fired"
    )

    # Let the second daemon drain so it does not outlive the test.
    release_second.set()


def test_summary_partial_streams_via_sink_and_clears(
    tr_client, _agent_state_clean, monkeypatch
):
    """The orchestrator feeds a per-run token sink into the agent; the streamed
    text is exposed by partial_text() and the GET /api/agent/summary poll while the
    run is in flight, then cleared once the run finishes."""
    orch = transcripts_server._orchestrator
    pid = "P01"

    transcripts_server._manifest = {
        "source_transcripts": {pid: {"segments": [{"id": "s0", "text": "x"}]}},
        "corrections": [],
        "marks": [],
    }

    summary_agent = thinking_agents.get_agent("summary")
    assert summary_agent is not None

    streamed = threading.Event()
    release = threading.Event()

    def streaming_run(snapshot, cancel_event, on_token=None):
        assert on_token is not None, "orchestrator must supply a token sink"
        on_token("Hello")
        on_token(" world")
        streamed.set()
        release.wait(timeout=5)
        return None  # don't commit; finally releases the slot

    monkeypatch.setitem(summary_agent, "run", streaming_run)

    orch.run_agent("summary", pid, force=True)
    try:
        assert streamed.wait(timeout=2), "agent never streamed its tokens"
        assert orch.partial_text(pid, "summary") == "Hello world"

        # The GET poll surfaces the same partial while generating.
        resp = tr_client.get(f"/transcripts/api/agent/summary/{pid}")
        data = resp.get_json()
        assert data["generating"] is True
        assert data["partial"] == "Hello world"
    finally:
        release.set()

    _join_orchestrator_threads(orch)
    # Slot released → buffer cleared, poll no longer reports it.
    assert orch.partial_text(pid, "summary") == ""


def test_summary_stream_emits_partial_then_done(
    tr_client, _agent_state_clean, monkeypatch
):
    """GET /api/agent/summary/<pid>/stream yields the accumulated partial text, then a
    done event once the run is no longer generating."""
    # Grace window only applies to the not-yet-started race; zero it so a
    # not-generating stream with buffered text returns immediately.
    monkeypatch.setattr(transcripts_server, "_SUMMARY_STREAM_START_GRACE", 0)
    pid = "P01"
    orch = transcripts_server._orchestrator
    # Buffer holds text but the run is not in flight → the generator emits the
    # partial once, sees not-generating, and finishes with done.
    with orch._partial_lock:
        orch._partial["summary"][pid] = ["Hello", " world"]

    resp = tr_client.get(f"/transcripts/api/agent/summary/{pid}/stream")
    assert resp.status_code == 200
    assert resp.mimetype == "text/event-stream"
    body = resp.get_data(as_text=True)
    assert '"partial": "Hello world"' in body
    assert '"done": true' in body


def test_summary_stream_done_when_idle(tr_client, _agent_state_clean, monkeypatch):
    """A stream opened when nothing is generating (empty buffer) ends with done
    after the start grace, without emitting a partial."""
    monkeypatch.setattr(transcripts_server, "_SUMMARY_STREAM_START_GRACE", 0)
    resp = tr_client.get("/transcripts/api/agent/summary/P01/stream")
    assert resp.status_code == 200
    body = resp.get_data(as_text=True)
    assert '"done": true' in body
    assert '"partial"' not in body


def test_summary_stop_clears_partial(tr_client, _agent_state_clean):
    """Stopping an in-flight summary run drops its partial buffer."""
    pid = "P01"
    orch = transcripts_server._orchestrator
    orch._in_flight["summary"].add(pid)
    orch._cancel_events["summary"][pid] = threading.Event()
    with orch._partial_lock:
        orch._partial["summary"][pid] = ["partial text"]

    assert orch.stop("summary", pid) is True
    assert orch.partial_text(pid, "summary") == ""


# ---- Model unload scheduling ----


def test_summary_stop_schedules_model_unload(
    tr_client, _agent_state_clean, monkeypatch
):
    """A successful summary /stop must schedule an unload timer for the
    summary model. We force a long delay so the timer is observable."""
    monkeypatch.setattr(config, "OLLAMA_UNLOAD_DELAY_SECONDS", 30.0)
    pid = "P01"
    evt = threading.Event()
    transcripts_server._orchestrator._in_flight["summary"].add(pid)
    transcripts_server._orchestrator._cancel_events["summary"][pid] = evt

    tr_client.post(f"/transcripts/api/agent/summary/{pid}/stop")

    model = config.OLLAMA_SUMMARY_MODEL
    assert model in transcripts_server._pending_model_unloads


def test_citations_stop_schedules_model_unload(
    tr_client, _agent_state_clean, monkeypatch
):
    """A successful citations /stop must schedule an unload timer."""
    monkeypatch.setattr(config, "OLLAMA_UNLOAD_DELAY_SECONDS", 30.0)
    pid = "P01"
    evt = threading.Event()
    transcripts_server._orchestrator._in_flight["citations"].add(pid)
    transcripts_server._orchestrator._cancel_events["citations"][pid] = evt

    tr_client.post(f"/transcripts/api/agent/citations/{pid}/stop")

    model = config.OLLAMA_SUMMARY_MODEL
    assert model in transcripts_server._pending_model_unloads


def test_starting_a_run_cancels_pending_unload(_agent_state_clean, monkeypatch):
    """When a new agent run starts, any pending unload for the same model is
    cancelled so we don't churn on rapid stop→run cycles."""
    monkeypatch.setattr(config, "OLLAMA_UNLOAD_DELAY_SECONDS", 30.0)
    model = config.OLLAMA_SUMMARY_MODEL
    transcripts_server._schedule_model_unload(model)
    assert model in transcripts_server._pending_model_unloads

    transcripts_server._cancel_pending_unload(model)
    assert model not in transcripts_server._pending_model_unloads


def test_unload_fires_after_delay(_agent_state_clean, monkeypatch):
    """With a tiny delay, the scheduled unload should actually fire and call
    ollama_client.unload_model with the right model."""
    monkeypatch.setattr(config, "OLLAMA_UNLOAD_DELAY_SECONDS", 0.05)
    calls: list[str] = []
    monkeypatch.setattr(
        transcripts_server.ollama_client,
        "unload_model",
        lambda m: calls.append(m) or True,
    )

    transcripts_server._schedule_model_unload("test-model")

    # Wait a bit longer than the delay
    import time

    time.sleep(0.2)
    assert calls == ["test-model"]


def test_zero_delay_unloads_immediately(_agent_state_clean, monkeypatch):
    """OLLAMA_UNLOAD_DELAY_SECONDS=0 unloads synchronously without scheduling."""
    monkeypatch.setattr(config, "OLLAMA_UNLOAD_DELAY_SECONDS", 0)
    calls: list[str] = []
    monkeypatch.setattr(
        transcripts_server.ollama_client,
        "unload_model",
        lambda m: calls.append(m) or True,
    )

    transcripts_server._schedule_model_unload("test-model")

    assert calls == ["test-model"]
    assert "test-model" not in transcripts_server._pending_model_unloads


# ---- Embed subtitle endpoints ----


def _seed_transcript(pid: str, video_path: str) -> None:
    """Insert a minimal transcript entry + participant for *pid*."""
    transcripts_server._manifest["source_transcripts"][pid] = {
        "segments": [{"id": "s1", "start": 0.0, "end": 1.0, "text": "hi"}],
        "language": "en",
        "model": "tiny",
        "source_file": video_path,
        "transcribed_at": "2026-01-01T00:00:00Z",
    }


def test_embed_subtitle_happy_path(tr_client, tmp_path, monkeypatch):
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    _seed_transcript("P01", str(video_path))
    monkeypatch.setattr(
        transcripts_server.utils, "get_effective_output_dir", lambda: tmp_path
    )

    captured = {}

    def fake_mux(input_video, srt_path, output_video, **kwargs):
        captured["args"] = (input_video, srt_path, output_video, kwargs)
        # Touch the output file so files.get_unique_filename treats a second
        # run as needing a -1 suffix.
        from pathlib import Path

        Path(output_video).write_bytes(b"\x00")
        return True

    monkeypatch.setattr(transcripts_server.video, "mux_subtitles", fake_mux)

    resp = tr_client.post("/transcripts/api/embed-subtitle/P01")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["output_filename"] == "study_P01-subtitled.mp4"
    assert (tmp_path / "study_P01-subtitled.mp4").is_file()
    # mux helper received correct args
    assert captured["args"][0] == str(video_path)
    assert captured["args"][3]["track_language"] == "en"


def test_embed_subtitle_404_without_transcript(tr_client, tmp_path, monkeypatch):
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    # No transcript seeded.
    monkeypatch.setattr(
        transcripts_server.utils, "get_effective_output_dir", lambda: tmp_path
    )
    monkeypatch.setattr(
        transcripts_server.video,
        "mux_subtitles",
        lambda *a, **kw: pytest.fail("mux_subtitles should not be called"),
    )
    resp = tr_client.post("/transcripts/api/embed-subtitle/P01")
    assert resp.status_code == 404
    assert resp.get_json()["ok"] is False


def test_embed_subtitle_500_when_ffmpeg_fails(tr_client, tmp_path, monkeypatch):
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    _seed_transcript("P01", str(video_path))
    monkeypatch.setattr(
        transcripts_server.utils, "get_effective_output_dir", lambda: tmp_path
    )
    monkeypatch.setattr(
        transcripts_server.video, "mux_subtitles", lambda *a, **kw: False
    )

    resp = tr_client.post("/transcripts/api/embed-subtitle/P01")
    assert resp.status_code == 500
    body = resp.get_json()
    assert body["ok"] is False


def test_embed_all_subtitles_mixed_participants(tr_client, tmp_path, monkeypatch):
    v1 = tmp_path / "study_P01.mp4"
    v2 = tmp_path / "study_P02.mp4"
    v1.write_bytes(b"\x00")
    v2.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(v1)], "has_video": True},
        {"id": "P02", "video_paths": [str(v2)], "has_video": True},
    ]
    _seed_transcript("P01", str(v1))
    # P02 has no transcript — should be skipped silently.
    monkeypatch.setattr(
        transcripts_server.utils, "get_effective_output_dir", lambda: tmp_path
    )

    def fake_mux(input_video, srt_path, output_video, **kwargs):
        from pathlib import Path

        Path(output_video).write_bytes(b"\x00")
        return True

    monkeypatch.setattr(transcripts_server.video, "mux_subtitles", fake_mux)

    resp = tr_client.post("/transcripts/api/embed-all-subtitles")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert len(data["results"]) == 1
    assert data["results"][0]["participant"] == "P01"
    assert data["results"][0]["ok"] is True


def test_embed_all_subtitles_404_when_no_transcripts(tr_client, tmp_path):
    resp = tr_client.post("/transcripts/api/embed-all-subtitles")
    assert resp.status_code == 404
    assert resp.get_json()["ok"] is False


# ---- Friction endpoints ----


def _seed_friction_entry(pid="P01", **extra):
    """Insert a transcript entry with a summary for *pid* into the manifest."""
    entry = {
        "segments": [{"id": f"{pid}:0", "start": 0.0, "end": 1.0, "text": "um where"}],
        "summary": "A session summary.",
    }
    entry.update(extra)
    transcripts_server._manifest["source_transcripts"][pid] = entry
    return entry


def _join_orchestrator_threads(orch, timeout=2.0):
    """Wait for all daemon agent threads (and any cascade they spawn) to finish.

    Two passes because a worker can spawn a cascade thread while we are mid-join;
    the second pass catches it.
    """
    for _ in range(2):
        for key in list(orch._threads):
            for t in list(orch._threads[key]):
                t.join(timeout)


def test_friction_get_deterministic_when_absent_and_idle(tr_client, _agent_state_clean):
    """With segments but no stored friction (and no run in flight), the endpoint
    surfaces the pure deterministic scorer so the heatmap/stats work before the
    summary-gated agent produces LLM moments."""
    _seed_friction_entry()  # segments present, no stored friction
    resp = tr_client.get("/transcripts/api/agent/friction/P01")
    assert resp.status_code == 200
    fr = resp.get_json()["friction"]
    assert fr["deterministic"] is True
    assert fr["moments"] == []
    assert len(fr["segments"]) == 1
    assert "score" in fr["segments"][0]
    assert "by_category" in fr["stats"]


def test_friction_get_404_when_no_segments(tr_client, _agent_state_clean):
    """No segments → nothing to score → 404 (the deterministic fallback needs
    segments)."""
    transcripts_server._manifest["source_transcripts"]["P01"] = {"summary": "s"}
    resp = tr_client.get("/transcripts/api/agent/friction/P01")
    assert resp.status_code == 404
    assert resp.get_json()["ok"] is False


def test_friction_get_returns_cached(tr_client, _agent_state_clean):
    fr = {"segments": [], "moments": [], "stats": {}, "stale": False, "model": "m"}
    _seed_friction_entry(friction=fr)
    resp = tr_client.get("/transcripts/api/agent/friction/P01")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["friction"]["model"] == "m"
    assert data["friction"]["stale"] is False


def test_friction_get_generating_when_in_flight(tr_client, _agent_state_clean):
    _seed_friction_entry()
    transcripts_server._orchestrator._in_flight["friction"].add("P01")
    resp = tr_client.get("/transcripts/api/agent/friction/P01")
    assert resp.status_code == 200
    assert resp.get_json()["generating"] is True


def _claim_slot(agent_key: str, pid: str, started_at: float) -> None:
    """Mark *agent_key* in flight for *pid* with a known start time, mirroring
    what run_agent stamps when it claims the slot."""
    transcripts_server._orchestrator._in_flight[agent_key].add(pid)
    transcripts_server._orchestrator._started_at[agent_key][pid] = started_at


def test_summary_get_generating_includes_started_at(tr_client, _agent_state_clean):
    """A generating summary must surface its server start time so the frontend
    elapsed clock survives page navigation instead of resetting to zero."""
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [{"id": "P01:0", "start": 0.0, "end": 1.0, "text": "x"}],
    }
    _claim_slot("summary", "P01", 1000.0)
    resp = tr_client.get("/transcripts/api/agent/summary/P01")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["generating"] is True
    assert data["started_at"] == 1000.0


def test_citations_get_generating_includes_started_at(tr_client, _agent_state_clean):
    _seed_friction_entry()  # summary present, no citations yet
    _claim_slot("citations", "P01", 2000.0)
    resp = tr_client.get("/transcripts/api/agent/citations/P01")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["generating"] is True
    assert data["started_at"] == 2000.0


def test_summary_get_includes_citations_started_at(tr_client, _agent_state_clean):
    """On a fresh load with the summary done but citations still running, the
    citations start rides on the summary response (the only call the UI makes)."""
    _seed_friction_entry()  # summary present, no citations yet
    _claim_slot("citations", "P01", 3000.0)
    resp = tr_client.get("/transcripts/api/agent/summary/P01")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["citations_generating"] is True
    assert data["citations_started_at"] == 3000.0


def test_friction_get_generating_includes_started_at(tr_client, _agent_state_clean):
    _seed_friction_entry()
    _claim_slot("friction", "P01", 4000.0)
    resp = tr_client.get("/transcripts/api/agent/friction/P01")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["generating"] is True
    assert data["started_at"] == 4000.0


def test_friction_regenerate_404_without_summary(tr_client, _agent_state_clean):
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [{"id": "P01:0", "start": 0.0, "end": 1.0, "text": "x"}],
    }
    resp = tr_client.post("/transcripts/api/agent/friction/P01/regenerate")
    assert resp.status_code == 404


def test_friction_regenerate_triggers(tr_client, _agent_state_clean, monkeypatch):
    monkeypatch.setattr(transcripts_server, "_persist_manifest", lambda: None)
    _seed_friction_entry(friction={"stale": False, "moments": []})

    done = threading.Event()

    def stub_run(snapshot, cancel_event, on_token=None):
        done.set()
        return None

    fr_agent = thinking_agents.get_agent("friction")
    assert fr_agent is not None
    monkeypatch.setitem(fr_agent, "run", stub_run)

    resp = tr_client.post("/transcripts/api/agent/friction/P01/regenerate")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["generating"] is True
    # friction was cleared before the re-trigger so the agent recomputes.
    assert "friction" not in transcripts_server._manifest["source_transcripts"]["P01"]
    done.wait(timeout=2)


def test_friction_regenerate_returns_generating_when_in_flight(
    tr_client, _agent_state_clean
):
    _seed_friction_entry(friction={"stale": False})
    transcripts_server._orchestrator._in_flight["friction"].add("P01")
    resp = tr_client.post("/transcripts/api/agent/friction/P01/regenerate")
    assert resp.status_code == 200
    assert resp.get_json()["generating"] is True
    # In-flight short-circuit must not clear the existing friction.
    assert "friction" in transcripts_server._manifest["source_transcripts"]["P01"]


def test_friction_stop_sets_cancel_event(tr_client, _agent_state_clean):
    pid = "P01"
    evt = threading.Event()
    transcripts_server._orchestrator._in_flight["friction"].add(pid)
    transcripts_server._orchestrator._cancel_events["friction"][pid] = evt
    resp = tr_client.post(f"/transcripts/api/agent/friction/{pid}/stop")
    assert resp.status_code == 200
    assert resp.get_json()["running"] is False
    assert evt.is_set()
    assert pid not in transcripts_server._orchestrator._in_flight["friction"]


def test_friction_stop_schedules_model_unload(
    tr_client, _agent_state_clean, monkeypatch
):
    monkeypatch.setattr(config, "OLLAMA_UNLOAD_DELAY_SECONDS", 30.0)
    pid = "P01"
    evt = threading.Event()
    transcripts_server._orchestrator._in_flight["friction"].add(pid)
    transcripts_server._orchestrator._cancel_events["friction"][pid] = evt
    tr_client.post(f"/transcripts/api/agent/friction/{pid}/stop")
    # Friction inherits the summary model (OLLAMA_FRICTION_MODEL is blank by
    # default), so the unload targets the resolved model.
    assert thinking_agents.friction_model() in transcripts_server._pending_model_unloads


def test_agent_get_unknown_key_404(tr_client, _agent_state_clean):
    """An unrecognized <agent_key> must return a JSON {ok: False} 404 (not Flask's
    HTML no-match) so the frontend's data.ok/.catch paths behave."""
    resp = tr_client.get("/transcripts/api/agent/bogus/P01")
    assert resp.status_code == 404
    assert resp.get_json()["ok"] is False


def test_agent_regenerate_unknown_key_404(tr_client, _agent_state_clean):
    resp = tr_client.post("/transcripts/api/agent/bogus/P01/regenerate")
    assert resp.status_code == 404
    assert resp.get_json()["ok"] is False


def test_agent_stop_unknown_key_404(tr_client, _agent_state_clean):
    resp = tr_client.post("/transcripts/api/agent/bogus/P01/stop")
    assert resp.status_code == 404
    assert resp.get_json()["ok"] is False


def test_friction_not_eligible_when_disabled(
    tr_client, _agent_state_clean, monkeypatch
):
    monkeypatch.setattr(config, "OLLAMA_FRICTION_ENABLED", False)
    # summary + citations present so those passes are skipped; friction is the
    # only remaining candidate.
    _seed_friction_entry(citations=[{"sentence": "s", "refs": []}])
    assert transcripts_server._orchestrator.next_eligible("P01") is None


def test_friction_eligible_when_enabled(tr_client, _agent_state_clean, monkeypatch):
    monkeypatch.setattr(config, "OLLAMA_FRICTION_ENABLED", True)
    _seed_friction_entry(citations=[{"sentence": "s", "refs": []}])
    agent = transcripts_server._orchestrator.next_eligible("P01")
    assert agent is not None and agent["key"] == "friction"


def test_friction_eligible_without_citations(
    tr_client, _agent_state_clean, monkeypatch
):
    """Friction depends only on summary — it runs even when citations is off."""
    monkeypatch.setattr(config, "OLLAMA_FRICTION_ENABLED", True)
    monkeypatch.setattr(config, "OLLAMA_CITATIONS_ENABLED", False)
    _seed_friction_entry()  # summary present, no citations
    agent = transcripts_server._orchestrator.next_eligible("P01")
    assert agent is not None and agent["key"] == "friction"


def test_friction_force_eligible_regardless_of_flag(
    tr_client, _agent_state_clean, monkeypatch
):
    monkeypatch.setattr(config, "OLLAMA_FRICTION_ENABLED", False)
    monkeypatch.setattr(config, "OLLAMA_SUMMARY_ENABLED", False)
    monkeypatch.setattr(config, "OLLAMA_CITATIONS_ENABLED", False)
    _seed_friction_entry(citations=[{"sentence": "s", "refs": []}])
    agent = transcripts_server._orchestrator.next_eligible("P01", force=True)
    assert agent is not None and agent["key"] == "friction"


def test_segment_edit_marks_friction_stale(tr_client, _agent_state_clean, monkeypatch):
    monkeypatch.setattr(transcripts_server, "_schedule_persist", lambda: None)
    _seed_friction_entry(friction={"stale": False, "moments": []})
    resp = tr_client.put(
        "/transcripts/api/transcript/P01/segment",
        json={"segment_id": "P01:0", "text": "completely new text"},
    )
    assert resp.status_code == 200
    entry = transcripts_server._manifest["source_transcripts"]["P01"]
    assert entry["friction"]["stale"] is True


def test_summary_put_marks_friction_stale(tr_client, _agent_state_clean, monkeypatch):
    monkeypatch.setattr(transcripts_server, "_schedule_persist", lambda: None)
    _seed_friction_entry(
        citations=[{"sentence": "s", "refs": []}],
        friction={"stale": False},
    )
    resp = tr_client.put(
        "/transcripts/api/agent/summary/P01", json={"summary": "an edited summary"}
    )
    assert resp.status_code == 200
    entry = transcripts_server._manifest["source_transcripts"]["P01"]
    assert entry["friction"]["stale"] is True
    assert "citations" not in entry  # citations invalidated alongside friction


def test_marks_add_debounces_persist_until_flush(tr_client, monkeypatch):
    """Rapid mark edits schedule a debounced write rather than blocking each
    request on disk; _flush_pending_persist forces the pending write to land
    exactly once."""
    saved = {"count": 0}

    def _spy_save(*args, **kwargs):
        saved["count"] += 1

    monkeypatch.setattr(transcripts, "save_transcripts_manifest", _spy_save)
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [{"id": "P01:0", "start": 0.0, "end": 1.0, "text": "x"}],
    }

    resp = tr_client.post(
        "/transcripts/api/marks",
        json={"segment_ids": ["P01:0"], "category": "friction"},
    )
    assert resp.status_code == 200
    # Debounced: the route armed a timer but has not written to disk yet.
    assert saved["count"] == 0
    # Flushing collapses the pending write into a single save.
    transcripts_server._flush_pending_persist()
    assert saved["count"] == 1
    # The flush cleared the dirty state: a second flush writes nothing.
    transcripts_server._flush_pending_persist()
    assert saved["count"] == 1


def test_marks_update_sets_severity(tr_client):
    """A mark's optional severity is settable via PUT and flows through to the
    resolved GET response. New marks start with no severity; an empty value clears it."""
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [{"id": "P01:0", "start": 0.0, "end": 1.0, "text": "x"}],
    }
    add = tr_client.post(
        "/transcripts/api/marks",
        json={"segment_ids": ["P01:0"], "category": "friction"},
    )
    assert add.status_code == 200
    mark_id = add.get_json()["marks"][0]["id"]
    # A freshly created mark carries no severity.
    assert transcripts_server._manifest["marks"][0].get("severity") is None

    upd = tr_client.put(
        f"/transcripts/api/marks/{mark_id}",
        json={"severity": "High"},
    )
    assert upd.status_code == 200
    assert upd.get_json()["mark"]["severity"] == "High"
    assert transcripts_server._manifest["marks"][0]["severity"] == "High"

    # _resolve_mark spreads **mark, so GET carries the new field through.
    listed = tr_client.get("/transcripts/api/marks")
    assert listed.status_code == 200
    assert listed.get_json()["marks"][0]["severity"] == "High"

    # An empty severity clears the value back to None.
    cleared = tr_client.put(
        f"/transcripts/api/marks/{mark_id}",
        json={"severity": ""},
    )
    assert cleared.status_code == 200
    assert transcripts_server._manifest["marks"][0]["severity"] is None


def test_intake_poll_combines_status_and_marks(tr_client, _agent_state_clean):
    """/api/intake-poll returns running-state booleans + resolved marks, collapsing
    the Studio intake client's four transcript polls into one request."""
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [{"id": "P01:0", "start": 0.0, "end": 1.0, "text": "hello"}],
    }
    transcripts_server._manifest["marks"] = [
        {"id": "m1", "segment_id": "P01:0", "category": "friction", "label": None}
    ]
    resp = tr_client.get("/transcripts/api/intake-poll")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["status"] == {
        "tasks_running": False,
        "model_warming": False,
        "agents_running": False,
    }
    assert data["marks"]["categories"] == config.MARK_CATEGORIES
    resolved = data["marks"]["marks"]
    assert len(resolved) == 1
    assert resolved[0]["segment_id"] == "P01:0"
    assert resolved[0]["valid"] is True


def test_intake_poll_agents_running_without_stat(
    tr_client, monkeypatch, _agent_state_clean
):
    """agents_running is read from the orchestrator's in-memory in-flight set, so the
    route must never stat() a video file (the expensive part of /api/participants)."""
    import pathlib

    def _boom(*_a, **_k):
        raise AssertionError("intake-poll must not stat video files")

    monkeypatch.setattr(pathlib.Path, "stat", _boom)
    transcripts_server._participants = [
        {"id": "P01", "video_paths": ["/tmp/study_P01.mp4"], "has_video": True}
    ]
    transcripts_server._orchestrator._in_flight["summary"].add("P01")

    resp = tr_client.get("/transcripts/api/intake-poll")
    assert resp.status_code == 200
    assert resp.get_json()["status"]["agents_running"] is True


def test_corrected_segments_cached_and_invalidated_on_correction(
    tr_client, monkeypatch
):
    """api_transcript memoizes corrected segments and recomputes only after a
    corrections mutation bumps the version."""
    calls = {"n": 0}
    real_apply = transcripts.apply_corrections

    def _counting_apply(segments, corrections):
        calls["n"] += 1
        return real_apply(segments, corrections)

    monkeypatch.setattr(transcripts, "apply_corrections", _counting_apply)
    monkeypatch.setattr(transcripts_server, "_schedule_persist", lambda: None)
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [{"id": "P01:0", "start": 0.0, "end": 1.0, "text": "teh cat"}],
    }

    # First read computes + caches; second read hits the cache (no recompute).
    assert tr_client.get("/transcripts/api/transcript/P01").status_code == 200
    assert calls["n"] == 1
    tr_client.get("/transcripts/api/transcript/P01")
    assert calls["n"] == 1

    # Adding a correction invalidates the cache, so the next read recomputes.
    resp = tr_client.post(
        "/transcripts/api/corrections", json={"from": "teh", "to": "the"}
    )
    assert resp.status_code == 200
    r3 = tr_client.get("/transcripts/api/transcript/P01")
    assert r3.status_code == 200
    assert calls["n"] == 2
    seg = r3.get_json()["segments"][0]
    assert seg["text"] == "the cat"
    assert seg["corrected"] is True


def test_vtt_shares_corrected_segments_cache(tr_client, monkeypatch):
    """api_vtt routes through the memoized cache, so a second request for the
    same participant does not recompute corrections."""
    calls = {"n": 0}
    real_apply = transcripts.apply_corrections

    def _counting_apply(segments, corrections):
        calls["n"] += 1
        return real_apply(segments, corrections)

    monkeypatch.setattr(transcripts, "apply_corrections", _counting_apply)
    transcripts_server._bump_corrections_version()  # start from a cold cache
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [{"id": "P01:0", "start": 0.0, "end": 1.0, "text": "teh cat"}],
    }
    transcripts_server._manifest["corrections"] = [{"from": "teh", "to": "the"}]

    r1 = tr_client.get("/transcripts/api/vtt/P01")
    assert r1.status_code == 200
    assert "the cat" in r1.get_data(as_text=True)
    assert calls["n"] == 1

    # Second request hits the cache — no recompute.
    r2 = tr_client.get("/transcripts/api/vtt/P01")
    assert r2.status_code == 200
    assert calls["n"] == 1


def test_persist_keeps_corrected_cache_for_unchanged_transcription(
    tr_client, monkeypatch
):
    """The corrected cache is bumped when a task's segments are first merged,
    but a later persist of the same (already-merged) task must not bump it
    again — otherwise the memoization would clear on every debounced write."""
    import copy as _copy

    completed = {
        "id": "t1",
        "participant": "P01",
        "status": transcripts.TASK_STATUS_COMPLETED,
        "result": {
            "segments": [{"id": "P01:0", "start": 0.0, "end": 1.0, "text": "hi"}],
            "transcribed_at": "2026-06-25T00:00:00+00:00",
        },
    }

    class _FakeWorker:
        def get_all_tasks(self):
            return [_copy.deepcopy(completed)]  # mirror the real deepcopy contract

    monkeypatch.setattr(transcripts_server, "_worker", _FakeWorker())
    monkeypatch.setattr(transcripts, "save_transcripts_manifest", lambda *a, **k: None)

    # First persist merges the transcription (one expected invalidation).
    transcripts_server._persist_manifest()
    v1 = transcripts_server._corrections_version
    # A second persist with the same completed task must NOT bump again.
    transcripts_server._persist_manifest()
    assert transcripts_server._corrections_version == v1


def test_participants_includes_friction_step_state(tr_client):
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [{"id": "P01:0", "start": 0.0, "end": 1.0, "text": "x"}],
        "summary": "A summary.",
    }
    resp = tr_client.get("/transcripts/api/participants")
    assert resp.status_code == 200
    by_id = {p["id"]: p for p in resp.get_json()["participants"]}
    assert by_id["P01"]["agents"]["friction"] == "idle"


def test_citations_regenerate_does_not_run_disabled_friction(
    tr_client, _agent_state_clean, monkeypatch
):
    """Regression (H1): a manual single-agent regenerate must not cascade into a
    disabled sibling. The post-success chain advance is force=False, so with
    friction off, regenerating citations recomputes only citations."""
    monkeypatch.setattr(config, "OLLAMA_FRICTION_ENABLED", False)
    monkeypatch.setattr(transcripts_server, "_persist_manifest", lambda: None)
    # summary + citations present, friction absent — friction is the only empty field.
    _seed_friction_entry(citations=[{"sentence": "s", "refs": []}])

    friction_ran = threading.Event()

    def cit_stub(snapshot, cancel_event, on_token=None):
        return [{"sentence": "s2", "refs": []}]  # non-None → commits → cascade fires

    def fr_stub(snapshot, cancel_event, on_token=None):
        friction_ran.set()
        return None

    cit_agent = thinking_agents.get_agent("citations")
    fr_agent = thinking_agents.get_agent("friction")
    assert cit_agent is not None and fr_agent is not None
    monkeypatch.setitem(cit_agent, "run", cit_stub)
    monkeypatch.setitem(fr_agent, "run", fr_stub)

    resp = tr_client.post("/transcripts/api/agent/citations/P01/regenerate")
    assert resp.status_code == 200

    _join_orchestrator_threads(transcripts_server._orchestrator)
    assert not friction_ran.is_set(), (
        "regenerating citations must not trigger the disabled friction agent"
    )
    assert "friction" not in transcripts_server._manifest["source_transcripts"]["P01"]


def test_summary_regenerate_marks_friction_stale(
    tr_client, _agent_state_clean, monkeypatch
):
    """Regression (M1): regenerating the summary invalidates friction (the new
    summary feeds the friction prompt), mirroring the summary-edit path."""
    monkeypatch.setattr(transcripts_server, "_persist_manifest", lambda: None)
    # Don't spawn real agent threads — we only assert the endpoint's synchronous
    # manifest mutation. The generic regenerate route triggers run_agent directly.
    monkeypatch.setattr(
        transcripts_server._orchestrator, "run_agent", lambda *a, **k: None
    )
    _seed_friction_entry(
        citations=[{"sentence": "s", "refs": []}],
        friction={"stale": False, "moments": []},
    )

    resp = tr_client.post("/transcripts/api/agent/summary/P01/regenerate")
    assert resp.status_code == 200
    entry = transcripts_server._manifest["source_transcripts"]["P01"]
    assert entry["friction"]["stale"] is True
    assert "summary" not in entry  # own field cleared (popped) before re-trigger
    assert "citations" not in entry  # citations invalidated alongside friction


def test_on_task_complete_registers_summary_before_disk_write(
    tr_client, _agent_state_clean, monkeypatch
):
    """Regression: when whisper finishes, the summary agent must read as
    in-flight (the UI's "generating" state) *before* the manifest disk write
    runs — not only after it. Otherwise the frontend can observe the task as
    completed with no agent running and stop polling until a manual reload.
    """
    monkeypatch.setattr(config, "OLLAMA_SUMMARY_ENABLED", True)

    pid = "P01"
    completed_task = {
        "id": "tr_x",
        "participant": pid,
        "status": transcripts.TASK_STATUS_COMPLETED,
        "result": {
            "segments": [{"id": f"{pid}:0", "start": 0.0, "end": 1.0, "text": "hi"}],
            "language": "en",
            "model": "m",
            "source_file": "study_P01.mp4",
            "transcribed_at": "2026-06-18T00:00:00Z",
        },
    }

    class _FakeWorker:
        def get_all_tasks(self):
            return [dict(completed_task)]

    monkeypatch.setattr(transcripts_server, "_worker", _FakeWorker())

    # When the (mocked) disk write runs, the summary agent must already be
    # registered as in-flight. This assertion fires while persist runs.
    persisted = {"flag": False}

    def _spy_save(*args, **kwargs):
        assert transcripts_server._orchestrator.is_generating(pid, "summary"), (
            "summary must be in-flight before the manifest disk write"
        )
        persisted["flag"] = True
        return None  # avoid touching disk

    monkeypatch.setattr(transcripts, "save_transcripts_manifest", _spy_save)

    # Block the summary agent so its in-flight slot stays claimed while we assert.
    started = threading.Event()
    release = threading.Event()

    def _blocking_summary(snapshot, cancel_event, on_token=None):
        started.set()
        release.wait(2.0)
        return None  # don't commit/chain; the finally{} block releases the slot

    summary_agent = thinking_agents.get_agent("summary")
    assert summary_agent is not None
    monkeypatch.setitem(summary_agent, "run", _blocking_summary)

    try:
        transcripts_server._on_task_complete()

        assert started.wait(2.0), "summary agent thread did not start"
        assert transcripts_server._orchestrator.is_generating(pid, "summary")
        assert persisted["flag"], "manifest persist did not run"
    finally:
        release.set()
        _join_orchestrator_threads(transcripts_server._orchestrator)


# ---------------------------------------------------------------------------
# Local-model install gating: Ollama pull endpoints + /api/models surface
# ---------------------------------------------------------------------------


def test_ollama_pull_requires_model(tr_client):
    resp = tr_client.post("/transcripts/api/models/ollama/pull", json={})
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


def test_ollama_pull_status_unknown_model(tr_client):
    resp = tr_client.get("/transcripts/api/models/ollama/pull-status?model=ghost:1b")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["found"] is False


def test_ollama_pull_starts_and_reports_success(tr_client, monkeypatch):
    import time

    import ollama_client

    def _fake_pull(model, on_progress=None):
        if on_progress:
            on_progress({"status": "downloading", "total": 10, "completed": 5})
        return True

    monkeypatch.setattr(ollama_client, "pull_model", _fake_pull)
    transcripts_server._ollama_pull_status.clear()

    resp = tr_client.post(
        "/transcripts/api/models/ollama/pull", json={"model": "tiny:1b"}
    )
    assert resp.status_code == 200
    assert resp.get_json()["started"] is True

    # The pull runs in a daemon thread; poll until it reports done.
    status = {}
    for _ in range(100):
        status = tr_client.get(
            "/transcripts/api/models/ollama/pull-status?model=tiny:1b"
        ).get_json()
        if status.get("found") and status.get("done"):
            break
        time.sleep(0.02)
    assert status["found"] is True
    assert status["done"] is True
    assert status["succeeded"] is True


def test_ollama_pull_reports_failure(tr_client, monkeypatch):
    import time

    import ollama_client

    monkeypatch.setattr(
        ollama_client, "pull_model", lambda model, on_progress=None: False
    )
    transcripts_server._ollama_pull_status.clear()

    tr_client.post("/transcripts/api/models/ollama/pull", json={"model": "bad:1b"})
    status = {}
    for _ in range(100):
        status = tr_client.get(
            "/transcripts/api/models/ollama/pull-status?model=bad:1b"
        ).get_json()
        if status.get("found") and status.get("done"):
            break
        time.sleep(0.02)
    assert status["done"] is True
    assert status["succeeded"] is False
    assert status["error"]


def test_api_models_includes_cached_and_agents(monkeypatch):
    import ollama_client
    import server as server_mod

    monkeypatch.setattr(
        transcripts, "is_whisper_model_cached", lambda n=None: n == "base"
    )
    monkeypatch.setattr(
        ollama_client,
        "list_models",
        lambda: [
            {
                "name": "qwen3.5:9b",
                "size_bytes": 0,
                "parameter_size": "",
                "family": "",
            }
        ],
    )

    app = server_mod.build_combined_app()
    with app.test_client() as c:
        data = c.get("/api/models").get_json()

    assert data["ok"] is True
    whisper = data["whisper"]["models"]
    assert all("cached" in m for m in whisper)
    base = next(m for m in whisper if m["name"] == "base")
    assert base["cached"] is True
    tiny = next(m for m in whisper if m["name"] == "tiny")
    assert tiny["cached"] is False

    agents = data["ollama"]["agents"]
    assert {a["key"] for a in agents} == {"summary", "citations", "friction"}
    # The configured summary/friction model is present in the faked install list.
    assert all(a["installed"] for a in agents)


# ---- Completed-task merge semantics ----


class _CompletedTasksWorker:
    """Minimal worker stub exposing get_all_tasks for merge tests."""

    def __init__(self, tasks):
        self._tasks = tasks

    def get_all_tasks(self):
        return self._tasks


def test_merge_completed_results_does_not_clobber_edited_segments(monkeypatch):
    """A completed task's frozen result is merged exactly once. A later persist
    must not re-apply its original segments over in-memory edits."""
    pid = "P01"
    task = {
        "id": "tr_abc123",
        "participant": pid,
        "status": transcripts.TASK_STATUS_COMPLETED,
        "result": {
            "segments": [{"id": "s0", "text": "original"}],
            "language": "en",
        },
    }
    monkeypatch.setattr(
        transcripts_server, "_manifest", {"source_transcripts": {}}, raising=False
    )
    monkeypatch.setattr(
        transcripts_server,
        "_worker",
        cast("transcripts.TranscriptWorker", _CompletedTasksWorker([task])),
        raising=False,
    )
    transcripts_server._merged_task_ids.clear()

    transcripts_server._merge_completed_results_locked()
    src = transcripts_server._manifest["source_transcripts"]
    assert src[pid]["segments"][0]["text"] == "original"
    assert "tr_abc123" in transcripts_server._merged_task_ids

    # Simulate an in-memory edit to the segments, then persist again.
    src[pid]["segments"][0]["text"] = "edited"
    transcripts_server._merge_completed_results_locked()

    # The edit survives — the task is not re-merged.
    assert src[pid]["segments"][0]["text"] == "edited"


def test_merge_completed_results_new_task_wins(monkeypatch):
    """Re-transcription mints a new task id, so its fresh segments merge once
    and overwrite the previous result."""
    pid = "P01"
    old = {
        "id": "tr_old",
        "participant": pid,
        "status": transcripts.TASK_STATUS_COMPLETED,
        "result": {"segments": [{"id": "s0", "text": "old"}]},
    }
    monkeypatch.setattr(
        transcripts_server, "_manifest", {"source_transcripts": {}}, raising=False
    )
    monkeypatch.setattr(
        transcripts_server,
        "_worker",
        cast("transcripts.TranscriptWorker", _CompletedTasksWorker([old])),
        raising=False,
    )
    transcripts_server._merged_task_ids.clear()
    transcripts_server._merge_completed_results_locked()

    new = {
        "id": "tr_new",
        "participant": pid,
        "status": transcripts.TASK_STATUS_COMPLETED,
        "result": {"segments": [{"id": "s0", "text": "new"}]},
    }
    transcripts_server._worker = cast(
        "transcripts.TranscriptWorker", _CompletedTasksWorker([old, new])
    )
    transcripts_server._merge_completed_results_locked()

    assert (
        transcripts_server._manifest["source_transcripts"][pid]["segments"][0]["text"]
        == "new"
    )


def test_on_task_complete_refreshes_agents_on_retranscription(monkeypatch):
    """A re-transcription (new task id) for a participant who already has AI
    outputs must clear those outputs and re-run the agent chain, not leave
    them stale against the old transcript."""
    pid = "P01"
    # Participant already has a full set of AI outputs from a prior run.
    monkeypatch.setattr(
        transcripts_server,
        "_manifest",
        {
            "source_transcripts": {
                pid: {
                    "segments": [{"id": "s0", "text": "old"}],
                    "summary": {"paragraph": "stale"},
                    "citations": {"claims": []},
                    "friction": {"moments": []},
                }
            }
        },
        raising=False,
    )
    new_task = {
        "id": "tr_new",
        "participant": pid,
        "status": transcripts.TASK_STATUS_COMPLETED,
        "result": {"segments": [{"id": "s0", "text": "fresh"}]},
    }
    monkeypatch.setattr(
        transcripts_server,
        "_worker",
        cast("transcripts.TranscriptWorker", _CompletedTasksWorker([new_task])),
        raising=False,
    )
    transcripts_server._merged_task_ids.clear()

    chained: list[str] = []
    monkeypatch.setattr(
        transcripts_server._orchestrator, "run_chain", lambda p: chained.append(p)
    )
    monkeypatch.setattr(transcripts, "save_transcripts_manifest", lambda *a, **k: None)

    transcripts_server._on_task_complete()

    entry = transcripts_server._manifest["source_transcripts"][pid]
    # New segments merged; stale agent outputs cleared.
    assert entry["segments"][0]["text"] == "fresh"
    for agent in thinking_agents.AGENTS:
        assert agent["manifest_field"] not in entry
    # Chain re-run for the participant.
    assert chained == [pid]
