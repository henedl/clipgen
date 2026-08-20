"""Smoke tests for Transcripts Flask API (prewarm + model status)."""

import json
import threading
import time
from pathlib import Path
from typing import cast

import pytest

Flask = pytest.importorskip("flask").Flask

import config
import thinking_agents
import transcripts
import transcripts_server
import video
import viewer


@pytest.fixture(scope="module")
def tr_app():
    """The Flask app, built once for the module.

    Registering the blueprint compiles ~35 Werkzeug URL rules, which dominates
    this fixture's cost — and the app object holds no per-test state: everything
    these tests touch lives in ``transcripts_server`` module globals, re-pinned
    per test by the function-scoped ``tr_client`` below.
    """
    app = Flask(__name__)
    app.register_blueprint(transcripts_server.transcripts_bp, url_prefix="/transcripts")
    return app


@pytest.fixture
def tr_client(tr_app, tmp_path, monkeypatch):
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
    # The media route reads the input dir live rather than from a module global,
    # so point config at tmp_path instead of pinning a snapshot.
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path))
    monkeypatch.setattr(transcripts_server, "_transcript_model_warming", False)
    # Fresh corrected-segments cache + merged-task set per test (auto-restored).
    monkeypatch.setattr(transcripts_server, "_corrected_cache", {})
    monkeypatch.setattr(transcripts_server, "_merged_task_ids", set())
    monkeypatch.setattr(transcripts_server, "_pending_chain_pids", [])

    monkeypatch.setattr(viewer, "load_manifest_artifacts", list)

    with tr_app.test_client() as c:
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
    assert data["has_sheet"] is False


def test_participants_payload_reports_in_sheet(tr_client, monkeypatch):
    # The route builds its response dict field-by-field, so in_sheet only
    # survives if it is explicitly copied across.
    monkeypatch.setattr(
        transcripts_server,
        "_participants",
        [
            {
                "id": "P01",
                "video_paths": ["/tmp/clipgen-test_P01.mp4"],
                "has_video": False,
                "in_sheet": True,
            },
            {
                "id": "P13",
                "video_paths": ["/tmp/study_P13.mp4"],
                "has_video": False,
                "in_sheet": False,
            },
        ],
    )
    data = tr_client.get("/transcripts/api/participants").get_json()
    assert [p["in_sheet"] for p in data["participants"]] == [True, False]


def test_audio_info_reports_tracks(tr_client, monkeypatch):
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda p: {
            "audio_tracks": [
                {"index": 0, "label": "System"},
                {"index": 1, "label": "Microphone"},
            ],
            "audio_track_count": 2,
        },
    )
    data = tr_client.get("/transcripts/api/audio-info/P01").get_json()
    assert data["ok"] is True
    assert data["audio_track_count"] == 2
    assert [t["label"] for t in data["audio_tracks"]] == ["System", "Microphone"]
    # The picker labels its Auto option from this rather than re-deriving the
    # heuristic in JS.
    assert data["auto_index"] == 1


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
    # Running entries carry the phase sub-state ("queued" here: the test flips
    # status by hand, so _execute_task never advanced it).
    assert entry["phase"] == "queued"

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


def test_model_status_warming_during_on_demand_load(tr_client, monkeypatch):
    """A load triggered by a transcription task (not the warmup endpoint) must
    read as warming — 'not loaded, not warming' renders as 'failed to load'."""
    monkeypatch.setattr(transcripts, "_model_loading", True)
    data = tr_client.get("/transcripts/api/transcribe/model-status").get_json()
    assert data["warming"] is True


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


def test_transcribe_returns_adoptable_task_records(tr_client, monkeypatch):
    """The enqueue response is the client's only view of a new task until the
    next 3 s status poll, and it gates the *next* enqueue on it (the server has
    no in-flight guard, so a duplicate POST would run the participant twice).
    A record missing created_at loses the newest-task-per-participant reducer to
    the participant's older completed task, so a re-transcribe would read as
    still-eligible and the pill would keep painting the stale done state."""

    class _StubWorker:
        def enqueue(self, task):
            pass

    transcripts_server._participants = [
        {"id": "P01", "video_paths": ["/tmp/P01.mp4"], "has_video": True}
    ]
    transcripts_server._worker = cast("transcripts.TranscriptWorker", _StubWorker())
    monkeypatch.setattr(transcripts, "is_whisper_model_cached", lambda n=None: True)

    resp = tr_client.post(
        "/transcripts/api/transcribe", json={"participants": ["P01"], "force": True}
    )
    assert resp.status_code == 200
    task = resp.get_json()["tasks"][0]
    # Same keys the status route serves, so the client can concat the two lists.
    assert set(task) == {
        "id",
        "participant",
        "status",
        "phase",
        "progress",
        "error",
        "created_at",
        "completed_at",
    }
    assert task["participant"] == "P01"
    assert task["status"] == transcripts.TASK_STATUS_QUEUED
    assert task["created_at"], "the reducer sorts on this; it must not be null"


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
    assert captured[0]["audio_index"] is None


def _stub_transcribe_worker(monkeypatch) -> list[dict]:
    """Point the transcribe route at a capture-only worker with a cached model."""
    captured: list[dict] = []

    class _StubWorker:
        def enqueue(self, task):
            captured.append(task)

    transcripts_server._participants = [
        {"id": "P01", "video_paths": ["/tmp/P01.mp4"], "has_video": True}
    ]
    transcripts_server._worker = cast("transcripts.TranscriptWorker", _StubWorker())
    monkeypatch.setattr(transcripts, "is_whisper_model_cached", lambda n=None: True)
    return captured


@pytest.mark.parametrize("sent", [0, "0"])
def test_transcribe_keeps_audio_track_zero(tr_client, monkeypatch, sent):
    """Track 0 is a real selection, not "no override" — a falsy test loses it."""
    captured = _stub_transcribe_worker(monkeypatch)

    resp = tr_client.post(
        "/transcripts/api/transcribe",
        json={
            "participants": ["P01"],
            "force": True,
            "overrides": {"P01": {"audio_index": sent}},
        },
    )
    assert resp.status_code == 200
    assert len(captured) == 1
    assert captured[0]["audio_index"] == 0


def test_transcribe_applies_audio_track_override(tr_client, monkeypatch):
    captured = _stub_transcribe_worker(monkeypatch)

    resp = tr_client.post(
        "/transcripts/api/transcribe",
        json={
            "participants": ["P01"],
            "force": True,
            "overrides": {"P01": {"audio_index": 2}},
        },
    )
    assert resp.status_code == 200
    assert captured[0]["audio_index"] == 2


@pytest.mark.parametrize("bad", ["x", -1])
def test_transcribe_rejects_invalid_audio_track(tr_client, monkeypatch, bad):
    captured = _stub_transcribe_worker(monkeypatch)

    resp = tr_client.post(
        "/transcripts/api/transcribe",
        json={
            "participants": ["P01"],
            "force": True,
            "overrides": {"P01": {"audio_index": bad}},
        },
    )
    assert resp.get_json()["ok"] is False
    assert captured == []


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


def _ndjson(resp) -> list[dict]:
    """Parse a streamed NDJSON response body into its lines."""
    return [json.loads(line) for line in resp.data.decode().strip().split("\n")]


def _writing_mux(captured: dict):
    """A mux_subtitles stub that records its args and touches the output file."""

    def fake_mux(input_video, srt_path, output_video, **kwargs):
        captured.setdefault("calls", []).append(
            (input_video, srt_path, output_video, kwargs)
        )
        # Touch the output so files.get_unique_filename treats a second run as
        # needing a -1 suffix.
        Path(output_video).write_bytes(b"\x00")
        return True

    return fake_mux


def test_embed_subtitles_happy_path(tr_client, tmp_path, monkeypatch):
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    _seed_transcript("P01", str(video_path))
    monkeypatch.setattr(
        transcripts_server.utils, "get_effective_output_dir", lambda: tmp_path
    )

    captured: dict = {}
    monkeypatch.setattr(
        transcripts_server.video, "mux_subtitles", _writing_mux(captured)
    )

    resp = tr_client.post(
        "/transcripts/api/embed-subtitles", json={"participants": ["P01"]}
    )
    assert resp.status_code == 200
    assert resp.mimetype == "application/x-ndjson"
    header, result, done = _ndjson(resp)
    assert done == {"done": True}
    assert header == {"total": 1, "output_dir": str(tmp_path)}
    assert result["index"] == 0
    assert result["participant"] == "P01"
    assert result["ok"] is True
    assert result["output_filename"] == "study_P01-subtitled.mp4"
    assert (tmp_path / "study_P01-subtitled.mp4").is_file()
    # mux helper received correct args
    assert captured["calls"][0][0] == str(video_path)
    assert captured["calls"][0][3]["track_language"] == "en"
    # Default disposition is on unless the request opts out.
    assert captured["calls"][0][3]["set_default"] is True


def test_embed_subtitles_forwards_default_track_false(tr_client, tmp_path, monkeypatch):
    """Unticking 'Set as default track' reaches mux_subtitles as set_default=False."""
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    _seed_transcript("P01", str(video_path))
    monkeypatch.setattr(
        transcripts_server.utils, "get_effective_output_dir", lambda: tmp_path
    )

    captured: dict = {}
    monkeypatch.setattr(
        transcripts_server.video, "mux_subtitles", _writing_mux(captured)
    )

    resp = tr_client.post(
        "/transcripts/api/embed-subtitles",
        json={"participants": ["P01"], "default_track": False},
    )
    assert resp.status_code == 200
    _ndjson(resp)
    assert captured["calls"][0][3]["set_default"] is False


def test_embed_subtitles_400_without_participants(tr_client):
    resp = tr_client.post("/transcripts/api/embed-subtitles", json={"participants": []})
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


def test_embed_subtitles_streams_failure_for_missing_transcript(
    tr_client, tmp_path, monkeypatch
):
    """A participant with no transcript is an ok=false line, not a failed request.

    One unusable id must not sink the rest of a whole-study batch.
    """
    v1 = tmp_path / "study_P01.mp4"
    v2 = tmp_path / "study_P02.mp4"
    v1.write_bytes(b"\x00")
    v2.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(v1)], "has_video": True},
        {"id": "P02", "video_paths": [str(v2)], "has_video": True},
    ]
    _seed_transcript("P01", str(v1))  # P02 deliberately left untranscribed
    monkeypatch.setattr(
        transcripts_server.utils, "get_effective_output_dir", lambda: tmp_path
    )
    monkeypatch.setattr(transcripts_server.video, "mux_subtitles", _writing_mux({}))

    resp = tr_client.post(
        "/transcripts/api/embed-subtitles", json={"participants": ["P01", "P02"]}
    )
    assert resp.status_code == 200
    lines = _ndjson(resp)
    assert lines[0]["total"] == 2
    assert [ln["participant"] for ln in lines[1:-1]] == ["P01", "P02"]
    assert lines[1]["ok"] is True
    assert lines[2]["ok"] is False
    assert lines[2]["error"] == "No transcript for participant"


def test_embed_subtitles_streams_failure_when_ffmpeg_fails(
    tr_client, tmp_path, monkeypatch
):
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

    resp = tr_client.post(
        "/transcripts/api/embed-subtitles", json={"participants": ["P01"]}
    )
    assert resp.status_code == 200
    _, result, done = _ndjson(resp)
    assert done == {"done": True}
    assert result["ok"] is False
    assert result["error"] == "ffmpeg failed to mux subtitles"
    # get_unique_filename reserves by creating an empty file; a failed mux that
    # does not release it leaves a 0-byte artifact looking like a real export.
    assert not list(tmp_path.glob("*-subtitled.mp4"))


def test_embed_subtitles_truncated_stream_has_no_done_sentinel(
    tr_client, tmp_path, monkeypatch
):
    """A generator that dies mid-run just truncates the body, which the client's
    NDJSON reader cannot distinguish from a clean end — so the terminal
    sentinel's absence is what marks the run as incomplete."""
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    _seed_transcript("P01", str(video_path))
    monkeypatch.setattr(
        transcripts_server.utils, "get_effective_output_dir", lambda: tmp_path
    )

    def _explode(*_a, **_kw):
        raise RuntimeError("mux blew up")

    monkeypatch.setattr(transcripts_server, "_embed_subtitle_for_participant", _explode)

    resp = tr_client.post(
        "/transcripts/api/embed-subtitles", json={"participants": ["P01"]}
    )
    with pytest.raises(RuntimeError, match="mux blew up"):
        resp.get_data()
    # The slot must still be free — the generator's finally runs on teardown.
    assert transcripts_server._embed_busy is False


def test_embed_subtitles_cancel_stops_before_the_next_file(
    tr_client, tmp_path, monkeypatch
):
    """A cancel mid-run skips the remaining participants and ends with the flag."""
    v1 = tmp_path / "study_P01.mp4"
    v2 = tmp_path / "study_P02.mp4"
    v1.write_bytes(b"\x00")
    v2.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(v1)], "has_video": True},
        {"id": "P02", "video_paths": [str(v2)], "has_video": True},
    ]
    _seed_transcript("P01", str(v1))
    _seed_transcript("P02", str(v2))
    monkeypatch.setattr(
        transcripts_server.utils, "get_effective_output_dir", lambda: tmp_path
    )

    def cancelling_mux(input_video, srt_path, output_video, **kwargs):
        Path(output_video).write_bytes(b"\x00")
        # Stand in for the user hitting Stop while the first file is muxing.
        transcripts_server._embed_cancel_event.set()
        return True

    monkeypatch.setattr(transcripts_server.video, "mux_subtitles", cancelling_mux)

    resp = tr_client.post(
        "/transcripts/api/embed-subtitles", json={"participants": ["P01", "P02"]}
    )
    lines = _ndjson(resp)
    assert lines[0]["total"] == 2
    assert lines[1]["participant"] == "P01"
    assert lines[-2] == {"cancelled": True}
    assert lines[-1] == {"done": True}
    assert not any(ln.get("participant") == "P02" for ln in lines)
    # The slot is released even on the cancel path, so the next run can claim it.
    assert transcripts_server._embed_busy is False


def test_embed_subtitles_409_while_a_run_holds_the_slot(tr_client, monkeypatch):
    """A second tab gets 409 rather than sharing the first run's cancel event."""
    monkeypatch.setattr(transcripts_server, "_embed_busy", True)
    resp = tr_client.post(
        "/transcripts/api/embed-subtitles", json={"participants": ["P01"]}
    )
    assert resp.status_code == 409
    assert resp.get_json()["ok"] is False


def test_embed_slot_is_freed_when_the_stream_is_never_consumed(tr_client, monkeypatch):
    """Closing a generator that was never *started* runs no body at all — not
    its finally — so a response discarded before the first read would have left
    the slot held and every later embed answering 409 until restart."""
    monkeypatch.setattr(transcripts_server, "_embed_busy", False)
    monkeypatch.setattr(transcripts_server, "_embed_owner", None)

    resp = tr_client.post(
        "/transcripts/api/embed-subtitles", json={"participants": ["P01"]}
    )
    assert resp.status_code == 200
    # Tear the response down without ever pulling a line from the body.
    resp.close()

    assert transcripts_server._embed_busy is False
    assert transcripts_server._embed_owner is None


def test_embed_slot_release_is_scoped_to_its_own_run(monkeypatch):
    """The release is attempted twice per run (the generator's finally and the
    response's call_on_close). An ungated release would let the late one clear
    a successor's claim — the 863edf8f pattern."""
    monkeypatch.setattr(transcripts_server, "_embed_busy", False)
    monkeypatch.setattr(transcripts_server, "_embed_owner", None)

    first = transcripts_server._claim_embed_slot()
    assert first is not None
    transcripts_server._release_embed_slot(first)

    second = transcripts_server._claim_embed_slot()
    assert second is not None and second != first

    # The first run's straggler release must not free the second run's slot.
    transcripts_server._release_embed_slot(first)
    assert transcripts_server._embed_busy is True
    assert transcripts_server._claim_embed_slot() is None

    transcripts_server._release_embed_slot(second)
    assert transcripts_server._embed_busy is False


def test_embed_subtitles_cancel_route_sets_the_event(tr_client):
    transcripts_server._embed_cancel_event.clear()
    resp = tr_client.post("/transcripts/api/embed-subtitles/cancel")
    assert resp.status_code == 200
    assert transcripts_server._embed_cancel_event.is_set()
    transcripts_server._embed_cancel_event.clear()


# ---- Normalize audio ----


def _audio_props(track_count: int) -> dict:
    """A probe_video_properties stub with *track_count* generic audio tracks."""
    return {
        "audio_track_count": track_count,
        "audio_tracks": [
            {"index": i, "codec": "aac", "channels": 2, "label": f"Track {i + 1}"}
            for i in range(track_count)
        ],
    }


def _capturing_normalize(captured: dict, results=None):
    """A normalize_audio_inplace stub recording (path, indices) per call.

    *results* maps a path's basename to a (ok, message) tuple; unlisted paths
    succeed.
    """
    captured["calls"] = []

    def stub(path, indices, **_kwargs):
        captured["calls"].append((path, indices))
        outcome = (results or {}).get(Path(path).name)
        return outcome if outcome else (True, "Audio normalized.")

    return stub


def test_normalize_audio_happy_path_auto_track(tr_client, tmp_path, monkeypatch):
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    monkeypatch.setattr(
        transcripts_server.video, "probe_video_properties", lambda _p: _audio_props(2)
    )
    monkeypatch.setattr(
        transcripts_server.video, "pick_speech_audio_track", lambda _tracks: 1
    )
    captured: dict = {}
    monkeypatch.setattr(
        transcripts_server.video,
        "normalize_audio_inplace",
        _capturing_normalize(captured),
    )

    resp = tr_client.post(
        "/transcripts/api/normalize-audio", json={"participants": ["P01"]}
    )
    assert resp.status_code == 200
    assert resp.mimetype == "application/x-ndjson"
    header, result, done = _ndjson(resp)
    assert header == {"total": 1}
    assert done == {"done": True}
    assert result["index"] == 0
    assert result["participant"] == "P01"
    assert result["ok"] is True
    assert result["parts"] == 1
    # The default "auto" spec resolves through the speech-track heuristic.
    assert captured["calls"] == [(str(video_path), [1])]


def test_normalize_audio_all_tracks(tr_client, tmp_path, monkeypatch):
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    monkeypatch.setattr(
        transcripts_server.video, "probe_video_properties", lambda _p: _audio_props(3)
    )
    captured: dict = {}
    monkeypatch.setattr(
        transcripts_server.video,
        "normalize_audio_inplace",
        _capturing_normalize(captured),
    )

    resp = tr_client.post(
        "/transcripts/api/normalize-audio",
        json={"participants": ["P01"], "tracks": "all"},
    )
    assert resp.status_code == 200
    _ndjson(resp)
    assert captured["calls"] == [(str(video_path), [0, 1, 2])]


def test_normalize_audio_explicit_list_is_intersected_per_file(
    tr_client, tmp_path, monkeypatch
):
    """Out-of-range indices are dropped, not fatal: on a multi-part participant
    part 2 may legitimately have fewer tracks than the part the dialog probed."""
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    monkeypatch.setattr(
        transcripts_server.video, "probe_video_properties", lambda _p: _audio_props(2)
    )
    captured: dict = {}
    monkeypatch.setattr(
        transcripts_server.video,
        "normalize_audio_inplace",
        _capturing_normalize(captured),
    )

    resp = tr_client.post(
        "/transcripts/api/normalize-audio",
        json={"participants": ["P01"], "tracks": [0, 5]},
    )
    assert resp.status_code == 200
    _ndjson(resp)
    assert captured["calls"] == [(str(video_path), [0])]


def test_normalize_audio_explicit_list_with_no_match_is_a_failure_line(
    tr_client, tmp_path, monkeypatch
):
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    monkeypatch.setattr(
        transcripts_server.video, "probe_video_properties", lambda _p: _audio_props(2)
    )
    captured: dict = {}
    monkeypatch.setattr(
        transcripts_server.video,
        "normalize_audio_inplace",
        _capturing_normalize(captured),
    )

    resp = tr_client.post(
        "/transcripts/api/normalize-audio",
        json={"participants": ["P01"], "tracks": [5]},
    )
    assert resp.status_code == 200
    _, result, done = _ndjson(resp)
    assert done == {"done": True}
    assert result["ok"] is False
    assert "None of the selected tracks" in result["error"]
    assert captured["calls"] == []


def test_normalize_audio_single_track_file_overrides_the_spec(
    tr_client, tmp_path, monkeypatch
):
    """A single-track file always normalizes track 0 — there is nothing to
    choose, so even an explicit index list must not fail it."""
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    monkeypatch.setattr(
        transcripts_server.video, "probe_video_properties", lambda _p: _audio_props(1)
    )
    captured: dict = {}
    monkeypatch.setattr(
        transcripts_server.video,
        "normalize_audio_inplace",
        _capturing_normalize(captured),
    )

    resp = tr_client.post(
        "/transcripts/api/normalize-audio",
        json={"participants": ["P01"], "tracks": [1]},
    )
    assert resp.status_code == 200
    _, result, _done = _ndjson(resp)
    assert result["ok"] is True
    assert captured["calls"] == [(str(video_path), [0])]


def test_normalize_audio_multi_part_normalizes_every_part(
    tr_client, tmp_path, monkeypatch
):
    """Unlike subtitle muxing, parts are independent files — each is rewritten,
    and the participant still gets exactly one aggregate NDJSON line."""
    p1 = tmp_path / "study_P01.mp4"
    p2 = tmp_path / "study_P01 2.mp4"
    p1.write_bytes(b"\x00")
    p2.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(p1), str(p2)], "has_video": True}
    ]
    monkeypatch.setattr(
        transcripts_server.video, "probe_video_properties", lambda _p: _audio_props(1)
    )
    captured: dict = {}
    monkeypatch.setattr(
        transcripts_server.video,
        "normalize_audio_inplace",
        _capturing_normalize(captured),
    )

    resp = tr_client.post(
        "/transcripts/api/normalize-audio", json={"participants": ["P01"]}
    )
    assert resp.status_code == 200
    lines = _ndjson(resp)
    assert [c[0] for c in captured["calls"]] == [str(p1), str(p2)]
    results = [ln for ln in lines if "participant" in ln]
    assert len(results) == 1
    assert results[0]["ok"] is True
    assert results[0]["parts"] == 2


def test_normalize_audio_partial_part_failure_names_the_part(
    tr_client, tmp_path, monkeypatch
):
    p1 = tmp_path / "study_P01.mp4"
    p2 = tmp_path / "study_P01 2.mp4"
    p1.write_bytes(b"\x00")
    p2.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(p1), str(p2)], "has_video": True}
    ]
    monkeypatch.setattr(
        transcripts_server.video, "probe_video_properties", lambda _p: _audio_props(1)
    )
    captured: dict = {}
    monkeypatch.setattr(
        transcripts_server.video,
        "normalize_audio_inplace",
        _capturing_normalize(captured, results={p2.name: (False, "ffmpeg failed")}),
    )

    resp = tr_client.post(
        "/transcripts/api/normalize-audio", json={"participants": ["P01"]}
    )
    assert resp.status_code == 200
    _, result, _done = _ndjson(resp)
    # Both parts were attempted; the aggregate line names only the failed one.
    assert len(captured["calls"]) == 2
    assert result["ok"] is False
    assert p2.name in result["error"]
    assert p1.name not in result["error"]
    # Part 1 was swapped on disk despite the participant-level failure — the
    # client's post-run reload keys on this count, not on ok.
    assert result["parts_done"] == 1


def test_normalize_audio_retry_skips_parts_with_kept_originals(
    tr_client, tmp_path, monkeypatch
):
    """A retry of a half-finished multi-part participant must finish the
    remaining parts, not collect a 'still kept' refusal for the ones that
    already succeeded (which would read as a failure forever)."""
    p1 = tmp_path / "study_P01.mp4"
    p2 = tmp_path / "study_P01 2.mp4"
    p1.write_bytes(b"\x00")
    p2.write_bytes(b"\x00")
    # Part 1 already rewritten by the failed run: its backup slot is occupied.
    Path(str(p1) + ".orig").write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(p1), str(p2)], "has_video": True}
    ]
    monkeypatch.setattr(
        transcripts_server.video, "probe_video_properties", lambda _p: _audio_props(1)
    )
    captured: dict = {}
    monkeypatch.setattr(
        transcripts_server.video,
        "normalize_audio_inplace",
        _capturing_normalize(captured),
    )

    resp = tr_client.post(
        "/transcripts/api/normalize-audio", json={"participants": ["P01"]}
    )
    assert resp.status_code == 200
    _, result, _done = _ndjson(resp)
    assert [c[0] for c in captured["calls"]] == [str(p2)]
    assert result["ok"] is True
    assert result["parts_done"] == 1
    assert "already-rewritten" in result["message"]


def test_normalize_audio_fully_kept_participant_is_a_clean_noop(
    tr_client, tmp_path, monkeypatch
):
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    Path(str(video_path) + ".orig").write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    captured: dict = {}
    monkeypatch.setattr(
        transcripts_server.video,
        "normalize_audio_inplace",
        _capturing_normalize(captured),
    )

    resp = tr_client.post(
        "/transcripts/api/normalize-audio", json={"participants": ["P01"]}
    )
    assert resp.status_code == 200
    _, result, _done = _ndjson(resp)
    assert captured["calls"] == []
    assert result["ok"] is True
    assert result["parts_done"] == 0
    assert "Already rewritten" in result["message"]


def test_normalize_audio_400_without_participants(tr_client):
    resp = tr_client.post("/transcripts/api/normalize-audio", json={"participants": []})
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


@pytest.mark.parametrize("bad_tracks", ["bogus", [1.5], [True], [], {"a": 1}])
def test_normalize_audio_400_on_malformed_tracks(tr_client, bad_tracks):
    resp = tr_client.post(
        "/transcripts/api/normalize-audio",
        json={"participants": ["P01"], "tracks": bad_tracks},
    )
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


def test_normalize_audio_missing_video_is_a_failure_line(
    tr_client, tmp_path, monkeypatch
):
    """An unknown participant is an ok=false line, not a failed request — one
    bad id must not sink the rest of the batch."""
    video_path = tmp_path / "study_P01.mp4"
    video_path.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(video_path)], "has_video": True}
    ]
    monkeypatch.setattr(
        transcripts_server.video, "probe_video_properties", lambda _p: _audio_props(1)
    )
    captured: dict = {}
    monkeypatch.setattr(
        transcripts_server.video,
        "normalize_audio_inplace",
        _capturing_normalize(captured),
    )

    resp = tr_client.post(
        "/transcripts/api/normalize-audio", json={"participants": ["P99", "P01"]}
    )
    assert resp.status_code == 200
    lines = _ndjson(resp)
    assert lines[1]["participant"] == "P99"
    assert lines[1]["ok"] is False
    assert lines[1]["error"] == "Source video not found"
    assert lines[2]["participant"] == "P01"
    assert lines[2]["ok"] is True


def test_normalize_audio_cancel_stops_before_the_next_participant(
    tr_client, tmp_path, monkeypatch
):
    v1 = tmp_path / "study_P01.mp4"
    v2 = tmp_path / "study_P02.mp4"
    v1.write_bytes(b"\x00")
    v2.write_bytes(b"\x00")
    transcripts_server._participants = [
        {"id": "P01", "video_paths": [str(v1)], "has_video": True},
        {"id": "P02", "video_paths": [str(v2)], "has_video": True},
    ]
    monkeypatch.setattr(
        transcripts_server.video, "probe_video_properties", lambda _p: _audio_props(1)
    )

    def cancelling_normalize(path, indices, **_kwargs):
        # Stand in for the user hitting Stop while the first file rewrites.
        transcripts_server._normalize_cancel_event.set()
        return True, "Audio normalized."

    monkeypatch.setattr(
        transcripts_server.video, "normalize_audio_inplace", cancelling_normalize
    )

    resp = tr_client.post(
        "/transcripts/api/normalize-audio", json={"participants": ["P01", "P02"]}
    )
    lines = _ndjson(resp)
    assert lines[0]["total"] == 2
    assert lines[1]["participant"] == "P01"
    assert lines[-2] == {"cancelled": True}
    assert lines[-1] == {"done": True}
    assert not any(ln.get("participant") == "P02" for ln in lines)
    # The slot is released even on the cancel path, so the next run can claim it.
    assert transcripts_server._normalize_busy is False


def test_normalize_audio_409_while_a_run_holds_the_slot(tr_client, monkeypatch):
    monkeypatch.setattr(transcripts_server, "_normalize_busy", True)
    resp = tr_client.post(
        "/transcripts/api/normalize-audio", json={"participants": ["P01"]}
    )
    assert resp.status_code == 409
    assert resp.get_json()["ok"] is False


def test_normalize_slot_is_freed_when_the_stream_is_never_consumed(
    tr_client, monkeypatch
):
    monkeypatch.setattr(transcripts_server, "_normalize_busy", False)
    monkeypatch.setattr(transcripts_server, "_normalize_owner", None)

    resp = tr_client.post(
        "/transcripts/api/normalize-audio", json={"participants": ["P01"]}
    )
    assert resp.status_code == 200
    # Tear the response down without ever pulling a line from the body.
    resp.close()

    assert transcripts_server._normalize_busy is False
    assert transcripts_server._normalize_owner is None


def test_normalize_slot_release_is_scoped_to_its_own_run(monkeypatch):
    monkeypatch.setattr(transcripts_server, "_normalize_busy", False)
    monkeypatch.setattr(transcripts_server, "_normalize_owner", None)

    first = transcripts_server._claim_normalize_slot()
    assert first is not None
    transcripts_server._release_normalize_slot(first)

    second = transcripts_server._claim_normalize_slot()
    assert second is not None and second != first

    # The first run's straggler release must not free the second run's slot.
    transcripts_server._release_normalize_slot(first)
    assert transcripts_server._normalize_busy is True
    assert transcripts_server._claim_normalize_slot() is None

    transcripts_server._release_normalize_slot(second)
    assert transcripts_server._normalize_busy is False


def test_normalize_audio_cancel_route_sets_the_event(tr_client):
    transcripts_server._normalize_cancel_event.clear()
    resp = tr_client.post("/transcripts/api/normalize-audio/cancel")
    assert resp.status_code == 200
    assert transcripts_server._normalize_cancel_event.is_set()
    transcripts_server._normalize_cancel_event.clear()


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


def test_friction_get_keeps_deterministic_scores_while_the_agent_runs(
    tr_client, _agent_state_clean
):
    """Regenerate pops the stored friction before the run, so a mid-run refetch
    used to answer with nothing at all — blanking the histogram, chips, transcript
    tinting and timeline band for the whole run. The scores are pure and owe
    nothing to the LLM, so they ride along with the generating flag."""
    _seed_friction_entry()  # segments present, no stored friction
    transcripts_server._orchestrator._in_flight["friction"].add("P01")

    resp = tr_client.get("/transcripts/api/agent/friction/P01")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is False and data["generating"] is True
    fr = data["friction"]
    assert fr["deterministic"] is True, (
        "the mid-run payload must be flagged so the poll does not mistake it "
        "for a completed run"
    )
    assert len(fr["segments"]) == 1 and "score" in fr["segments"][0]
    assert fr["moments"] == [], "only the LLM half waits on the agent"


def test_friction_generating_without_segments_carries_no_scores(
    tr_client, _agent_state_clean
):
    """Nothing to score → the generating response stays bare rather than
    shipping an empty friction object the client would render as a result."""
    transcripts_server._manifest["source_transcripts"]["P01"] = {"summary": "s"}
    transcripts_server._orchestrator._in_flight["friction"].add("P01")

    data = tr_client.get("/transcripts/api/agent/friction/P01").get_json()
    assert data["generating"] is True
    assert "friction" not in data


def test_summary_generating_never_carries_friction_scores(
    tr_client, _agent_state_clean
):
    """The deterministic ride-along is friction-only; it must not leak onto the
    other agents' generating responses."""
    # No stored summary, or the result branch would answer before the flag.
    _seed_friction_entry(summary="")
    transcripts_server._orchestrator._in_flight["summary"].add("P01")

    data = tr_client.get("/transcripts/api/agent/summary/P01").get_json()
    assert data["generating"] is True
    assert "friction" not in data


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
    citations start rides on the summary response, so the UI can seed the
    elapsed clock without waiting a poll interval. Only the *status* rides
    along; a settled load fetches the payload from the citations endpoint (see
    test_summary_get_omits_citations_payload)."""
    _seed_friction_entry()  # summary present, no citations yet
    _claim_slot("citations", "P01", 3000.0)
    resp = tr_client.get("/transcripts/api/agent/summary/P01")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["citations_generating"] is True
    assert data["citations_started_at"] == 3000.0


def test_summary_get_omits_citations_payload(tr_client, _agent_state_clean):
    """The summary response carries citations' status but not their payload —
    the generic agent GET returns only its own manifest field, and inlining the
    dependents would put the whole friction blob on the 1.2s summary poll.

    The frontend therefore re-fetches stored citations from their own endpoint
    after a settled summary load (_loadStoredCitations). Without that second
    call, every loadSummary renders the summary with its superscripts stripped.
    """
    cites = [
        {
            "sentence": "A session summary.",
            "refs": [{"start": 0.0, "end": 1.0, "segment_index": 0}],
        }
    ]
    _seed_friction_entry(citations=cites)

    summary = tr_client.get("/transcripts/api/agent/summary/P01").get_json()
    assert summary["ok"] is True
    assert summary["citations_generating"] is False
    assert "citations" not in summary

    stored = tr_client.get("/transcripts/api/agent/citations/P01").get_json()
    assert stored["ok"] is True
    assert stored["citations"] == cites


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
    # The stub stores nothing, so the chain advances — without this the seeded
    # entry (summary, no citations) would spawn the REAL citations agent, a
    # live Ollama generation leaking past the end of the test.
    monkeypatch.setattr(config, "OLLAMA_CITATIONS_ENABLED", False)
    monkeypatch.setattr(transcripts_server, "_persist_manifest", lambda: None)
    _seed_friction_entry(friction={"stale": False, "moments": []})

    done = threading.Event()

    def stub_run(snapshot, cancel_event, on_token=None):
        done.set()

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
    _join_orchestrator_threads(transcripts_server._orchestrator)


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


def test_report_regenerate_404_without_summary(tr_client, _agent_state_clean):
    """The report agent depends on summary; the regenerate route must refuse
    until one exists (the Overview tab offers the summary trigger first)."""
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [{"id": "P01:0", "start": 0.0, "end": 1.0, "text": "x"}],
    }
    resp = tr_client.post("/transcripts/api/agent/report/P01/regenerate")
    assert resp.status_code == 404


def test_report_regenerate_runs_when_disabled(
    tr_client, _agent_state_clean, monkeypatch
):
    """Manual-trigger contract: OLLAMA_REPORT_ENABLED=False keeps report out of
    the auto-chain, but the regenerate route runs it anyway (force=True). The
    orchestrator's snapshot must also carry the participant id for the report
    agent's injected getters."""
    monkeypatch.setattr(config, "OLLAMA_REPORT_ENABLED", False)
    # The seeded entry has no citations, and the post-run chain advance would
    # otherwise spawn the REAL citations agent (enabled by default) — a live
    # Ollama call in a daemon thread that outlives this test and corrupts the
    # shared orchestrator for whichever chain test runs next.
    monkeypatch.setattr(config, "OLLAMA_CITATIONS_ENABLED", False)
    monkeypatch.setattr(transcripts_server, "_persist_manifest", lambda: None)
    _seed_friction_entry()

    done = threading.Event()
    seen: dict = {}

    def stub_run(snapshot, cancel_event, on_token=None):
        seen["participant"] = snapshot.get("participant")
        done.set()

    rp_agent = thinking_agents.get_agent("report")
    assert rp_agent is not None
    monkeypatch.setitem(rp_agent, "run", stub_run)

    resp = tr_client.post("/transcripts/api/agent/report/P01/regenerate")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["generating"] is True
    assert done.wait(timeout=2)
    assert seen["participant"] == "P01"
    _join_orchestrator_threads(transcripts_server._orchestrator)


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


def test_next_eligible_honors_skip(tr_client, _agent_state_clean, monkeypatch):
    """skip lets the chain advance past an agent that stored nothing.

    Its manifest field is still empty, so without skip it is picked straight
    back and the chain spins on it instead of reaching its siblings.
    """
    monkeypatch.setattr(config, "OLLAMA_FRICTION_ENABLED", True)
    monkeypatch.setattr(config, "OLLAMA_CITATIONS_ENABLED", True)
    _seed_friction_entry()  # summary present; citations + friction both pending

    agent = transcripts_server._orchestrator.next_eligible("P01")
    assert agent is not None and agent["key"] == "citations"

    agent = transcripts_server._orchestrator.next_eligible("P01", skip={"citations"})
    assert agent is not None and agent["key"] == "friction"


def test_failed_citations_still_chains_to_friction(
    tr_client, _agent_state_clean, monkeypatch
):
    """A citations run that stores nothing must not strand friction.

    find_citations returns None when the model call fails, so nothing is
    committed. Friction depends only on the summary, so the chain has to carry
    on past the failure rather than stopping at the uncommitted step.
    """
    monkeypatch.setattr(config, "OLLAMA_FRICTION_ENABLED", True)
    monkeypatch.setattr(config, "OLLAMA_CITATIONS_ENABLED", True)
    monkeypatch.setattr(transcripts_server, "_persist_manifest", lambda: None)
    _seed_friction_entry()

    started: list[str] = []
    orch = transcripts_server._orchestrator
    real_run_agent = orch.run_agent

    def _spy(agent_key, participant, force=False, skip=None):
        started.append(agent_key)
        if agent_key == "citations":
            # Simulate the model-call failure: run for real, but with the agent's
            # run callable returning None so nothing is committed.
            monkeypatch.setitem(
                thinking_agents.get_agent("citations"),  # type: ignore[arg-type]
                "run",
                lambda entry, cancel, on_token=None: None,
            )
            real_run_agent(agent_key, participant, force=force, skip=skip)

    monkeypatch.setattr(orch, "run_agent", _spy)
    orch.run_chain("P01")
    _join_orchestrator_threads(orch)

    assert started[0] == "citations"
    assert "friction" in started, f"chain stalled after citations: {started}"
    assert "citations" not in transcripts_server._manifest["source_transcripts"]["P01"]


def test_raising_agent_still_chains_past_it(tr_client, _agent_state_clean, monkeypatch):
    """An agent that raises stores nothing, so the chain must route around it.

    Same stall as the uncommitted-result path — friction depends only on the
    summary and would otherwise never start after a citations exception.
    """
    monkeypatch.setattr(config, "OLLAMA_FRICTION_ENABLED", True)
    monkeypatch.setattr(config, "OLLAMA_CITATIONS_ENABLED", True)
    monkeypatch.setattr(transcripts_server, "_persist_manifest", lambda: None)
    _seed_friction_entry()

    def _boom(entry, cancel, on_token=None):
        raise RuntimeError("model exploded")

    monkeypatch.setitem(
        thinking_agents.get_agent("citations"),  # type: ignore[arg-type]
        "run",
        _boom,
    )
    # The spy passes friction through to the real run_agent; without a stub
    # that is a REAL Ollama generation on machines where Ollama runs. It
    # outlives the 2 s thread join, then commits into a later test's manifest
    # and squats the shared in-flight slot — the suite's flakiest interleaving.
    monkeypatch.setitem(
        thinking_agents.get_agent("friction"),  # type: ignore[arg-type]
        "run",
        lambda entry, cancel, on_token=None: None,
    )

    started: list[str] = []
    orch = transcripts_server._orchestrator
    real_run_agent = orch.run_agent

    def _spy(agent_key, participant, force=False, skip=None):
        started.append(agent_key)
        real_run_agent(agent_key, participant, force=force, skip=skip)

    monkeypatch.setattr(orch, "run_agent", _spy)
    orch.run_chain("P01")
    _join_orchestrator_threads(orch)

    assert started[0] == "citations"
    assert "friction" in started, f"chain stalled after the exception: {started}"


def test_chain_terminates_when_every_agent_stores_nothing(
    tr_client, _agent_state_clean, monkeypatch
):
    """Agents that all store nothing must not re-qualify each other forever.

    Their manifest fields stay empty and each leaves the in-flight set when its
    thread ends, so a skip set carrying only the *last* agent would let two of
    them take turns indefinitely — one live model call per lap. The skip set
    accumulates across the chain to make each agent run at most once.
    """
    monkeypatch.setattr(config, "OLLAMA_FRICTION_ENABLED", True)
    monkeypatch.setattr(config, "OLLAMA_CITATIONS_ENABLED", True)
    monkeypatch.setattr(transcripts_server, "_persist_manifest", lambda: None)
    _seed_friction_entry()

    # Citations returns immediately; friction lingers. That ordering is what
    # makes the regression deterministic rather than a race: citations' thread
    # reaches its `finally` and frees its in-flight slot while friction is still
    # working, so by the time friction advances the chain, citations looks
    # eligible again (empty field, not in flight) to a skip set that only
    # remembers friction.
    monkeypatch.setitem(
        thinking_agents.get_agent("citations"),  # type: ignore[arg-type]
        "run",
        lambda entry, cancel, on_token=None: None,
    )

    orch = transcripts_server._orchestrator
    saw_citations_retire: list[bool] = []

    def _wait_out_citations(entry, cancel, on_token=None):
        # Falls off the end, i.e. returns None: "ran, stored nothing".
        #
        # Hold friction open until citations has actually left the in-flight set
        # — that overlap *is* the regression window, so wait on the real signal
        # rather than sleeping a fixed span past it. run_chain fires from inside
        # citations' `try` (transcripts_server.py:2011) while its slot is still
        # claimed, and the `finally` (:2026) releases it a moment later, so this
        # normally returns in well under a millisecond.
        deadline = time.monotonic() + 2.0
        while time.monotonic() < deadline:
            if not orch.is_generating("P01", "citations"):
                saw_citations_retire.append(True)
                return
            time.sleep(0.001)

    monkeypatch.setitem(
        thinking_agents.get_agent("friction"),  # type: ignore[arg-type]
        "run",
        _wait_out_citations,
    )

    started: list[str] = []
    real_run_agent = orch.run_agent

    def _spy(agent_key, participant, force=False, skip=None):
        started.append(agent_key)
        # Circuit-breaker so a regression fails the assert below instead of
        # spinning the test suite forever.
        if len(started) > 6:
            return
        real_run_agent(agent_key, participant, force=force, skip=skip)

    monkeypatch.setattr(orch, "run_agent", _spy)
    orch.run_chain("P01")
    _join_orchestrator_threads(orch)

    # Assert the window held before reading anything into the chain order. If
    # friction never observed citations retire, the state this test exists to
    # cover never occurred and `started` proves nothing about the skip set.
    assert saw_citations_retire, (
        "friction ran but never saw citations leave the in-flight set, so the "
        "re-qualification window never opened; this assertion is the only thing "
        "standing between a real regression check and a vacuous one"
    )
    assert started == ["citations", "friction"]


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


def test_transcript_includes_words_when_present(tr_client):
    words = [
        {"start": 0.4, "end": 0.7, "text": "the"},
        {"start": 0.75, "end": 1.0, "text": "cat"},
    ]
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [
            {
                "id": "P01:0",
                "start": 0.4,
                "end": 1.0,
                "text": "the cat",
                "words": words,
            },
            {"id": "P01:1", "start": 2.0, "end": 3.0, "text": "old shape"},
        ],
    }
    resp = tr_client.get("/transcripts/api/transcript/P01")
    assert resp.status_code == 200
    segments = resp.get_json()["segments"]
    assert segments[0]["words"] == words
    # Segments from manifests predating word timing degrade to an empty list.
    assert segments[1]["words"] == []


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


def test_participants_includes_report_step_state(tr_client):
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [{"id": "P01:0", "start": 0.0, "end": 1.0, "text": "x"}],
        "summary": "A summary.",
        "report": {"text": "## Overview\nFine.", "model": "m"},
    }
    resp = tr_client.get("/transcripts/api/participants")
    assert resp.status_code == 200
    by_id = {p["id"]: p for p in resp.get_json()["participants"]}
    assert by_id["P01"]["agents"]["report"] == "done"


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

    monkeypatch.setattr(transcripts, "save_transcripts_manifest", _spy_save)

    # Block the summary agent so its in-flight slot stays claimed while we assert.
    started = threading.Event()
    release = threading.Event()

    def _blocking_summary(snapshot, cancel_event, on_token=None):
        started.set()
        release.wait(2.0)

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


def test_ollama_install_rejected_where_unsupported(tr_client, monkeypatch):
    import ollama_client

    monkeypatch.setattr(ollama_client, "can_install_managed", lambda: False)
    resp = tr_client.post("/transcripts/api/models/ollama/install", json={})
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


def test_ollama_install_short_circuits_when_installed(tr_client, monkeypatch):
    """The idempotent path: a working install must not re-download.

    Stubs is_working_install, which is what the route gates on — stubbing only
    is_installed leaves the real predicate running, and on a machine with no
    ollama on PATH that answers False, so the route falls through and spawns a
    genuine 145 MB install_managed() on a daemon thread.
    """
    import ollama_client

    monkeypatch.setattr(ollama_client, "can_install_managed", lambda: True)
    monkeypatch.setattr(ollama_client, "is_working_install", lambda: True)

    def _must_not_run(on_progress=None):
        raise AssertionError("install_managed must not run for a working install")

    monkeypatch.setattr(ollama_client, "install_managed", _must_not_run)
    monkeypatch.setattr(transcripts_server, "_ollama_install_status", None)

    resp = tr_client.post("/transcripts/api/models/ollama/install", json={})
    assert resp.status_code == 200
    assert resp.get_json()["already_installed"] is True


def test_ollama_install_status_before_any_install(tr_client, monkeypatch):
    monkeypatch.setattr(transcripts_server, "_ollama_install_status", None)
    resp = tr_client.get("/transcripts/api/models/ollama/install-status")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["found"] is False


def test_ollama_install_starts_and_reports_success(tr_client, monkeypatch):
    import time

    import ollama_client

    monkeypatch.setattr(ollama_client, "can_install_managed", lambda: True)
    # The route gates on is_working_install (is_installed only asks whether a
    # file exists, which a half-extracted tree also satisfies).
    monkeypatch.setattr(ollama_client, "is_working_install", lambda: False)

    def _fake_install(on_progress=None):
        if on_progress:
            on_progress({"status": "downloading Ollama", "total": 100, "completed": 40})
        return True

    monkeypatch.setattr(ollama_client, "install_managed", _fake_install)
    monkeypatch.setattr(transcripts_server, "_ollama_install_status", None)

    resp = tr_client.post("/transcripts/api/models/ollama/install", json={})
    assert resp.status_code == 200
    assert resp.get_json()["started"] is True

    status = {}
    for _ in range(100):
        status = tr_client.get(
            "/transcripts/api/models/ollama/install-status"
        ).get_json()
        if status.get("found") and status.get("done"):
            break
        time.sleep(0.02)
    assert status["found"] is True
    assert status["done"] is True
    assert status["succeeded"] is True
    assert status["status"] == "success"


def test_ollama_install_retries_a_broken_install(tr_client, monkeypatch):
    """A half-extracted tree satisfies is_installed(), and this route is the
    only way to repair one — so gating on it answered already_installed
    forever and stranded the user with no in-app recovery."""
    import ollama_client

    monkeypatch.setattr(ollama_client, "can_install_managed", lambda: True)
    monkeypatch.setattr(ollama_client, "is_installed", lambda: True)
    monkeypatch.setattr(ollama_client, "is_working_install", lambda: False)
    monkeypatch.setattr(ollama_client, "install_managed", lambda on_progress=None: True)
    monkeypatch.setattr(transcripts_server, "_ollama_install_status", None)

    body = tr_client.post("/transcripts/api/models/ollama/install", json={}).get_json()
    assert body.get("already_installed") is not True
    assert body["started"] is True


def test_ollama_install_second_post_attaches(tr_client, monkeypatch):
    import threading as threading_mod

    import ollama_client

    monkeypatch.setattr(ollama_client, "can_install_managed", lambda: True)
    # The route gates on is_working_install (is_installed only asks whether a
    # file exists, which a half-extracted tree also satisfies).
    monkeypatch.setattr(ollama_client, "is_working_install", lambda: False)
    release = threading_mod.Event()
    monkeypatch.setattr(
        ollama_client,
        "install_managed",
        lambda on_progress=None: release.wait(2.0),
    )
    monkeypatch.setattr(transcripts_server, "_ollama_install_status", None)
    try:
        first = tr_client.post("/transcripts/api/models/ollama/install", json={})
        assert first.get_json()["started"] is True
        second = tr_client.post("/transcripts/api/models/ollama/install", json={})
        assert second.get_json()["already_installing"] is True
    finally:
        release.set()


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
    assert {a["key"] for a in agents} == {"summary", "citations", "friction", "report"}
    # The configured summary model is present in the faked install list, and the
    # blank friction/report models resolve to it.
    assert all(a["installed"] for a in agents)


def _models_payload(monkeypatch, *, binary_present, server_answers):
    import ollama_client
    import server as server_mod

    monkeypatch.setattr(transcripts, "is_whisper_model_cached", lambda n=None: True)
    monkeypatch.setattr(ollama_client, "is_installed", lambda: binary_present)
    monkeypatch.setattr(
        ollama_client, "list_models", lambda: [] if server_answers else None
    )
    app = server_mod.build_combined_app()
    with app.test_client() as c:
        return c.get("/api/models").get_json()["ollama"]


def test_api_models_separates_not_installed_from_not_running(monkeypatch):
    """`available` alone cannot tell the two apart, and they need opposite
    advice — telling someone who never installed Ollama to "start it" was the
    bug this field exists to fix."""
    missing = _models_payload(monkeypatch, binary_present=False, server_answers=False)
    assert missing["installed"] is False
    assert missing["available"] is False

    stopped = _models_payload(monkeypatch, binary_present=True, server_answers=False)
    assert stopped["installed"] is True
    assert stopped["available"] is False

    running = _models_payload(monkeypatch, binary_present=True, server_answers=True)
    assert running["installed"] is True
    assert running["available"] is True


def test_api_models_ships_install_guidance(monkeypatch):
    """The platform install commands used to reach only the terminal, which a
    desktop-bundle user never sees."""
    payload = _models_payload(monkeypatch, binary_present=False, server_answers=False)
    hint = payload["install_hint"]
    assert hint and all(isinstance(line, str) for line in hint)
    assert any("ollama.com" in line for line in hint)


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


class TestMediaRouteFollowsTheInputDir:
    """``/media/<file>`` must resolve the input directory per request.

    The Start overlay sets the input directory through ``POST /api/dirs``, which
    moves ``config.INPUT_DIR`` without re-running ``_init_transcripts_state``.
    While the route served a snapshot taken at init, choosing a directory in the
    overlay left the page listing participants from the new one while every
    video 404'd — with the page's own ``?v=`` mtime proving the file was there.
    """

    def test_serves_from_the_directory_chosen_after_startup(
        self, tr_client, tmp_path, monkeypatch
    ):
        first = tmp_path / "before"
        first.mkdir()
        (first / "study_P01.mp4").write_bytes(b"first")
        monkeypatch.setattr(config, "INPUT_DIR", str(first))
        assert tr_client.get("/transcripts/media/study_P01.mp4").status_code == 200

        # The user picks a different folder in the Start overlay.
        second = tmp_path / "after"
        second.mkdir()
        (second / "study_P02.mp4").write_bytes(b"second")
        monkeypatch.setattr(config, "INPUT_DIR", str(second))

        assert tr_client.get("/transcripts/media/study_P02.mp4").status_code == 200
        # ...and the old directory is no longer served.
        assert tr_client.get("/transcripts/media/study_P01.mp4").status_code == 404

    def test_serves_after_the_very_first_directory_choice(
        self, tr_client, tmp_path, monkeypatch
    ):
        """The shape every desktop launch hits.

        Nothing configures an input directory before the Start overlay does, so
        ``get_effective_input_dir()`` falls back to ``Path.cwd()`` — which
        ``cli.main()`` has chdir'd to the folder holding the app. Snapshotting
        that gave the media route a directory that exists but holds no videos,
        so the first transcription's playback 404'd rather than reporting a
        missing configuration.
        """
        monkeypatch.setattr(config, "INPUT_DIR", "")
        chosen = tmp_path / "chosen"
        chosen.mkdir()
        (chosen / "study_P03.mp4").write_bytes(b"chosen")
        monkeypatch.setattr(config, "INPUT_DIR", str(chosen))
        assert tr_client.get("/transcripts/media/study_P03.mp4").status_code == 200
