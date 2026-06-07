"""Smoke tests for Transcripts Flask API (prewarm + model status)."""

import threading
from typing import cast

import pytest

Flask = pytest.importorskip("flask").Flask

import config  # noqa: E402
import thinking_agents  # noqa: E402
import transcripts  # noqa: E402
import transcripts_server  # noqa: E402
import viewer  # noqa: E402


@pytest.fixture
def tr_client(tmp_path, monkeypatch):
    app = Flask(__name__)
    app.register_blueprint(transcripts_server.transcripts_bp, url_prefix="/transcripts")

    transcripts_server._manifest = {
        "source_transcripts": {},
        "corrections": [],
        "marks": [],
    }
    transcripts_server._participants = [
        {
            "id": "P01",
            "video_path": str(tmp_path / "study_P01.mp4"),
            "has_video": False,
        }
    ]
    transcripts_server._worker = None
    transcripts_server._input_dir = str(tmp_path)
    transcripts_server._transcript_model_warming = False

    monkeypatch.setattr(viewer, "load_manifest_artifacts", lambda: [])

    with app.test_client() as c:
        yield c


def test_participants_includes_transcribe_prewarm(tr_client):
    resp = tr_client.get("/transcripts/api/participants")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["transcribe_prewarm"] in ("off", "queue_open", "page_load")


def test_prewarm_invalid_config_normalized_in_participants(tr_client, monkeypatch):
    monkeypatch.setattr(config, "TRANSCRIBE_PREWARM", "bogus")
    resp = tr_client.get("/transcripts/api/participants")
    assert resp.status_code == 200
    assert resp.get_json()["transcribe_prewarm"] == "queue_open"


def test_participants_includes_video_version(tr_client, tmp_path):
    """video_version (mtime_ns) drives the frontend's media/<file>?v=… cache-bust."""
    video_file = tmp_path / "study_P09.mp4"
    video_file.write_bytes(b"\x00data")
    transcripts_server._participants = [
        {"id": "P09", "video_path": str(video_file), "has_video": True},
        {"id": "P10", "video_path": str(tmp_path / "missing.mp4"), "has_video": False},
    ]
    resp = tr_client.get("/transcripts/api/participants")
    assert resp.status_code == 200
    by_id = {p["id"]: p for p in resp.get_json()["participants"]}
    assert by_id["P09"]["video_version"] == video_file.stat().st_mtime_ns
    assert by_id["P10"]["video_version"] is None


def test_participants_stale_artifact_detection(tr_client, monkeypatch):
    transcripts_server._participants = [
        {"id": "P01", "video_path": "study_P01.mp4", "has_video": True},
        {"id": "P02", "video_path": "study_P02.mp4", "has_video": True},
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


def test_transcribe_applies_per_participant_overrides(tr_client, monkeypatch):
    """POST /api/transcribe threads {model, language} overrides onto the task."""

    captured: list[dict] = []

    class _StubWorker:
        def enqueue(self, task):
            captured.append(task)

    transcripts_server._participants = [
        {"id": "P01", "video_path": "/tmp/P01.mp4", "has_video": True}
    ]
    transcripts_server._worker = cast("transcripts.TranscriptWorker", _StubWorker())

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


def test_transcribe_without_overrides_defaults_to_none(tr_client):
    """Missing overrides → task uses None (worker falls back to config defaults)."""

    captured: list[dict] = []

    class _StubWorker:
        def enqueue(self, task):
            captured.append(task)

    transcripts_server._participants = [
        {"id": "P01", "video_path": "/tmp/P01.mp4", "has_video": True}
    ]
    transcripts_server._worker = cast("transcripts.TranscriptWorker", _StubWorker())

    resp = tr_client.post(
        "/transcripts/api/transcribe",
        json={"participants": ["P01"], "force": True},
    )
    assert resp.status_code == 200
    assert len(captured) == 1
    assert captured[0]["model"] is None
    assert captured[0]["language"] is None


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
        with transcripts_server._pending_model_unloads_lock:
            for timer in transcripts_server._pending_model_unloads.values():
                timer.cancel()
            transcripts_server._pending_model_unloads.clear()

    _reset()
    yield
    _reset()


def test_summary_stop_sets_cancel_event(tr_client, _agent_state_clean):
    """POST /api/summary/<pid>/stop must set the registered event and remove
    the participant from the in-flight set so the UI flips to idle."""
    pid = "P01"
    evt = threading.Event()
    transcripts_server._orchestrator._in_flight["summary"].add(pid)
    transcripts_server._orchestrator._cancel_events["summary"][pid] = evt

    resp = tr_client.post(f"/transcripts/api/summary/{pid}/stop")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["running"] is False
    assert evt.is_set()
    assert pid not in transcripts_server._orchestrator._in_flight["summary"]


def test_citations_stop_sets_cancel_event(tr_client, _agent_state_clean):
    """POST /api/citations/<pid>/stop must set the registered event and
    remove the participant from the in-flight set."""
    pid = "P01"
    evt = threading.Event()
    transcripts_server._orchestrator._in_flight["citations"].add(pid)
    transcripts_server._orchestrator._cancel_events["citations"][pid] = evt

    resp = tr_client.post(f"/transcripts/api/citations/{pid}/stop")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["running"] is False
    assert evt.is_set()
    assert pid not in transcripts_server._orchestrator._in_flight["citations"]


def test_summary_stop_when_not_running_is_noop(tr_client, _agent_state_clean):
    """When nothing is in flight, /stop is a no-op returning running: False."""
    resp = tr_client.post("/transcripts/api/summary/P01/stop")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["running"] is False
    assert "P01" not in transcripts_server._orchestrator._cancel_events["summary"]


def test_citations_stop_when_not_running_is_noop(tr_client, _agent_state_clean):
    """When nothing is in flight, /stop is a no-op returning running: False."""
    resp = tr_client.post("/transcripts/api/citations/P01/stop")
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

    def blocking_run_first(snapshot, cancel_event):
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

    def blocking_run_second(snapshot, cancel_event):
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

    tr_client.post(f"/transcripts/api/summary/{pid}/stop")

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

    tr_client.post(f"/transcripts/api/citations/{pid}/stop")

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
        {"id": "P01", "video_path": str(video_path), "has_video": True}
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
        {"id": "P01", "video_path": str(video_path), "has_video": True}
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
        {"id": "P01", "video_path": str(video_path), "has_video": True}
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
        {"id": "P01", "video_path": str(v1), "has_video": True},
        {"id": "P02", "video_path": str(v2), "has_video": True},
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


def test_friction_get_404_when_absent_and_idle(tr_client, _agent_state_clean):
    _seed_friction_entry()
    resp = tr_client.get("/transcripts/api/friction/P01")
    assert resp.status_code == 404
    assert resp.get_json()["ok"] is False


def test_friction_get_returns_cached(tr_client, _agent_state_clean):
    fr = {"segments": [], "moments": [], "stats": {}, "stale": False, "model": "m"}
    _seed_friction_entry(friction=fr)
    resp = tr_client.get("/transcripts/api/friction/P01")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["friction"]["model"] == "m"
    assert data["friction"]["stale"] is False


def test_friction_get_generating_when_in_flight(tr_client, _agent_state_clean):
    _seed_friction_entry()
    transcripts_server._orchestrator._in_flight["friction"].add("P01")
    resp = tr_client.get("/transcripts/api/friction/P01")
    assert resp.status_code == 200
    assert resp.get_json()["generating"] is True


def test_friction_regenerate_404_without_summary(tr_client, _agent_state_clean):
    transcripts_server._manifest["source_transcripts"]["P01"] = {
        "segments": [{"id": "P01:0", "start": 0.0, "end": 1.0, "text": "x"}],
    }
    resp = tr_client.post("/transcripts/api/friction/P01/regenerate")
    assert resp.status_code == 404


def test_friction_regenerate_triggers(tr_client, _agent_state_clean, monkeypatch):
    monkeypatch.setattr(transcripts_server, "_persist_manifest", lambda: None)
    _seed_friction_entry(friction={"stale": False, "moments": []})

    done = threading.Event()

    def stub_run(snapshot, cancel_event):
        done.set()
        return None

    fr_agent = thinking_agents.get_agent("friction")
    assert fr_agent is not None
    monkeypatch.setitem(fr_agent, "run", stub_run)

    resp = tr_client.post("/transcripts/api/friction/P01/regenerate")
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
    resp = tr_client.post("/transcripts/api/friction/P01/regenerate")
    assert resp.status_code == 200
    assert resp.get_json()["generating"] is True
    # In-flight short-circuit must not clear the existing friction.
    assert "friction" in transcripts_server._manifest["source_transcripts"]["P01"]


def test_friction_stop_sets_cancel_event(tr_client, _agent_state_clean):
    pid = "P01"
    evt = threading.Event()
    transcripts_server._orchestrator._in_flight["friction"].add(pid)
    transcripts_server._orchestrator._cancel_events["friction"][pid] = evt
    resp = tr_client.post(f"/transcripts/api/friction/{pid}/stop")
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
    tr_client.post(f"/transcripts/api/friction/{pid}/stop")
    # Friction inherits the summary model (OLLAMA_FRICTION_MODEL is blank by
    # default), so the unload targets the resolved model.
    assert thinking_agents.friction_model() in transcripts_server._pending_model_unloads


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
    monkeypatch.setattr(transcripts_server, "_persist_manifest", lambda: None)
    _seed_friction_entry(friction={"stale": False, "moments": []})
    resp = tr_client.put(
        "/transcripts/api/transcript/P01/segment",
        json={"segment_id": "P01:0", "text": "completely new text"},
    )
    assert resp.status_code == 200
    entry = transcripts_server._manifest["source_transcripts"]["P01"]
    assert entry["friction"]["stale"] is True


def test_summary_put_marks_friction_stale(tr_client, _agent_state_clean, monkeypatch):
    monkeypatch.setattr(transcripts_server, "_persist_manifest", lambda: None)
    _seed_friction_entry(
        citations=[{"sentence": "s", "refs": []}],
        friction={"stale": False},
    )
    resp = tr_client.put(
        "/transcripts/api/summary/P01", json={"summary": "an edited summary"}
    )
    assert resp.status_code == 200
    entry = transcripts_server._manifest["source_transcripts"]["P01"]
    assert entry["friction"]["stale"] is True
    assert "citations" not in entry  # citations invalidated alongside friction


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

    def cit_stub(snapshot, cancel_event):
        return [{"sentence": "s2", "refs": []}]  # non-None → commits → cascade fires

    def fr_stub(snapshot, cancel_event):
        friction_ran.set()
        return None

    cit_agent = thinking_agents.get_agent("citations")
    fr_agent = thinking_agents.get_agent("friction")
    assert cit_agent is not None and fr_agent is not None
    monkeypatch.setitem(cit_agent, "run", cit_stub)
    monkeypatch.setitem(fr_agent, "run", fr_stub)

    resp = tr_client.post("/transcripts/api/citations/P01/regenerate")
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
    # manifest mutation.
    monkeypatch.setattr(
        transcripts_server._orchestrator, "run_chain", lambda *a, **k: None
    )
    _seed_friction_entry(
        citations=[{"sentence": "s", "refs": []}],
        friction={"stale": False, "moments": []},
    )

    resp = tr_client.post("/transcripts/api/summary/P01/regenerate")
    assert resp.status_code == 200
    entry = transcripts_server._manifest["source_transcripts"]["P01"]
    assert entry["friction"]["stale"] is True
    assert entry["summary"] == ""
    assert "citations" not in entry  # citations invalidated alongside friction
