"""Smoke tests for Transcripts Flask API (prewarm + model status)."""

import threading
from typing import cast

import pytest

Flask = pytest.importorskip("flask").Flask

import config  # noqa: E402
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
        for key in transcripts_server._agent_in_flight:
            transcripts_server._agent_in_flight[key].clear()
        for key in transcripts_server._agent_cancel_events:
            transcripts_server._agent_cancel_events[key].clear()
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
    transcripts_server._agent_in_flight["summary"].add(pid)
    transcripts_server._agent_cancel_events["summary"][pid] = evt

    resp = tr_client.post(f"/transcripts/api/summary/{pid}/stop")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["running"] is False
    assert evt.is_set()
    assert pid not in transcripts_server._agent_in_flight["summary"]


def test_citations_stop_sets_cancel_event(tr_client, _agent_state_clean):
    """POST /api/citations/<pid>/stop must set the registered event and
    remove the participant from the in-flight set."""
    pid = "P01"
    evt = threading.Event()
    transcripts_server._agent_in_flight["citations"].add(pid)
    transcripts_server._agent_cancel_events["citations"][pid] = evt

    resp = tr_client.post(f"/transcripts/api/citations/{pid}/stop")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["running"] is False
    assert evt.is_set()
    assert pid not in transcripts_server._agent_in_flight["citations"]


def test_summary_stop_when_not_running_is_noop(tr_client, _agent_state_clean):
    """When nothing is in flight, /stop is a no-op returning running: False."""
    resp = tr_client.post("/transcripts/api/summary/P01/stop")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["running"] is False
    assert "P01" not in transcripts_server._agent_cancel_events["summary"]


def test_citations_stop_when_not_running_is_noop(tr_client, _agent_state_clean):
    """When nothing is in flight, /stop is a no-op returning running: False."""
    resp = tr_client.post("/transcripts/api/citations/P01/stop")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["running"] is False
    assert "P01" not in transcripts_server._agent_cancel_events["citations"]


# ---- Model unload scheduling ----


def test_summary_stop_schedules_model_unload(
    tr_client, _agent_state_clean, monkeypatch
):
    """A successful summary /stop must schedule an unload timer for the
    summary model. We force a long delay so the timer is observable."""
    monkeypatch.setattr(config, "OLLAMA_UNLOAD_DELAY_SECONDS", 30.0)
    pid = "P01"
    evt = threading.Event()
    transcripts_server._agent_in_flight["summary"].add(pid)
    transcripts_server._agent_cancel_events["summary"][pid] = evt

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
    transcripts_server._agent_in_flight["citations"].add(pid)
    transcripts_server._agent_cancel_events["citations"][pid] = evt

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
