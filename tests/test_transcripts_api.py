"""Smoke tests for Transcripts Flask API (prewarm + model status)."""

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
