import pytest

Flask = pytest.importorskip("flask").Flask
import server


@pytest.fixture
def client(monkeypatch):
    app = Flask(__name__)
    app.register_blueprint(server.studio_bp, url_prefix="/studio")

    # Default: no worksheet/context loaded (error state)
    monkeypatch.setattr(server, "_worksheet", None)
    monkeypatch.setattr(server, "_sheet_context", None)
    monkeypatch.setattr(server, "_generated_artifacts", [])
    monkeypatch.setattr(server, "_generated_reels", [])

    with app.test_client() as c:
        yield c


def test_api_sheet_500_when_no_context(client):
    resp = client.get("/studio/api/sheet")
    assert resp.status_code == 500
    data = resp.get_json()
    assert data["ok"] is False
    assert "No sheet loaded" in data["error"]


def test_api_generate_500_when_no_worksheet(client):
    resp = client.post("/studio/api/generate", json={"cells": ["P01.3"]})
    assert resp.status_code == 500
    data = resp.get_json()
    assert data["ok"] is False


def test_api_generate_400_when_no_cells(client, monkeypatch):
    monkeypatch.setattr(server, "_worksheet", object())
    resp = client.post("/studio/api/generate", json={"cells": []})
    assert resp.status_code == 400
    data = resp.get_json()
    assert "No cells" in data["error"]


def test_api_generate_400_for_invalid_format(client, monkeypatch):
    monkeypatch.setattr(server, "_worksheet", object())
    resp = client.post("/studio/api/generate", json={"cells": ["P01.3"], "format": "pdf"})
    assert resp.status_code == 400
    data = resp.get_json()
    assert "Invalid format" in data["error"]


def test_api_reel_400_when_no_cells(client, monkeypatch):
    monkeypatch.setattr(server, "_worksheet", object())
    resp = client.post("/studio/api/reel", json={"cells": []})
    assert resp.status_code == 400
    data = resp.get_json()
    assert "No cells" in data["error"]


def test_api_reel_highlights_duration_override(client, monkeypatch):
    """highlights_duration temporarily overrides config and is restored after."""
    import config

    monkeypatch.setattr(server, "_worksheet", object())
    original = config.HIGHLIGHTS_REEL_DURATION_SECONDS
    captured = {}

    def fake_generate_list(ws, mode, *, reel_input, skip_prompts):
        captured["duration"] = config.HIGHLIGHTS_REEL_DURATION_SECONDS
        return []

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)

    resp = client.post(
        "/studio/api/reel",
        json={"cells": ["highlights", "batch"], "highlights_duration": 120},
    )
    assert resp.status_code == 400  # no clips → 400
    assert captured["duration"] == 120
    assert config.HIGHLIGHTS_REEL_DURATION_SECONDS == original


def test_api_reel_highlights_duration_restored_on_error(client, monkeypatch):
    """Config is restored even if generate_list raises."""
    import config

    monkeypatch.setattr(server, "_worksheet", object())
    original = config.HIGHLIGHTS_REEL_DURATION_SECONDS

    def raise_generate_list(ws, mode, *, reel_input, skip_prompts):
        raise RuntimeError("boom")

    monkeypatch.setattr("spreadsheet.generate_list", raise_generate_list)

    resp = client.post(
        "/studio/api/reel",
        json={"cells": ["highlights"], "highlights_duration": 999},
    )
    assert resp.status_code == 500
    assert config.HIGHLIGHTS_REEL_DURATION_SECONDS == original


def test_api_viewer_400_when_no_artifacts(client):
    resp = client.post("/studio/api/viewer")
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False
    assert "No artifacts" in data["error"]
