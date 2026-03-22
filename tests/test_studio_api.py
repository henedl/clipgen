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


def test_api_viewer_400_when_no_artifacts(client):
    resp = client.post("/studio/api/viewer")
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False
    assert "No artifacts" in data["error"]
