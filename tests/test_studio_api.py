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


def test_api_thumbnail_500_when_no_context(client):
    resp = client.get("/studio/api/thumbnail/P01/0")
    assert resp.status_code == 500
    data = resp.get_json()
    assert data["ok"] is False
    assert "No sheet loaded" in data["error"]


def test_api_thumbnail_returns_jpeg(client, monkeypatch, tmp_path):
    import types

    import video

    fake_jpeg = b"\xff\xd8\xff\xe0fake-jpeg-data"
    dummy_video = tmp_path / "study_P01.mp4"
    dummy_video.write_bytes(b"not a real video")

    ctx = types.SimpleNamespace(
        header_row=["ID", "P01"],
        id_cell=types.SimpleNamespace(col=1),
        num_participants=1,
        study_name="study",
        filename_row_idx=None,
        sheet_data=[],
    )
    monkeypatch.setattr(server, "_sheet_context", ctx)
    monkeypatch.setattr(server, "_thumbnail_cache", {})
    monkeypatch.setattr("utils.resolve_input_path", lambda name: dummy_video)
    monkeypatch.setattr(video, "extract_thumbnail_bytes", lambda *a, **kw: fake_jpeg)

    resp = client.get("/studio/api/thumbnail/P01/10")
    assert resp.status_code == 200
    assert resp.content_type == "image/jpeg"
    assert resp.data == fake_jpeg


def test_api_thumbnail_caches(client, monkeypatch, tmp_path):
    import types

    import video

    call_count = [0]
    fake_jpeg = b"\xff\xd8\xff\xe0cached"
    dummy_video = tmp_path / "study_P01.mp4"
    dummy_video.write_bytes(b"x")

    ctx = types.SimpleNamespace(
        header_row=["ID", "P01"],
        id_cell=types.SimpleNamespace(col=1),
        num_participants=1,
        study_name="study",
        filename_row_idx=None,
        sheet_data=[],
    )
    monkeypatch.setattr(server, "_sheet_context", ctx)
    monkeypatch.setattr(server, "_thumbnail_cache", {})
    monkeypatch.setattr("utils.resolve_input_path", lambda name: dummy_video)

    def counting_extract(*a, **kw):
        call_count[0] += 1
        return fake_jpeg

    monkeypatch.setattr(video, "extract_thumbnail_bytes", counting_extract)

    resp1 = client.get("/studio/api/thumbnail/P01/5")
    resp2 = client.get("/studio/api/thumbnail/P01/5")
    assert resp1.status_code == 200
    assert resp2.status_code == 200
    assert call_count[0] == 1


def test_api_manifest_get_returns_artifacts(client, monkeypatch):
    import viewer

    fake_artifacts = [{"id": "a5c2s0", "type": "clip", "participant": "P01", "cellRow": 5}]
    monkeypatch.setattr(viewer, "load_manifest_artifacts", lambda: fake_artifacts)
    monkeypatch.setattr(viewer, "load_manifest_reels", lambda: [])
    resp = client.get("/studio/api/manifest")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert len(data["artifacts"]) == 1
    assert data["artifacts"][0]["id"] == "a5c2s0"
    assert data["reels"] == []


def test_api_manifest_get_empty(client, monkeypatch):
    import viewer

    monkeypatch.setattr(viewer, "load_manifest_artifacts", lambda: [])
    monkeypatch.setattr(viewer, "load_manifest_reels", lambda: [])
    resp = client.get("/studio/api/manifest")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["artifacts"] == []
    assert data["reels"] == []


def test_api_manifest_post_still_works(client, monkeypatch):
    monkeypatch.setattr(server, "_generated_artifacts", [])
    monkeypatch.setattr(server, "_generated_reels", [])
    resp = client.post("/studio/api/manifest")
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False
    assert "No artifacts" in data["error"]


def test_api_viewer_400_when_no_artifacts(client):
    resp = client.post("/studio/api/viewer")
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False
    assert "No artifacts" in data["error"]
