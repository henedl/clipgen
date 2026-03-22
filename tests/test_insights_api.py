import pytest

Flask = pytest.importorskip("flask").Flask
import insights
import insights_server
import viewer


@pytest.fixture
def client(tmp_path, monkeypatch):
    app = Flask(__name__)
    app.register_blueprint(insights_server.insights_bp, url_prefix="/insights")

    # Set module state directly — no disk reads or sprite generation
    insights_server._insights_data = {"meta": {}, "insights": []}
    insights_server._artifacts = []
    insights_server._output_dir = str(tmp_path)

    # No-op manifest save to avoid disk I/O (roundtrip covered in test_insights.py)
    monkeypatch.setattr(insights, "save_insights_manifest", lambda meta, ins: None)

    with app.test_client() as c:
        yield c


def test_list_insights_empty(client):
    resp = client.get("/insights/api/insights")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["insights"] == []


def test_create_insight_via_api(client):
    resp = client.post(
        "/insights/api/insights",
        json={"title": "Navigation bug", "severity": "High"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    ins = data["insight"]
    assert ins["id"].startswith("ins_")
    assert ins["title"] == "Navigation bug"
    assert ins["severity"] == "High"
    assert ins["status"] == "draft"

    # Verify it appears in the list
    list_resp = client.get("/insights/api/insights")
    assert len(list_resp.get_json()["insights"]) == 1


def test_get_single_insight(client):
    create_resp = client.post("/insights/api/insights", json={"title": "A"})
    insight_id = create_resp.get_json()["insight"]["id"]

    resp = client.get(f"/insights/api/insights/{insight_id}")
    assert resp.status_code == 200
    assert resp.get_json()["insight"]["id"] == insight_id


def test_get_nonexistent_returns_404(client):
    resp = client.get("/insights/api/insights/ins_nonexistent")
    assert resp.status_code == 404
    assert resp.get_json()["ok"] is False


def test_update_insight_via_api(client):
    create_resp = client.post("/insights/api/insights", json={"title": "Original"})
    insight_id = create_resp.get_json()["insight"]["id"]

    resp = client.put(
        f"/insights/api/insights/{insight_id}",
        json={"title": "Updated", "severity": "Critical"},
    )
    assert resp.status_code == 200
    updated = resp.get_json()["insight"]
    assert updated["title"] == "Updated"
    assert updated["severity"] == "Critical"
    assert "updatedAt" in updated


def test_update_nonexistent_returns_404(client):
    resp = client.put("/insights/api/insights/ins_missing", json={"title": "X"})
    assert resp.status_code == 404


def test_delete_insight_via_api(client):
    create_resp = client.post("/insights/api/insights", json={"title": "Doomed"})
    insight_id = create_resp.get_json()["insight"]["id"]

    resp = client.delete(f"/insights/api/insights/{insight_id}")
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True

    list_resp = client.get("/insights/api/insights")
    assert list_resp.get_json()["insights"] == []


def test_delete_nonexistent_returns_404(client):
    resp = client.delete("/insights/api/insights/ins_missing")
    assert resp.status_code == 404


def test_sprite_404_for_missing_clip(client):
    resp = client.get("/insights/api/sprites/nonexistent.mp4")
    assert resp.status_code == 404


def test_artifacts_include_sprite_metadata(client, monkeypatch):
    test_artifacts = [
        {
            "id": "a1",
            "type": "clip",
            "file": "test.mp4",
            "start": 0,
            "end": 60,
            "study": "s",
            "participant": "P01",
            "category": "",
            "description": "",
        }
    ]
    monkeypatch.setattr(viewer, "load_manifest_artifacts", lambda: test_artifacts)
    resp = client.get("/insights/api/artifacts")
    data = resp.get_json()
    art = data["artifacts"][0]
    assert "spriteData" in art
    assert art["spriteData"]["frameCount"] > 0
    assert art["spriteData"]["cols"] > 0


def test_sprites_generate_endpoint_removed(client):
    resp = client.post("/insights/api/sprites/generate")
    assert resp.status_code in (404, 405)
