"""Tests for Screenspace server API endpoints."""

import json

import pytest

Flask = pytest.importorskip("flask").Flask

import config
import screenspace
import screenspace_server


@pytest.fixture
def client(tmp_path, monkeypatch):
    app = Flask(__name__)
    app.register_blueprint(
        screenspace_server.screenspace_bp, url_prefix="/screenspace"
    )

    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    screenspace_server._manifest = {"regions": {}, "tasks": []}
    screenspace_server._participants = [
        {"id": "P01", "video_path": "/tmp/test_P01.mp4", "has_video": False},
        {"id": "P02", "video_path": "/tmp/test_P02.mp4", "has_video": False},
    ]
    screenspace_server._output_dir = str(tmp_path)
    screenspace_server._worker = screenspace.ScreenspaceWorker()

    monkeypatch.setattr(
        screenspace, "save_screenspace_manifest", lambda r, t: tmp_path / "m.json"
    )

    with app.test_client() as c:
        yield c


# ---- Participants ----


def test_list_participants(client):
    resp = client.get("/screenspace/api/participants")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert len(data["participants"]) == 2
    assert data["participants"][0]["id"] == "P01"


# ---- Regions ----


def test_list_regions_empty(client):
    resp = client.get("/screenspace/api/regions")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["regions"] == {}


def test_create_region(client):
    resp = client.post(
        "/screenspace/api/regions",
        json={"name": "healthbar", "x": 100, "y": 20, "w": 300, "h": 30},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["region"]["x"] == 100

    list_resp = client.get("/screenspace/api/regions")
    assert "healthbar" in list_resp.get_json()["regions"]


def test_create_region_with_description(client):
    resp = client.post(
        "/screenspace/api/regions",
        json={
            "name": "score",
            "x": 0,
            "y": 0,
            "w": 50,
            "h": 20,
            "description": "Score display",
        },
    )
    assert resp.status_code == 200
    regions = client.get("/screenspace/api/regions").get_json()["regions"]
    assert regions["score"]["description"] == "Score display"


def test_create_region_missing_name(client):
    resp = client.post(
        "/screenspace/api/regions",
        json={"x": 0, "y": 0, "w": 10, "h": 10},
    )
    assert resp.status_code == 400


def test_create_region_missing_coords(client):
    resp = client.post(
        "/screenspace/api/regions",
        json={"name": "test", "x": 0, "y": 0},
    )
    assert resp.status_code == 400


def test_delete_region(client):
    client.post(
        "/screenspace/api/regions",
        json={"name": "temp", "x": 0, "y": 0, "w": 10, "h": 10},
    )
    resp = client.delete("/screenspace/api/regions/temp")
    assert resp.status_code == 200

    list_resp = client.get("/screenspace/api/regions")
    assert "temp" not in list_resp.get_json()["regions"]


def test_delete_nonexistent_region(client):
    resp = client.delete("/screenspace/api/regions/nope")
    assert resp.status_code == 404


# ---- Tasks ----


def test_list_tasks_empty(client):
    resp = client.get("/screenspace/api/tasks")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["tasks"] == []


def test_create_task_missing_region(client):
    resp = client.post(
        "/screenspace/api/tasks",
        json={"type": "color", "participant": "P01"},
    )
    assert resp.status_code == 400
    assert "region" in resp.get_json()["error"].lower()


def test_create_task_unknown_region(client):
    resp = client.post(
        "/screenspace/api/tasks",
        json={"type": "color", "participant": "P01", "region": "nonexistent"},
    )
    assert resp.status_code == 400


def test_create_task_invalid_type(client):
    resp = client.post(
        "/screenspace/api/tasks",
        json={"type": "invalid", "participant": "P01", "region": "r"},
    )
    assert resp.status_code == 400


def test_create_task_no_video(client):
    # P01 has_video=False
    screenspace_server._manifest["regions"]["r"] = {
        "x": 0,
        "y": 0,
        "w": 10,
        "h": 10,
    }
    resp = client.post(
        "/screenspace/api/tasks",
        json={"type": "color", "participant": "P01", "region": "r"},
    )
    assert resp.status_code == 400
    assert "video" in resp.get_json()["error"].lower()


def test_create_task_numbers_type_accepted(client):
    """'numbers' passes type validation (fails at video check, not type check)."""
    screenspace_server._manifest["regions"]["r"] = {
        "x": 0,
        "y": 0,
        "w": 10,
        "h": 10,
    }
    resp = client.post(
        "/screenspace/api/tasks",
        json={"type": "numbers", "participant": "P01", "region": "r"},
    )
    data = resp.get_json()
    assert resp.status_code == 400
    assert "video" in data["error"].lower()


def test_get_task_not_found(client):
    resp = client.get("/screenspace/api/tasks/ss_nonexist")
    assert resp.status_code == 404


def test_cancel_task_not_found(client):
    resp = client.delete("/screenspace/api/tasks/ss_nonexist")
    assert resp.status_code == 400


def test_task_results_not_found(client):
    resp = client.get("/screenspace/api/tasks/ss_nonexist/results")
    assert resp.status_code == 404


def test_reorder_missing_body(client):
    resp = client.put("/screenspace/api/tasks/reorder", json={})
    assert resp.status_code == 400


def test_reorder_valid(client):
    resp = client.put(
        "/screenspace/api/tasks/reorder",
        json={"task_ids": ["ss_a", "ss_b"]},
    )
    assert resp.status_code == 200


# ---- Video ----


def test_video_frame_no_video(client):
    resp = client.get("/screenspace/api/video/frame/P99/0.0")
    assert resp.status_code == 404


def test_video_info_no_video(client):
    resp = client.get("/screenspace/api/video/info/P99")
    assert resp.status_code == 404


def test_video_info_participant_without_video(client):
    resp = client.get("/screenspace/api/video/info/P01")
    assert resp.status_code == 404


# ---- Static serving ----


def test_serve_index(client):
    resp = client.get("/screenspace/")
    assert resp.status_code == 200
    assert b"Screenspace" in resp.data
