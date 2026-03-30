"""Tests for Screenspace server API endpoints."""

import pytest

Flask = pytest.importorskip("flask").Flask

import config  # noqa: E402
import screenspace  # noqa: E402
import screenspace_server  # noqa: E402


@pytest.fixture
def client(tmp_path, monkeypatch):
    app = Flask(__name__)
    app.register_blueprint(screenspace_server.screenspace_bp, url_prefix="/screenspace")

    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    screenspace_server._manifest = {"regions": {}, "tasks": [], "events": []}
    screenspace_server._participants = [
        {"id": "P01", "video_path": "/tmp/test_P01.mp4", "has_video": False},
        {"id": "P02", "video_path": "/tmp/test_P02.mp4", "has_video": False},
    ]
    screenspace_server._output_dir = str(tmp_path)
    screenspace_server._worker = screenspace.ScreenspaceWorker()

    monkeypatch.setattr(
        screenspace,
        "save_screenspace_manifest",
        lambda r, t, e=None, stashes=None: tmp_path / "m.json",
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
        json={
            "name": "healthbar",
            "x": 100,
            "y": 20,
            "w": 300,
            "h": 30,
            "canvas_width": 1920,
            "canvas_height": 1080,
        },
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["region"]["source_width"] == 1920
    assert abs(data["region"]["x"] - 100 / 1920) < 1e-9

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
            "canvas_width": 1920,
            "canvas_height": 1080,
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


def test_create_region_normalizes_coords(client):
    resp = client.post(
        "/screenspace/api/regions",
        json={
            "name": "center",
            "x": 960,
            "y": 540,
            "w": 192,
            "h": 108,
            "canvas_width": 1920,
            "canvas_height": 1080,
        },
    )
    data = resp.get_json()
    assert data["ok"] is True
    r = data["region"]
    assert abs(r["x"] - 0.5) < 1e-9
    assert abs(r["y"] - 0.5) < 1e-9
    assert abs(r["w"] - 0.1) < 1e-9
    assert abs(r["h"] - 0.1) < 1e-9
    assert r["source_width"] == 1920
    assert r["source_height"] == 1080


def test_create_region_legacy_no_canvas_dims(client):
    resp = client.post(
        "/screenspace/api/regions",
        json={"name": "legacy", "x": 100, "y": 20, "w": 300, "h": 30},
    )
    data = resp.get_json()
    assert data["ok"] is True
    r = data["region"]
    assert r["x"] == 100
    assert r["y"] == 20
    assert r["w"] == 300
    assert r["h"] == 30
    assert "source_width" not in r


def test_denormalize_region():
    from screenspace_server import _denormalize_region

    region = {
        "x": 0.5,
        "y": 0.5,
        "w": 0.1,
        "h": 0.1,
        "source_width": 1920,
        "source_height": 1080,
    }
    px = _denormalize_region(region, 1920, 1080)
    assert px == {"x": 960, "y": 540, "w": 192, "h": 108}


def test_denormalize_legacy_region():
    from screenspace_server import _denormalize_region

    region = {"x": 100, "y": 20, "w": 300, "h": 30}
    px = _denormalize_region(region, 1280, 720)
    assert px == {"x": 100, "y": 20, "w": 300, "h": 30}


def test_denormalize_cross_resolution():
    from screenspace_server import _denormalize_region

    region = {
        "x": 0.5,
        "y": 0.5,
        "w": 0.1,
        "h": 0.1,
        "source_width": 1920,
        "source_height": 1080,
    }
    px = _denormalize_region(region, 1280, 720)
    assert px == {"x": 640, "y": 360, "w": 128, "h": 72}


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


def test_dismiss_queued_task(client):
    worker = screenspace_server._worker
    assert worker is not None
    task = screenspace.create_task(
        "color", "P01", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
    )
    worker.enqueue(task)
    screenspace_server._manifest["tasks"].append(task)
    resp = client.delete(f"/screenspace/api/tasks/{task['id']}?dismiss=true")
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True
    # Task removed from worker
    assert worker.get_task(task["id"]) is None
    # Task removed from manifest
    assert all(t["id"] != task["id"] for t in screenspace_server._manifest["tasks"])


def test_dismiss_completed_task(client):
    worker = screenspace_server._worker
    assert worker is not None
    task = screenspace.create_task(
        "color", "P01", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
    )
    worker.enqueue(task)
    # Simulate completed status
    with worker._lock:
        worker._tasks[task["id"]]["status"] = "completed"
    screenspace_server._manifest["tasks"].append(task)
    resp = client.delete(f"/screenspace/api/tasks/{task['id']}?dismiss=true")
    assert resp.status_code == 200
    assert worker.get_task(task["id"]) is None
    assert all(t["id"] != task["id"] for t in screenspace_server._manifest["tasks"])


def test_dismiss_nonexistent_task(client):
    resp = client.delete("/screenspace/api/tasks/ss_nonexist?dismiss=true")
    assert resp.status_code == 404


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


# ---- Events ----


def _sample_event(
    event_id="ev_test1234", excluded=False, participant="P01", task_id="ss_t1"
):
    return {
        "id": event_id,
        "source_video": "study_P01.mp4",
        "participant": participant,
        "detector": "change",
        "event_type": "change: hud",
        "time_in": 10.0,
        "time_out": 10.0,
        "confidence": 0.85,
        "metadata": {"magnitude": 0.12},
        "excluded": excluded,
        "task_id": task_id,
        "region": "hud",
    }


def test_events_list_empty(client):
    resp = client.get("/screenspace/api/events")
    data = resp.get_json()
    assert data["ok"] is True
    assert data["events"] == []


def test_events_list_all(client):
    screenspace_server._manifest["events"] = [
        _sample_event("ev_1"),
        _sample_event("ev_2", excluded=True),
    ]
    resp = client.get("/screenspace/api/events")
    data = resp.get_json()
    assert len(data["events"]) == 2


def test_events_filter_excluded_false(client):
    screenspace_server._manifest["events"] = [
        _sample_event("ev_1"),
        _sample_event("ev_2", excluded=True),
    ]
    resp = client.get("/screenspace/api/events?excluded=false")
    data = resp.get_json()
    assert len(data["events"]) == 1
    assert data["events"][0]["id"] == "ev_1"


def test_events_filter_by_participant(client):
    screenspace_server._manifest["events"] = [
        _sample_event("ev_1", participant="P01"),
        _sample_event("ev_2", participant="P02"),
    ]
    resp = client.get("/screenspace/api/events?participant=P02")
    data = resp.get_json()
    assert len(data["events"]) == 1
    assert data["events"][0]["participant"] == "P02"


def test_events_filter_by_task_id(client):
    screenspace_server._manifest["events"] = [
        _sample_event("ev_1", task_id="ss_a"),
        _sample_event("ev_2", task_id="ss_b"),
    ]
    resp = client.get("/screenspace/api/events?task_id=ss_a")
    data = resp.get_json()
    assert len(data["events"]) == 1
    assert data["events"][0]["task_id"] == "ss_a"


def test_event_exclude(client):
    screenspace_server._manifest["events"] = [_sample_event("ev_1")]
    resp = client.put("/screenspace/api/events/ev_1/exclude")
    data = resp.get_json()
    assert data["ok"] is True
    assert screenspace_server._manifest["events"][0]["excluded"] is True


def test_event_include(client):
    screenspace_server._manifest["events"] = [_sample_event("ev_1", excluded=True)]
    resp = client.put("/screenspace/api/events/ev_1/include")
    data = resp.get_json()
    assert data["ok"] is True
    assert screenspace_server._manifest["events"][0]["excluded"] is False


def test_event_exclude_not_found(client):
    screenspace_server._manifest["events"] = []
    resp = client.put("/screenspace/api/events/ev_missing/exclude")
    assert resp.status_code == 404


def test_events_bulk_exclude(client):
    screenspace_server._manifest["events"] = [
        _sample_event("ev_1"),
        _sample_event("ev_2"),
        _sample_event("ev_3"),
    ]
    resp = client.put(
        "/screenspace/api/events/bulk-exclude",
        json={"ids": ["ev_1", "ev_3"]},
    )
    data = resp.get_json()
    assert data["ok"] is True
    assert data["updated"] == 2
    assert screenspace_server._manifest["events"][0]["excluded"] is True
    assert screenspace_server._manifest["events"][1]["excluded"] is False
    assert screenspace_server._manifest["events"][2]["excluded"] is True


def test_events_bulk_include(client):
    screenspace_server._manifest["events"] = [
        _sample_event("ev_1", excluded=True),
        _sample_event("ev_2", excluded=True),
    ]
    resp = client.put(
        "/screenspace/api/events/bulk-include",
        json={"ids": ["ev_1", "ev_2"]},
    )
    data = resp.get_json()
    assert data["ok"] is True
    assert data["updated"] == 2
    assert screenspace_server._manifest["events"][0]["excluded"] is False


def test_events_bulk_exclude_empty_ids(client):
    resp = client.put(
        "/screenspace/api/events/bulk-exclude",
        json={"ids": []},
    )
    assert resp.status_code == 400
