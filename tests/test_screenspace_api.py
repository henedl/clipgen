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
        json={
            "name": "temp",
            "x": 0,
            "y": 0,
            "w": 10,
            "h": 10,
            "canvas_width": 1920,
            "canvas_height": 1080,
        },
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


def test_create_region_rejects_missing_canvas_dims(client):
    resp = client.post(
        "/screenspace/api/regions",
        json={"name": "no_canvas", "x": 100, "y": 20, "w": 300, "h": 30},
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False
    assert "canvas_width" in data["error"]


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


@pytest.mark.parametrize("task_type", ["template", "flow", "scene"])
def test_create_task_new_types_accepted(client, task_type):
    """New phase-4 types pass type validation (fail at video, not type)."""
    screenspace_server._manifest["regions"]["r"] = {
        "x": 0,
        "y": 0,
        "w": 10,
        "h": 10,
    }
    resp = client.post(
        "/screenspace/api/tasks",
        json={"type": task_type, "participant": "P01", "region": "r"},
    )
    data = resp.get_json()
    assert resp.status_code == 400
    assert "video" in data["error"].lower()


def test_create_template_task_no_region_with_upload(client):
    """Template task with uploaded image skips region validation."""
    import base64

    # Minimal 1x1 red PNG
    png_b64 = base64.b64encode(
        b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01"
        b"\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00"
        b"\x00\x00\x0cIDATx\x9cc\xf8\x0f\x00\x00\x01\x01\x00"
        b"\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82"
    ).decode()
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "template",
            "participant": "P01",
            "parameters": {"template_image_data": png_b64},
        },
    )
    data = resp.get_json()
    # Should fail at video check, NOT at "region is required"
    assert resp.status_code == 400
    assert "region" not in data["error"].lower()
    assert "video" in data["error"].lower()


def test_api_preview_template_capture_includes_template_panel(
    client, monkeypatch
) -> None:
    """GET template preview with ref+region yields a wider composite than without ref."""
    import cv2
    import numpy as np
    import video

    _enable_video_task_setup(monkeypatch, "P01")
    frame = np.zeros((240, 320, 3), dtype=np.uint8)
    frame[50:110, 80:180] = (10, 180, 90)
    monkeypatch.setattr(
        video, "extract_frame_at_timestamp", lambda _path, _ts: frame.copy()
    )

    region = "0.25,0.2083333333,0.3125,0.25"
    r_noref = client.get(
        f"/screenspace/api/preview/P01/0.500?tool=template&region={region}"
    )
    r_with = client.get(
        f"/screenspace/api/preview/P01/0.500?tool=template&region={region}&ref=0.0"
    )
    assert r_noref.status_code == 200
    assert r_with.status_code == 200
    a = cv2.imdecode(np.frombuffer(r_noref.data, np.uint8), cv2.IMREAD_COLOR)
    b = cv2.imdecode(np.frombuffer(r_with.data, np.uint8), cv2.IMREAD_COLOR)
    assert a is not None and b is not None
    assert b.shape[1] > a.shape[1]


def test_api_preview_template_post_upload(client, monkeypatch) -> None:
    """POST template preview with template_image_data includes the template patch."""
    import base64

    import cv2
    import numpy as np
    import video

    _enable_video_task_setup(monkeypatch, "P01")
    frame = np.zeros((240, 320, 3), dtype=np.uint8)
    monkeypatch.setattr(
        video, "extract_frame_at_timestamp", lambda _path, _ts: frame.copy()
    )

    tile = np.zeros((40, 50, 3), dtype=np.uint8)
    tile[:] = (40, 50, 60)
    ok, buf = cv2.imencode(".png", tile)
    assert ok
    b64 = base64.b64encode(buf.tobytes()).decode()

    resp = client.post(
        "/screenspace/api/preview/P01/0.0?tool=template",
        json={"template_image_data": b64},
    )
    assert resp.status_code == 200
    assert resp.mimetype == "image/png"
    out = cv2.imdecode(np.frombuffer(resp.data, np.uint8), cv2.IMREAD_COLOR)
    assert out is not None
    assert out.shape[1] > 400


def test_api_preview_template_post_invalid_image(client, monkeypatch) -> None:
    _enable_video_task_setup(monkeypatch, "P01")
    import numpy as np
    import video

    monkeypatch.setattr(
        video,
        "extract_frame_at_timestamp",
        lambda _path, _ts: np.zeros((240, 320, 3), dtype=np.uint8),
    )
    resp = client.post(
        "/screenspace/api/preview/P01/0.0?tool=template",
        json={"template_image_data": "qqqq"},
    )
    assert resp.status_code == 400


def test_api_preview_layers_catalog(client) -> None:
    """The catalog endpoint returns each tool's overlay layers."""
    resp = client.get("/screenspace/api/preview/layers")
    assert resp.status_code == 200
    payload = resp.get_json()
    assert payload["ok"] is True
    layers = payload["layers"]
    # Tools with no overlay-eligible layers must not appear.
    assert "timelapse" not in layers
    assert "inactivity" not in layers
    # Multi-layer tools list each layer with id/label/scope.
    change_layers = layers["change"]
    assert {layer["id"] for layer in change_layers} == {"gray_blur", "abs_diff", "mask"}
    assert all(layer["scope"] == "region" for layer in change_layers)
    # Template's match heatmap is frame-scoped.
    template_layers = layers["template"]
    assert any(
        layer["id"] == "match_heatmap" and layer["scope"] == "frame"
        for layer in template_layers
    )


def test_api_preview_layer_returns_native_resolution_png(client, monkeypatch) -> None:
    """layer= returns a region-sized PNG (no max-width cap)."""
    import cv2
    import numpy as np
    import video

    _enable_video_task_setup(monkeypatch, "P01")
    monkeypatch.setattr(
        video,
        "extract_frame_at_timestamp",
        lambda _path, _ts: np.full((240, 320, 3), 50, dtype=np.uint8),
    )

    region = "0.25,0.2083333333,0.3125,0.25"  # ~ 100x60 px
    resp = client.get(
        f"/screenspace/api/preview/P01/0.500?tool=color&region={region}&layer=region"
    )
    assert resp.status_code == 200
    assert resp.mimetype == "image/png"
    img = cv2.imdecode(np.frombuffer(resp.data, np.uint8), cv2.IMREAD_COLOR)
    assert img is not None
    # Native size matches the requested region in pixels.
    assert img.shape[:2] == (60, 100)


def test_api_preview_layer_invalid_returns_400(client, monkeypatch) -> None:
    """Asking for a layer the tool doesn't expose returns 400."""
    import numpy as np
    import video

    _enable_video_task_setup(monkeypatch, "P01")
    monkeypatch.setattr(
        video,
        "extract_frame_at_timestamp",
        lambda _path, _ts: np.zeros((240, 320, 3), dtype=np.uint8),
    )
    region = "0.25,0.2083333333,0.3125,0.25"
    # timelapse is intentionally excluded from OVERLAY_LAYERS.
    resp = client.get(
        f"/screenspace/api/preview/P01/0.500?tool=timelapse&region={region}&layer=anything"
    )
    assert resp.status_code == 400
    # bogus layer for a valid tool also rejected.
    resp = client.get(
        f"/screenspace/api/preview/P01/0.500?tool=color&region={region}&layer=bogus"
    )
    assert resp.status_code == 400


def test_create_non_template_task_still_requires_region(client):
    """Non-template tasks still require a region."""
    resp = client.post(
        "/screenspace/api/tasks",
        json={"type": "change", "participant": "P01"},
    )
    assert resp.status_code == 400
    assert "region" in resp.get_json()["error"].lower()


def test_create_similarity_task_invalid_reference_timestamp(client, monkeypatch):
    _create_region(client, "r")
    _enable_video_task_setup(monkeypatch, "P01")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "similarity",
            "participant": "P01",
            "region": "r",
            "parameters": {"reference_timestamp": "not-a-number"},
        },
    )
    assert resp.status_code == 400
    assert "reference_timestamp" in resp.get_json()["error"]


def test_create_template_task_invalid_reference_timestamp(client, monkeypatch):
    _create_region(client, "r")
    _enable_video_task_setup(monkeypatch, "P01")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "template",
            "participant": "P01",
            "region": "r",
            "parameters": {"reference_timestamp": "not-a-number"},
        },
    )
    assert resp.status_code == 400
    assert "reference_timestamp" in resp.get_json()["error"]


def test_create_scene_task_invalid_scene_references(client, monkeypatch):
    _create_region(client, "r")
    _enable_video_task_setup(monkeypatch, "P01")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "scene",
            "participant": "P01",
            "region": "r",
            "parameters": {
                "scene_references": [{"name": "Main menu", "timestamp": "x"}]
            },
        },
    )
    assert resp.status_code == 400
    assert "scene_references[0].timestamp" in resp.get_json()["error"]


def test_create_multitool_task_invalid_step_reference_timestamp(client, monkeypatch):
    _create_region(client, "healthbar")
    _create_region(client, "statusbar", x=0, y=0, w=100, h=20)
    _enable_video_task_setup(monkeypatch, "P01")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "multitool",
            "participant": "P01",
            "region": "healthbar",
            "parameters": {
                "steps": [
                    {
                        "type": "similarity",
                        "region": "healthbar",
                        "reference_timestamp": "bad",
                    },
                    {"type": "change", "region": "statusbar"},
                ]
            },
        },
    )
    assert resp.status_code == 400
    assert "Step 0: reference_timestamp" in resp.get_json()["error"]


def test_create_task_persists_manifest(client, monkeypatch):
    _create_region(client, "r")
    _enable_video_task_setup(monkeypatch, "P01")
    calls = _install_persist_spy(monkeypatch)
    resp = client.post(
        "/screenspace/api/tasks",
        json={"type": "color", "participant": "P01", "region": "r"},
    )
    assert resp.status_code == 200
    assert calls == [{"drain_events": False}]


def test_get_task_not_found(client):
    resp = client.get("/screenspace/api/tasks/ss_nonexist")
    assert resp.status_code == 404


def test_cancel_task_not_found(client):
    resp = client.delete("/screenspace/api/tasks/ss_nonexist")
    assert resp.status_code == 400


def test_cancel_task_persists_manifest(client, monkeypatch):
    worker = screenspace_server._worker
    assert worker is not None
    task = screenspace.create_task(
        "color", "P01", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
    )
    worker.enqueue(task)
    calls = _install_persist_spy(monkeypatch)
    resp = client.delete(f"/screenspace/api/tasks/{task['id']}")
    assert resp.status_code == 200
    assert calls == [{"drain_events": False}]


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
    worker = screenspace_server._worker
    assert worker is not None
    task = screenspace.create_task(
        "color", "P01", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
    )
    worker.enqueue(task)
    resp = client.put(
        "/screenspace/api/tasks/reorder",
        json={"task_ids": [task["id"]]},
    )
    assert resp.status_code == 200


def test_reorder_persists_manifest(client, monkeypatch):
    worker = screenspace_server._worker
    assert worker is not None
    task = screenspace.create_task(
        "color", "P01", "s.mp4", "/v.mp4", "r", {"x": 0, "y": 0, "w": 1, "h": 1}
    )
    worker.enqueue(task)
    calls = _install_persist_spy(monkeypatch)
    resp = client.put(
        "/screenspace/api/tasks/reorder",
        json={"task_ids": [task["id"]]},
    )
    assert resp.status_code == 200
    assert calls == [{"drain_events": False}]


def test_pause_persists_manifest(client, monkeypatch):
    calls = _install_persist_spy(monkeypatch)
    resp = client.post("/screenspace/api/tasks/pause")
    assert resp.status_code == 200
    assert calls == [{"drain_events": False}]


def test_resume_persists_manifest(client, monkeypatch):
    worker = screenspace_server._worker
    assert worker is not None
    task = screenspace.create_task(
        "color",
        "P01",
        "s.mp4",
        "/v.mp4",
        "r",
        {"x": 0, "y": 0, "w": 1, "h": 1},
        parameters={"start_seconds": 0.0, "end_seconds": 10.0},
    )
    worker.enqueue(task)
    with worker._lock:
        worker._tasks[task["id"]]["status"] = screenspace.TASK_STATUS_PAUSED
    calls = _install_persist_spy(monkeypatch)
    resp = client.post("/screenspace/api/tasks/resume")
    assert resp.status_code == 200
    assert calls == [{"drain_events": False}]


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


def test_event_exclude_uses_persist_manifest_helper(client, monkeypatch):
    screenspace_server._manifest["events"] = [_sample_event("ev_1")]
    calls = _install_persist_spy(monkeypatch)
    resp = client.put("/screenspace/api/events/ev_1/exclude")
    assert resp.status_code == 200
    assert calls == [{"drain_events": False}]


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


# ---- Multitool tasks ----


def test_create_multitool_task_too_few_steps(client):
    _create_region(client, "healthbar")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "multitool",
            "participant": "P01",
            "region": "healthbar",
            "parameters": {"steps": [{"type": "color"}]},
        },
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "at least 2" in data["error"]


def test_create_multitool_task_invalid_step_type(client):
    _create_region(client, "healthbar")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "multitool",
            "participant": "P01",
            "region": "healthbar",
            "parameters": {
                "steps": [
                    {"type": "color", "target_color": {"h": 0, "s": 0, "v": 0}},
                    {"type": "timelapse"},
                ]
            },
        },
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "invalid type" in data["error"]


def test_create_multitool_task_invalid_step_logic(client):
    _create_region(client, "healthbar")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "multitool",
            "participant": "P01",
            "region": "healthbar",
            "parameters": {
                "steps": [
                    {"type": "color", "region": "healthbar"},
                    {"type": "change", "region": "healthbar", "logic": "MAYBE"},
                ]
            },
        },
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "logic must be" in data["error"]


def test_create_multitool_task_no_steps(client):
    _create_region(client, "healthbar")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "multitool",
            "participant": "P01",
            "region": "healthbar",
            "parameters": {},
        },
    )
    assert resp.status_code == 400


def test_create_multitool_task_missing_step_region(client):
    _create_region(client, "healthbar")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "multitool",
            "participant": "P01",
            "region": "healthbar",
            "parameters": {
                "steps": [
                    {"type": "color", "region": "healthbar"},
                    {"type": "change"},
                ]
            },
        },
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "region is required" in data["error"]


def test_create_multitool_task_unknown_step_region(client):
    _create_region(client, "healthbar")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "multitool",
            "participant": "P01",
            "region": "healthbar",
            "parameters": {
                "steps": [
                    {"type": "color", "region": "healthbar"},
                    {"type": "change", "region": "nonexistent"},
                ]
            },
        },
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "not found" in data["error"]


def test_create_multitool_task_no_global_region_ok(client):
    """Multitool tasks should not require a global region when steps have per-step regions."""
    _create_region(client, "healthbar")
    _create_region(client, "statusbar", x=0, y=0, w=100, h=20)
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "multitool",
            "participant": "P01",
            "region": "",
            "parameters": {
                "steps": [
                    {"type": "color", "region": "healthbar"},
                    {"type": "change", "region": "statusbar"},
                ]
            },
        },
    )
    # Should fail on video lookup (has_video=False), not region validation
    assert resp.status_code == 400
    data = resp.get_json()
    assert "No video" in data["error"]


def test_fast_scan_mode_passes_validation(client):
    """scan_mode in parameters should not cause a validation error.

    P01 has no video, so task creation fails at the video-existence check
    (not at parameter validation).  This confirms scan_mode is accepted
    and passes through to the task parameters.
    """
    screenspace_server._manifest["regions"]["fstest"] = {
        "x": 0,
        "y": 0,
        "w": 10,
        "h": 10,
    }
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "color",
            "participant": "P01",
            "region": "fstest",
            "parameters": {
                "scan_mode": "fast",
                "target_color": {"h": 0, "s": 0, "v": 0},
                "tolerance": {"h": 10, "s": 50, "v": 50},
                "interval": 1.0,
            },
        },
    )
    data = resp.get_json()
    # Should fail at video check, not at parameter validation
    assert resp.status_code == 400
    assert "video" in data["error"].lower()


def test_clean_task_preserves_scan_mode():
    """_clean_task should keep scan_mode in parameters (not stripped)."""
    task = {
        "id": "ss_test",
        "type": "color",
        "parameters": {
            "scan_mode": "fast",
            "target_color": {"h": 0, "s": 0, "v": 0},
            "interval": 1.0,
        },
    }
    cleaned = screenspace_server._clean_task(task)
    assert cleaned["parameters"]["scan_mode"] == "fast"


def test_create_region_uses_persist_manifest_helper(client, monkeypatch):
    calls = _install_persist_spy(monkeypatch)
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
    assert calls == [{"drain_events": False}]


def test_create_stash_uses_persist_manifest_helper(client, monkeypatch):
    _create_region(client, "stashme")
    calls = _install_persist_spy(monkeypatch)
    resp = client.post("/screenspace/api/stashes")
    assert resp.status_code == 200
    assert calls == [{"drain_events": False}]


def _create_region(client, name, x=100, y=20, w=300, h=30):
    client.post(
        "/screenspace/api/regions",
        json={
            "name": name,
            "x": x,
            "y": y,
            "w": w,
            "h": h,
            "canvas_width": 1920,
            "canvas_height": 1080,
        },
    )


def _enable_video_task_setup(monkeypatch, participant_id):
    for participant in screenspace_server._participants:
        if participant["id"] == participant_id:
            participant["has_video"] = True
    monkeypatch.setattr(
        screenspace_server.video,
        "probe_video_properties",
        lambda _path: {"width": 1920, "height": 1080},
    )


def _install_persist_spy(monkeypatch):
    calls = []

    def spy(*, drain_events=True):
        calls.append({"drain_events": drain_events})

    monkeypatch.setattr(screenspace_server, "_do_persist", spy)
    return calls
