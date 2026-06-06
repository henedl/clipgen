"""Tests for Screenspace server API endpoints."""

import pytest

Flask = pytest.importorskip("flask").Flask

import config  # noqa: E402
import screenspace  # noqa: E402
import screenspace_server  # noqa: E402


@pytest.fixture
def client(tmp_path, monkeypatch):
    app = Flask(__name__)
    app.json.sort_keys = False  # mirror start_combined_server: preserve manifest order
    app.register_blueprint(screenspace_server.screenspace_bp, url_prefix="/screenspace")

    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    screenspace_server._manifest = {
        "regions": {},
        "tasks": [],
        "events": [],
        "stashes": [],
        "per_participant": {},
    }
    screenspace_server._participants = [
        {"id": "P01", "video_path": "/tmp/test_P01.mp4", "has_video": False},
        {"id": "P02", "video_path": "/tmp/test_P02.mp4", "has_video": False},
    ]
    screenspace_server._output_dir = str(tmp_path)
    screenspace_server._worker = screenspace.ScreenspaceWorker()

    monkeypatch.setattr(
        screenspace,
        "save_screenspace_manifest",
        lambda r, t, e=None, stashes=None, per_participant=None: tmp_path / "m.json",
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


def test_participant_notes_default_empty(client):
    resp = client.get("/screenspace/api/participants/P01/notes")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["notes"] == ""


def test_participant_notes_round_trip(client):
    put_resp = client.put(
        "/screenspace/api/participants/P01/notes", json={"notes": "watch HUD glitch"}
    )
    assert put_resp.status_code == 200
    assert put_resp.get_json()["ok"] is True

    get_resp = client.get("/screenspace/api/participants/P01/notes")
    assert get_resp.get_json()["notes"] == "watch HUD glitch"


def test_participant_notes_unknown_participant(client):
    resp = client.get("/screenspace/api/participants/PNOPE/notes")
    assert resp.status_code == 404


def test_participant_notes_too_large(client):
    huge = "x" * (64 * 1024 + 1)
    resp = client.put("/screenspace/api/participants/P01/notes", json={"notes": huge})
    assert resp.status_code == 413


def test_participant_notes_isolated_per_participant(client):
    client.put("/screenspace/api/participants/P01/notes", json={"notes": "p1 note"})
    client.put("/screenspace/api/participants/P02/notes", json={"notes": "p2 note"})
    assert (
        client.get("/screenspace/api/participants/P01/notes").get_json()["notes"]
        == "p1 note"
    )
    assert (
        client.get("/screenspace/api/participants/P02/notes").get_json()["notes"]
        == "p2 note"
    )


def test_participant_issues_no_sheet(client, monkeypatch):
    import server

    monkeypatch.setattr(server, "_sheet_context", None, raising=False)
    resp = client.get("/screenspace/api/participants/P01/issues")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["issues"] == []


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


@pytest.mark.parametrize(
    "payload,expected_err",
    [
        ({"x": 0, "y": 0, "w": 10, "h": 10}, None),
        ({"name": "test", "x": 0, "y": 0}, None),
        (
            {"name": "no_canvas", "x": 100, "y": 20, "w": 300, "h": 30},
            "canvas_width",
        ),
    ],
    ids=["missing_name", "missing_coords", "missing_canvas_dims"],
)
def test_create_region_400_for_invalid_payload(client, payload, expected_err):
    resp = client.post("/screenspace/api/regions", json=payload)
    assert resp.status_code == 400
    if expected_err:
        data = resp.get_json()
        assert data["ok"] is False
        assert expected_err in data["error"]


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


def test_denormalize_region():
    from screenspace import denormalize_region

    region = {
        "x": 0.5,
        "y": 0.5,
        "w": 0.1,
        "h": 0.1,
        "source_width": 1920,
        "source_height": 1080,
    }
    px = denormalize_region(region, 1920, 1080)
    assert px == {"x": 960, "y": 540, "w": 192, "h": 108}


def test_denormalize_cross_resolution():
    from screenspace import denormalize_region

    region = {
        "x": 0.5,
        "y": 0.5,
        "w": 0.1,
        "h": 0.1,
        "source_width": 1920,
        "source_height": 1080,
    }
    px = denormalize_region(region, 1280, 720)
    assert px == {"x": 640, "y": 360, "w": 128, "h": 72}


# ---- Region reorder ----


def test_reorder_regions(client):
    _create_region(client, "a")
    _create_region(client, "b")
    _create_region(client, "c")
    resp = client.put(
        "/screenspace/api/regions/reorder", json={"names": ["c", "a", "b"]}
    )
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True
    keys = list(client.get("/screenspace/api/regions").get_json()["regions"].keys())
    assert keys == ["c", "a", "b"]


def test_reorder_regions_rejects_mismatched_names(client):
    _create_region(client, "a")
    _create_region(client, "b")
    # Unknown name and a subset both rejected — the active set must match exactly.
    assert (
        client.put(
            "/screenspace/api/regions/reorder", json={"names": ["a", "x"]}
        ).status_code
        == 400
    )
    assert (
        client.put(
            "/screenspace/api/regions/reorder", json={"names": ["a"]}
        ).status_code
        == 400
    )
    keys = list(client.get("/screenspace/api/regions").get_json()["regions"].keys())
    assert keys == ["a", "b"]


def test_reorder_regions_rejects_duplicate_names(client):
    _create_region(client, "a")
    _create_region(client, "b")
    resp = client.put(
        "/screenspace/api/regions/reorder", json={"names": ["a", "a"]}
    )
    assert resp.status_code == 400
    keys = list(client.get("/screenspace/api/regions").get_json()["regions"].keys())
    assert keys == ["a", "b"]


def test_reorder_regions_requires_names_list(client):
    _create_region(client, "a")
    assert client.put("/screenspace/api/regions/reorder", json={}).status_code == 400
    assert (
        client.put("/screenspace/api/regions/reorder", json={"names": "a"}).status_code
        == 400
    )


def test_reorder_regions_uses_persist_manifest_helper(client, monkeypatch):
    _create_region(client, "a")
    _create_region(client, "b")
    calls = _install_persist_spy(monkeypatch)
    resp = client.put("/screenspace/api/regions/reorder", json={"names": ["b", "a"]})
    assert resp.status_code == 200
    assert calls == [{"drain_events": False}]


# ---- Stash: copy a region in ----


def test_add_region_to_stash_copies_and_keeps_active(client):
    _create_region(client, "a")
    stash = client.post("/screenspace/api/stashes").get_json()[
        "stash"
    ]  # stashes "a", clears active
    _create_region(client, "b")
    resp = client.post(
        "/screenspace/api/stashes/" + stash["id"] + "/regions", json={"name": "b"}
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    # Copy, not move: "b" stays in the active set.
    active = client.get("/screenspace/api/regions").get_json()["regions"]
    assert "b" in active
    # And "b" now lives in the stash too (response + manifest).
    assert "b" in data["stash"]["regions"]
    assert "b" in screenspace_server._manifest["stashes"][0]["regions"]


def test_add_region_to_stash_name_collision_overwrites(client):
    screenspace_server._manifest["stashes"] = [
        {
            "id": "stash_x",
            "name": "S",
            "regions": {
                "a": {
                    "x": 0.1,
                    "y": 0.1,
                    "w": 0.1,
                    "h": 0.1,
                    "source_width": 1920,
                    "source_height": 1080,
                }
            },
        }
    ]
    _create_region(client, "a", x=960, y=540, w=192, h=108)  # different coords
    active_a = client.get("/screenspace/api/regions").get_json()["regions"]["a"]
    resp = client.post("/screenspace/api/stashes/stash_x/regions", json={"name": "a"})
    assert resp.status_code == 200
    # Active definition overwrites the stale stash entry (last-write-wins).
    assert resp.get_json()["stash"]["regions"]["a"] == active_a
    assert "a" in client.get("/screenspace/api/regions").get_json()["regions"]


def test_add_region_to_stash_missing_region(client):
    screenspace_server._manifest["stashes"] = [
        {"id": "stash_x", "name": "S", "regions": {}}
    ]
    resp = client.post(
        "/screenspace/api/stashes/stash_x/regions", json={"name": "nope"}
    )
    assert resp.status_code == 404


def test_add_region_to_stash_missing_stash(client):
    _create_region(client, "a")
    resp = client.post("/screenspace/api/stashes/bogus/regions", json={"name": "a"})
    assert resp.status_code == 404


def test_add_region_to_stash_requires_name(client):
    screenspace_server._manifest["stashes"] = [
        {"id": "stash_x", "name": "S", "regions": {}}
    ]
    assert (
        client.post("/screenspace/api/stashes/stash_x/regions", json={}).status_code
        == 400
    )
    assert (
        client.post(
            "/screenspace/api/stashes/stash_x/regions", json={"name": "  "}
        ).status_code
        == 400
    )


def test_add_region_to_stash_uses_persist_manifest_helper(client, monkeypatch):
    _create_region(client, "a")
    stash = client.post("/screenspace/api/stashes").get_json()["stash"]
    _create_region(client, "b")
    calls = _install_persist_spy(monkeypatch)
    resp = client.post(
        "/screenspace/api/stashes/" + stash["id"] + "/regions", json={"name": "b"}
    )
    assert resp.status_code == 200
    assert calls == [{"drain_events": False}]


# ---- Tasks ----


def test_list_tasks_empty(client):
    resp = client.get("/screenspace/api/tasks")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["tasks"] == []


@pytest.mark.parametrize(
    "payload,expected_err",
    [
        ({"type": "color", "participant": "P01"}, "region"),
        (
            {"type": "color", "participant": "P01", "region": "nonexistent"},
            None,
        ),
        ({"type": "invalid", "participant": "P01", "region": "r"}, None),
    ],
    ids=["missing_region", "unknown_region", "invalid_type"],
)
def test_create_task_400_for_invalid_payload(client, payload, expected_err):
    resp = client.post("/screenspace/api/tasks", json=payload)
    assert resp.status_code == 400
    if expected_err:
        assert expected_err in resp.get_json()["error"].lower()


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


@pytest.mark.parametrize("value", [None, -0.1, 1.1, "not-a-number"])
def test_create_ocr_task_rejects_invalid_confidence_threshold(
    client, monkeypatch, value
):
    _create_region(client, "r")
    _enable_video_task_setup(monkeypatch, "P01")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "text",
            "participant": "P01",
            "region": "r",
            "parameters": {
                "search_string": "score",
                "ocr_confidence_threshold": value,
            },
        },
    )
    assert resp.status_code == 400
    assert "ocr_confidence_threshold" in resp.get_json()["error"]


def test_create_ocr_task_accepts_zero_confidence_threshold(client, monkeypatch):
    _create_region(client, "r")
    _enable_video_task_setup(monkeypatch, "P01")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "text",
            "participant": "P01",
            "region": "r",
            "parameters": {
                "search_string": "score",
                "ocr_confidence_threshold": 0,
            },
        },
    )
    assert resp.status_code == 200
    params = resp.get_json()["task"]["parameters"]
    assert params["ocr_confidence_threshold"] == 0.0


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


def test_api_preview_text_preprocess_changes_ocr_input(client, monkeypatch) -> None:
    import cv2
    import numpy as np
    import video

    _enable_video_task_setup(monkeypatch, "P01")
    frame = np.zeros((80, 120, 3), dtype=np.uint8)
    monkeypatch.setattr(
        video, "extract_frame_at_timestamp", lambda _path, _ts: frame.copy()
    )
    monkeypatch.setattr(
        screenspace,
        "_preprocess_for_ocr",
        lambda _pixels: np.full((40, 40, 3), 255, dtype=np.uint8),
    )

    region = "0.0,0.0,0.5,0.5"
    raw = client.get(f"/screenspace/api/preview/P01/0.0?tool=text&region={region}")
    enhanced = client.get(
        f"/screenspace/api/preview/P01/0.0?tool=text&region={region}&ocr_preprocess=1"
    )
    assert raw.status_code == 200
    assert enhanced.status_code == 200
    raw_img = cv2.imdecode(np.frombuffer(raw.data, np.uint8), cv2.IMREAD_COLOR)
    enhanced_img = cv2.imdecode(
        np.frombuffer(enhanced.data, np.uint8), cv2.IMREAD_COLOR
    )
    assert raw_img is not None and enhanced_img is not None
    assert float(enhanced_img.mean()) > float(raw_img.mean())


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


def test_create_task_region_ref_prefers_active_duplicate(client, monkeypatch):
    _create_region(client, "target", x=96, y=54, w=192, h=108)
    screenspace_server._manifest["stashes"] = [
        {
            "id": "stash_a",
            "name": "Stashed Regions",
            "regions": {
                "target": {
                    "x": 0.5,
                    "y": 0.5,
                    "w": 0.25,
                    "h": 0.25,
                    "source_width": 1920,
                    "source_height": 1080,
                }
            },
        }
    ]
    _enable_video_task_setup(monkeypatch, "P01")

    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "color",
            "participant": "P01",
            "region": "target",
            "region_ref": {"source": "active", "name": "target"},
        },
    )

    assert resp.status_code == 200
    task = resp.get_json()["task"]
    assert task["region"] == "target"
    assert task["region_coords"] == {"x": 96, "y": 54, "w": 192, "h": 108}


def test_create_task_region_ref_resolves_stash_duplicate(client, monkeypatch):
    _create_region(client, "target", x=96, y=54, w=192, h=108)
    screenspace_server._manifest["stashes"] = [
        {
            "id": "stash_a",
            "name": "Stashed Regions",
            "regions": {
                "target": {
                    "x": 0.5,
                    "y": 0.5,
                    "w": 0.25,
                    "h": 0.25,
                    "source_width": 1920,
                    "source_height": 1080,
                }
            },
        }
    ]
    _enable_video_task_setup(monkeypatch, "P01")

    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "color",
            "participant": "P01",
            "region": "target",
            "region_ref": {
                "source": "stash",
                "stash_id": "stash_a",
                "name": "target",
            },
        },
    )

    assert resp.status_code == 200
    task = resp.get_json()["task"]
    assert task["region"] == "target"
    assert task["region_coords"] == {"x": 960, "y": 540, "w": 480, "h": 270}


def test_create_multitool_step_region_ref_resolves_stash_duplicate(client, monkeypatch):
    _create_region(client, "target", x=96, y=54, w=192, h=108)
    _create_region(client, "other", x=10, y=10, w=100, h=50)
    screenspace_server._manifest["stashes"] = [
        {
            "id": "stash_a",
            "name": "Stashed Regions",
            "regions": {
                "target": {
                    "x": 0.5,
                    "y": 0.5,
                    "w": 0.25,
                    "h": 0.25,
                    "source_width": 1920,
                    "source_height": 1080,
                }
            },
        }
    ]
    _enable_video_task_setup(monkeypatch, "P01")

    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "multitool",
            "participant": "P01",
            "region": "",
            "parameters": {
                "steps": [
                    {
                        "type": "color",
                        "region": "target",
                        "region_ref": {
                            "source": "stash",
                            "stash_id": "stash_a",
                            "name": "target",
                        },
                    },
                    {"type": "change", "region": "other"},
                ]
            },
        },
    )

    assert resp.status_code == 200
    worker = screenspace_server._worker
    assert worker is not None
    task = worker.get_all_tasks()[0]
    assert task["parameters"]["steps"][0]["region_coords"] == {
        "x": 960,
        "y": 540,
        "w": 480,
        "h": 270,
    }


def test_create_task_full_frame_region_ref(client, monkeypatch):
    """A region_ref with source 'full_frame' should denormalize to the full video frame."""
    _enable_video_task_setup(monkeypatch, "P01")

    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "color",
            "participant": "P01",
            "region": "full_frame",
            "region_ref": {"source": "full_frame"},
        },
    )

    assert resp.status_code == 200
    task = resp.get_json()["task"]
    assert task["region"] == "full_frame"
    assert task["region_coords"] == {"x": 0, "y": 0, "w": 1920, "h": 1080}


def test_create_multitool_step_full_frame_region_ref(client, monkeypatch):
    """A multitool step can target the full frame via source 'full_frame'."""
    _create_region(client, "other", x=10, y=10, w=100, h=50)
    _enable_video_task_setup(monkeypatch, "P01")

    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "multitool",
            "participant": "P01",
            "region": "",
            "parameters": {
                "steps": [
                    {
                        "type": "color",
                        "region": "full_frame",
                        "region_ref": {"source": "full_frame"},
                    },
                    {"type": "change", "region": "other"},
                ]
            },
        },
    )

    assert resp.status_code == 200
    worker = screenspace_server._worker
    assert worker is not None
    task = worker.get_all_tasks()[0]
    steps = task["parameters"]["steps"]
    assert steps[0]["region"] == "full_frame"
    assert steps[0]["region_coords"] == {"x": 0, "y": 0, "w": 1920, "h": 1080}


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
    # Task removed from manifest (flush the debounced persist first)
    screenspace_server._flush_pending_persist()
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
    screenspace_server._flush_pending_persist()
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


def test_participants_payload_includes_version(client, tmp_path, monkeypatch):
    """/api/participants enriches has_video entries with the file's mtime_ns."""
    video_file = tmp_path / "study_P03.mp4"
    video_file.write_bytes(b"\x00fake")
    monkeypatch.setattr(
        screenspace_server,
        "_participants",
        [
            {"id": "P01", "video_path": "/tmp/none.mp4", "has_video": False},
            {"id": "P03", "video_path": str(video_file), "has_video": True},
        ],
    )

    resp = client.get("/screenspace/api/participants")
    assert resp.status_code == 200
    by_id = {p["id"]: p for p in resp.get_json()["participants"]}
    assert by_id["P01"].get("version") is None
    assert by_id["P03"]["version"] == video_file.stat().st_mtime_ns


def test_video_frame_cache_invalidates_on_mtime_change(client, tmp_path, monkeypatch):
    """Re-encoding the source file gives a fresh frame; same URL+v hits the cache."""
    import video as video_mod

    video_file = tmp_path / "study_P04.mp4"
    video_file.write_bytes(b"\x00original")
    monkeypatch.setattr(
        screenspace_server,
        "_participants",
        [{"id": "P04", "video_path": str(video_file), "has_video": True}],
    )
    monkeypatch.setattr(
        screenspace_server, "_frame_cache", type(screenspace_server._frame_cache)()
    )

    calls = []

    def _fake_extract(path, ts):
        calls.append((path, ts, video_file.stat().st_mtime_ns))
        # Return a 1x1 BGR ndarray-like; cv2.imencode needs a real ndarray.
        import numpy as np

        return np.zeros((1, 1, 3), dtype=np.uint8)

    monkeypatch.setattr(video_mod, "extract_frame_at_timestamp", _fake_extract)

    first = client.get("/screenspace/api/video/frame/P04/0.0")
    assert first.status_code == 200
    first_bytes = first.data
    assert len(calls) == 1

    # Same URL — should hit the cache, no new extraction.
    again = client.get("/screenspace/api/video/frame/P04/0.0")
    assert again.status_code == 200
    assert again.data == first_bytes
    assert len(calls) == 1

    # Bump mtime forward by 1 second to simulate a re-encode. The cache key
    # includes mtime_ns, so this is a fresh entry and we re-extract.
    new_mtime = video_file.stat().st_mtime_ns + 10**9
    import os

    os.utime(video_file, ns=(new_mtime, new_mtime))
    third = client.get("/screenspace/api/video/frame/P04/0.0")
    assert third.status_code == 200
    assert len(calls) == 2


def test_video_info_reprobes_on_mtime_change(client, tmp_path, monkeypatch):
    """info response carries current mtime as version and re-probes on change."""
    import video as video_mod

    video_file = tmp_path / "study_P05.mp4"
    video_file.write_bytes(b"\x00v1")
    monkeypatch.setattr(
        screenspace_server,
        "_participants",
        [{"id": "P05", "video_path": str(video_file), "has_video": True}],
    )
    monkeypatch.setattr(screenspace_server, "_video_metadata_cache", {})

    probe_calls = []

    def _fake_probe(path):
        probe_calls.append(path)
        # Simulate that duration grew after the user re-encoded.
        duration = 30.0 if len(probe_calls) == 1 else 45.0
        return {
            "width": 1920,
            "height": 1080,
            "video_codec": "h264",
            "audio_codec": "aac",
            "fps": 30.0,
            "duration": duration,
            "nb_frames": int(duration * 30),
        }

    monkeypatch.setattr(video_mod, "probe_video_properties", _fake_probe)

    first = client.get("/screenspace/api/video/info/P05").get_json()
    assert first["ok"] is True
    first_version = first["info"]["version"]
    assert first_version == video_file.stat().st_mtime_ns
    assert first["info"]["duration"] == 30
    assert len(probe_calls) == 1

    # Cache hit (same mtime).
    cached = client.get("/screenspace/api/video/info/P05").get_json()
    assert cached["info"]["version"] == first_version
    assert len(probe_calls) == 1

    # Mtime change forces a re-probe and a new version in the response.
    new_mtime = first_version + 10**9
    import os

    os.utime(video_file, ns=(new_mtime, new_mtime))
    fresh = client.get("/screenspace/api/video/info/P05").get_json()
    assert fresh["info"]["version"] == new_mtime
    assert fresh["info"]["version"] != first_version
    assert fresh["info"]["duration"] == 45
    assert len(probe_calls) == 2


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


@pytest.mark.parametrize(
    "parameters,expected_err",
    [
        ({"steps": [{"type": "color"}]}, "at least 2"),
        (
            {
                "steps": [
                    {"type": "color", "target_color": {"h": 0, "s": 0, "v": 0}},
                    {"type": "timelapse"},
                ]
            },
            "invalid type",
        ),
        (
            {
                "steps": [
                    {"type": "color", "region": "healthbar"},
                    {"type": "change", "region": "healthbar", "logic": "MAYBE"},
                ]
            },
            "logic must be",
        ),
        ({}, None),
        (
            {
                "steps": [
                    {"type": "color", "region": "healthbar"},
                    {"type": "change"},
                ]
            },
            "region is required",
        ),
        (
            {
                "steps": [
                    {"type": "color", "region": "healthbar"},
                    {"type": "change", "region": "nonexistent"},
                ]
            },
            "not found",
        ),
    ],
    ids=[
        "too_few_steps",
        "invalid_step_type",
        "invalid_step_logic",
        "no_steps",
        "missing_step_region",
        "unknown_step_region",
    ],
)
def test_create_multitool_task_400_for_invalid_payload(
    client, parameters, expected_err
):
    _create_region(client, "healthbar")
    resp = client.post(
        "/screenspace/api/tasks",
        json={
            "type": "multitool",
            "participant": "P01",
            "region": "healthbar",
            "parameters": parameters,
        },
    )
    assert resp.status_code == 400
    if expected_err:
        assert expected_err in resp.get_json()["error"]


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
    # Debounced task-queue routes go through _schedule_persist → Timer → _do_persist.
    # Collapse the debounce here so assertions don't race the timer.
    monkeypatch.setattr(
        screenspace_server,
        "_schedule_persist",
        lambda: spy(drain_events=False),
    )
    return calls


# ---- Export endpoint ----


def test_export_events_json(client):
    screenspace_server._manifest["events"] = [
        _sample_event("ev_1"),
        _sample_event("ev_2", excluded=True),
    ]
    resp = client.get("/screenspace/api/export/events?format=json")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert len(data["events"]) == 2
    # Metadata should be hoisted to top-level "magnitude"
    assert any("magnitude" in e for e in data["events"])


def test_export_events_csv(client):
    import csv as _csv
    import io as _io

    screenspace_server._manifest["events"] = [
        _sample_event("ev_1"),
        _sample_event("ev_2", participant="P02"),
    ]
    resp = client.get("/screenspace/api/export/events?format=csv")
    assert resp.status_code == 200
    assert resp.content_type.startswith("text/csv")
    assert 'attachment; filename="screenspace_events.csv"' in resp.headers.get(
        "Content-Disposition", ""
    )
    text = resp.get_data(as_text=True)
    reader = _csv.DictReader(_io.StringIO(text))
    rows = list(reader)
    assert len(rows) == 2
    fieldnames = reader.fieldnames
    assert fieldnames is not None
    assert "magnitude" in fieldnames
    assert fieldnames[0] == "id"


def test_export_events_filter_excluded_false(client):
    screenspace_server._manifest["events"] = [
        _sample_event("ev_1"),
        _sample_event("ev_2", excluded=True),
    ]
    resp = client.get("/screenspace/api/export/events?format=json&excluded=false")
    data = resp.get_json()
    assert len(data["events"]) == 1
    assert data["events"][0]["id"] == "ev_1"


def test_export_events_filter_participant(client):
    screenspace_server._manifest["events"] = [
        _sample_event("ev_1", participant="P01"),
        _sample_event("ev_2", participant="P02"),
    ]
    resp = client.get("/screenspace/api/export/events?format=json&participant=P02")
    data = resp.get_json()
    assert len(data["events"]) == 1
    assert data["events"][0]["participant"] == "P02"


def test_export_events_unsupported_format(client):
    resp = client.get("/screenspace/api/export/events?format=xml")
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False


def test_notify_sse_clients_coalesces_on_full_queue():
    """A saturated client queue keeps a fresh 'update' marker instead of
    silently dropping the change, so a slow SSE client still converges to
    current task state once it catches up."""
    import queue as queue_mod

    q: queue_mod.Queue = queue_mod.Queue(maxsize=4)
    for _ in range(4):
        q.put_nowait("stale")

    saved = list(screenspace_server._sse_clients)
    screenspace_server._sse_clients[:] = [q]
    try:
        screenspace_server._notify_sse_clients("task_created")
    finally:
        screenspace_server._sse_clients[:] = saved

    drained = []
    while not q.empty():
        drained.append(q.get_nowait())
    assert len(drained) == 4
    assert "update" in drained


def test_events_list_stable_under_concurrent_writes(client):
    """GET /api/events keeps returning 200 while events are bulk-toggled — the
    handler snapshots the manifest under _manifest_lock instead of letting
    jsonify iterate event dicts another request is mutating."""
    import concurrent.futures
    import threading

    # Fresh app so each thread can hold its own test client; module state
    # (_manifest etc.) set up by the fixture is shared across all of them.
    app = Flask(__name__)
    app.register_blueprint(screenspace_server.screenspace_bp, url_prefix="/screenspace")

    ids = [f"ev_{i}" for i in range(40)]
    screenspace_server._manifest["events"] = [_sample_event(i) for i in ids]

    stop = threading.Event()
    errors: list[object] = []

    def reader() -> None:
        c = app.test_client()
        while not stop.is_set():
            try:
                resp = c.get("/screenspace/api/events")
                if resp.status_code != 200:
                    errors.append(resp.status_code)
            except Exception as exc:  # pragma: no cover - failure path
                errors.append(exc)

    def writer() -> None:
        c = app.test_client()
        c.put("/screenspace/api/events/bulk-exclude", json={"ids": ids})
        c.put("/screenspace/api/events/bulk-include", json={"ids": ids})

    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as pool:
        readers = [pool.submit(reader) for _ in range(3)]
        for _ in range(200):
            pool.submit(writer).result()
        stop.set()
        for r in readers:
            r.result()

    assert errors == []
