import json

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


def test_api_generate_skips_existing_artifacts(client, monkeypatch, tmp_path):
    """Already-generated artifacts are returned without re-running process_clips."""
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))

    # Create the artifact file on disk
    (tmp_path / "clip.mp4").write_bytes(b"video")

    existing = [
        {"id": "a5c2s0", "type": "clip", "file": "clip.mp4", "cellRow": 5, "cellCol": 2}
    ]
    monkeypatch.setattr(server, "_generated_artifacts", list(existing))

    cell = types.SimpleNamespace(row=5, col=2, value="1:00")

    def fake_generate_list(ws, mode, *, cell_specs, skip_prompts):
        return [{"participant": "P01", "cell": cell}]

    def fake_parse_cell_specs(text):
        return [("P01", 5)]

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)
    monkeypatch.setattr("spreadsheet.parse_cell_specifications", fake_parse_cell_specs)

    process_called = []
    monkeypatch.setattr(
        "clipgen.process_clips",
        lambda *a, **kw: process_called.append(1) or (1, []),
    )

    resp = client.post("/studio/api/generate", json={"cells": ["P01.5"], "format": "clip"})
    assert resp.status_code == 200
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert lines[0]["ok"] is True
    assert lines[0]["skipped"] is True
    assert lines[0]["artifacts"] == existing
    assert process_called == []


def test_api_generate_regenerates_when_file_missing(client, monkeypatch, tmp_path):
    """If artifact file is missing from disk, regeneration proceeds normally."""
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))

    # Artifact record exists but file does NOT
    stale = [
        {"id": "a5c2s0", "type": "clip", "file": "gone.mp4", "cellRow": 5, "cellCol": 2}
    ]
    monkeypatch.setattr(server, "_generated_artifacts", list(stale))

    cell = types.SimpleNamespace(row=5, col=2, value="1:00")

    def fake_generate_list(ws, mode, *, cell_specs, skip_prompts):
        return [{"participant": "P01", "cell": cell}]

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)
    monkeypatch.setattr("spreadsheet.parse_cell_specifications", lambda t: [("P01", 5)])

    new_artifact = {"id": "a5c2s0", "type": "clip", "file": "new.mp4", "cellRow": 5, "cellCol": 2}
    monkeypatch.setattr(
        "clipgen.process_clips",
        lambda *a, **kw: (1, [new_artifact]),
    )

    resp = client.post("/studio/api/generate", json={"cells": ["P01.5"], "format": "clip"})
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert lines[0]["ok"] is True
    assert "skipped" not in lines[0]
    assert lines[0]["artifacts"] == [new_artifact]


def test_api_reel_skips_existing_reel(client, monkeypatch, tmp_path):
    """An identical reel is returned without re-running process_reel."""
    import types

    import clipgen as clipgen_mod

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))

    # Create the reel file on disk
    (tmp_path / "study_reel.mp4").write_bytes(b"reel")

    cell = types.SimpleNamespace(row=5, col=2, value="1:00-1:30")

    # Compute expected reel ID using the same function the server will use
    components = [{"cellRow": 5, "cellCol": 2, "start": 60.0, "end": 90.0}]
    expected_id = clipgen_mod.compute_reel_id(components)

    existing_reel = {
        "id": expected_id,
        "file": "study_reel.mp4",
        "study": "study",
        "components": components,
    }
    monkeypatch.setattr(server, "_generated_reels", [existing_reel])

    def fake_generate_list(ws, mode, *, reel_input, skip_prompts):
        return [
            {
                "participant": "P01",
                "cell": cell,
                "desc": "test",
                "category": "cat",
                "study": "study",
                "severity": "",
            }
        ]

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)

    process_called = []
    monkeypatch.setattr(
        "clipgen.process_reel",
        lambda *a, **kw: process_called.append(1) or (1, []),
    )

    resp = client.post("/studio/api/reel", json={"cells": ["P01.5"]})
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["skipped"] is True
    assert data["reels"] == [existing_reel]
    assert process_called == []


def test_api_gallery_500_when_no_context(client):
    resp = client.post("/studio/api/gallery", json={"participant": "P01"})
    assert resp.status_code == 500
    data = resp.get_json()
    assert data["ok"] is False
    assert "No sheet loaded" in data["error"]


def test_api_gallery_400_when_no_participant(client, monkeypatch):
    monkeypatch.setattr(server, "_sheet_context", object())
    resp = client.post("/studio/api/gallery", json={"participant": "", "format": "screen"})
    assert resp.status_code == 400
    data = resp.get_json()
    assert "No participant" in data["error"]


def test_api_gallery_400_for_invalid_format(client, monkeypatch):
    monkeypatch.setattr(server, "_sheet_context", object())
    resp = client.post("/studio/api/gallery", json={"participant": "P01", "format": "clip"})
    assert resp.status_code == 400
    data = resp.get_json()
    assert "Invalid format" in data["error"]


def test_api_gallery_404_when_video_not_found(client, monkeypatch):
    monkeypatch.setattr(server, "_sheet_context", object())
    monkeypatch.setattr(server, "_resolve_source_video", lambda p: None)
    resp = client.post(
        "/studio/api/gallery",
        json={"participant": "P01", "format": "screen", "interval": 10},
    )
    assert resp.status_code == 404
    data = resp.get_json()
    assert "not found" in data["error"]


def test_api_viewer_400_when_no_artifacts(client):
    resp = client.post("/studio/api/viewer")
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False
    assert "No artifacts" in data["error"]


# ---- Stash API tests ----


def test_api_stashes_get_empty(client, monkeypatch):
    monkeypatch.setattr(server, "_load_stashes", lambda: [])
    resp = client.get("/studio/api/stashes")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["stashes"] == []


def test_api_stashes_create(client, monkeypatch):
    saved = []
    monkeypatch.setattr(server, "_load_stashes", lambda: [])
    monkeypatch.setattr(server, "_save_stashes", lambda s: saved.append(s))

    items = [
        {"participant": "P01", "row": 5, "segDuration": 30},
        {"participant": "P02", "row": 7, "segDuration": 45},
    ]
    resp = client.post(
        "/studio/api/stashes",
        json={"action": "create", "items": items, "name": "My reel"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    stash = data["stash"]
    assert stash["name"] == "My reel"
    assert stash["count"] == 2
    assert stash["totalDuration"] == 75
    assert stash["id"].startswith("stash_")
    assert "createdAt" in stash
    assert len(saved) == 1
    assert len(saved[0]) == 1


def test_api_stashes_create_default_name(client, monkeypatch):
    monkeypatch.setattr(server, "_load_stashes", lambda: [{"id": "stash_old"}])
    monkeypatch.setattr(server, "_save_stashes", lambda s: None)

    resp = client.post(
        "/studio/api/stashes",
        json={"action": "create", "items": [{"segDuration": 10}]},
    )
    data = resp.get_json()
    assert data["stash"]["name"] == "Stash 2"


def test_api_stashes_create_empty_items_400(client, monkeypatch):
    monkeypatch.setattr(server, "_load_stashes", lambda: [])
    resp = client.post(
        "/studio/api/stashes",
        json={"action": "create", "items": []},
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "No items" in data["error"]


def test_api_stashes_update_name(client, monkeypatch):
    existing = [{"id": "stash_abc", "name": "Old name", "items": [], "count": 0}]
    saved = []
    monkeypatch.setattr(server, "_load_stashes", lambda: list(existing))
    monkeypatch.setattr(server, "_save_stashes", lambda s: saved.append(s))

    resp = client.post(
        "/studio/api/stashes",
        json={"action": "update", "id": "stash_abc", "name": "New name"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["stash"]["name"] == "New name"
    assert saved[0][0]["name"] == "New name"


def test_api_stashes_update_404(client, monkeypatch):
    monkeypatch.setattr(server, "_load_stashes", lambda: [])
    resp = client.post(
        "/studio/api/stashes",
        json={"action": "update", "id": "stash_nope", "name": "x"},
    )
    assert resp.status_code == 404


def test_api_stashes_delete(client, monkeypatch):
    existing = [
        {"id": "stash_aaa", "name": "A"},
        {"id": "stash_bbb", "name": "B"},
    ]
    saved = []
    monkeypatch.setattr(server, "_load_stashes", lambda: list(existing))
    monkeypatch.setattr(server, "_save_stashes", lambda s: saved.append(s))

    resp = client.post(
        "/studio/api/stashes",
        json={"action": "delete", "id": "stash_aaa"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert len(saved[0]) == 1
    assert saved[0][0]["id"] == "stash_bbb"


def test_api_stashes_delete_not_found_404(client, monkeypatch):
    monkeypatch.setattr(server, "_load_stashes", lambda: [])
    resp = client.post(
        "/studio/api/stashes",
        json={"action": "delete", "id": "stash_nope"},
    )
    assert resp.status_code == 404


def test_api_stashes_unknown_action_400(client, monkeypatch):
    monkeypatch.setattr(server, "_load_stashes", lambda: [])
    resp = client.post(
        "/studio/api/stashes",
        json={"action": "bogus"},
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "Unknown action" in data["error"]


# ---- Artifact stash API tests ----


def test_api_artifact_stashes_get_empty(client, monkeypatch):
    monkeypatch.setattr(server, "_load_artifact_stashes", lambda: [])
    resp = client.get("/studio/api/artifact-stashes")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["stashes"] == []


def test_api_artifact_stashes_create(client, monkeypatch):
    saved = []
    monkeypatch.setattr(server, "_load_artifact_stashes", lambda: [])
    monkeypatch.setattr(server, "_save_artifact_stashes", lambda s: saved.append(s))

    items = [
        {"participant": "P01", "row": 5, "segDuration": 30},
        {"participant": "P02", "row": 7, "segDuration": 45},
    ]
    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "create", "items": items, "name": "My artifacts"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    stash = data["stash"]
    assert stash["name"] == "My artifacts"
    assert stash["count"] == 2
    assert stash["totalDuration"] == 75
    assert stash["id"].startswith("astash_")
    assert "createdAt" in stash
    assert len(saved) == 1
    assert len(saved[0]) == 1


def test_api_artifact_stashes_create_default_name(client, monkeypatch):
    monkeypatch.setattr(
        server, "_load_artifact_stashes", lambda: [{"id": "astash_old"}]
    )
    monkeypatch.setattr(server, "_save_artifact_stashes", lambda s: None)

    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "create", "items": [{"segDuration": 10}]},
    )
    data = resp.get_json()
    assert data["stash"]["name"] == "Stash 2"


def test_api_artifact_stashes_create_empty_items_400(client, monkeypatch):
    monkeypatch.setattr(server, "_load_artifact_stashes", lambda: [])
    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "create", "items": []},
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "No items" in data["error"]


def test_api_artifact_stashes_update_name(client, monkeypatch):
    existing = [{"id": "astash_abc", "name": "Old name", "items": [], "count": 0}]
    saved = []
    monkeypatch.setattr(server, "_load_artifact_stashes", lambda: list(existing))
    monkeypatch.setattr(server, "_save_artifact_stashes", lambda s: saved.append(s))

    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "update", "id": "astash_abc", "name": "New name"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["stash"]["name"] == "New name"
    assert saved[0][0]["name"] == "New name"


def test_api_artifact_stashes_update_404(client, monkeypatch):
    monkeypatch.setattr(server, "_load_artifact_stashes", lambda: [])
    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "update", "id": "astash_nope", "name": "x"},
    )
    assert resp.status_code == 404


def test_api_artifact_stashes_delete(client, monkeypatch):
    existing = [
        {"id": "astash_aaa", "name": "A"},
        {"id": "astash_bbb", "name": "B"},
    ]
    saved = []
    monkeypatch.setattr(server, "_load_artifact_stashes", lambda: list(existing))
    monkeypatch.setattr(server, "_save_artifact_stashes", lambda s: saved.append(s))

    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "delete", "id": "astash_aaa"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert len(saved[0]) == 1
    assert saved[0][0]["id"] == "astash_bbb"


def test_api_artifact_stashes_delete_not_found_404(client, monkeypatch):
    monkeypatch.setattr(server, "_load_artifact_stashes", lambda: [])
    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "delete", "id": "astash_nope"},
    )
    assert resp.status_code == 404


def test_api_artifact_stashes_unknown_action_400(client, monkeypatch):
    monkeypatch.setattr(server, "_load_artifact_stashes", lambda: [])
    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "bogus"},
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "Unknown action" in data["error"]


# ---- Titlecard settings tests ----


def test_api_sheet_returns_titlecard_defaults(client, monkeypatch):
    """api/sheet includes titlecardsEnabled and titlecardDuration from config."""
    import types

    import config

    ctx = types.SimpleNamespace(
        header_row=["ID", "P01"],
        id_cell=types.SimpleNamespace(row=1, col=1),
        num_participants=1,
        study_name="study",
        observation_cell=types.SimpleNamespace(col=3),
        category_cell=types.SimpleNamespace(col=4),
        severity_cell=None,
        baseline_row_idx=None,
        filename_row_idx=None,
        first_data_row_idx=2,
        sheet_data=[["study"], ["ID", "P01", "Observation", "Category"]],
    )
    monkeypatch.setattr(server, "_sheet_context", ctx)

    resp = client.get("/studio/api/sheet")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["titlecardsEnabled"] == config.TITLECARDS_ENABLED
    assert data["titlecardDuration"] == config.TITLECARD_DURATION_SECONDS


def test_api_generate_titlecard_override(client, monkeypatch):
    """titlecards_enabled and titlecard_duration temporarily override config during generate."""
    import types

    import config

    monkeypatch.setattr(server, "_worksheet", object())
    original_enabled = config.TITLECARDS_ENABLED
    original_duration = config.TITLECARD_DURATION_SECONDS
    captured = {}

    cell = types.SimpleNamespace(row=5, col=2, value="1:00")

    def fake_generate_list(ws, mode, *, cell_specs, skip_prompts):
        return [{"participant": "P01", "cell": cell}]

    def fake_process_clips(clips, *, output_format):
        captured["enabled"] = config.TITLECARDS_ENABLED
        captured["duration"] = config.TITLECARD_DURATION_SECONDS
        return (1, [{"id": "a1", "type": "clip"}])

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)
    monkeypatch.setattr("spreadsheet.parse_cell_specifications", lambda t: [("P01", 5)])
    monkeypatch.setattr("clipgen.process_clips", fake_process_clips)
    monkeypatch.setattr(server, "_generated_artifacts", [])

    resp = client.post(
        "/studio/api/generate",
        json={
            "cells": ["P01.5"],
            "format": "clip",
            "titlecards_enabled": True,
            "titlecard_duration": 5,
        },
    )
    lines = [json.loads(ln) for ln in resp.data.decode().strip().split("\n")]
    assert resp.status_code == 200
    assert captured["enabled"] is True
    assert captured["duration"] == 5
    assert config.TITLECARDS_ENABLED == original_enabled
    assert config.TITLECARD_DURATION_SECONDS == original_duration


def test_api_generate_titlecard_restored_on_error(client, monkeypatch):
    """Titlecard config is restored even when process_clips raises."""
    import types

    import config

    monkeypatch.setattr(server, "_worksheet", object())
    original_enabled = config.TITLECARDS_ENABLED
    original_duration = config.TITLECARD_DURATION_SECONDS

    cell = types.SimpleNamespace(row=5, col=2, value="1:00")

    def fake_generate_list(ws, mode, *, cell_specs, skip_prompts):
        return [{"participant": "P01", "cell": cell}]

    def fake_process_clips(clips, *, output_format):
        raise RuntimeError("boom")

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)
    monkeypatch.setattr("spreadsheet.parse_cell_specifications", lambda t: [("P01", 5)])
    monkeypatch.setattr("clipgen.process_clips", fake_process_clips)
    monkeypatch.setattr(server, "_generated_artifacts", [])

    resp = client.post(
        "/studio/api/generate",
        json={
            "cells": ["P01.5"],
            "format": "clip",
            "titlecards_enabled": True,
            "titlecard_duration": 10,
        },
    )
    _ = resp.data  # consume streaming response to trigger finally
    assert resp.status_code == 200  # streaming response still 200
    assert config.TITLECARDS_ENABLED == original_enabled
    assert config.TITLECARD_DURATION_SECONDS == original_duration


def test_api_reel_titlecard_override(client, monkeypatch):
    """titlecards_enabled and titlecard_duration temporarily override config during reel build."""
    import config

    monkeypatch.setattr(server, "_worksheet", object())
    original_enabled = config.TITLECARDS_ENABLED
    original_duration = config.TITLECARD_DURATION_SECONDS
    captured = {}

    def fake_generate_list(ws, mode, *, reel_input, skip_prompts):
        captured["enabled"] = config.TITLECARDS_ENABLED
        captured["duration"] = config.TITLECARD_DURATION_SECONDS
        return []

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)

    resp = client.post(
        "/studio/api/reel",
        json={
            "cells": ["P01.5"],
            "titlecards_enabled": True,
            "titlecard_duration": 4,
        },
    )
    assert resp.status_code == 400  # no clips → 400
    assert captured["enabled"] is True
    assert captured["duration"] == 4
    assert config.TITLECARDS_ENABLED == original_enabled
    assert config.TITLECARD_DURATION_SECONDS == original_duration


def test_api_reel_titlecard_restored_on_error(client, monkeypatch):
    """Titlecard config is restored even when generate_list raises."""
    import config

    monkeypatch.setattr(server, "_worksheet", object())
    original_enabled = config.TITLECARDS_ENABLED
    original_duration = config.TITLECARD_DURATION_SECONDS

    def raise_generate_list(ws, mode, *, reel_input, skip_prompts):
        raise RuntimeError("boom")

    monkeypatch.setattr("spreadsheet.generate_list", raise_generate_list)

    resp = client.post(
        "/studio/api/reel",
        json={
            "cells": ["P01.5"],
            "titlecards_enabled": True,
            "titlecard_duration": 7,
        },
    )
    assert resp.status_code == 500
    assert config.TITLECARDS_ENABLED == original_enabled
    assert config.TITLECARD_DURATION_SECONDS == original_duration


# ---- Settings API tests ----


def test_api_settings_get(client):
    """GET /api/settings returns all Studio-exposed settings with metadata."""
    resp = client.get("/studio/api/settings")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    settings = data["settings"]
    assert len(settings) > 0

    names = {s["name"] for s in settings}
    assert "REENCODING" in names
    assert "TITLECARDS_ENABLED" in names
    assert "HIGHLIGHTS_REEL_DURATION_SECONDS" in names

    for s in settings:
        assert "name" in s
        assert "value" in s
        assert "default" in s
        assert "description" in s
        assert "group" in s
        assert "type" in s


def test_api_settings_put_applies_values(client, monkeypatch):
    """PUT /api/settings applies values to config and returns applied dict."""
    import config

    monkeypatch.setattr(server, "_save_studio_settings", lambda o: None)
    original = config.REENCODING

    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"REENCODING": True}},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["applied"]["REENCODING"] is True
    assert config.REENCODING is True

    # Restore
    config.REENCODING = original


def test_api_settings_put_ignores_unknown(client, monkeypatch):
    """Unknown setting names are silently ignored."""
    monkeypatch.setattr(server, "_save_studio_settings", lambda o: None)

    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"FAKE_SETTING": 42, "REENCODING": False}},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert "FAKE_SETTING" not in data["applied"]
    assert "REENCODING" in data["applied"]


def test_api_settings_put_type_coercion(client, monkeypatch):
    """Int and bool values are coerced correctly."""
    import config

    monkeypatch.setattr(server, "_save_studio_settings", lambda o: None)
    original = config.HIGHLIGHTS_REEL_DURATION_SECONDS

    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"HIGHLIGHTS_REEL_DURATION_SECONDS": "120"}},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["applied"]["HIGHLIGHTS_REEL_DURATION_SECONDS"] == 120
    assert config.HIGHLIGHTS_REEL_DURATION_SECONDS == 120

    config.HIGHLIGHTS_REEL_DURATION_SECONDS = original


def test_api_settings_put_invalid_payload(client):
    """PUT /api/settings with no settings dict returns 400."""
    resp = client.put(
        "/studio/api/settings",
        json={"not_settings": {}},
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False


def test_load_studio_settings(monkeypatch, tmp_path):
    """_load_studio_settings reads file and applies to config."""
    import config

    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    original = config.REENCODING

    settings_file = tmp_path / config.STUDIO_SETTINGS_FILENAME
    settings_file.write_text(json.dumps({"REENCODING": True}))

    applied = server._load_studio_settings()
    assert applied["REENCODING"] is True
    assert config.REENCODING is True

    config.REENCODING = original


def test_load_studio_settings_missing_file(monkeypatch, tmp_path):
    """Missing settings file returns empty dict without error."""
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    applied = server._load_studio_settings()
    assert applied == {}


def test_save_studio_settings_non_defaults_only(monkeypatch, tmp_path):
    """Only non-default values are written; all-defaults deletes the file."""
    import config

    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    settings_file = tmp_path / config.STUDIO_SETTINGS_FILENAME

    # Save a non-default value
    result = server._save_studio_settings({"REENCODING": True})
    assert result is not None
    assert settings_file.is_file()
    data = json.loads(settings_file.read_text())
    assert data["REENCODING"] is True

    # Save all defaults — file should be removed
    result = server._save_studio_settings(
        {"REENCODING": server._settings_defaults["REENCODING"]}
    )
    assert result is None
    assert not settings_file.is_file()
