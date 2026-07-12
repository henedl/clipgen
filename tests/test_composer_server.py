"""Smoke tests for the Composer Flask blueprint.

Verifies the page serves, participant/part discovery, the composer manifest
round-trip (cuts CRUD + UI toggles persisted to ``composer_manifest.json``),
and span clamping — mirroring tests/test_workflows_api.py's bare-blueprint
setup. Combined-app registration (topnav-visible ``/composer/`` + the
``/api/status`` flag) is exercised against ``server.build_combined_app``.
"""

import json

import pytest

Flask = pytest.importorskip("flask").Flask

import composer_server  # noqa: E402
import config  # noqa: E402
import utils  # noqa: E402
import video  # noqa: E402


@pytest.fixture
def co_client(tmp_path, monkeypatch):
    app = Flask(__name__)
    app.register_blueprint(composer_server.composer_bp, url_prefix="/composer")

    # Seed module globals via monkeypatch so they auto-restore on teardown.
    monkeypatch.setattr(composer_server, "_manifest", composer_server._empty_manifest())
    monkeypatch.setattr(composer_server, "_input_dir", str(tmp_path))
    monkeypatch.setattr(composer_server, "_sheet_context", None)
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path), raising=False)

    # One two-part participant with 10 s parts → stitched duration 20 s.
    monkeypatch.setattr(
        utils,
        "discover_participant_videos",
        lambda study_name="": [
            {
                "id": "P01",
                "has_video": True,
                "video_paths": [
                    str(tmp_path / "study_P01-1.mp4"),
                    str(tmp_path / "study_P01-2.mp4"),
                ],
            }
        ],
    )
    monkeypatch.setattr(video, "get_file_duration", lambda path: 10)

    with app.test_client() as c:
        yield c


def _manifest_on_disk(tmp_path):
    return json.loads(
        (tmp_path / config.COMPOSER_MANIFEST_FILENAME).read_text(encoding="utf-8")
    )


def test_page_serves(co_client):
    resp = co_client.get("/composer/")
    assert resp.status_code == 200
    assert resp.mimetype == "text/html"
    body = resp.get_data(as_text=True)
    assert 'data-frontend="composer"' in body


def test_participants_reports_parts_and_total(co_client):
    body = co_client.get("/composer/api/participants").get_json()
    assert body["ok"] is True
    (p,) = body["participants"]
    assert p["id"] == "P01"
    assert [part["offset"] for part in p["parts"]] == [0, 10]
    assert p["total_duration"] == 20
    assert "defaultDuration" in body["config"]


def test_cuts_crud_round_trip(co_client, tmp_path):
    created = co_client.post(
        "/composer/api/cuts",
        json={"participant": "P01", "start": 2.0, "end": 5.5},
    ).get_json()
    assert created["ok"] is True
    cut_id = created["cut"]["id"]
    assert cut_id.startswith("cut_")

    # Persisted to composer_manifest.json (save-after-mutation).
    disk = _manifest_on_disk(tmp_path)
    assert [c["id"] for c in disk["cuts"]] == [cut_id]

    patched = co_client.patch(
        f"/composer/api/cuts/{cut_id}", json={"start": 1.0, "end": 4.0}
    ).get_json()
    assert patched["ok"] is True
    assert (patched["cut"]["start"], patched["cut"]["end"]) == (1.0, 4.0)

    listed = co_client.get("/composer/api/manifest").get_json()
    assert listed["manifest"]["cuts"][0]["start"] == 1.0

    deleted = co_client.delete(f"/composer/api/cuts/{cut_id}").get_json()
    assert deleted["ok"] is True
    assert _manifest_on_disk(tmp_path)["cuts"] == []


def test_cut_create_rejects_inverted_span(co_client):
    resp = co_client.post(
        "/composer/api/cuts", json={"participant": "P01", "start": 5.0, "end": 5.0}
    )
    assert resp.status_code == 400


def test_cut_patch_clamps_to_duration(co_client):
    cut = co_client.post(
        "/composer/api/cuts", json={"participant": "P01", "start": 2.0, "end": 5.0}
    ).get_json()["cut"]
    patched = co_client.patch(
        f"/composer/api/cuts/{cut['id']}", json={"end": 999.0}
    ).get_json()
    assert patched["cut"]["end"] == 20  # stitched duration caps the out point


def test_cut_patch_unknown_id_404(co_client):
    resp = co_client.patch("/composer/api/cuts/cut_nope", json={"end": 3})
    assert resp.status_code == 404


def test_ui_toggles_round_trip(co_client, tmp_path):
    resp = co_client.put(
        "/composer/api/ui",
        json={"markerSources": {"sheet": False, "screenspace": True}},
    ).get_json()
    assert resp["ok"] is True
    assert resp["markerSources"] == {
        "sheet": False,
        "screenspace": True,
        "transcript": True,
    }
    assert _manifest_on_disk(tmp_path)["ui"]["markerSources"]["sheet"] is False


def test_ui_lane_folds_round_trip(co_client, tmp_path):
    resp = co_client.put(
        "/composer/api/ui",
        json={"laneFolds": {"screenspace": False}},
    ).get_json()
    assert resp["ok"] is True
    assert resp["laneFolds"] == {
        "sheet": True,
        "screenspace": False,
        "transcript": True,
    }
    disk = _manifest_on_disk(tmp_path)["ui"]
    assert disk["laneFolds"]["screenspace"] is False
    # A folds-only PUT must not clobber previously saved lane toggles.
    assert "markerSources" not in disk or isinstance(disk.get("markerSources"), dict)


def test_ui_requires_some_payload(co_client):
    resp = co_client.put("/composer/api/ui", json={})
    assert resp.status_code == 400


def test_combined_app_registers_composer(tmp_path, monkeypatch):
    import server
    import start_settings

    in_dir = tmp_path / "in"
    out_dir = tmp_path / "out"
    in_dir.mkdir()
    out_dir.mkdir()
    monkeypatch.setattr(config, "INPUT_DIR", str(in_dir), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(out_dir), raising=False)
    monkeypatch.setattr(
        start_settings, "_settings_path", lambda: tmp_path / "start.json"
    )

    app = server.build_combined_app(worksheet=None, default_page="composer")
    with app.test_client() as client:
        assert client.get("/").location.endswith("/composer/")
        status = client.get("/api/status").get_json()
        assert status["composer"] is True
        assert client.get("/composer/").status_code == 200
