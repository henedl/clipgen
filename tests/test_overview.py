"""Tests for overview: the /overview/ blueprint and its convergence-offset routes.

Route smokes mount the blueprint on a bare Flask app with a tmp output dir
(mirroring tests/test_workflows_api.py).
"""

import json

import pytest

import config
import overview


@pytest.fixture
def seeded_output_dir(tmp_path, monkeypatch):
    """Point the manifest loaders at a tmp output dir; return it for seeding."""
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path), raising=False)
    return tmp_path


@pytest.fixture(scope="module")
def overview_app():
    """The Flask app, built once for the module — see ``client`` in
    tests/test_screenspace_api.py for why the build is safe to share."""
    Flask = pytest.importorskip("flask").Flask
    app = Flask(__name__)
    app.register_blueprint(overview.overview_bp, url_prefix="/overview")
    return app


@pytest.fixture
def overview_client(overview_app, seeded_output_dir):
    with overview_app.test_client() as c:
        yield c


def test_overview_page_serves(overview_client):
    resp = overview_client.get("/overview/")
    assert resp.status_code == 200
    assert "text/html" in resp.content_type
    assert 'data-frontend="overview"' in resp.get_data(as_text=True)


# ---- Convergence offsets routes (moved here with the Convergence tab) ----


def test_api_convergence_offsets_get_empty(overview_client, seeded_output_dir):
    # Doubles as the "API routes are registered before register_static_routes,
    # so the catch-all /<path:filename> can never shadow them" guard.
    resp = overview_client.get("/overview/api/convergence/offsets")
    assert resp.status_code == 200
    assert "application/json" in resp.content_type
    data = resp.get_json()
    assert data["ok"] is True
    assert data["offsets"] == {}


def test_api_convergence_offsets_put_persists(overview_client, seeded_output_dir):
    # Nested per-lane shape: the transcript: 0 lane is dropped on cleaning.
    payload = {
        "offsets": {
            "P01": {"sheet": 12.5, "screenspace": 12.5, "transcript": 0},
            "P03": {"sheet": -7.0},
        }
    }
    expected = {"P01": {"sheet": 12.5, "screenspace": 12.5}, "P03": {"sheet": -7.0}}
    resp = overview_client.put("/overview/api/convergence/offsets", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["offsets"] == expected

    saved = json.loads(
        (seeded_output_dir / config.CONVERGENCE_OFFSETS_FILENAME).read_text()
    )
    assert saved == {"offsets": expected}

    resp2 = overview_client.get("/overview/api/convergence/offsets")
    assert resp2.get_json()["offsets"] == expected


def test_api_convergence_offsets_put_strips_zeros_and_garbage(
    overview_client, seeded_output_dir
):
    payload = {
        "offsets": {
            # Surviving participant: only the non-zero, known-source lane stays.
            "P01": {"sheet": 5.0, "screenspace": 0, "transcript": "nope"},
            # All lanes zero/garbage -> participant dropped entirely.
            "P02": {"sheet": 0, "screenspace": float("inf")},
            # Unknown source key dropped -> participant has no lanes -> dropped.
            "P03": {"audio": 5.0},
            # Non-dict participant value -> dropped (not a 400).
            "P04": 5,
            # Empty participant id -> dropped.
            "": {"sheet": 9.0},
        }
    }
    resp = overview_client.put("/overview/api/convergence/offsets", json=payload)
    assert resp.status_code == 200
    assert resp.get_json()["offsets"] == {"P01": {"sheet": 5.0}}


def test_api_convergence_offsets_round_trip_per_lane(
    overview_client, seeded_output_dir
):
    """A participant with one non-zero and one zero lane round-trips to just the
    non-zero lane through PUT then GET, keeping the get/put cleaners in sync."""
    payload = {"offsets": {"P02": {"sheet": 0, "transcript": -3.5}}}
    put = overview_client.put("/overview/api/convergence/offsets", json=payload)
    assert put.status_code == 200
    assert put.get_json()["offsets"] == {"P02": {"transcript": -3.5}}

    get = overview_client.get("/overview/api/convergence/offsets")
    assert get.get_json()["offsets"] == {"P02": {"transcript": -3.5}}


def test_api_convergence_offsets_put_empty_removes_file(
    overview_client, seeded_output_dir
):
    # Seed a manifest file via an earlier put.
    overview_client.put(
        "/overview/api/convergence/offsets",
        json={"offsets": {"P01": {"sheet": 3.0}}},
    )
    settings_file = seeded_output_dir / config.CONVERGENCE_OFFSETS_FILENAME
    assert settings_file.is_file()

    resp = overview_client.put(
        "/overview/api/convergence/offsets",
        json={"offsets": {"P01": {"sheet": 0}}},
    )
    assert resp.status_code == 200
    assert resp.get_json()["offsets"] == {}
    assert not settings_file.is_file()


def test_api_convergence_offsets_put_empty_sweeps_stale_tmp(
    overview_client, seeded_output_dir
):
    """An interrupted save's .tmp sibling is removed along with the manifest."""
    stale_tmp = seeded_output_dir / (config.CONVERGENCE_OFFSETS_FILENAME + ".tmp")
    stale_tmp.write_text("{}")

    resp = overview_client.put(
        "/overview/api/convergence/offsets", json={"offsets": {}}
    )
    assert resp.status_code == 200
    assert not stale_tmp.exists()


def test_api_convergence_offsets_put_rejects_non_dict(
    overview_client, seeded_output_dir
):
    resp = overview_client.put(
        "/overview/api/convergence/offsets", json={"offsets": "nope"}
    )
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False
