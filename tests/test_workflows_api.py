"""Smoke tests for the Workflows Flask blueprint.

Verifies the page serves, the node catalog endpoint, and blueprint CRUD (the
canvas autosave target), mirroring tests/test_transcripts_api.py. Run endpoints
get their own tests when that milestone (M4) lands.
"""

import json

import pytest

Flask = pytest.importorskip("flask").Flask

import config  # noqa: E402
import workflows  # noqa: E402
import workflows_server  # noqa: E402


@pytest.fixture
def wf_client(tmp_path, monkeypatch):
    app = Flask(__name__)
    app.register_blueprint(workflows_server.workflows_bp, url_prefix="/workflows")

    workflows_server._manifest = workflows.empty_workflows_manifest()
    workflows_server._input_dir = str(tmp_path)
    workflows_server._sheet_context = None
    # Sandbox save_workflows_manifest's write into tmp (it targets the output dir).
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path), raising=False)

    with app.test_client() as c:
        yield c


def test_page_serves(wf_client):
    resp = wf_client.get("/workflows/")
    assert resp.status_code == 200
    assert resp.mimetype == "text/html"
    # The shared TopNav mount + page hub script must be present.
    body = resp.get_data(as_text=True)
    assert 'data-frontend="workflows"' in body
    assert "workflows.js" in body


def test_blueprints_empty_by_default(wf_client):
    resp = wf_client.get("/workflows/api/blueprints")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["blueprints"] == []


def test_blueprints_reflects_manifest(wf_client):
    workflows_server._manifest = {
        "blueprints": [{"id": "bp1", "name": "Demo", "nodes": [], "edges": []}],
        "stashes": [],
        "runs": [],
    }
    resp = wf_client.get("/workflows/api/blueprints")
    assert resp.status_code == 200
    blueprints = resp.get_json()["blueprints"]
    assert len(blueprints) == 1
    assert blueprints[0]["id"] == "bp1"


def test_empty_manifest_has_all_keys():
    m = workflows.empty_workflows_manifest()
    assert set(m) == {"blueprints", "stashes", "runs"}
    assert all(m[k] == [] for k in m)


# ---- Catalog ----


def test_catalog_returns_serializable_node_types(wf_client):
    resp = wf_client.get("/workflows/api/catalog")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    catalog = data["catalog"]
    assert isinstance(catalog, list) and catalog
    ids = {node["id"] for node in catalog}
    # A few headline nodes the cross-domain path needs.
    assert {"transcribe", "ss_scan", "make_clips", "gate"} <= ids
    # The serialized catalog must carry no `execute` callable (JSON-safe).
    assert all("execute" not in node for node in catalog)
    # Launch-context flags drive palette grey-out; participants populate the
    # Video-Source dropdown (reusing the videoDir discovery call).
    assert set(data["context"]) == {"sheet", "videoDir", "participants"}
    assert data["context"]["sheet"] is False  # fixture sets _sheet_context=None
    assert isinstance(data["context"]["participants"], list)


def test_serialize_catalog_is_json_safe():
    # serialize_catalog must round-trip through json without TypeErrors.
    json.dumps(workflows.serialize_catalog())


def test_every_node_type_has_a_callable_executor():
    # M3 wires an execute callable onto every catalog node; a miss would mean a
    # placed node the runner can't execute. serialize_catalog still strips it.
    missing = [
        nid
        for nid, node in workflows.NODE_TYPES.items()
        if not callable(node.get("execute"))
    ]
    assert missing == []
    assert all("execute" not in node for node in workflows.serialize_catalog())


# ---- Blueprint CRUD ----


def test_blueprint_create_returns_id_and_defaults(wf_client):
    resp = wf_client.post("/workflows/api/blueprints", json={"name": "Graph A"})
    assert resp.status_code == 200
    bp = resp.get_json()["blueprint"]
    assert bp["id"].startswith("bp_")
    assert bp["name"] == "Graph A"
    assert set(bp) >= {
        "id",
        "name",
        "nodes",
        "edges",
        "viewport",
        "trigger",
        "createdAt",
    }
    assert bp["trigger"] is None
    # Reflected by the list endpoint.
    listed = wf_client.get("/workflows/api/blueprints").get_json()["blueprints"]
    assert any(b["id"] == bp["id"] for b in listed)


def test_blueprint_update_round_trips_nodes_and_viewport(wf_client):
    bp = wf_client.post("/workflows/api/blueprints", json={}).get_json()["blueprint"]
    payload = {
        "name": "Renamed",
        "nodes": [
            {
                "id": "n_1",
                "type": "transcribe",
                "params": {},
                "position": {"x": 40, "y": 90},
            }
        ],
        "viewport": {"x": -120, "y": 30, "zoom": 1.5},
    }
    resp = wf_client.put(f"/workflows/api/blueprints/{bp['id']}", json=payload)
    assert resp.status_code == 200

    fetched = wf_client.get("/workflows/api/blueprints").get_json()["blueprints"]
    saved = next(b for b in fetched if b["id"] == bp["id"])
    assert saved["name"] == "Renamed"
    assert saved["nodes"][0]["position"] == {"x": 40, "y": 90}
    assert saved["viewport"] == {"x": -120, "y": 30, "zoom": 1.5}


def test_blueprint_delete_removes_it(wf_client):
    bp = wf_client.post("/workflows/api/blueprints", json={}).get_json()["blueprint"]
    assert wf_client.delete(f"/workflows/api/blueprints/{bp['id']}").status_code == 200
    listed = wf_client.get("/workflows/api/blueprints").get_json()["blueprints"]
    assert all(b["id"] != bp["id"] for b in listed)


def test_blueprint_update_and_delete_404_on_missing(wf_client):
    assert (
        wf_client.put(
            "/workflows/api/blueprints/bp_missing", json={"name": "x"}
        ).status_code
        == 404
    )
    assert wf_client.delete("/workflows/api/blueprints/bp_missing").status_code == 404
