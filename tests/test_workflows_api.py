"""Smoke tests for the Workflows Flask blueprint (M0 scaffold).

Verifies the blueprint serves its page and the read-only blueprint-list
endpoint, mirroring tests/test_transcripts_api.py. Canvas CRUD, the node
catalog, and run endpoints get their own tests as those milestones land.
"""

import pytest

Flask = pytest.importorskip("flask").Flask

import workflows  # noqa: E402
import workflows_server  # noqa: E402


@pytest.fixture
def wf_client(tmp_path):
    app = Flask(__name__)
    app.register_blueprint(workflows_server.workflows_bp, url_prefix="/workflows")

    workflows_server._manifest = workflows.empty_workflows_manifest()
    workflows_server._input_dir = str(tmp_path)
    workflows_server._sheet_context = None

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
