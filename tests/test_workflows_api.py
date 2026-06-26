"""Smoke tests for the Workflows Flask blueprint.

Verifies the page serves, the node catalog endpoint, blueprint CRUD (the canvas
autosave target), and the M4 run lifecycle (create/list/cancel/SSE), mirroring
tests/test_transcripts_api.py. The engine itself is covered in
tests/test_workflows_runner.py; here we exercise the HTTP surface.
"""

import json
import time

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
    workflows_server._worksheet = None
    # Reset run state so runners/SSE clients never leak across tests.
    workflows_server._runs = {}
    workflows_server._sse_clients = []
    # Sandbox save_workflows_manifest's write into tmp (it targets the output dir).
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path), raising=False)

    with app.test_client() as c:
        yield c


def _make_blueprint(client, nodes, edges=None):
    bp = client.post("/workflows/api/blueprints", json={}).get_json()["blueprint"]
    client.put(
        f"/workflows/api/blueprints/{bp['id']}",
        json={"nodes": nodes, "edges": edges or []},
    )
    return bp["id"]


def _wait_terminal(client, run_id, timeout=5.0):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        run = client.get(f"/workflows/api/runs/{run_id}").get_json()["run"]
        if run["status"] in ("completed", "failed", "cancelled"):
            return run
        time.sleep(0.02)
    raise AssertionError(f"run {run_id} did not finish within {timeout}s")


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
    # A few headline nodes the cross-domain path needs (P2 replaced the single
    # ss_scan with ten per-detector nodes; ss_scan must be gone).
    assert {"transcribe", "ss_text", "ss_color", "make_clips", "gate"} <= ids
    assert "ss_scan" not in ids
    # P2 catalog tranche additions.
    assert {"highlights", "multitool", "timelapse", "heatmap", "measure"} <= ids
    # P2 follow-ups: manual time-range source + the clipRecords→timeRange adapter.
    assert "time_range" in ids
    assert ("clipRecords", "timeRange") in workflows.ADAPTERS
    # The serialized catalog must carry no `execute` callable (JSON-safe).
    assert all("execute" not in node for node in catalog)
    # Launch-context flags drive palette grey-out; participants populate the
    # Video-Source dropdown (reusing the videoDir discovery call).
    assert set(data["context"]) == {"sheet", "videoDir", "participants"}
    assert data["context"]["sheet"] is False  # fixture sets _sheet_context=None
    assert isinstance(data["context"]["participants"], list)


def test_catalog_serves_adapter_pairs(wf_client):
    # The catalog endpoint serves the runner's ADAPTERS table (top-level, so the
    # context-shape assertion above is unaffected) as JSON-safe [src, dst] pairs,
    # so the frontend's canConnect accepts exactly the coercions the runner
    # applies in _gather_inputs. This is the UI↔runner parity guard.
    resp = wf_client.get("/workflows/api/catalog")
    adapters = resp.get_json()["adapters"]
    assert isinstance(adapters, list) and adapters
    assert all(isinstance(p, list) and len(p) == 2 for p in adapters)
    assert {tuple(p) for p in adapters} == set(workflows.ADAPTERS)


def test_serialize_catalog_is_json_safe():
    # serialize_catalog must round-trip through json without TypeErrors.
    json.dumps(workflows.serialize_catalog())


def test_serialize_adapters_matches_table():
    # serialize_adapters must be JSON-safe (no callables) and cover every key in
    # ADAPTERS — the frontend Set is then correct by construction.
    pairs = workflows.serialize_adapters()
    json.dumps(pairs)
    assert {tuple(p) for p in pairs} == set(workflows.ADAPTERS)


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


# ---- Run lifecycle (M4) ----


def test_run_executes_small_dag(wf_client, monkeypatch):
    # No real media/Whisper/Ollama: empty video paths short-circuit transcribe to
    # a stub, and Ollama is forced unavailable so summarize returns "".
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    import ollama_client

    monkeypatch.setattr(ollama_client, "is_available", lambda: False)

    bp_id = _make_blueprint(
        wf_client,
        nodes=[
            {"id": "v", "type": "video_source", "params": {"participant": "P01"}},
            {"id": "t", "type": "transcribe", "params": {}},
            {"id": "s", "type": "summarize", "params": {}},
        ],
        edges=[
            {"from": "v", "fromPort": "video", "to": "t", "toPort": "video"},
            {"from": "t", "fromPort": "transcript", "to": "s", "toPort": "transcript"},
        ],
    )
    created = wf_client.post("/workflows/api/runs", json={"blueprintId": bp_id})
    assert created.status_code == 200
    run = created.get_json()["run"]
    assert run["id"].startswith("run_")
    assert run["blueprintId"] == bp_id

    final = _wait_terminal(wf_client, run["id"])
    assert final["status"] == "completed"
    assert {n["status"] for n in final["nodeStates"].values()} == {"completed"}


def test_run_rejects_cycle_with_400(wf_client):
    bp_id = _make_blueprint(
        wf_client,
        nodes=[
            {"id": "a", "type": "gate", "params": {}},
            {"id": "b", "type": "gate", "params": {}},
        ],
        edges=[
            {"from": "a", "fromPort": "pass", "to": "b", "toPort": "value"},
            {"from": "b", "fromPort": "pass", "to": "a", "toPort": "value"},
        ],
    )
    resp = wf_client.post("/workflows/api/runs", json={"blueprintId": bp_id})
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


def test_run_missing_blueprint_404(wf_client):
    resp = wf_client.post("/workflows/api/runs", json={"blueprintId": "bp_nope"})
    assert resp.status_code == 404


def test_runs_list_and_filter(wf_client):
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "g", "type": "gate", "params": {}}]
    )
    run = wf_client.post("/workflows/api/runs", json={"blueprintId": bp_id}).get_json()[
        "run"
    ]
    _wait_terminal(wf_client, run["id"])

    listed = wf_client.get("/workflows/api/runs").get_json()["runs"]
    assert any(r["id"] == run["id"] for r in listed)
    filtered = wf_client.get(f"/workflows/api/runs?blueprintId={bp_id}").get_json()[
        "runs"
    ]
    assert filtered and all(r["blueprintId"] == bp_id for r in filtered)
    # An unrelated blueprint filter excludes it.
    assert (
        wf_client.get("/workflows/api/runs?blueprintId=bp_other").get_json()["runs"]
        == []
    )


def test_run_get_404_when_unknown(wf_client):
    assert wf_client.get("/workflows/api/runs/run_nope").status_code == 404


def test_run_cancel_404_when_finished(wf_client):
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "g", "type": "gate", "params": {}}]
    )
    run = wf_client.post("/workflows/api/runs", json={"blueprintId": bp_id}).get_json()[
        "run"
    ]
    _wait_terminal(wf_client, run["id"])
    # Still cancellable while the runner lives in _runs (no-op on a finished run),
    # but an unknown id 404s.
    assert wf_client.post("/workflows/api/runs/run_nope/cancel").status_code == 404


def test_run_stream_is_event_stream(wf_client):
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "g", "type": "gate", "params": {}}]
    )
    run = wf_client.post("/workflows/api/runs", json={"blueprintId": bp_id}).get_json()[
        "run"
    ]
    _wait_terminal(wf_client, run["id"])
    resp = wf_client.get(f"/workflows/api/runs/{run['id']}/stream")
    assert resp.mimetype == "text/event-stream"
    # Pull only the immediate first payload (the loop then blocks on keepalive).
    first = next(iter(resp.response))
    text = first.decode() if isinstance(first, bytes) else first
    assert text.startswith("data:")
    assert run["id"] in text
    resp.close()
