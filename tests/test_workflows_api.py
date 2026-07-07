"""Smoke tests for the Workflows Flask blueprint.

Verifies the page serves, the node catalog endpoint, blueprint CRUD (the canvas
autosave target), and the M4 run lifecycle (create/list/cancel/SSE), mirroring
tests/test_transcripts_api.py. The engine itself is covered in
tests/test_workflows_runner.py; here we exercise the HTTP surface.
"""

import json
import threading
import time

import pytest

Flask = pytest.importorskip("flask").Flask

import config  # noqa: E402
import utils  # noqa: E402
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
    # Reset run state so runners/SSE clients never leak across tests. The SSE
    # registries are the channel's live lists (make_sse_channel) — clear in place,
    # never rebind, or the notify/stream closures detach from the module name.
    workflows_server._runs = {}
    workflows_server._sse_clients.clear()
    workflows_server._batches = {}
    workflows_server._batch_sse_clients.clear()
    # Reset watch-dir trigger state (P6) so seen pids never leak across tests.
    workflows_server._watch_seen = set()
    workflows_server._watch_pending = {}
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
    assert set(data["context"]) == {"sheet", "videoDir", "participants", "outputDir"}
    assert data["context"]["sheet"] is False  # fixture sets _sheet_context=None
    assert isinstance(data["context"]["participants"], list)
    assert isinstance(data["context"]["outputDir"], str)  # run panel shows it


def test_catalog_serves_node_descriptions(wf_client):
    # Every node carries a non-empty description string — it drives the palette
    # tooltip and the on-card `?` help glyph, so a blank one ships a broken tip.
    catalog = wf_client.get("/workflows/api/catalog").get_json()["catalog"]
    missing = [n["id"] for n in catalog if not str(n.get("description", "")).strip()]
    assert not missing, f"nodes without a description: {missing}"


def test_catalog_participants_filtered_to_has_video(wf_client, monkeypatch):
    # The Video-Source dropdown (and the "All participants" fan-out) must only see
    # participants with an actual video file, matching the batch endpoint's filter.
    entries = [
        {"id": "P01", "has_video": True, "video_paths": ["a.mp4"]},
        {"id": "P02", "has_video": False, "video_paths": []},
    ]
    monkeypatch.setattr(utils, "discover_participant_videos", lambda *a, **k: entries)
    ctx = wf_client.get("/workflows/api/catalog").get_json()["context"]
    assert ctx["participants"] == ["P01"]
    assert ctx["videoDir"] is True


def test_catalog_serves_required_param_flag(wf_client):
    # P5: params whose empty value is a guaranteed no-op carry required:true, so
    # the client-side validation panel can flag them and disable Run.
    catalog = wf_client.get("/workflows/api/catalog").get_json()["catalog"]
    by_id = {n["id"]: n for n in catalog}

    def _param(node_id, name):
        return next(p for p in by_id[node_id]["params"] if p["name"] == name)

    assert _param("find_word", "word").get("required") is True
    assert _param("video_source", "participant").get("required") is True
    assert _param("ss_text", "search_string").get("required") is True
    # A param with a sane default stays optional (no required flag).
    assert "required" not in _param("find_word", "pad")


def test_catalog_flags_multitool_step_detectors(wf_client):
    # The Multitool step editor derives its step types from the catalog's
    # multitoolStep flag (no hardcoded JS list): the six per-frame detectors carry
    # it; the reference-based ones (similarity/template/scene) don't.
    catalog = wf_client.get("/workflows/api/catalog").get_json()["catalog"]
    by_id = {n["id"]: n for n in catalog}
    steps = {n["id"][3:] for n in catalog if n.get("multitoolStep")}
    assert steps == {"color", "change", "flow", "text", "numbers", "inactivity"}
    assert by_id["ss_template"].get("multitoolStep") is False


def test_catalog_serves_collection_ops(wf_client):
    # The collection-algebra control nodes (filter/merge/partition/limit/dedup) are
    # per-type families grouped under "Collection". Ports must be exact-typed
    # (same wire type in/out) so no adapter is needed, and the predicate value /
    # limit take params carry required:true for the validation panel.
    catalog = wf_client.get("/workflows/api/catalog").get_json()["catalog"]
    by_id = {n["id"]: n for n in catalog}

    expected = {
        f"{op}_{k}"
        for op in ("filter", "partition", "merge", "limit")
        for k in ("events", "clips", "segments", "timerange", "artifacts")
    }
    # dedup is span-based -> events + clips + time ranges (no segments/artifacts).
    expected |= {"dedup_events", "dedup_clips", "dedup_timerange"}
    assert expected <= set(by_id)

    # Every collection node is grouped under the "Collection" palette category.
    assert all(by_id[nid]["category"] == "Collection" for nid in expected)

    # The families must wire to the established pipeline by exact type: make_clips
    # outputs `artifacts` (so limit/filter_artifacts accept it), and timeRange
    # families sit between time-range sources and make_clips / ss detectors.
    assert by_id["limit_artifacts"]["inputs"][0]["type"] == "artifacts"
    assert by_id["make_clips"]["outputs"][0]["type"] == "artifacts"
    assert by_id["filter_timerange"]["outputs"][0]["type"] == "timeRange"

    # filter_events: events -> events (exact type, no coercion), value required.
    fe = by_id["filter_events"]
    assert [p["type"] for p in fe["inputs"]] == ["events"]
    assert [p["type"] for p in fe["outputs"]] == ["events"]
    assert next(p for p in fe["params"] if p["name"] == "value")["required"] is True

    # partition emits two outputs of the input's type (the gate's data-level else).
    pc = by_id["partition_clips"]
    assert [p["name"] for p in pc["outputs"]] == ["matched", "unmatched"]
    assert {p["type"] for p in pc["outputs"]} == {"clipRecords"}

    # merge takes 2-3 same-typed inputs (in2/in3 optional) into one output.
    ms = by_id["merge_segments"]
    assert [p["name"] for p in ms["inputs"]] == ["in1", "in2", "in3"]
    assert ms["inputs"][1].get("optional") and ms["inputs"][2].get("optional")
    assert {p["type"] for p in ms["inputs"]} == {"segments"}

    # limit's take is required; dedup is span-based (no segments family).
    assert (
        next(p for p in by_id["limit_events"]["params"] if p["name"] == "take")[
            "required"
        ]
        is True
    )
    assert "dedup_segments" not in by_id


def test_catalog_serves_adapter_pairs(wf_client):
    # The catalog endpoint serves the runner's ADAPTERS table (top-level, so the
    # context-shape assertion above is unaffected) as JSON-safe [src, dst] pairs,
    # so the frontend's canConnect accepts exactly the coercions the runner
    # applies in _gather_inputs. This is the UI↔runner parity guard.
    resp = wf_client.get("/workflows/api/catalog")
    adapters = resp.get_json()["adapters"]
    assert isinstance(adapters, list) and adapters
    # [src, dst, description] — description drives the coerced-wire tooltip.
    assert all(isinstance(p, list) and len(p) == 3 for p in adapters)
    assert {(p[0], p[1]) for p in adapters} == set(workflows.ADAPTERS)
    assert all(p[2].strip() for p in adapters)  # every coercion is explained


def test_serialize_catalog_is_json_safe():
    # serialize_catalog must round-trip through json without TypeErrors.
    json.dumps(workflows.serialize_catalog())


def test_serialize_adapters_matches_table():
    # serialize_adapters must be JSON-safe (no callables) and cover every key in
    # ADAPTERS — the frontend Set is then correct by construction.
    pairs = workflows.serialize_adapters()
    json.dumps(pairs)
    assert {(p[0], p[1]) for p in pairs} == set(workflows.ADAPTERS)
    # Every adapter carries a non-empty plain-language description.
    assert all(p[2].strip() for p in pairs)


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


def test_zero_interaction_launch_writes_no_manifest(wf_client, tmp_path):
    # The frontend auto-creates a bare "Untitled" blueprint on first load; that
    # node-less blueprint must not drop a manifest file in the output dir.
    resp = wf_client.post("/workflows/api/blueprints", json={})
    assert resp.status_code == 200
    assert not (tmp_path / config.WORKFLOWS_MANIFEST_FILENAME).exists()


def test_blueprint_with_node_persists_manifest(wf_client, tmp_path):
    _make_blueprint(wf_client, nodes=[{"id": "n1", "type": "video_source"}])
    assert (tmp_path / config.WORKFLOWS_MANIFEST_FILENAME).is_file()


def test_emptying_blueprint_removes_manifest(wf_client, tmp_path):
    bp_id = _make_blueprint(wf_client, nodes=[{"id": "n1", "type": "video_source"}])
    manifest = tmp_path / config.WORKFLOWS_MANIFEST_FILENAME
    assert manifest.is_file()
    # Clear the graph back to empty (the autosave PUT shape) → file reclaimed.
    wf_client.put(f"/workflows/api/blueprints/{bp_id}", json={"nodes": [], "edges": []})
    assert not manifest.exists()


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


# ---- Stash CRUD (M5) + built-in recipes (P4) ----


def test_stashes_default_to_builtins_only(wf_client):
    resp = wf_client.get("/workflows/api/stashes")
    assert resp.status_code == 200
    stashes = resp.get_json()["stashes"]
    # No user stashes saved yet → only the read-only built-in recipes, all first.
    assert stashes == workflows.BUILTIN_STASHES
    assert stashes and all(s["builtin"] for s in stashes)


def test_builtin_recipes_reference_only_real_nodes_and_ports(wf_client):
    # Each built-in recipe must wire real catalog node ids/ports, or instantiating
    # it would stamp a broken graph onto the canvas.
    for stash in workflows.BUILTIN_STASHES:
        by_id = {n["id"]: n for n in stash["nodes"]}
        for node in stash["nodes"]:
            assert node["type"] in workflows.NODE_TYPES
        for edge in stash["edges"]:
            src = by_id[edge["from"]]
            dst = by_id[edge["to"]]
            out_ports = {
                p["name"] for p in workflows.NODE_TYPES[src["type"]]["outputs"]
            }
            in_ports = {p["name"] for p in workflows.NODE_TYPES[dst["type"]]["inputs"]}
            assert edge["fromPort"] in out_ports
            assert edge["toPort"] in in_ports


def test_stash_create_returns_id_and_shape(wf_client):
    nodes = [
        {"id": "s1", "type": "transcribe", "params": {}, "position": {"x": 0, "y": 0}},
        {"id": "s2", "type": "find_word", "params": {}, "position": {"x": 200, "y": 0}},
    ]
    edges = [
        {
            "id": "e1",
            "from": "s1",
            "fromPort": "segments",
            "to": "s2",
            "toPort": "segments",
        }
    ]
    resp = wf_client.post(
        "/workflows/api/stashes",
        json={"name": "My chain", "nodes": nodes, "edges": edges},
    )
    assert resp.status_code == 200
    stash = resp.get_json()["stash"]
    assert stash["id"].startswith("stash_")
    assert stash["name"] == "My chain"
    assert stash["builtin"] is False
    assert "createdAt" in stash
    assert stash["nodes"] == nodes
    assert stash["edges"] == edges
    # Appears after the leading built-ins.
    listed = wf_client.get("/workflows/api/stashes").get_json()["stashes"]
    assert listed[: len(workflows.BUILTIN_STASHES)] == workflows.BUILTIN_STASHES
    assert any(s["id"] == stash["id"] for s in listed)


def test_stash_create_rejects_empty_nodes(wf_client):
    resp = wf_client.post("/workflows/api/stashes", json={"name": "Empty", "nodes": []})
    assert resp.status_code == 400


def test_stash_rename_round_trips(wf_client):
    nodes = [
        {"id": "s1", "type": "transcribe", "params": {}, "position": {"x": 0, "y": 0}}
    ]
    stash = wf_client.post(
        "/workflows/api/stashes", json={"name": "Before", "nodes": nodes}
    ).get_json()["stash"]
    resp = wf_client.put(
        f"/workflows/api/stashes/{stash['id']}", json={"name": "After"}
    )
    assert resp.status_code == 200
    listed = wf_client.get("/workflows/api/stashes").get_json()["stashes"]
    saved = next(s for s in listed if s["id"] == stash["id"])
    assert saved["name"] == "After"


def test_stash_delete_removes_it(wf_client):
    nodes = [
        {"id": "s1", "type": "transcribe", "params": {}, "position": {"x": 0, "y": 0}}
    ]
    stash = wf_client.post(
        "/workflows/api/stashes", json={"name": "Doomed", "nodes": nodes}
    ).get_json()["stash"]
    assert wf_client.delete(f"/workflows/api/stashes/{stash['id']}").status_code == 200
    listed = wf_client.get("/workflows/api/stashes").get_json()["stashes"]
    assert all(s["id"] != stash["id"] for s in listed)


def test_builtin_stash_is_read_only(wf_client):
    builtin_id = workflows.BUILTIN_STASHES[0]["id"]
    assert (
        wf_client.put(
            f"/workflows/api/stashes/{builtin_id}", json={"name": "Hijack"}
        ).status_code
        == 403
    )
    assert wf_client.delete(f"/workflows/api/stashes/{builtin_id}").status_code == 403


def test_stash_update_and_delete_404_on_missing(wf_client):
    assert (
        wf_client.put(
            "/workflows/api/stashes/stash_missing", json={"name": "x"}
        ).status_code
        == 404
    )
    assert wf_client.delete("/workflows/api/stashes/stash_missing").status_code == 404


def test_stash_write_preserves_blueprints_and_runs(wf_client):
    # The combined manifest holds blueprints + stashes + runs; a stash write must
    # read-modify-write under the shared lock without clobbering its siblings.
    bp = wf_client.post(
        "/workflows/api/blueprints", json={"name": "Keep me"}
    ).get_json()["blueprint"]
    workflows_server._manifest.setdefault("runs", []).append(
        {"id": "run_1", "blueprintId": bp["id"], "status": "completed"}
    )
    nodes = [
        {"id": "s1", "type": "transcribe", "params": {}, "position": {"x": 0, "y": 0}}
    ]
    wf_client.post("/workflows/api/stashes", json={"name": "Side", "nodes": nodes})
    # Blueprint and run survive the stash write.
    listed = wf_client.get("/workflows/api/blueprints").get_json()["blueprints"]
    assert any(b["id"] == bp["id"] for b in listed)
    assert any(r["id"] == "run_1" for r in workflows_server._manifest["runs"])


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


def test_run_rejects_unknown_target_node_with_400(wf_client):
    # A stale "run to here" target must error, not silently run the whole graph.
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "g", "type": "gate", "params": {}}]
    )
    resp = wf_client.post(
        "/workflows/api/runs",
        json={"blueprintId": bp_id, "targetNodeId": "ghost"},
    )
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


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
    # A finished runner is evicted from _runs (P3 leak fix), so cancelling it — or
    # an unknown id — 404s.
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


def test_runner_evicted_from_runs_after_terminal(wf_client):
    # The leak fix (P3): a terminal runner is dropped from _runs once persisted;
    # the snapshot is still served from the manifest via _run_snapshot.
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "g", "type": "gate", "params": {}}]
    )
    run = wf_client.post("/workflows/api/runs", json={"blueprintId": bp_id}).get_json()[
        "run"
    ]
    final = _wait_terminal(wf_client, run["id"])
    deadline = time.monotonic() + 2.0  # let the finalize thread run the eviction
    while time.monotonic() < deadline and run["id"] in workflows_server._runs:
        time.sleep(0.02)
    assert run["id"] not in workflows_server._runs
    served = wf_client.get(f"/workflows/api/runs/{run['id']}").get_json()["run"]
    assert served["status"] == final["status"]


# ---- Batch lifecycle (P3: whole-study fan-out) ----


def _mock_participants(monkeypatch, ids=("P01", "P02", "P03")):
    """Force participant discovery to a fixed set with (empty) video paths.

    Empty ``video_paths`` keep the executors in their DEBUGGING stub path — no real
    media is needed to fan a blueprint out across participants.
    """
    entries = [{"id": i, "has_video": True, "video_paths": []} for i in ids]
    monkeypatch.setattr(utils, "discover_participant_videos", lambda *a, **k: entries)
    return list(ids)


def _wait_batch_terminal(client, batch_id, timeout=5.0):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        batch = client.get(f"/workflows/api/batches/{batch_id}").get_json()["batch"]
        if batch["status"] in ("completed", "failed", "cancelled"):
            return batch
        time.sleep(0.02)
    raise AssertionError(f"batch {batch_id} did not finish within {timeout}s")


def test_batch_runs_every_participant(wf_client, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    parts = _mock_participants(monkeypatch)
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "v", "type": "video_source", "params": {}}]
    )
    created = wf_client.post("/workflows/api/batches", json={"blueprintId": bp_id})
    assert created.status_code == 200
    batch = created.get_json()["batch"]
    assert batch["id"].startswith("batch_")
    assert set(batch["participants"]) == set(parts)

    final = _wait_batch_terminal(wf_client, batch["id"])
    assert final["status"] == "completed"
    assert final["counts"].get("completed") == len(parts)
    # Drill-in returns one run per participant, each tagged with the batch.
    detail = wf_client.get(f"/workflows/api/batches/{batch['id']}").get_json()
    runs = detail["runs"]
    assert {r["participant"] for r in runs} == set(parts)
    assert all(r["batchId"] == batch["id"] for r in runs)


def test_batch_honors_participant_subset(wf_client, monkeypatch):
    # The multi-select widget can POST an explicit `participants` subset; the
    # batch must fan out over only those (intersected with with-video ids), not
    # every participant. Regression guard for the Phase 5 frontend wiring.
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    _mock_participants(monkeypatch, ids=("P01", "P02", "P03"))
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "v", "type": "video_source", "params": {}}]
    )
    subset = ["P01", "P03"]
    created = wf_client.post(
        "/workflows/api/batches",
        json={"blueprintId": bp_id, "participants": subset},
    )
    assert created.status_code == 200
    batch = created.get_json()["batch"]
    assert set(batch["participants"]) == set(subset)

    final = _wait_batch_terminal(wf_client, batch["id"])
    assert final["status"] == "completed"
    assert final["counts"].get("completed") == len(subset)
    detail = wf_client.get(f"/workflows/api/batches/{batch['id']}").get_json()
    assert {r["participant"] for r in detail["runs"]} == set(subset)


def test_batch_seeds_sheet_selection_once(wf_client, monkeypatch):
    # sheet_selection is participant-independent and hits the rate-limited Sheets
    # API; a batch must compute it once, not once per participant.
    import spreadsheet

    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    parts = _mock_participants(monkeypatch, ids=("P01", "P02", "P03"))

    class _Sheet:
        study_name = "study"

    monkeypatch.setattr(workflows_server, "_sheet_context", _Sheet())
    calls = {"n": 0}

    def _fake_generate_list(*a, **k):
        calls["n"] += 1
        return []

    monkeypatch.setattr(spreadsheet, "generate_list", _fake_generate_list)

    bp_id = _make_blueprint(
        wf_client,
        nodes=[
            {"id": "v", "type": "video_source", "params": {}},
            {"id": "s", "type": "sheet_selection", "params": {"selector": "1"}},
        ],
    )
    batch = wf_client.post(
        "/workflows/api/batches", json={"blueprintId": bp_id}
    ).get_json()["batch"]
    final = _wait_batch_terminal(wf_client, batch["id"])
    assert final["status"] == "completed"
    assert len(parts) == 3
    # Seeded once for the whole batch, not once per participant.
    assert calls["n"] == 1


def test_batch_continues_when_one_participant_fails(wf_client, monkeypatch):
    # Continue-on-error is mandatory: one bad participant must not sink the batch.
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    _mock_participants(monkeypatch, ids=("P01", "P02", "P03"))
    orig = workflows.NODE_TYPES["video_source"]["execute"]

    def flaky(ctx, inputs, params):
        if params.get("participant") == "P02":
            raise RuntimeError("boom")
        return orig(ctx, inputs, params)

    monkeypatch.setitem(workflows.NODE_TYPES["video_source"], "execute", flaky)
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "v", "type": "video_source", "params": {}}]
    )
    batch = wf_client.post(
        "/workflows/api/batches", json={"blueprintId": bp_id}
    ).get_json()["batch"]
    final = _wait_batch_terminal(wf_client, batch["id"])
    by_part = {c["participant"]: c["status"] for c in final["children"]}
    assert by_part["P01"] == "completed"
    assert by_part["P03"] == "completed"
    assert by_part["P02"] == "failed"
    assert final["status"] == "failed"


def test_batch_cancel_short_circuits_remaining(wf_client, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    _mock_participants(monkeypatch, ids=("P01", "P02", "P03"))
    started = threading.Event()
    orig = workflows.NODE_TYPES["video_source"]["execute"]

    def blocker(ctx, inputs, params):
        # Hold the first child until the batch is cancelled (cancel sets the ctx
        # cancel event via the runner), so P02/P03 never start.
        if params.get("participant") == "P01":
            started.set()
            for _ in range(500):
                if ctx.cancel_flag():
                    break
                time.sleep(0.01)
        return orig(ctx, inputs, params)

    monkeypatch.setitem(workflows.NODE_TYPES["video_source"], "execute", blocker)
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "v", "type": "video_source", "params": {}}]
    )
    batch = wf_client.post(
        "/workflows/api/batches", json={"blueprintId": bp_id}
    ).get_json()["batch"]
    assert started.wait(timeout=5)
    assert (
        wf_client.post(f"/workflows/api/batches/{batch['id']}/cancel").status_code
        == 200
    )
    final = _wait_batch_terminal(wf_client, batch["id"])
    assert final["status"] == "cancelled"
    by_part = {c["participant"]: c["status"] for c in final["children"]}
    assert by_part["P02"] == "cancelled"
    assert by_part["P03"] == "cancelled"


def test_batch_400_when_no_video_source(wf_client, monkeypatch):
    _mock_participants(monkeypatch)
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "g", "type": "gate", "params": {}}]
    )
    resp = wf_client.post("/workflows/api/batches", json={"blueprintId": bp_id})
    assert resp.status_code == 400


def test_batch_400_when_no_participants(wf_client, monkeypatch):
    monkeypatch.setattr(utils, "discover_participant_videos", lambda *a, **k: [])
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "v", "type": "video_source", "params": {}}]
    )
    resp = wf_client.post("/workflows/api/batches", json={"blueprintId": bp_id})
    assert resp.status_code == 400


def test_batch_404_when_blueprint_missing(wf_client):
    resp = wf_client.post("/workflows/api/batches", json={"blueprintId": "bp_nope"})
    assert resp.status_code == 404


def test_batches_list_and_filter(wf_client, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    _mock_participants(monkeypatch, ids=("P01", "P02"))
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "v", "type": "video_source", "params": {}}]
    )
    batch = wf_client.post(
        "/workflows/api/batches", json={"blueprintId": bp_id}
    ).get_json()["batch"]
    _wait_batch_terminal(wf_client, batch["id"])

    listed = wf_client.get("/workflows/api/batches").get_json()["batches"]
    assert any(b["id"] == batch["id"] for b in listed)
    filtered = wf_client.get(f"/workflows/api/batches?blueprintId={bp_id}").get_json()[
        "batches"
    ]
    assert filtered and all(b["blueprintId"] == bp_id for b in filtered)
    assert (
        wf_client.get("/workflows/api/batches?blueprintId=bp_other").get_json()[
            "batches"
        ]
        == []
    )


def test_batch_summary_historical_from_persisted_runs(wf_client):
    """A finished batch (no live record) is rebuilt from persisted runs by reading
    scalars only, and its summary stays stable as unrelated history grows."""
    workflows_server._manifest["runs"] = [
        {
            "id": "run_a",
            "batchId": "batch_x",
            "participant": "P01",
            "blueprintId": "bp1",
            "startedAt": "2026-01-01T00:00:00",
            "status": workflows.RUN_STATUS_COMPLETED,
        },
        {
            "id": "run_b",
            "batchId": "batch_x",
            "participant": "P02",
            "blueprintId": "bp1",
            "startedAt": "2026-01-01T00:00:01",
            "status": workflows.RUN_STATUS_FAILED,
        },
    ]
    summary = workflows_server._batch_summary("batch_x")
    assert summary is not None
    assert summary["blueprintId"] == "bp1"
    assert summary["participants"] == ["P01", "P02"]
    assert summary["counts"] == {
        workflows.RUN_STATUS_COMPLETED: 1,
        workflows.RUN_STATUS_FAILED: 1,
    }
    child_status = {c["runId"]: c["status"] for c in summary["children"]}
    assert child_status == {
        "run_a": workflows.RUN_STATUS_COMPLETED,
        "run_b": workflows.RUN_STATUS_FAILED,
    }

    # Growing unrelated run history must not perturb this batch's summary.
    workflows_server._manifest["runs"].append(
        {
            "id": "run_c",
            "batchId": "batch_other",
            "participant": "P09",
            "blueprintId": "bp2",
            "startedAt": "2026-01-02T00:00:00",
            "status": workflows.RUN_STATUS_RUNNING,
        }
    )
    assert workflows_server._batch_summary("batch_x") == summary
    assert workflows_server._batch_summary("batch_missing") is None


def test_batch_get_and_cancel_404_when_unknown(wf_client):
    assert wf_client.get("/workflows/api/batches/batch_nope").status_code == 404
    assert wf_client.post("/workflows/api/batches/batch_nope/cancel").status_code == 404


def test_batch_stream_is_event_stream(wf_client, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    _mock_participants(monkeypatch, ids=("P01",))
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "v", "type": "video_source", "params": {}}]
    )
    batch = wf_client.post(
        "/workflows/api/batches", json={"blueprintId": bp_id}
    ).get_json()["batch"]
    _wait_batch_terminal(wf_client, batch["id"])
    resp = wf_client.get(f"/workflows/api/batches/{batch['id']}/stream")
    assert resp.mimetype == "text/event-stream"
    first = next(iter(resp.response))
    text = first.decode() if isinstance(first, bytes) else first
    assert text.startswith("data:")
    assert batch["id"] in text
    resp.close()


def test_batch_children_survive_run_history_cap(wf_client, monkeypatch):
    # The history cap evicts whole units (a batch's children together), never
    # splitting a batch — so a batch larger than the cap keeps all its children and
    # older loose runs are dropped instead. (Regression for the per-record cap that
    # evicted a batch's own not-yet-finished children.)
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    monkeypatch.setattr(workflows_server, "_MAX_RUN_HISTORY", 3, raising=False)

    # Three loose single runs fill the cap first.
    gate_bp = _make_blueprint(
        wf_client, nodes=[{"id": "g", "type": "gate", "params": {}}]
    )
    loose_ids = []
    for _ in range(3):
        r = wf_client.post(
            "/workflows/api/runs", json={"blueprintId": gate_bp}
        ).get_json()["run"]
        _wait_terminal(wf_client, r["id"])
        loose_ids.append(r["id"])

    # A 2-participant batch then runs; it must stay intact past the cap.
    parts = _mock_participants(monkeypatch, ids=("P01", "P02"))
    bp_id = _make_blueprint(
        wf_client, nodes=[{"id": "v", "type": "video_source", "params": {}}]
    )
    batch = wf_client.post(
        "/workflows/api/batches", json={"blueprintId": bp_id}
    ).get_json()["batch"]
    final = _wait_batch_terminal(wf_client, batch["id"])

    # All children survived and are individually retrievable.
    assert {c["participant"] for c in final["children"]} == set(parts)
    detail = wf_client.get(f"/workflows/api/batches/{batch['id']}").get_json()
    assert len(detail["runs"]) == len(parts)
    for child in final["children"]:
        assert wf_client.get(f"/workflows/api/runs/{child['runId']}").status_code == 200
    # The oldest loose runs were evicted to make room (cap held by dropping units).
    assert wf_client.get(f"/workflows/api/runs/{loose_ids[0]}").status_code == 404


# ---- Per-node result sidecars (P5) ----


def test_node_result_endpoint_serves_sidecar(wf_client, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    # A lone make_clips node completes with an (empty) artifacts result — an
    # inspectable port, so a sidecar is written and flagged on the snapshot.
    bp = _make_blueprint(
        wf_client, nodes=[{"id": "m", "type": "make_clips", "params": {}}]
    )
    run = wf_client.post("/workflows/api/runs", json={"blueprintId": bp}).get_json()[
        "run"
    ]
    final = _wait_terminal(wf_client, run["id"])
    assert final["nodeStates"]["m"]["hasResult"] is True

    res = wf_client.get(f"/workflows/api/runs/{run['id']}/nodes/m/result")
    assert res.status_code == 200
    payload = res.get_json()
    assert payload["ok"] is True
    assert "artifacts" in payload["result"]
    # Unknown node / unknown run → 404 (no sidecar on disk).
    assert (
        wf_client.get(f"/workflows/api/runs/{run['id']}/nodes/ghost/result").status_code
        == 404
    )
    assert wf_client.get("/workflows/api/runs/nope/nodes/m/result").status_code == 404


def test_node_result_sidecars_pruned_with_history(wf_client, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    monkeypatch.setattr(workflows_server, "_MAX_RUN_HISTORY", 2, raising=False)
    bp = _make_blueprint(
        wf_client, nodes=[{"id": "m", "type": "make_clips", "params": {}}]
    )
    base = utils.get_effective_output_dir()
    run_ids = []
    for _ in range(3):
        r = wf_client.post("/workflows/api/runs", json={"blueprintId": bp}).get_json()[
            "run"
        ]
        _wait_terminal(wf_client, r["id"])
        run_ids.append(r["id"])
    # The oldest run was evicted past the cap; its sidecar dir is pruned in
    # lockstep, while the newest survives.
    assert not workflows.run_results_dir(base, run_ids[0]).exists()
    assert workflows.run_results_dir(base, run_ids[-1]).exists()
    assert (
        wf_client.get(f"/workflows/api/runs/{run_ids[0]}/nodes/m/result").status_code
        == 404
    )


# ---- Watch-dir triggers (P6) ----


def _video_source_blueprint(client, participant="P01"):
    """An armable blueprint: a lone Video Source (has a source, no cycle)."""
    return _make_blueprint(
        client,
        nodes=[
            {"id": "v", "type": "video_source", "params": {"participant": participant}}
        ],
    )


def _entry(pid, has_video=True):
    """A discover_participant_videos entry whose first path equals the pid."""
    return {"id": pid, "has_video": has_video, "video_paths": [pid]}


def _record_launches(monkeypatch):
    """Replace _launch_run with a recorder so no real runner threads spawn."""
    calls = []

    def _fake(blueprint, participant="", triggered=False):
        calls.append((blueprint, participant, triggered))
        return {}

    monkeypatch.setattr(workflows_server, "_launch_run", _fake)
    return calls


def _mock_discovery(monkeypatch, entries, stats):
    monkeypatch.setattr(utils, "discover_participant_videos", lambda *a, **k: entries)
    monkeypatch.setattr(
        workflows_server,
        "_stat_first_video",
        lambda paths: stats.get(paths[0]) if paths else None,
    )


def test_trigger_arm_and_disarm_round_trip(wf_client):
    bp = _video_source_blueprint(wf_client)
    res = wf_client.put(
        f"/workflows/api/blueprints/{bp}/trigger", json={"enabled": True}
    )
    assert res.status_code == 200
    assert res.get_json()["blueprint"]["trigger"] == {
        "type": "watch_dir",
        "enabled": True,
    }
    listed = wf_client.get("/workflows/api/blueprints").get_json()["blueprints"]
    assert next(b for b in listed if b["id"] == bp)["trigger"]["enabled"] is True

    res2 = wf_client.put(
        f"/workflows/api/blueprints/{bp}/trigger", json={"enabled": False}
    )
    assert res2.get_json()["blueprint"]["trigger"]["enabled"] is False


def test_trigger_is_single_active(wf_client):
    a = _video_source_blueprint(wf_client)
    b = _video_source_blueprint(wf_client)
    wf_client.put(f"/workflows/api/blueprints/{a}/trigger", json={"enabled": True})
    wf_client.put(f"/workflows/api/blueprints/{b}/trigger", json={"enabled": True})
    blueprints = {
        x["id"]: x
        for x in wf_client.get("/workflows/api/blueprints").get_json()["blueprints"]
    }
    # Arming b disarmed a (only one blueprint may watch at a time).
    assert blueprints[a]["trigger"]["enabled"] is False
    assert blueprints[b]["trigger"]["enabled"] is True


def test_trigger_arm_rejects_cycle(wf_client):
    bp = _make_blueprint(
        wf_client,
        nodes=[
            {"id": "v", "type": "video_source", "params": {"participant": "P01"}},
            {"id": "a", "type": "gate", "params": {}},
            {"id": "b", "type": "gate", "params": {}},
        ],
        edges=[
            {"from": "a", "fromPort": "pass", "to": "b", "toPort": "value"},
            {"from": "b", "fromPort": "pass", "to": "a", "toPort": "value"},
        ],
    )
    res = wf_client.put(
        f"/workflows/api/blueprints/{bp}/trigger", json={"enabled": True}
    )
    assert res.status_code == 400


def test_trigger_arm_rejects_no_video_source(wf_client):
    bp = _make_blueprint(wf_client, nodes=[{"id": "g", "type": "gate", "params": {}}])
    res = wf_client.put(
        f"/workflows/api/blueprints/{bp}/trigger", json={"enabled": True}
    )
    assert res.status_code == 400


def test_trigger_404_on_missing_blueprint(wf_client):
    res = wf_client.put(
        "/workflows/api/blueprints/bp_nope/trigger", json={"enabled": True}
    )
    assert res.status_code == 404


def test_watch_seed_marks_existing_seen(wf_client, monkeypatch):
    monkeypatch.setattr(
        utils,
        "discover_participant_videos",
        lambda *a, **k: [_entry("P01"), _entry("P02")],
    )
    workflows_server._seed_watch_seen()
    assert workflows_server._watch_seen == {"P01", "P02"}


def test_watch_fires_after_two_stable_polls(wf_client, monkeypatch):
    calls = _record_launches(monkeypatch)
    bp = _video_source_blueprint(wf_client)
    wf_client.put(f"/workflows/api/blueprints/{bp}/trigger", json={"enabled": True})
    _mock_discovery(monkeypatch, [_entry("P01")], {"P01": (100, 1.0)})

    workflows_server._watch_poll_once()  # records pending — no fire on first sight
    assert calls == []
    workflows_server._watch_poll_once()  # stable -> fires once
    assert len(calls) == 1
    assert calls[0][1] == "P01" and calls[0][2] is True
    workflows_server._watch_poll_once()  # already seen -> no refire
    assert len(calls) == 1


def test_watch_unstable_stat_does_not_fire(wf_client, monkeypatch):
    calls = _record_launches(monkeypatch)
    bp = _video_source_blueprint(wf_client)
    wf_client.put(f"/workflows/api/blueprints/{bp}/trigger", json={"enabled": True})
    sizes = iter([(100, 1.0), (200, 2.0)])  # still being copied between polls
    monkeypatch.setattr(
        utils, "discover_participant_videos", lambda *a, **k: [_entry("P01")]
    )
    monkeypatch.setattr(
        workflows_server, "_stat_first_video", lambda paths: next(sizes)
    )
    workflows_server._watch_poll_once()
    workflows_server._watch_poll_once()
    assert calls == []  # never stable across two polls


def test_watch_gated_on_single_armed(wf_client, monkeypatch):
    calls = _record_launches(monkeypatch)
    # No blueprint armed at all -> the poll skips all work (no glob/stat) and fires
    # nothing. No-retro-fire is handled by the arm-time re-seed, not here.
    _mock_discovery(monkeypatch, [_entry("P01")], {"P01": (100, 1.0)})
    workflows_server._watch_poll_once()
    workflows_server._watch_poll_once()
    assert calls == []


def test_arming_reseeds_so_backlog_never_fires(wf_client, monkeypatch):
    # A pid already present when the blueprint is armed is re-baselined as seen at
    # arm time, so it never retro-fires (the poll no longer maintains the seen-set
    # while disarmed).
    calls = _record_launches(monkeypatch)
    bp = _video_source_blueprint(wf_client)
    _mock_discovery(monkeypatch, [_entry("P01")], {"P01": (100, 1.0)})
    wf_client.put(f"/workflows/api/blueprints/{bp}/trigger", json={"enabled": True})
    assert "P01" in workflows_server._watch_seen  # seeded at arm time
    workflows_server._watch_poll_once()
    workflows_server._watch_poll_once()
    assert calls == []


def test_watch_two_armed_is_ambiguous_and_skips(wf_client, monkeypatch):
    calls = _record_launches(monkeypatch)
    a = _video_source_blueprint(wf_client)
    b = _video_source_blueprint(wf_client)
    # Force both armed directly (bypassing the single-active write) to prove the
    # watcher defends against an inconsistent manifest.
    for bid in (a, b):
        bp = next(x for x in workflows_server._manifest["blueprints"] if x["id"] == bid)
        bp["trigger"] = {"type": "watch_dir", "enabled": True}
    _mock_discovery(monkeypatch, [_entry("P01")], {"P01": (100, 1.0)})
    workflows_server._watch_poll_once()
    workflows_server._watch_poll_once()
    assert calls == []


def test_watch_disarmed_arrival_not_retrofired(wf_client, monkeypatch):
    calls = _record_launches(monkeypatch)
    bp = _video_source_blueprint(wf_client)  # not armed
    _mock_discovery(monkeypatch, [_entry("P01")], {"P01": (100, 1.0)})
    workflows_server._watch_poll_once()
    workflows_server._watch_poll_once()  # disarmed -> no-op
    assert calls == []
    # Arming re-seeds the now-present P01, so it isn't retro-fired.
    wf_client.put(f"/workflows/api/blueprints/{bp}/trigger", json={"enabled": True})
    workflows_server._watch_poll_once()
    assert calls == []


def test_watch_seeded_existing_never_fires(wf_client, monkeypatch):
    calls = _record_launches(monkeypatch)
    bp = _video_source_blueprint(wf_client)
    wf_client.put(f"/workflows/api/blueprints/{bp}/trigger", json={"enabled": True})
    _mock_discovery(monkeypatch, [_entry("P01")], {"P01": (100, 1.0)})
    workflows_server._seed_watch_seen()  # P01 present at startup
    workflows_server._watch_poll_once()
    workflows_server._watch_poll_once()
    assert calls == []  # the backlog is never auto-run


def test_watch_multipart_fires_once(wf_client, monkeypatch):
    calls = _record_launches(monkeypatch)
    bp = _video_source_blueprint(wf_client)
    wf_client.put(f"/workflows/api/blueprints/{bp}/trigger", json={"enabled": True})
    entry = {"id": "P01", "has_video": True, "video_paths": ["P01-1", "P01-2"]}
    _mock_discovery(monkeypatch, [entry], {"P01-1": (100, 1.0)})
    for _ in range(4):
        workflows_server._watch_poll_once()
    assert len(calls) == 1  # one entry per pid -> a single fire
