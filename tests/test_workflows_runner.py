"""Engine tests for the Workflows run engine (M4).

Exercises ``topo_order`` + ``WorkflowRunner`` directly (synchronously, no HTTP),
using ``config.DEBUGGING`` + a forced-unavailable Ollama so no Whisper / ffmpeg /
Ollama is needed. The HTTP surface is covered in tests/test_workflows_api.py.
"""

import json
import threading
from typing import Any

import pytest

import config
import workflows


def _ctx(tmp_path, **kw):
    kw.setdefault("cancel_event", threading.Event())
    return workflows.NodeContext(input_dir=tmp_path, output_dir=tmp_path, **kw)


def _runner(tmp_path, nodes, edges, **ctx_kw):
    return workflows.WorkflowRunner(
        "run_test",
        {"id": "bp", "nodes": nodes, "edges": edges},
        _ctx(tmp_path, **ctx_kw),
    )


# ---- topo_order ----


def test_topo_order_is_dependency_respecting():
    nodes = [{"id": "a"}, {"id": "b"}, {"id": "c"}]
    edges = [
        {"from": "a", "fromPort": "o", "to": "b", "toPort": "i"},
        {"from": "b", "fromPort": "o", "to": "c", "toPort": "i"},
    ]
    assert workflows.topo_order(nodes, edges) == ["a", "b", "c"]


def test_topo_order_rejects_cycle():
    nodes = [{"id": "a"}, {"id": "b"}]
    edges = [
        {"from": "a", "fromPort": "o", "to": "b", "toPort": "i"},
        {"from": "b", "fromPort": "o", "to": "a", "toPort": "i"},
    ]
    with pytest.raises(workflows.WorkflowCycleError):
        workflows.topo_order(nodes, edges)


def test_topo_order_ignores_stale_edges():
    # An edge to a node that no longer exists must not block the run.
    nodes = [{"id": "a"}]
    edges = [{"from": "ghost", "fromPort": "o", "to": "a", "toPort": "i"}]
    assert workflows.topo_order(nodes, edges) == ["a"]


# ---- bind_participant (P3 whole-study fan-out) ----


def test_bind_participant_rebinds_every_video_source_without_mutating_original():
    blueprint: dict[str, Any] = {
        "id": "bp",
        "nodes": [
            {"id": "v1", "type": "video_source", "params": {"participant": "P01"}},
            {"id": "v2", "type": "video_source", "params": {}},
            {"id": "t", "type": "transcribe", "params": {"foo": "bar"}},
        ],
        "edges": [],
    }
    bound = workflows.bind_participant(blueprint, "P07")
    # Every video_source is rebound; non-source nodes are untouched.
    by_id = {n["id"]: n for n in bound["nodes"]}
    assert by_id["v1"]["params"]["participant"] == "P07"
    assert by_id["v2"]["params"]["participant"] == "P07"
    assert by_id["t"]["params"] == {"foo": "bar"}
    # Pure: the original blueprint is never mutated (deep copy).
    orig_nodes: Any = blueprint["nodes"]
    orig_by_id = {n["id"]: n for n in orig_nodes}
    assert orig_by_id["v1"]["params"]["participant"] == "P01"
    assert "participant" not in orig_by_id["v2"]["params"]


def test_blueprint_participant_nodes_finds_only_video_sources():
    blueprint = {
        "nodes": [
            {"id": "v", "type": "video_source", "params": {}},
            {"id": "s", "type": "sheet_selection", "params": {}},
        ]
    }
    found = workflows.blueprint_participant_nodes(blueprint)
    assert [n["id"] for n in found] == ["v"]
    assert workflows.blueprint_participant_nodes({"nodes": []}) == []


def test_runner_snapshot_carries_batch_identity(tmp_path):
    runner = workflows.WorkflowRunner(
        "run_x",
        {"id": "bp", "nodes": [], "edges": []},
        _ctx(tmp_path),
        participant="P03",
        batch_id="batch_abc",
    )
    snap = runner.snapshot()
    assert snap["participant"] == "P03"
    assert snap["batchId"] == "batch_abc"
    # A normal single run carries blank batch identity and is not triggered.
    plain = workflows.WorkflowRunner(
        "run_y",
        {"id": "bp", "nodes": [], "edges": []},
        _ctx(tmp_path),
    )
    assert plain.snapshot()["batchId"] == ""
    assert plain.snapshot()["participant"] == ""
    assert plain.snapshot()["triggered"] is False


def test_runner_snapshot_carries_triggered_flag(tmp_path):
    runner = workflows.WorkflowRunner(
        "run_t",
        {"id": "bp", "nodes": [], "edges": []},
        _ctx(tmp_path),
        participant="P01",
        triggered=True,
    )
    assert runner.snapshot()["triggered"] is True


# ---- WorkflowRunner ----


def test_runner_executes_chain_to_completion(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    import ollama_client

    monkeypatch.setattr(ollama_client, "is_available", lambda: False)

    runner = _runner(
        tmp_path,
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
    runner.run()
    assert runner.status == "completed"
    assert {n["status"] for n in runner.node_states.values()} == {"completed"}


# ---- Per-node result sidecars (P5) ----


def test_runner_writes_node_result_sidecars(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    import ollama_client

    monkeypatch.setattr(ollama_client, "is_available", lambda: False)
    runner = _runner(
        tmp_path,
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
    runner.run()
    rdir = workflows.run_results_dir(tmp_path, "run_test")
    # Inspectable nodes (segments, summary) get a sidecar; the plumbing-only
    # video source (video/participant ports) does not.
    assert (rdir / "t.json").exists()
    assert (rdir / "s.json").exists()
    assert not (rdir / "v.json").exists()
    # The snapshot flags which nodes have a fetchable result.
    states = runner.snapshot()["nodeStates"]
    assert states["t"]["hasResult"] is True
    assert states["s"]["hasResult"] is True
    assert states["v"]["hasResult"] is False
    # The sidecar holds the inspectable payload, JSON-loadable.
    assert "summary" in json.loads((rdir / "s.json").read_text(encoding="utf-8"))


def test_inspectable_result_projects_and_filters_ports():
    # ss_text declares one output port `events` (inspectable). The heavy
    # raw_results rider is projected out; the events list survives.
    keep = workflows._inspectable_result(
        "ss_text",
        {"events": {"events": [{"time_in": 1}], "source": {}, "raw_results": [1, 2]}},
    )
    assert keep["events"]["events"] == [{"time_in": 1}]
    assert "raw_results" not in keep["events"]
    # sheet_selection outputs `clips` (clipRecords) — dropped, so gspread Cells
    # never reach a sidecar.
    assert (
        workflows._inspectable_result(
            "sheet_selection", {"clips": {"records": ["x"], "study": "s"}}
        )
        == {}
    )


def test_write_node_sidecar_guards_and_empty_payload(tmp_path):
    arts = {"artifacts": {"artifacts": [], "study": "", "count": 0}}
    # A node id that isn't a bare basename is rejected (no file, returns False).
    assert not workflows.write_node_sidecar(
        tmp_path, "run_x", "../escape", "make_clips", arts
    )
    # A node with no inspectable port writes nothing.
    assert not workflows.write_node_sidecar(
        tmp_path, "run_x", "v", "video_source", {"video": {"participant": "P01"}}
    )
    # A real inspectable payload writes a loadable sidecar.
    assert workflows.write_node_sidecar(tmp_path, "run_x", "m", "make_clips", arts)
    written = workflows.run_results_dir(tmp_path, "run_x") / "m.json"
    assert "artifacts" in json.loads(written.read_text(encoding="utf-8"))


def test_failed_node_skips_its_downstream(tmp_path, monkeypatch):
    # Force the summarize executor to raise; the run is 'failed' and any
    # downstream node is 'skipped' (none here, but the failed status propagates).
    def _boom(ctx, inputs, params):
        raise RuntimeError("kaboom")

    monkeypatch.setitem(workflows.NODE_TYPES["summarize"], "execute", _boom)
    runner = _runner(
        tmp_path,
        nodes=[
            {"id": "v", "type": "video_source", "params": {"participant": "P01"}},
            {"id": "t", "type": "transcribe", "params": {}},
            {"id": "s", "type": "summarize", "params": {}},
            {"id": "view", "type": "timeline_viewer", "params": {}},
        ],
        edges=[
            {"from": "v", "fromPort": "video", "to": "t", "toPort": "video"},
            {"from": "t", "fromPort": "transcript", "to": "s", "toPort": "transcript"},
            {"from": "s", "fromPort": "summary", "to": "view", "toPort": "segments"},
        ],
    )
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    runner.run()
    assert runner.status == "failed"
    assert runner.node_states["s"]["status"] == "failed"
    assert runner.node_states["s"]["error"]
    assert runner.node_states["view"]["status"] == "skipped"


def test_disabled_node_skips_itself_and_downstream(tmp_path, monkeypatch):
    # Muting 't' skips it and propagates SKIPPED to its dependent 's'; the run
    # still completes (a mute is not a failure) and the unrelated 'v' runs.
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)
    runner = _runner(
        tmp_path,
        nodes=[
            {"id": "v", "type": "video_source", "params": {"participant": "P01"}},
            {"id": "t", "type": "transcribe", "params": {}, "disabled": True},
            {"id": "s", "type": "summarize", "params": {}},
        ],
        edges=[
            {"from": "v", "fromPort": "video", "to": "t", "toPort": "video"},
            {"from": "t", "fromPort": "transcript", "to": "s", "toPort": "transcript"},
        ],
    )
    runner.run()
    assert runner.node_states["t"]["status"] == "skipped"
    assert runner.node_states["s"]["status"] == "skipped"
    assert runner.status == "completed"


def _gate_graph():
    return (
        [
            {"id": "v", "type": "video_source", "params": {"participant": "P01"}},
            {"id": "g", "type": "gate", "params": {"op": ">=", "threshold": 0}},
            {"id": "m", "type": "make_clips", "params": {}},
        ],
        [
            {"from": "v", "fromPort": "video", "to": "g", "toPort": "value"},
            {"from": "g", "fromPort": "pass", "to": "m", "toPort": "__gate__"},
            {"from": "v", "fromPort": "video", "to": "m", "toPort": "video"},
        ],
    )


def test_gate_false_skips_downstream_branch(tmp_path):
    nodes, edges = _gate_graph()
    # No real video -> duration 0; threshold 10 -> gate False -> 'm' skipped.
    nodes[1]["params"]["threshold"] = 10
    runner = _runner(tmp_path, nodes, edges)
    runner.run()
    assert runner.node_states["g"]["status"] == "completed"
    assert runner.node_states["m"]["status"] == "skipped"


def test_gate_true_runs_downstream_branch(tmp_path):
    nodes, edges = _gate_graph()
    nodes[1]["params"]["threshold"] = 0  # 0 >= 0 -> pass -> 'm' runs
    runner = _runner(tmp_path, nodes, edges)
    runner.run()
    assert runner.node_states["m"]["status"] == "completed"


def test_cancel_skips_remaining_nodes(tmp_path):
    nodes, edges = _gate_graph()
    cancel = threading.Event()
    cancel.set()  # pre-cancelled: every node is skipped, run is 'cancelled'
    runner = _runner(tmp_path, nodes, edges, cancel_event=cancel)
    runner.run()
    assert runner.status == "cancelled"
    assert all(s["status"] == "skipped" for s in runner.node_states.values())


def test_control_edges_are_excluded_from_inputs(tmp_path):
    # A gate's control edge establishes a dependency but never feeds data: 'm'
    # gets its 'video' input but not '__gate__'.
    nodes, edges = _gate_graph()
    runner = _runner(tmp_path, nodes, edges)
    runner._results["v"] = {"video": {"participant": "P01", "video_paths": []}}
    runner._results["g"] = {"pass": True}
    inputs = runner._gather_inputs({"id": "m", "type": "make_clips"})
    assert "video" in inputs
    assert "__gate__" not in inputs


def test_adapter_is_applied_on_type_mismatch(tmp_path):
    # ss_color emits `events`; make_clips consumes `clips` (clipRecords). The
    # runner must apply the events->clipRecords adapter when gathering inputs.
    nodes = [
        {"id": "a", "type": "ss_color", "params": {}},
        {"id": "b", "type": "make_clips", "params": {}},
    ]
    edges = [{"from": "a", "fromPort": "events", "to": "b", "toPort": "clips"}]
    runner = _runner(tmp_path, nodes, edges)
    runner._results["a"] = {
        "events": {
            "events": [
                {
                    "time_in": 1.0,
                    "time_out": 2.0,
                    "source_video": "study_P01.mp4",
                    "participant": "P01",
                }
            ],
            "source": {},
        }
    }
    inputs = runner._gather_inputs(nodes[1])
    assert "clips" in inputs
    assert "records" in inputs["clips"]  # adapter produced ClipRecords


def test_snapshot_is_json_safe_and_summarized(tmp_path):
    nodes, edges = _gate_graph()
    runner = _runner(tmp_path, nodes, edges)
    runner.run()
    snap = runner.snapshot()
    # Round-trips through json without TypeErrors (no raw frames / ndarrays).
    json.dumps(snap)
    assert set(snap) >= {"id", "blueprintId", "status", "nodeStates", "results"}
    assert all("started_at" in ns for ns in snap["nodeStates"].values())
