"""Direct-invocation tests for the Workflows node executors and typed-port
adapters (M3).

The M4 ``WorkflowRunner`` does not exist yet, so each executor is exercised by
calling ``NODE_TYPES[id]["execute"](ctx, inputs, params)`` directly. Whisper /
ffmpeg paths run under ``config.DEBUGGING`` or with ffmpeg mocked; Ollama and
Screenspace scans are monkeypatched so no model/subprocess/network is touched.
"""

from pathlib import Path
from types import SimpleNamespace
from typing import Any
from unittest.mock import Mock

import config
import files
import screenspace
import utils
import video
import viewer
import workflows


def _ctx(tmp_path, **kw):
    return workflows.NodeContext(input_dir=tmp_path, output_dir=tmp_path, **kw)


def _run(node_id, ctx, inputs, params):
    return workflows.NODE_TYPES[node_id]["execute"](ctx, inputs, params)


# ---- Pure executors ----


def test_gate_compares_scalar(tmp_path):
    ctx = _ctx(tmp_path)
    assert _run("gate", ctx, {"value": 300.0}, {"op": ">=", "threshold": 120})["pass"]
    assert not _run("gate", ctx, {"value": 60.0}, {"op": ">=", "threshold": 120})[
        "pass"
    ]
    # Non-numeric value fails closed rather than raising.
    assert not _run("gate", ctx, {"value": None}, {"op": ">", "threshold": 0})["pass"]


def test_gate_collection_measures_and_gates(tmp_path):
    ctx = _ctx(tmp_path)
    evs = {
        "events": {
            "events": [{"time_in": 0, "time_out": 1}, {"time_in": 2, "time_out": 3}]
        }
    }
    assert _run(
        "gate_collection", ctx, evs, {"metric": "count", "op": ">=", "threshold": 2}
    )["pass"]
    assert not _run(
        "gate_collection", ctx, evs, {"metric": "count", "op": ">=", "threshold": 3}
    )["pass"]
    # Nothing wired -> value 0 -> fails a positive threshold (fails closed).
    assert not _run(
        "gate_collection", ctx, {}, {"metric": "count", "op": ">", "threshold": 0}
    )["pass"]


def test_find_word_filters_segments_and_pads(tmp_path):
    ctx = _ctx(tmp_path)
    segs = {
        "segments": [
            {"start": 10.0, "end": 12.0, "text": "please open the Menu"},
            {"start": 30.0, "end": 32.0, "text": "no match here"},
        ],
        "source": {"participant": "P01"},
    }
    out = _run("find_word", ctx, {"segments": segs}, {"word": "menu", "pad": 2})
    assert out["timeRange"]["ranges"] == [(8.0, 14.0)]
    assert out["timestamps"]["times"] == [10.0]
    # Source descriptor carried through for downstream clip cutting.
    assert out["timeRange"]["source"] == {"participant": "P01"}


def test_video_source_resolves_participant(tmp_path, monkeypatch):
    monkeypatch.setattr(
        workflows.utils,
        "discover_participant_videos",
        lambda *a, **k: [
            {"id": "P01", "video_paths": ["/v/study_P01.mp4"], "has_video": True}
        ],
    )
    out = _run("video_source", _ctx(tmp_path), {}, {"participant": "P01"})
    assert out["participant"] == "P01"
    assert out["video"]["source_filename"] == "study_P01.mp4"
    assert out["video"]["study"] == "study"
    assert out["video"]["video_paths"] == ["/v/study_P01.mp4"]


# ---- Transcript ----


def test_transcribe_returns_segments_and_carries_source(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True)
    src = {
        "participant": "P01",
        "study": "study",
        "source_filename": "study_P01.mp4",
        "video_paths": [str(tmp_path / "study_P01.mp4")],
    }
    out = _run("transcribe", _ctx(tmp_path), {"video": src}, {"language": "auto"})
    assert "segments" in out["transcript"]
    # Both outputs carry the originating source descriptor (same object).
    assert out["transcript"]["source"] is src
    assert out["segments"]["source"] is src


# ---- Thinking (Ollama) ----


def test_thinking_executors_empty_when_ollama_unavailable(tmp_path, monkeypatch):
    import ollama_client

    monkeypatch.setattr(ollama_client, "is_available", lambda: False)
    ctx = _ctx(tmp_path)
    segs = {"segments": [{"start": 0, "end": 1, "text": "hi"}], "source": {}}
    tr = {"segments": [{"start": 0, "end": 1, "text": "hi"}]}
    assert _run("summarize", ctx, {"transcript": tr}, {})["summary"] == ""
    assert (
        _run("citations", ctx, {"summary": "s", "segments": segs}, {})["citations"]
        == []
    )
    assert _run("friction", ctx, {"segments": segs}, {})["friction"] == []


def test_summarize_wires_thinking_agent(tmp_path, monkeypatch):
    import ollama_client
    import thinking_agents

    monkeypatch.setattr(ollama_client, "is_available", lambda: True)
    monkeypatch.setattr(
        thinking_agents, "summarize_transcript", lambda segments, **kw: "the summary"
    )
    out = _run(
        "summarize", _ctx(tmp_path), {"transcript": {"segments": [{"text": "x"}]}}, {}
    )
    assert out["summary"] == "the summary"


def test_thinking_executors_thread_model_param(tmp_path, monkeypatch):
    """The ``model`` node param reaches each thinking agent; blank → None."""
    import ollama_client
    import thinking_agents

    monkeypatch.setattr(ollama_client, "is_available", lambda: True)
    seen: dict[str, Any] = {}

    monkeypatch.setattr(
        thinking_agents,
        "summarize_transcript",
        lambda segments, **kw: seen.__setitem__("summary", kw.get("model")) or "s",
    )
    monkeypatch.setattr(
        thinking_agents,
        "find_citations",
        lambda summary, segments, **kw: (
            seen.__setitem__("citations", kw.get("model")) or []
        ),
    )
    monkeypatch.setattr(
        thinking_agents,
        "find_friction_moments",
        lambda summary, segments, candidates, **kw: (
            seen.__setitem__("friction", kw.get("model")) or []
        ),
    )

    ctx = _ctx(tmp_path)
    tr = {"transcript": {"segments": [{"text": "x"}]}}
    segs = {"segments": [{"start": 0, "end": 1, "text": "hi"}], "source": {}}

    _run("summarize", ctx, tr, {"model": "llama3"})
    _run("citations", ctx, {"summary": "s", "segments": segs}, {"model": "llama3"})
    _run("friction", ctx, {"segments": segs}, {})  # blank → None

    assert seen["summary"] == "llama3"
    assert seen["citations"] == "llama3"
    assert seen["friction"] is None


# ---- Screenspace ----


class _FakeTool:
    """Minimal AnalysisTool stand-in returning one fixed result per scan."""

    calls: list = []

    def scan(self, video_path, region_coords, params, **kw):
        _FakeTool.calls.append(
            (video_path, params.get("start_seconds"), params.get("end_seconds"))
        )
        return [{"timestamp": 5.0, "_confidence": 0.9}]


def test_detect_dispatches_to_selected_detector(tmp_path, monkeypatch):
    # The unified Detect node routes to the tool named by its `detector` param,
    # reusing the same scan body as the (hidden) ss_<tool> nodes.
    _FakeTool.calls = []
    monkeypatch.setitem(screenspace.TOOLS, "color", _FakeTool())
    monkeypatch.setattr(video, "timeline_or_none", lambda paths: None)
    src = {
        "participant": "P01",
        "study": "study",
        "source_filename": "study_P01.mp4",
        "video_paths": ["study_P01.mp4"],
    }
    out = _run("detect", _ctx(tmp_path), {"video": src}, {"detector": "color"})
    assert len(_FakeTool.calls) == 1
    assert out["events"]["events"][0]["time_in"] == 5.0


def test_detect_node_is_palette_facing_and_ss_nodes_hidden():
    # Detect is the visible node; the ten ss_<tool> nodes stay in the catalog
    # (multitool/detect read their specs) but are hidden from the palette.
    assert "detect" in workflows.NODE_TYPES
    assert not workflows.NODE_TYPES["detect"].get("hidden")
    for tool in ("text", "color", "change"):
        node = workflows.NODE_TYPES["ss_" + tool]
        assert node.get("hidden") is True
        # Still executable so old blueprints keep running.
        assert workflows.NODE_TYPES["ss_" + tool].get("execute") is not None


def test_ss_detector_generates_events_with_window(tmp_path, monkeypatch):
    _FakeTool.calls = []
    monkeypatch.setitem(screenspace.TOOLS, "color", _FakeTool())
    monkeypatch.setattr(video, "timeline_or_none", lambda paths: None)
    src = {
        "participant": "P01",
        "study": "study",
        "source_filename": "study_P01.mp4",
        "video_paths": ["study_P01.mp4"],
    }
    tr = {"ranges": [(10.0, 20.0)], "source": src}
    out = _run("ss_color", _ctx(tmp_path), {"video": src, "timeRange": tr}, {})
    # The timeRange window is forwarded as the scan's local [start, end].
    assert _FakeTool.calls == [("study_P01.mp4", 10.0, 20.0)]
    events = out["events"]["events"]
    assert len(events) == 1
    assert events[0]["time_in"] == 5.0
    assert events[0]["participant"] == "P01"
    assert out["events"]["source"] is src


def test_ss_detector_progress_monotonic_across_windows(tmp_path, monkeypatch):
    progress = []

    class _ProgressTool:
        def scan(self, video_path, region_coords, params, **kw):
            cb = kw["on_progress"]
            cb(0.0)
            cb(0.5)
            cb(1.0)
            return [{"timestamp": 1.0, "_confidence": 0.9}]

    monkeypatch.setitem(screenspace.TOOLS, "color", _ProgressTool())
    monkeypatch.setattr(video, "timeline_or_none", lambda paths: None)
    ctx = workflows.NodeContext(
        input_dir=tmp_path, output_dir=tmp_path, on_progress=progress.append
    )
    src = {
        "participant": "P01",
        "study": "study",
        "source_filename": "study_P01.mp4",
        "video_paths": ["study_P01.mp4"],
    }
    tr = {"ranges": [(0.0, 10.0), (10.0, 30.0)], "source": src}
    _run("ss_color", ctx, {"video": src, "timeRange": tr}, {})
    # Job-level progress advances monotonically; only the final window reaches
    # 1.0 (the first window's own 0->1 maps into [0, 1/3], not back to 0).
    assert progress == sorted(progress)
    assert progress[-1] == 1.0
    assert progress.count(1.0) == 1


def test_ss_detector_multipart_offsets_event_times(tmp_path, monkeypatch):
    _FakeTool.calls = []
    monkeypatch.setitem(screenspace.TOOLS, "color", _FakeTool())
    monkeypatch.setattr(
        video, "timeline_or_none", lambda paths: [("a.mp4", 80, 0), ("b.mp4", 120, 80)]
    )
    src = {
        "participant": "P01",
        "study": "study",
        "source_filename": "a.mp4",
        "video_paths": ["a.mp4", "b.mp4"],
    }
    tr = {"ranges": [(60.0, 90.0)], "source": src}
    out = _run("ss_color", _ctx(tmp_path), {"video": src, "timeRange": tr}, {})
    # 60..90 straddles the 80s boundary; result times shift back to the global
    # timeline (5.0 from part a, 5.0+80 from part b).
    times = sorted(e["time_in"] for e in out["events"]["events"])
    assert times == [5.0, 85.0]


# ---- Artifact ----


def test_make_clips_cuts_from_timerange(tmp_path, monkeypatch):
    monkeypatch.setattr(Path, "is_file", lambda self: True)
    monkeypatch.setattr(utils, "create_progress_bar", lambda: None)
    monkeypatch.setattr(config, "CLIP_PARALLEL_WORKERS", 1)
    monkeypatch.setattr(
        files,
        "get_unique_filename",
        lambda _t, file_format=None: str(tmp_path / f"out{file_format or '.mp4'}"),
    )
    monkeypatch.setattr(video, "run_ffmpeg", Mock(return_value=True))
    src = {
        "participant": "P01",
        "study": "study",
        "source_filename": "study_P01.mp4",
        "video_paths": ["study_P01.mp4"],
    }
    tr = {"ranges": [(10.0, 20.0)], "source": src}
    out = _run("make_clips", _ctx(tmp_path), {"timeRange": tr}, {"description": "hit"})
    assert out["artifacts"]["count"] == 1
    assert len(out["artifacts"]["artifacts"]) == 1
    assert out["artifacts"]["study"] == "study"


def test_build_reel_honors_name_param(tmp_path, monkeypatch):
    import pipeline

    seen = {}

    def fake_unique(template, file_format=None):
        seen["template"] = template
        return str(tmp_path / template)

    monkeypatch.setattr(files, "get_unique_filename", fake_unique)
    monkeypatch.setattr(
        pipeline,
        "process_reel",
        lambda records, output_file=None, cancel_flag=None, **kw: (1, [{"id": "r"}]),
    )
    out = _run(
        "build_reel",
        _ctx(tmp_path),
        {"clips": {"records": [{"x": 1}], "study": "study"}},
        {"name": "My Reel"},
    )
    # The (sanitized) reel name drives the reserved output filename.
    assert seen["template"].lower().startswith("my")
    assert out["artifacts"]["count"] == 1


def test_build_reel_chronological_sorts_records(tmp_path, monkeypatch):
    import pipeline

    seen: dict[str, Any] = {}
    monkeypatch.setattr(
        files, "get_unique_filename", lambda template, **kw: str(tmp_path / template)
    )
    monkeypatch.setattr(
        pipeline,
        "process_reel",
        lambda records, **kw: (
            seen.__setitem__("records", records) or (1, [{"id": "r"}])
        ),
    )
    out_of_order = [
        {"id": "b", "times": [("0:30", "0:35")]},
        {"id": "a", "times": [("0:05", "0:10")]},
    ]
    _run(
        "build_reel",
        _ctx(tmp_path),
        {"clips": {"records": out_of_order, "study": "study"}},
        {"name": "reel", "chronological": True},
    )
    assert [r["id"] for r in seen["records"]] == ["a", "b"]


def test_transcribe_threads_model_param(tmp_path, monkeypatch):
    import transcripts

    seen: dict[str, Any] = {}
    monkeypatch.setattr(
        transcripts,
        "transcribe_video",
        lambda path, **kw: (
            seen.__setitem__("model", kw.get("model_name")) or {"segments": []}
        ),
    )
    src = {"participant": "P01", "video_paths": [str(tmp_path / "study_P01.mp4")]}
    _run("transcribe", _ctx(tmp_path), {"video": src}, {"model": "small"})
    assert seen["model"] == "small"


def test_timeline_viewer_generates_path(tmp_path, monkeypatch):
    monkeypatch.setattr(
        viewer,
        "generate_timeline_viewer",
        lambda data, **kw: tmp_path / "workflow_viewer.html",
    )
    out = _run(
        "timeline_viewer",
        _ctx(tmp_path),
        {"artifacts": {"artifacts": [], "study": "study"}},
        {},
    )
    assert out["viewer"]["path"].endswith("workflow_viewer.html")


def test_timeline_viewer_routes_reels_to_reels_slot(tmp_path, monkeypatch):
    # A reel record (carries `components`, no start/end) must reach the viewer's
    # `reels` slot, not the timeline `artifacts` slot — else it's filtered for
    # lack of start/end and the viewer is empty (the Build Reel → Viewer bug).
    captured = {}

    def fake_finalize(artifacts, **kw):
        captured["artifacts"] = artifacts
        captured["reels"] = kw.get("reels")
        return {"artifacts": artifacts}

    monkeypatch.setattr(viewer, "finalize_timeline_data", fake_finalize)
    monkeypatch.setattr(
        viewer, "generate_timeline_viewer", lambda data, **kw: tmp_path / "v.html"
    )
    reel = {"id": "r1", "file": "reel.mp4", "components": [{"start": "0:00"}]}
    clip = {"id": "c1", "file": "c.mp4", "type": "clip", "start": 1.0, "end": 2.0}
    out = _run(
        "timeline_viewer",
        _ctx(tmp_path),
        {"artifacts": {"artifacts": [reel, clip], "study": "study"}},
        {},
    )
    assert captured["reels"] == [reel]
    assert captured["artifacts"] == [clip]
    assert out["viewer"]["path"]


# ---- Sheet selection (real pure path) ----


def test_sheet_selection_generates_records(tmp_path):
    from spreadsheet import SheetContext

    sheet_data = [
        ["study", "", "", "", ""],
        ["ID", "P01", "P02", "Observation", "Category"],
        ["1", "00:10-00:20", "", "Obs one", "CatA"],
    ]
    sctx = SheetContext(
        sheet_data=sheet_data,
        id_cell=SimpleNamespace(row=2, col=1),
        observation_cell=SimpleNamespace(row=2, col=4),
        category_cell=SimpleNamespace(row=2, col=5),
        num_participants=2,
        study_name="study",
    )
    out = _run(
        "sheet_selection", _ctx(tmp_path, sheet_context=sctx), {}, {"selector": "P01"}
    )
    assert out["clips"]["study"] == "study"
    records = out["clips"]["records"]
    assert records and records[0]["participant"] == "P01"


def test_sheet_selection_empty_without_context(tmp_path):
    out = _run("sheet_selection", _ctx(tmp_path), {}, {"selector": "P01"})
    assert out["clips"]["records"] == []


# ---- Adapters (pure value -> value) ----


def test_adapter_transcript_to_segments():
    out = workflows.ADAPTERS[("transcript", "segments")](
        {"segments": [1, 2], "source": {"participant": "P01"}}
    )
    assert out == {"segments": [1, 2], "source": {"participant": "P01"}}


def test_adapter_segments_to_timerange_projects_all():
    val = {
        "segments": [{"start": 1.0, "end": 2.0}, {"start": 5.0, "end": 6.0}],
        "source": {"x": 1},
    }
    out = workflows.ADAPTERS[("segments", "timeRange")](val)
    assert out["ranges"] == [(1.0, 2.0), (5.0, 6.0)]
    assert out["source"] == {"x": 1}


def test_adapter_video_to_scalar_duration(monkeypatch):
    monkeypatch.setattr(video, "probe_video_properties", lambda p: {"duration": 123.0})
    assert workflows.ADAPTERS[("video", "scalar")]({"video_paths": ["a.mp4"]}) == 123.0


def test_adapter_timerange_to_cliprecords_builds_records():
    val = {
        "ranges": [(10.0, 20.0)],
        "source": {
            "participant": "P01",
            "study": "study",
            "source_filename": "study_P01.mp4",
            "video_paths": ["study_P01.mp4"],
        },
    }
    out = workflows.ADAPTERS[("timeRange", "clipRecords")](val)
    recs = out["records"]
    assert len(recs) == 1
    assert recs[0]["participant"] == "P01"
    assert recs[0]["cell"].col == files._WORKFLOW_CELL_COL
    assert recs[0]["times"] == [("0:00:10", "0:00:20")]


def test_adapter_events_to_cliprecords_clusters():
    events = [
        {
            "time_in": 10.0,
            "time_out": 11.0,
            "participant": "P01",
            "source_video": "study_P01.mp4",
        },
        {
            "time_in": 12.0,
            "time_out": 13.0,
            "participant": "P01",
            "source_video": "study_P01.mp4",
        },
    ]
    val = {
        "events": events,
        "source": {
            "participant": "P01",
            "study": "study",
            "source_filename": "study_P01.mp4",
            "video_paths": ["study_P01.mp4"],
        },
    }
    out = workflows.ADAPTERS[("events", "clipRecords")](val)
    # gap 5.0 merges the two adjacent events into one clip record.
    assert len(out["records"]) == 1


def test_adapter_events_to_timerange_derives_source():
    events = [
        {
            "time_in": 1.0,
            "time_out": 2.0,
            "participant": "P01",
            "source_video": "study_P01.mp4",
        }
    ]
    out = workflows.ADAPTERS[("events", "timeRange")]({"events": events})
    assert out["ranges"] == [(1.0, 2.0)]
    assert out["source"]["participant"] == "P01"
    assert out["source"]["study"] == "study"


# ---- P2 catalog tranche ----


def test_make_clips_passes_output_format_and_titlecards(tmp_path, monkeypatch):
    import pipeline

    captured = {}

    def fake_process_clips(records, output_format="clip", include_severity=False, **kw):
        captured["output_format"] = output_format
        captured["titlecards_enabled"] = kw.get("titlecards_enabled")
        captured["titlecard_duration"] = kw.get("titlecard_duration_seconds")
        return (len(records), [{"id": "a", "type": output_format}])

    monkeypatch.setattr(pipeline, "process_clips", fake_process_clips)
    src = {
        "participant": "P01",
        "study": "study",
        "source_filename": "study_P01.mp4",
        "video_paths": ["study_P01.mp4"],
    }
    tr = {"ranges": [(10.0, 20.0)], "source": src}
    out = _run(
        "make_clips",
        _ctx(tmp_path),
        {"timeRange": tr},
        {"output_format": "gif", "titlecards": True, "titlecard_duration": 3},
    )
    assert captured["output_format"] == "gif"
    assert captured["titlecards_enabled"] is True
    assert captured["titlecard_duration"] == 3
    assert out["artifacts"]["count"] == 1


def test_highlights_truncates_to_budget(tmp_path, monkeypatch):
    import spreadsheet

    seen = {}

    def fake_score(records, existing, budget):
        seen["budget"] = budget
        seen["existing"] = existing
        return records[:1]

    monkeypatch.setattr(spreadsheet, "score_and_truncate_clips", fake_score)
    monkeypatch.setattr(files, "discover_clips", lambda: ["old.mp4"])
    clips = {"records": [{"a": 1}, {"a": 2}, {"a": 3}], "study": "study"}
    out = _run("highlights", _ctx(tmp_path), {"clips": clips}, {"budget": 90})
    assert seen["budget"] == 90
    assert seen["existing"] == {"old.mp4"}
    assert len(out["clips"]["records"]) == 1
    assert out["clips"]["study"] == "study"


def test_measure_counts_events_and_durations(tmp_path):
    ctx = _ctx(tmp_path)
    events = {
        "events": [
            {"time_in": 1.0, "time_out": 3.0, "confidence": 0.4},
            {"time_in": 5.0, "time_out": 6.0, "confidence": 0.9},
        ]
    }
    assert _run("measure", ctx, {"events": events}, {"metric": "count"})["value"] == 2.0
    assert (
        _run("measure", ctx, {"events": events}, {"metric": "max_confidence"})["value"]
        == 0.9
    )
    assert (
        _run("measure", ctx, {"events": events}, {"metric": "total_duration"})["value"]
        == 3.0
    )


def test_measure_drives_gate(tmp_path):
    ctx = _ctx(tmp_path)
    events = {"events": [{"time_in": 0, "time_out": 1} for _ in range(5)]}
    measured = _run("measure", ctx, {"events": events}, {"metric": "count"})
    assert _run(
        "gate", ctx, {"value": measured["value"]}, {"op": ">=", "threshold": 3}
    )["pass"]


def test_ss_detector_reshapes_real_params(tmp_path, monkeypatch):
    captured = {}

    class _CaptureTool:
        def scan(self, video_path, region_coords, params, **kw):
            captured.update(params)
            return [{"timestamp": 1.0, "_confidence": 0.5}]

    monkeypatch.setitem(screenspace.TOOLS, "color", _CaptureTool())
    monkeypatch.setattr(video, "timeline_or_none", lambda paths: None)
    src = {
        "participant": "P01",
        "study": "study",
        "source_filename": "study_P01.mp4",
        "video_paths": ["study_P01.mp4"],
    }
    _run(
        "ss_color",
        _ctx(tmp_path),
        {"video": src},
        {"color_h": 120, "color_s": 200, "color_v": 150, "color_mode": "presence"},
    )
    # The dead-params gap is fixed: the node's flat params are reshaped into the
    # nested scan params the tool expects (not an empty {}).
    assert captured["target_color"] == {"h": 120.0, "s": 200.0, "v": 150.0}
    assert captured["color_mode"] == "presence"


def test_timelapse_emits_attachment(tmp_path, monkeypatch):
    import screenspace_scans

    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda p: {"width": 1920, "height": 1080, "duration": 60.0},
    )
    monkeypatch.setattr(
        files,
        "get_unique_filename",
        lambda name, file_format=None: str(tmp_path / "timelapse.mp4"),
    )
    monkeypatch.setattr(
        screenspace_scans, "generate_timelapse", lambda *a, **kw: str(a[3])
    )
    src = {
        "participant": "P01",
        "study": "study",
        "source_filename": "study_P01.mp4",
        "video_paths": ["study_P01.mp4"],
    }
    out = _run("timelapse", _ctx(tmp_path), {"video": src}, {"output_format": "mp4"})
    arts = out["artifacts"]["artifacts"]
    assert len(arts) == 1
    assert arts[0]["type"] == "timelapse"
    # Attachment artifacts sit off the timeline (start/end 0).
    assert arts[0]["start"] == 0 and arts[0]["end"] == 0


def test_heatmap_consumes_raw_results(tmp_path, monkeypatch):
    import screenspace_heatmap

    monkeypatch.setattr(
        video, "probe_video_properties", lambda p: {"width": 640, "height": 480}
    )
    monkeypatch.setattr(
        files,
        "get_unique_filename",
        lambda name, file_format=None: str(tmp_path / "heatmap.png"),
    )
    called = {}

    def fake_change(results, w, h, out_path):
        called["results"] = results
        called["dims"] = (w, h)
        return out_path

    monkeypatch.setattr(screenspace_heatmap, "generate_change_heatmap", fake_change)
    events_in = {
        "events": [],
        "source": {
            "participant": "P01",
            "study": "study",
            "source_filename": "study_P01.mp4",
            "video_paths": ["study_P01.mp4"],
        },
        "raw_results": [{"change_grid": []}],
    }
    out = _run("heatmap", _ctx(tmp_path), {"events": events_in}, {"style": "change"})
    assert called["results"] == [{"change_grid": []}]
    assert called["dims"] == (640, 480)
    assert out["artifacts"]["artifacts"][0]["type"] == "heatmap"


# ---- Full-frame fallback + time-range authoring ----


def test_ss_detector_defaults_to_full_frame_when_region_unwired(tmp_path, monkeypatch):
    seen = {}

    class _RegionTool:
        def scan(self, video_path, region_coords, params, **kw):
            seen["region"] = dict(region_coords)
            return []

    monkeypatch.setitem(screenspace.TOOLS, "color", _RegionTool())
    monkeypatch.setattr(video, "timeline_or_none", lambda paths: None)
    monkeypatch.setattr(
        video, "probe_video_properties", lambda p: {"width": 1920, "height": 1080}
    )
    src = {
        "participant": "P01",
        "study": "study",
        "source_filename": "study_P01.mp4",
        "video_paths": ["study_P01.mp4"],
    }
    # No region input wired → scan the whole frame (not a zero-size no-op).
    _run("ss_color", _ctx(tmp_path), {"video": src}, {})
    assert seen["region"] == {"x": 0, "y": 0, "w": 1920, "h": 1080}


def test_resolve_region_coords_full_frame_fallback(monkeypatch):
    monkeypatch.setattr(
        video, "probe_video_properties", lambda p: {"width": 1280, "height": 720}
    )
    name, coords = workflows._resolve_region_coords({}, "study_P01.mp4")
    assert name == ""
    assert coords == {"x": 0, "y": 0, "w": 1280, "h": 720}


def test_time_range_parses_manual_ranges(tmp_path):
    out = _run("time_range", _ctx(tmp_path), {}, {"ranges": "1:23-1:45, 2:00-2:30"})
    assert out["timeRange"]["ranges"] == [(83.0, 105.0), (120.0, 150.0)]
    assert out["timeRange"]["source"] == {}


def test_time_range_empty_is_safe(tmp_path):
    out = _run("time_range", _ctx(tmp_path), {}, {"ranges": ""})
    assert out["timeRange"]["ranges"] == []


def test_adapter_cliprecords_to_timerange():
    # Sheet cells → SS scan windows: records carry resolved `times` here.
    value = {
        "records": [
            {"times": [("0:00:10", "0:00:20")], "participant": "P01", "study": "s"},
            {"times": [("0:01:00", "0:01:05")], "participant": "P01", "study": "s"},
        ],
        "study": "s",
    }
    out = workflows.ADAPTERS[("clipRecords", "timeRange")](value)
    assert out["ranges"] == [(10.0, 20.0), (60.0, 65.0)]
    assert out["source"]["participant"] == "P01"
