"""Direct-invocation tests for the Workflows node executors and typed-port
adapters (M3).

The M4 ``WorkflowRunner`` does not exist yet, so each executor is exercised by
calling ``NODE_TYPES[id]["execute"](ctx, inputs, params)`` directly. Whisper /
ffmpeg paths run under ``config.DEBUGGING`` or with ffmpeg mocked; the AI server and
Screenspace scans are monkeypatched so no model/subprocess/network is touched.
"""

import json
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


def test_region_notes_missing_name(tmp_path, monkeypatch):
    # A named-but-missing region degrades to a full-frame scan downstream; the
    # node must say so instead of completing silently.
    monkeypatch.setattr(
        screenspace, "load_screenspace_manifest", lambda: {"regions": {}}
    )
    out = _run("region", _ctx(tmp_path), {}, {"name": "timer"})
    assert out["region"] == {"name": "timer", "coords": None}
    assert "not found" in out["__note__"]

    monkeypatch.setattr(
        screenspace,
        "load_screenspace_manifest",
        lambda: {"regions": {"timer": {"x": 1, "y": 2, "w": 3, "h": 4}}},
    )
    out = _run("region", _ctx(tmp_path), {}, {"name": "timer"})
    assert out["region"]["coords"] == {"x": 1, "y": 2, "w": 3, "h": 4}
    assert "__note__" not in out


def test_transcript_marks_resolves_categories_and_pads(tmp_path, monkeypatch):
    import transcripts

    manifest = {
        "source_transcripts": {
            "P01": {
                "segments": [
                    {"start": 10.0, "end": 12.0, "text": "a"},
                    {"start": 30.0, "end": 33.0, "text": "b"},
                ]
            }
        },
        "marks": [
            {"segment_id": "P01:0", "category": "pain_point"},
            {"segment_id": "P01:1", "category": "quote"},
            {"segment_id": "P02:0", "category": "pain_point"},  # other participant
            {"segment_id": "P01:9", "category": "pain_point"},  # stale index
        ],
    }
    monkeypatch.setattr(transcripts, "load_transcripts_manifest", lambda: manifest)
    src = {"participant": "P01", "video_paths": ["v.mp4"]}

    out = _run("transcript_marks", _ctx(tmp_path), {"video": src}, {"pad": 2})
    assert out["timeRange"]["ranges"] == [(8.0, 14.0), (28.0, 35.0)]
    assert out["timestamps"]["times"] == [10.0, 30.0]

    out = _run(
        "transcript_marks",
        _ctx(tmp_path),
        {"video": src},
        {"pad": 0, "category": "quote"},
    )
    assert out["timeRange"]["ranges"] == [(30.0, 33.0)]

    out = _run(
        "transcript_marks",
        _ctx(tmp_path),
        {"video": src},
        {"category": "bookmark"},
    )
    assert out["timeRange"]["ranges"] == []
    assert "No marks" in out["__note__"]


def test_report_builds_from_summary_and_sources(tmp_path, monkeypatch):
    import llm_client
    import thinking_agents

    monkeypatch.setattr(llm_client, "is_available", lambda: True)
    monkeypatch.setattr(
        thinking_agents,
        "report_source_lines",
        lambda pid: ([f"- obs for {pid}"], ["[0:10] (Quote) hello"]),
    )
    seen = {}

    def fake_build(summary, obs, marks, *, participant, model, cancel_event):
        seen.update(
            {"summary": summary, "obs": obs, "marks": marks, "pid": participant}
        )
        return "the report"

    monkeypatch.setattr(thinking_agents, "build_report", fake_build)
    src = {"participant": "P01", "video_paths": ["v.mp4"]}
    out = _run("report", _ctx(tmp_path), {"summary": "the summary", "video": src}, {})
    assert out["report"] == "the report"
    assert "__note__" not in out
    assert seen == {
        "summary": "the summary",
        "obs": "- obs for P01",
        "marks": "[0:10] (Quote) hello",
        "pid": "P01",
    }
    # No summary → skipped with a note, no model call.
    out = _run("report", _ctx(tmp_path), {}, {})
    assert out["report"] == ""
    assert "No summary" in out["__note__"]


def test_post_process_remux_skips_already_seekable(tmp_path, monkeypatch):
    called = []
    monkeypatch.setattr(
        video, "probe_container_seekability", lambda p: {"browser_seekable": True}
    )
    monkeypatch.setattr(
        video, "remux_to_faststart", lambda p, **kw: called.append(p) or (True, "ok")
    )
    src = {"participant": "P01", "study": "s", "video_paths": ["a.mp4", "b.mp4"]}
    out = _run(
        "post_process",
        _ctx(tmp_path),
        {"video": src},
        {"operation": "remux_faststart"},
    )
    assert called == []
    assert out["video"] is src
    assert "Already browser-seekable" in out["__note__"]


def test_post_process_embed_subtitles_muxes_a_copy(tmp_path, monkeypatch):
    import transcripts

    monkeypatch.setattr(transcripts, "write_transcript", lambda *a, **kw: True)
    muxed = {}
    monkeypatch.setattr(
        video,
        "mux_subtitles",
        lambda vid, srt, out, **kw: muxed.update({"vid": vid, "out": out}) or True,
    )
    monkeypatch.setattr(
        files,
        "get_unique_filename",
        lambda name, file_format=None: str(tmp_path / name),
    )
    src = {
        "participant": "P01",
        "study": "s",
        "source_filename": "s_P01.mp4",
        "video_paths": [str(tmp_path / "s_P01.mp4")],
    }
    transcript = {"segments": [{"start": 0, "end": 1, "text": "hi"}], "source": src}
    out = _run(
        "post_process",
        _ctx(tmp_path),
        {"video": src, "transcript": transcript},
        {"operation": "embed_subtitles"},
    )
    assert muxed["out"].endswith("s_P01-subtitled.mp4")
    assert out["video"]["video_paths"] == [muxed["out"]]
    assert out["artifacts"]["count"] == 1
    # No transcript wired → note, no mux.
    out = _run(
        "post_process", _ctx(tmp_path), {"video": src}, {"operation": "embed_subtitles"}
    )
    assert out["artifacts"]["count"] == 0
    assert "transcript" in out["__note__"]


def test_gallery_viewer_filters_stills(tmp_path, monkeypatch):
    seen = {}
    monkeypatch.setattr(
        viewer,
        "finalize_gallery_data",
        lambda arts, **kw: seen.update({"arts": arts, **kw}) or {"artifacts": arts},
    )
    monkeypatch.setattr(
        viewer,
        "generate_gallery_viewer",
        lambda data, output_basename="": tmp_path / output_basename,
    )
    arts = {
        "artifacts": [
            {"type": "screen", "file": "a.png"},
            {"type": "clip", "file": "c.mp4"},
            {"type": "gif", "file": "b.gif"},
        ],
        "study": "s",
    }
    out = _run("gallery_viewer", _ctx(tmp_path), {"artifacts": arts}, {})
    assert [a["file"] for a in seen["arts"]] == ["a.png", "b.gif"]
    assert seen["output_format"] == "gif"
    assert out["viewer"]["path"].endswith("workflow_gallery.html")
    # Clips only → note instead of an empty gallery.
    out = _run(
        "gallery_viewer",
        _ctx(tmp_path),
        {"artifacts": {"artifacts": [{"type": "clip", "file": "c.mp4"}]}},
        {},
    )
    assert out["viewer"]["path"] is None
    assert "Timeline Viewer" in out["__note__"]


def test_video_source_coerces_single_element_list(tmp_path, monkeypatch):
    # The canvas multi-select stores a list; a one-element list is a plain
    # single-participant run and must not stringify to "['P01']".
    monkeypatch.setattr(
        workflows.utils,
        "discover_participant_videos",
        lambda *a, **k: [
            {"id": "P01", "video_paths": ["/v/study_P01.mp4"], "has_video": True}
        ],
    )
    out = _run("video_source", _ctx(tmp_path), {}, {"participant": ["P01"]})
    assert out["participant"] == "P01"
    assert out["video"]["video_paths"] == ["/v/study_P01.mp4"]


def test_video_source_rejects_multi_selection_on_direct_run(tmp_path):
    # A multi-selection (or the __all__ sentinel) only makes sense as a batch;
    # a direct run must fail loudly instead of resolving zero videos.
    import pytest

    with pytest.raises(RuntimeError, match="batch"):
        _run("video_source", _ctx(tmp_path), {}, {"participant": ["P01", "P02"]})
    with pytest.raises(RuntimeError, match="batch"):
        _run("video_source", _ctx(tmp_path), {}, {"participant": "__all__"})


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


def test_transcribe_notes_when_nothing_wired(tmp_path):
    out = _run("transcribe", _ctx(tmp_path), {}, {})
    assert out["transcript"]["segments"] == []
    assert "wire" in out["__note__"].lower()


def test_transcribe_raises_when_decode_fails(tmp_path, monkeypatch):
    import transcripts

    monkeypatch.setattr(transcripts, "transcribe_video", lambda *a, **k: None)
    src = {"participant": "P01", "video_paths": [str(tmp_path / "study_P01.mp4")]}
    try:
        _run("transcribe", _ctx(tmp_path), {"video": src}, {})
    except RuntimeError as exc:
        assert "transcribe" in str(exc).lower()
    else:
        raise AssertionError("decode failure must fail the node, not return empty")


# ---- Thinking (local LLM) ----


def test_thinking_executors_empty_when_llm_unavailable(tmp_path, monkeypatch):
    import llm_client

    monkeypatch.setattr(llm_client, "is_available", lambda: False)
    ctx = _ctx(tmp_path)
    segs = {"segments": [{"start": 0, "end": 1, "text": "hi"}], "source": {}}
    tr = {"segments": [{"start": 0, "end": 1, "text": "hi"}]}
    summarize = _run("summarize", ctx, {"transcript": tr}, {})
    citations = _run("citations", ctx, {"summary": "s", "segments": segs}, {})
    friction = _run("friction", ctx, {"segments": segs}, {})
    assert summarize["summary"] == ""
    assert citations["citations"] == []
    assert friction["friction"] == []
    # Degraded-but-completed: a __note__ explains the empty output (AI server down).
    assert "AI server" in summarize["__note__"]
    assert "AI server" in citations["__note__"]
    assert "AI server" in friction["__note__"]


def test_make_clips_notes_when_nothing_wired(tmp_path):
    # No clips / time range / video wired → completes empty with a note saying why.
    result = _run("make_clips", _ctx(tmp_path), {}, {})
    assert result["artifacts"]["count"] == 0
    assert "wire" in result["__note__"].lower()


def test_multitool_step_flag_matches_tool_capability():
    # The catalog's multitoolStep flag must match the actual tool capability —
    # overrides check_frame AND isn't reference-based — so _MULTITOOL_STEP_TOOLS
    # can't drift from screenspace_tools.
    import screenspace
    import screenspace_tools as st

    def _cls(t):
        return t if isinstance(t, type) else type(t)

    derived = {
        name
        for name, tool in screenspace.TOOLS.items()
        if _cls(tool).check_frame is not st.AnalysisTool.check_frame
        and name not in workflows._SS_REFERENCE_DETECTORS
    }
    tagged = {
        n["id"][3:]
        for n in workflows.NODE_TYPES.values()
        if n.get("multitoolStep") and n["id"].startswith("ss_")
    }
    assert tagged == derived


def test_summarize_wires_thinking_agent(tmp_path, monkeypatch):
    import llm_client
    import thinking_agents

    monkeypatch.setattr(llm_client, "is_available", lambda: True)
    monkeypatch.setattr(
        thinking_agents, "summarize_transcript", lambda segments, **kw: "the summary"
    )
    out = _run(
        "summarize", _ctx(tmp_path), {"transcript": {"segments": [{"text": "x"}]}}, {}
    )
    assert out["summary"] == "the summary"


def test_thinking_executors_thread_model_param(tmp_path, monkeypatch):
    """The ``model`` node param reaches each thinking agent; blank → None."""
    import llm_client
    import thinking_agents

    monkeypatch.setattr(llm_client, "is_available", lambda: True)
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


def test_interval_captures_samples_each_range(tmp_path, monkeypatch):
    import pipeline

    seen: dict[str, Any] = {}
    monkeypatch.setattr(
        pipeline,
        "process_clips",
        lambda records, **kw: (
            seen.update({"records": records, "fmt": kw.get("output_format")}),
            (len(records), [{"id": i} for i in range(len(records))]),
        )[1],
    )
    src = {
        "participant": "P01",
        "study": "study",
        "source_filename": "study_P01.mp4",
        "video_paths": ["study_P01.mp4"],
    }
    tr = {"ranges": [(0.0, 30.0)], "source": src}
    out = _run(
        "interval_captures",
        _ctx(tmp_path),
        {"video": src, "timeRange": tr},
        {"interval": 10, "output_format": "screen"},
    )
    # 0, 10, 20 → three screenshots (unlike Make Clips, which makes one per range).
    assert out["artifacts"]["count"] == 3
    assert len(seen["records"]) == 3
    assert seen["fmt"] == "screen"


def test_interval_captures_whole_video_uses_duration(tmp_path, monkeypatch):
    import pipeline
    import video as video_mod

    monkeypatch.setattr(video_mod, "get_file_duration", lambda p: 25)
    seen: dict[str, Any] = {}
    monkeypatch.setattr(
        pipeline,
        "process_clips",
        lambda records, **kw: seen.__setitem__("n", len(records)) or (len(records), []),
    )
    src = {
        "participant": "P01",
        "study": "s",
        "source_filename": "s_P01.mp4",
        "video_paths": ["s_P01.mp4"],
    }
    # No timeRange wired → sample the whole video: duration 25, interval 10 → 0,10,20.
    _run("interval_captures", _ctx(tmp_path), {"video": src}, {"interval": 10})
    assert seen["n"] == 3


def test_interval_captures_keeps_fractional_interval(tmp_path, monkeypatch):
    # The interval is seconds as a float; 0.5 must sample every half second,
    # not silently truncate to 1 s.
    import pipeline

    seen: dict[str, Any] = {}
    monkeypatch.setattr(
        pipeline,
        "process_clips",
        lambda records, **kw: seen.__setitem__("n", len(records)) or (len(records), []),
    )
    src = {
        "participant": "P01",
        "study": "s",
        "source_filename": "s_P01.mp4",
        "video_paths": ["s_P01.mp4"],
    }
    tr = {"ranges": [(0.0, 2.0)], "source": src}
    _run(
        "interval_captures",
        _ctx(tmp_path),
        {"video": src, "timeRange": tr},
        {"interval": 0.5, "output_format": "screen"},
    )
    assert seen["n"] == 4  # 0, 0.5, 1.0, 1.5


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


def test_make_clips_forwards_padding_kwargs(tmp_path, monkeypatch):
    import pipeline

    seen: dict[str, Any] = {}
    monkeypatch.setattr(
        pipeline,
        "process_clips",
        lambda records, **kw: seen.update(kw) or (len(records), []),
    )
    src = {
        "participant": "P01",
        "study": "study",
        "source_filename": "study_P01.mp4",
        "video_paths": ["study_P01.mp4"],
    }
    tr = {"ranges": [(10.0, 20.0)], "source": src}
    _run(
        "make_clips",
        _ctx(tmp_path),
        {"timeRange": tr},
        {"pad_start": 3, "pad_end": -2, "max_duration": 15},
    )
    assert seen["pad_pre"] == 3.0
    assert seen["pad_post"] == -2.0
    assert seen["max_duration"] == 15.0


def test_build_reel_forwards_padding_kwargs(tmp_path, monkeypatch):
    import pipeline

    seen: dict[str, Any] = {}
    monkeypatch.setattr(
        files, "get_unique_filename", lambda template, **kw: str(tmp_path / template)
    )
    monkeypatch.setattr(
        pipeline,
        "process_reel",
        lambda records, output_file=None, cancel_flag=None, **kw: (
            seen.update(kw) or (1, [{"id": "r"}])
        ),
    )
    _run(
        "build_reel",
        _ctx(tmp_path),
        {"clips": {"records": [{"x": 1}], "study": "study"}},
        {"pad_start": 1, "pad_end": 2, "max_duration": 0},
    )
    assert seen["pad_pre"] == 1.0
    assert seen["pad_post"] == 2.0
    assert seen["max_duration"] == 0.0


def test_artifact_padding_params_defaults_are_noop():
    assert workflows._artifact_padding_params({}) == (0.0, 0.0, 0.0)


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


def test_gate_collection_covers_every_metric(tmp_path):
    # gate_collection subsumed the standalone measure node; its three metrics
    # must keep reducing correctly (thresholds chosen to pin exact values).
    ctx = _ctx(tmp_path)
    events = {
        "events": {
            "events": [
                {"time_in": 1.0, "time_out": 3.0, "confidence": 0.4},
                {"time_in": 5.0, "time_out": 6.0, "confidence": 0.9},
            ]
        }
    }
    for metric, value in (
        ("count", 2.0),
        ("max_confidence", 0.9),
        ("total_duration", 3.0),
    ):
        assert _run(
            "gate_collection",
            ctx,
            events,
            {"metric": metric, "op": "==", "threshold": value},
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


def test_heatmap_auto_infers_style_from_payload_keys(tmp_path, monkeypatch):
    # "auto" (the default) picks the style from the detector data actually
    # present, so the user no longer mirrors the upstream detector by hand.
    import screenspace_heatmap

    monkeypatch.setattr(
        video, "probe_video_properties", lambda p: {"width": 640, "height": 480}
    )
    monkeypatch.setattr(
        files,
        "get_unique_filename",
        lambda name, file_format=None: str(tmp_path / "heatmap.png"),
    )
    styles = {
        "attention": [{"saliency_grid": []}],
        "flow": [{"flow_grid": []}],
        "template": [{"matches": [{"x": 0, "y": 0, "w": 1, "h": 1}]}],
    }
    for style, raw in styles.items():
        called = {}
        monkeypatch.setattr(
            screenspace_heatmap,
            f"generate_{style}_heatmap",
            lambda results, w, h, out_path, _c=called: _c.update(hit=True) or out_path,
        )
        src = {"participant": "P01", "study": "s", "video_paths": ["v.mp4"]}
        events_in = {"events": [], "source": src, "raw_results": raw}
        out = _run("heatmap", _ctx(tmp_path), {"events": events_in}, {"style": "auto"})
        assert called.get("hit"), f"auto did not route to {style}"
        assert out["artifacts"]["count"] == 1
    # Events with no heatmap payload (e.g. a text detector) → note, not a wrong map.
    src = {"participant": "P01", "study": "s", "video_paths": ["v.mp4"]}
    events_in = {"events": [], "source": src, "raw_results": [{"text": "hi"}]}
    out = _run("heatmap", _ctx(tmp_path), {"events": events_in}, {})
    assert out["artifacts"]["count"] == 0
    assert "no heatmap data" in out["__note__"]


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


# ---- transcript_export / data_export / animated heatmap ----

_SRC = {
    "participant": "P01",
    "study": "study",
    "source_filename": "study_P01.mp4",
    "video_paths": ["study_P01.mp4"],
}


def _redirect_unique_filenames(monkeypatch, tmp_path):
    monkeypatch.setattr(
        files,
        "get_unique_filename",
        lambda name, file_format=None: str(tmp_path / Path(name).name),
    )


def test_transcript_export_writes_vtt(tmp_path, monkeypatch):
    _redirect_unique_filenames(monkeypatch, tmp_path)
    transcript = {
        "segments": [{"start": 0.0, "end": 1.5, "text": "hello"}],
        "language": "en",
        "model": "base",
        "source_file": "study_P01.mp4",
        "source": _SRC,
    }
    out = _run(
        "transcript_export",
        _ctx(tmp_path),
        {"transcript": transcript},
        {"format": "vtt"},
    )
    arts = out["artifacts"]["artifacts"]
    assert len(arts) == 1
    assert arts[0]["type"] == "export"
    assert arts[0]["start"] == 0 and arts[0]["end"] == 0
    written = tmp_path / "transcript_P01.vtt"
    text = written.read_text(encoding="utf-8")
    assert text.startswith("WEBVTT")
    assert "-->" in text and "hello" in text


def test_transcript_export_falls_back_to_segments_wire(tmp_path, monkeypatch):
    _redirect_unique_filenames(monkeypatch, tmp_path)
    seg_in = {
        "segments": [{"start": 0.0, "end": 2.0, "text": "from segments"}],
        "source": _SRC,
    }
    out = _run(
        "transcript_export", _ctx(tmp_path), {"segments": seg_in}, {"format": "md"}
    )
    assert out["artifacts"]["count"] == 1
    assert "from segments" in (tmp_path / "transcript_P01.md").read_text(
        encoding="utf-8"
    )


def test_transcript_export_notes_when_nothing_wired(tmp_path):
    out = _run("transcript_export", _ctx(tmp_path), {}, {"format": "md"})
    assert out["artifacts"]["artifacts"] == []
    assert "__note__" in out


def test_data_export_events_json_and_csv(tmp_path, monkeypatch):
    _redirect_unique_filenames(monkeypatch, tmp_path)
    events_in = {
        "events": [
            {
                "id": "ev1",
                "participant": "P01",
                "detector": "change",
                "time_in": 1.0,
                "time_out": 2.5,
                "confidence": 0.9,
                "metadata": {"magnitude": 0.4},
            }
        ],
        "source": _SRC,
        "raw_results": [],
    }
    out = _run("data_export", _ctx(tmp_path), {"events": events_in}, {"format": "both"})
    arts = out["artifacts"]["artifacts"]
    assert [a["type"] for a in arts] == ["export", "export"]
    payload = json.loads((tmp_path / "export_events_P01.json").read_text())
    assert payload["records"][0]["id"] == "ev1"
    assert payload["records"][0]["magnitude"] == 0.4  # metadata hoisted
    csv_text = (tmp_path / "export_events_P01.csv").read_text()
    header = csv_text.splitlines()[0]
    # Preferred column order leads the CSV header.
    assert header.startswith("id,participant")
    assert "magnitude" in header


def test_data_export_segments_only_csv(tmp_path, monkeypatch):
    _redirect_unique_filenames(monkeypatch, tmp_path)
    seg_in = {
        "segments": [{"start": 0.0, "end": 2.0, "text": "hi there"}],
        "source": _SRC,
    }
    out = _run("data_export", _ctx(tmp_path), {"segments": seg_in}, {"format": "csv"})
    assert out["artifacts"]["count"] == 1
    csv_text = (tmp_path / "export_segments_P01.csv").read_text()
    assert "hi there" in csv_text
    assert "P01" in csv_text
    assert not (tmp_path / "export_segments_P01.json").exists()


def test_data_export_both_surfaces(tmp_path, monkeypatch):
    _redirect_unique_filenames(monkeypatch, tmp_path)
    events_in = {
        "events": [{"id": "e", "time_in": 0, "time_out": 1}],
        "source": _SRC,
        "raw_results": [],
    }
    seg_in = {"segments": [{"start": 0, "end": 1, "text": "t"}], "source": _SRC}
    out = _run(
        "data_export",
        _ctx(tmp_path),
        {"events": events_in, "segments": seg_in},
        {"format": "json"},
    )
    assert out["artifacts"]["count"] == 2
    assert (tmp_path / "export_events_P01.json").exists()
    assert (tmp_path / "export_segments_P01.json").exists()


def test_data_export_notes_when_nothing_wired(tmp_path):
    out = _run("data_export", _ctx(tmp_path), {}, {})
    assert out["artifacts"]["artifacts"] == []
    assert "__note__" in out


def test_heatmap_rolling_gif_passes_window(tmp_path, monkeypatch):
    import screenspace_heatmap

    monkeypatch.setattr(
        video, "probe_video_properties", lambda p: {"width": 640, "height": 480}
    )
    _redirect_unique_filenames(monkeypatch, tmp_path)
    called = {}

    def fake_rolling(results, w, h, out_path, heatmap_type, num_frames, window_frames):
        called["window_frames"] = window_frames
        called["num_frames"] = num_frames
        called["heatmap_type"] = heatmap_type
        return {"path": out_path, "frames": num_frames}

    monkeypatch.setattr(
        screenspace_heatmap, "generate_rolling_heatmap_gif", fake_rolling
    )
    events_in = {"events": [], "source": _SRC, "raw_results": [{"change_grid": []}]}
    out = _run(
        "heatmap",
        _ctx(tmp_path),
        {"events": events_in},
        {"style": "change", "output": "rolling_gif", "frames": 12, "window": 3},
    )
    assert called == {"window_frames": 3, "num_frames": 12, "heatmap_type": "change"}
    assert out["artifacts"]["artifacts"][0]["file"] == "heatmap.gif"


def test_heatmap_gif_none_releases_reservation(tmp_path, monkeypatch):
    import screenspace_heatmap

    monkeypatch.setattr(
        video, "probe_video_properties", lambda p: {"width": 640, "height": 480}
    )
    _redirect_unique_filenames(monkeypatch, tmp_path)
    released = {}
    monkeypatch.setattr(
        files, "release_reservation", lambda p: released.setdefault("path", p)
    )
    monkeypatch.setattr(
        screenspace_heatmap, "generate_heatmap_gif", lambda *a, **kw: None
    )
    events_in = {"events": [], "source": _SRC, "raw_results": [{"change_grid": []}]}
    out = _run(
        "heatmap",
        _ctx(tmp_path),
        {"events": events_in},
        {"style": "change", "output": "gif"},
    )
    assert out["artifacts"]["artifacts"] == []
    assert "animated heatmap" in out["__note__"]
    assert released["path"].endswith("heatmap.gif")


def test_heatmap_default_output_stays_image(tmp_path, monkeypatch):
    import screenspace_heatmap

    monkeypatch.setattr(
        video, "probe_video_properties", lambda p: {"width": 640, "height": 480}
    )
    _redirect_unique_filenames(monkeypatch, tmp_path)
    monkeypatch.setattr(
        screenspace_heatmap, "generate_change_heatmap", lambda r, w, h, p: p
    )
    events_in = {"events": [], "source": _SRC, "raw_results": [{"change_grid": []}]}
    out = _run("heatmap", _ctx(tmp_path), {"events": events_in}, {"style": "change"})
    assert out["artifacts"]["artifacts"][0]["file"] == "heatmap.png"


def test_data_export_partial_write_failure_rolls_back(tmp_path, monkeypatch):
    # format "both" writes two files; if the second write fails, the first must
    # not orphan behind an artifact-less failed result.
    _redirect_unique_filenames(monkeypatch, tmp_path)
    calls = {"n": 0}
    real_write = Path.write_text

    def flaky_write(self, *args, **kwargs):
        calls["n"] += 1
        if calls["n"] == 2:
            raise OSError("disk full")
        return real_write(self, *args, **kwargs)

    monkeypatch.setattr(Path, "write_text", flaky_write)
    events_in = {
        "events": [{"id": "e", "time_in": 0, "time_out": 1}],
        "source": _SRC,
        "raw_results": [],
    }
    out = _run("data_export", _ctx(tmp_path), {"events": events_in}, {"format": "both"})
    assert out["artifacts"]["artifacts"] == []
    assert "__note__" in out
    # Neither the failed CSV nor the previously-written JSON remains.
    assert not (tmp_path / "export_events_P01.json").exists()
    assert not (tmp_path / "export_events_P01.csv").exists()
