"""Unit tests for the collection-algebra control nodes.

Exercises the filter / merge / partition / limit / dedup executors directly
(synchronously, no HTTP, no ffmpeg). These are pure single-pass nodes — none
touches ``NodeContext`` — so they're called with ``ctx=None``. The catalog/HTTP
surface is covered in tests/test_workflows_api.py.
"""

import workflows


def _exec(node_id, inputs, params=None):
    return workflows.NODE_TYPES[node_id]["execute"](None, inputs, params or {})


def _events(*specs):
    # specs: (confidence, time_in, time_out)
    return {
        "events": [
            {"confidence": c, "time_in": ti, "time_out": to} for c, ti, to in specs
        ],
        "source": {"participant": "P01"},
        "raw_results": [{"frame": 0}],
    }


def _clips(*specs):
    # specs: (category, start_str, end_str)
    return {
        "records": [
            {"category": cat, "desc": "d", "times": [(s, e)]} for cat, s, e in specs
        ],
        "study": "mystudy",
    }


def _segments(*specs):
    # specs: (start, end, text)
    return {
        "segments": [{"start": s, "end": e, "text": t} for s, e, t in specs],
        "source": {"participant": "P01"},
    }


# ---- filter ----


def test_filter_events_by_confidence_keeps_matches_and_preserves_envelope():
    env = _events((0.9, 1, 2), (0.3, 5, 6), (0.85, 8, 9))
    out = _exec(
        "filter_events",
        {"in": env},
        {"field": "confidence", "op": ">=", "value": "0.8"},
    )["out"]
    assert [e["confidence"] for e in out["events"]] == [0.9, 0.85]
    # Source lineage and raw_results survive the filter (downstream heatmap needs them).
    assert out["source"] == {"participant": "P01"}
    assert "raw_results" in out


def test_filter_events_by_duration():
    env = _events((0.5, 0, 1), (0.5, 0, 5))
    out = _exec(
        "filter_events", {"in": env}, {"field": "duration", "op": ">", "value": "2"}
    )["out"]
    assert [e["time_out"] for e in out["events"]] == [5]


def test_filter_clips_by_category_is_case_insensitive():
    env = _clips(("nav", "0:01", "0:05"), ("error", "0:10", "0:12"))
    out = _exec(
        "filter_clips", {"in": env}, {"field": "category", "op": "==", "value": "ERROR"}
    )["out"]
    assert len(out["records"]) == 1
    assert out["records"][0]["category"] == "error"
    assert out["study"] == "mystudy"  # envelope preserved


def test_filter_clips_by_duration_sums_record_spans():
    env = _clips(("a", "0:00", "0:01"), ("b", "0:00", "0:10"))
    out = _exec(
        "filter_clips", {"in": env}, {"field": "duration", "op": ">=", "value": "5"}
    )["out"]
    assert [r["category"] for r in out["records"]] == ["b"]


def test_filter_segments_contains_substring():
    env = _segments((0, 2, "hello world"), (3, 4, "goodbye"))
    out = _exec(
        "filter_segments",
        {"in": env},
        {"field": "text", "op": "contains", "value": "WORLD"},
    )["out"]
    assert [s["text"] for s in out["segments"]] == ["hello world"]


def test_filter_empty_value_or_unwired_input_is_safe():
    # An unwired input yields no items, not a crash.
    out = _exec(
        "filter_events", {}, {"field": "confidence", "op": ">=", "value": "0.5"}
    )
    assert out["out"]["events"] == []


# ---- filter's unmatched branch (subsumes the old partition family) ----


def test_filter_events_unmatched_is_complementary_and_total():
    env = _events((0.9, 1, 2), (0.3, 5, 6), (0.7, 8, 9))
    res = _exec(
        "filter_events",
        {"in": env},
        {"field": "confidence", "op": ">=", "value": "0.7"},
    )
    matched = res["out"]["events"]
    unmatched = res["unmatched"]["events"]
    assert [e["confidence"] for e in matched] == [0.9, 0.7]
    assert [e["confidence"] for e in unmatched] == [0.3]
    # The two branches cover the input exactly (no item lost or duplicated).
    assert len(matched) + len(unmatched) == len(env["events"])
    assert res["out"]["source"] == res["unmatched"]["source"] == env["source"]


# ---- merge ----


def test_merge_events_concatenates_and_unions_raw_results():
    a = {
        "events": [{"confidence": 0.9, "time_in": 1, "time_out": 2}],
        "raw_results": [1],
    }
    b = {
        "events": [{"confidence": 0.3, "time_in": 5, "time_out": 6}],
        "raw_results": [2],
    }
    c = {
        "events": [{"confidence": 0.6, "time_in": 9, "time_out": 10}],
        "raw_results": [3],
    }
    out = _exec("merge_events", {"in1": a, "in2": b, "in3": c}, {})["out"]
    assert len(out["events"]) == 3
    assert out["raw_results"] == [1, 2, 3]


def test_merge_clips_preserves_first_wired_study_and_skips_unwired():
    a = {"records": [{"times": [("0:01", "0:02")]}], "study": "s1"}
    b = {"records": [{"times": [("0:03", "0:04")]}], "study": "s2"}
    out = _exec("merge_clips", {"in1": a, "in3": b}, {})["out"]
    assert len(out["records"]) == 2
    assert out["study"] == "s1"  # first wired input's lineage


# ---- limit ----


def test_limit_events_sorts_desc_then_takes_n():
    env = _events((0.3, 1, 2), (0.9, 3, 4), (0.6, 5, 6))
    out = _exec(
        "limit_events",
        {"in": env},
        {"sort_by": "confidence", "order": "desc", "take": 2},
    )["out"]
    assert [e["confidence"] for e in out["events"]] == [0.9, 0.6]


def test_limit_ascending_and_no_sort_keeps_order():
    env = _events((0.3, 1, 2), (0.9, 3, 4), (0.6, 5, 6))
    asc = _exec(
        "limit_events",
        {"in": env},
        {"sort_by": "confidence", "order": "asc", "take": 2},
    )["out"]
    assert [e["confidence"] for e in asc["events"]] == [0.3, 0.6]
    # sort_by "none" leaves input order untouched, take still applies.
    none = _exec("limit_events", {"in": env}, {"sort_by": "none", "take": 1})["out"]
    assert [e["confidence"] for e in none["events"]] == [0.3]


# ---- dedup ----


def test_dedup_events_merges_overlapping_spans_keeping_max_confidence():
    env = {
        "events": [
            {"confidence": 0.5, "time_in": 1, "time_out": 3},
            {"confidence": 0.9, "time_in": 2, "time_out": 4},
            {"confidence": 0.4, "time_in": 10, "time_out": 11},
        ]
    }
    out = _exec("dedup_events", {"in": env}, {"gap": 0})["out"]
    assert len(out["events"]) == 2
    first = out["events"][0]
    assert (first["time_in"], first["time_out"], first["confidence"]) == (1, 4, 0.9)


def test_dedup_events_gap_bridges_nearby_spans():
    env = {
        "events": [
            {"confidence": 0.5, "time_in": 0, "time_out": 1},
            {"confidence": 0.6, "time_in": 2, "time_out": 3},
        ]
    }
    # gap 0 keeps them separate; gap 1.5 bridges the 1s hole.
    assert len(_exec("dedup_events", {"in": env}, {"gap": 0})["out"]["events"]) == 2
    assert len(_exec("dedup_events", {"in": env}, {"gap": 1.5})["out"]["events"]) == 1


def test_dedup_clips_drops_overlapping_records_keeping_first():
    env = {
        "records": [
            {"category": "keep", "times": [("0:00", "0:05")]},
            {"category": "drop", "times": [("0:03", "0:06")]},
            {"category": "keep2", "times": [("0:20", "0:25")]},
        ],
        "study": "s",
    }
    out = _exec("dedup_clips", {"in": env}, {"gap": 0})["out"]
    assert [r["category"] for r in out["records"]] == ["keep", "keep2"]
    assert out["study"] == "s"


# ---- timeRange family (tuple items) ----


def _ranges(*spans):
    return {"ranges": [tuple(s) for s in spans], "source": {"participant": "P01"}}


def test_filter_timerange_by_duration_preserves_source():
    env = _ranges((0.0, 2.0), (5.0, 5.5), (10.0, 14.0))
    out = _exec(
        "filter_timerange", {"in": env}, {"field": "duration", "op": ">=", "value": "2"}
    )["out"]
    assert out["ranges"] == [(0.0, 2.0), (10.0, 14.0)]
    assert out["source"] == {"participant": "P01"}


def test_limit_timerange_sorts_by_duration():
    env = _ranges((0.0, 2.0), (5.0, 5.5), (10.0, 14.0))
    out = _exec(
        "limit_timerange",
        {"in": env},
        {"sort_by": "duration", "order": "desc", "take": 1},
    )["out"]
    assert out["ranges"] == [(10.0, 14.0)]


def test_merge_timerange_concatenates_and_keeps_first_source():
    a = {"ranges": [(0.0, 1.0)], "source": {"participant": "P01"}}
    b = {"ranges": [(5.0, 6.0)]}
    out = _exec("merge_timerange", {"in1": a, "in2": b}, {})["out"]
    assert out["ranges"] == [(0.0, 1.0), (5.0, 6.0)]
    assert out["source"] == {"participant": "P01"}


def test_dedup_timerange_merges_overlapping_windows():
    env = {"ranges": [(0.0, 3.0), (2.0, 4.0), (10.0, 11.0)]}
    out = _exec("dedup_timerange", {"in": env}, {"gap": 0})["out"]
    assert out["ranges"] == [(0.0, 4.0), (10.0, 11.0)]


# ---- artifacts family (output-side, with count) ----


def _arts(*specs):
    # specs: (type, start, end, category)
    return {
        "artifacts": [
            {"type": t, "start": s, "end": e, "category": c} for t, s, e, c in specs
        ],
        "study": "s",
        "count": 0,  # stale on purpose — must be recomputed to the kept length
    }


def test_filter_artifacts_by_type_recounts_and_preserves_study():
    env = _arts(("clip", 0, 5, "nav"), ("clip", 0, 1, "err"), ("timelapse", 0, 0, ""))
    out = _exec(
        "filter_artifacts", {"in": env}, {"field": "type", "op": "==", "value": "clip"}
    )["out"]
    assert len(out["artifacts"]) == 2
    assert out["count"] == 2  # recomputed, not the stale 0
    assert out["study"] == "s"


def test_limit_artifacts_top_n_by_duration():
    env = _arts(("clip", 0, 1, "a"), ("clip", 0, 9, "b"), ("clip", 0, 3, "c"))
    out = _exec(
        "limit_artifacts",
        {"in": env},
        {"sort_by": "duration", "order": "desc", "take": 1},
    )["out"]
    assert [a["category"] for a in out["artifacts"]] == ["b"]
    assert out["count"] == 1


def test_filter_artifacts_branches_by_duration():
    env = _arts(("clip", 0, 5, "long"), ("clip", 0, 1, "short"))
    res = _exec(
        "filter_artifacts",
        {"in": env},
        {"field": "duration", "op": ">", "value": "2"},
    )
    assert [a["category"] for a in res["out"]["artifacts"]] == ["long"]
    assert [a["category"] for a in res["unmatched"]["artifacts"]] == ["short"]
    assert res["out"]["count"] == 1 and res["unmatched"]["count"] == 1


def test_dedup_clips_untimed_record_does_not_reset_overlap_tracker():
    # A record with no parseable timestamps (span None) must pass through without
    # clobbering last_span — otherwise a later overlapping duplicate would leak.
    env = {
        "records": [
            {"category": "keep", "times": [("0:00", "0:05")]},
            {"category": "untimed", "times": []},
            {"category": "drop", "times": [("0:02", "0:06")]},
        ],
        "study": "s",
    }
    out = _exec("dedup_clips", {"in": env}, {"gap": 0})["out"]
    cats = [r["category"] for r in out["records"]]
    assert "drop" not in cats  # the later overlapping record is still merged away
    assert "keep" in cats and "untimed" in cats


# ---- compound clause (combine / field2 / op2 / value2) ----


def test_filter_events_and_clause_narrows():
    env = _events((0.9, 1, 2), (0.9, 8, 9), (0.3, 1, 2))
    out = _exec(
        "filter_events",
        {"in": env},
        {
            "field": "confidence",
            "op": ">=",
            "value": "0.8",
            "combine": "AND",
            "field2": "start",
            "op2": "<",
            "value2": "5",
        },
    )["out"]
    # Only the high-confidence event that also starts before 5s survives.
    assert [(e["confidence"], e["time_in"]) for e in out["events"]] == [(0.9, 1)]


def test_filter_events_or_clause_widens():
    env = _events((0.9, 1, 2), (0.3, 8, 9), (0.2, 1, 2))
    out = _exec(
        "filter_events",
        {"in": env},
        {
            "field": "confidence",
            "op": ">=",
            "value": "0.8",
            "combine": "OR",
            "field2": "start",
            "op2": ">=",
            "value2": "5",
        },
    )["out"]
    # High confidence OR late start; the low-confidence early event is dropped.
    assert [(e["confidence"], e["time_in"]) for e in out["events"]] == [
        (0.9, 1),
        (0.3, 8),
    ]


def test_filter_combine_off_ignores_second_clause():
    env = _events((0.9, 1, 2), (0.3, 5, 6))
    out = _exec(
        "filter_events",
        {"in": env},
        {
            "field": "confidence",
            "op": ">=",
            "value": "0.8",
            "combine": "off",
            "field2": "start",
            "op2": ">=",
            "value2": "999",  # would kill everything if evaluated
        },
    )["out"]
    assert [e["confidence"] for e in out["events"]] == [0.9]


def test_filter_segments_mixed_text_and_numeric_clauses():
    env = _segments(
        (0, 10, "checkout was confusing"), (0, 1, "checkout ok"), (0, 10, "fine")
    )
    out = _exec(
        "filter_segments",
        {"in": env},
        {
            "field": "text",
            "op": "contains",
            "value": "checkout",
            "combine": "AND",
            "field2": "duration",
            "op2": ">=",
            "value2": "5",
        },
    )["out"]
    assert [s["text"] for s in out["segments"]] == ["checkout was confusing"]


def test_filter_compound_clause_is_complementary():
    env = _events((0.9, 1, 2), (0.5, 8, 9), (0.2, 3, 4))
    res = _exec(
        "filter_events",
        {"in": env},
        {
            "field": "confidence",
            "op": ">=",
            "value": "0.8",
            "combine": "OR",
            "field2": "start",
            "op2": ">=",
            "value2": "5",
        },
    )
    matched = res["out"]["events"]
    unmatched = res["unmatched"]["events"]
    assert len(matched) + len(unmatched) == 3
    assert [e["confidence"] for e in matched] == [0.9, 0.5]
    assert [e["confidence"] for e in unmatched] == [0.2]
