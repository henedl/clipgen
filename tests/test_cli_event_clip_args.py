"""Tests for the event-driven clip-cutting CLI flags (--ss-clips, --transcript-clips)."""

import pytest

import cli
import files

# Reuse the comprehensive Namespace factory from the screenspace test file so that
# any newly added argparse keys flow through one place.
from test_cli_screenspace_args import _ss_args


# ---- Argparse parsing ----


def test_parse_ss_clips_minimal(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--ss-clips"])
    args = cli.parse_arguments()
    assert args.ss_clips is True
    assert args.transcript_clips is False
    # Defaults match the user-confirmed plan choices.
    assert args.cluster_gap == 5.0
    assert args.clip_pre == 5.0
    assert args.clip_post == 5.0
    assert args.max_clip_duration == 0.0


def test_parse_ss_clips_with_filters(monkeypatch):
    monkeypatch.setattr(
        "sys.argv",
        [
            "clipgen.py",
            "--ss-clips",
            "--ss-clips-detector",
            "change,color",
            "--ss-clips-min-confidence",
            "0.7",
            "--ss-clips-region",
            "dialog",
            "--cluster-gap",
            "3",
            "--clip-pre",
            "1",
            "--clip-post",
            "2",
        ],
    )
    args = cli.parse_arguments()
    assert args.ss_clips_detector == "change,color"
    assert args.ss_clips_min_confidence == 0.7
    assert args.ss_clips_region == "dialog"
    assert args.cluster_gap == 3.0
    assert args.clip_pre == 1.0
    assert args.clip_post == 2.0


def test_parse_transcript_clips_with_mark(monkeypatch):
    monkeypatch.setattr(
        "sys.argv",
        [
            "clipgen.py",
            "--transcript-clips",
            "--transcript-clips-mark",
            "insight",
            "--transcript-clips-text",
            "checkout",
        ],
    )
    args = cli.parse_arguments()
    assert args.transcript_clips is True
    assert args.transcript_clips_mark == "insight"
    assert args.transcript_clips_text == "checkout"


# ---- Conflict validation ----


@pytest.mark.parametrize(
    "args_overrides,expects_exit",
    [
        ({"ss_clips": True, "transcript_clips": True}, True),
        ({"ss_clips": True, "studio": True}, True),
        (
            {
                "ss_clips": True,
                "ss_task": ["color", "P01", "btn"],
                "ss_target_color": "#FF0000",
                "ss_tolerance": "20,30,30",
                "ss_threshold": 0.85,
            },
            True,
        ),
        ({"transcript_clips": True, "screenspace": True}, True),
        # Previously-uncaught pairs: --ss-run-task was in no other mode's
        # conflict list and vice-versa. The all-pairs derivation closes the gap.
        ({"ss_run_task": "ss_1", "ss_clips": True}, True),
        ({"ss_run_task": "ss_1", "transcript_mark": "insight"}, True),
        ({"ss_clips": True}, False),
    ],
    ids=[
        "ss_clips_with_transcript_clips",
        "ss_clips_with_studio",
        "ss_clips_with_ss_task",
        "transcript_clips_with_screenspace",
        "ss_run_task_with_ss_clips",
        "ss_run_task_with_transcript_mark",
        "ss_clips_alone",
    ],
)
def test_clip_mode_conflicts(args_overrides, expects_exit):
    args = _ss_args(**args_overrides)
    if expects_exit:
        with pytest.raises(SystemExit):
            cli._validate_mode_conflicts(args)
    else:
        modes = cli._validate_mode_conflicts(args)
        assert modes["ss_clips"] is True
        assert modes["transcript_clips"] is False


# ---- Filter helpers ----


def _ev(**overrides):
    base = {
        "id": "ev_1",
        "source_video": "study_P01.mp4",
        "participant": "P01",
        "detector": "change",
        "event_type": "change: dialog",
        "time_in": 10.0,
        "time_out": 10.0,
        "confidence": 0.9,
        "metadata": {},
        "excluded": False,
        "task_id": "ss_1",
        "region": "dialog",
    }
    base.update(overrides)
    return base


_NO_FILTERS = dict(
    detectors=None,
    regions=None,
    participants=None,
    min_confidence=None,
    event_type_substr=None,
)


@pytest.mark.parametrize(
    "events,filter_overrides,expected_ids",
    [
        (
            [_ev(id="a"), _ev(id="b", excluded=True)],
            {},
            ["a"],
        ),
        (
            [_ev(id="lo", confidence=0.5), _ev(id="hi", confidence=0.95)],
            {"min_confidence": 0.8},
            ["hi"],
        ),
        (
            [
                _ev(id="a", detector="change", region="dialog"),
                _ev(id="b", detector="color", region="dialog"),
                _ev(id="c", detector="change", region="header"),
            ],
            {"detectors": {"change"}, "regions": {"dialog"}},
            ["a"],
        ),
        (
            [
                _ev(id="a", event_type="Login attempt"),
                _ev(id="b", event_type="logout"),
            ],
            {"event_type_substr": "LOGIN"},
            ["a"],
        ),
    ],
    ids=[
        "drops_excluded",
        "min_confidence",
        "detector_and_region",
        "event_type_substr",
    ],
)
def test_filter_ss_events(events, filter_overrides, expected_ids):
    out = cli._filter_screenspace_events(events, **{**_NO_FILTERS, **filter_overrides})
    assert [e["id"] for e in out] == expected_ids


def test_filter_transcript_segments_by_mark_category():
    manifest = {
        "source_transcripts": {
            "P01": {
                "segments": [
                    {"id": "P01:0", "start": 0.0, "end": 5.0, "text": "hello"},
                    {"id": "P01:1", "start": 5.0, "end": 10.0, "text": "world"},
                ]
            }
        },
        "marks": [
            {"id": "m1", "segment_id": "P01:0", "category": "insight"},
            {"id": "m2", "segment_id": "P01:1", "category": "action"},
        ],
    }
    rows = cli._filter_transcript_segments(
        manifest,
        participants=None,
        mark_categories={"insight"},
        text_substr=None,
    )
    assert len(rows) == 1
    pid, seg, marks = rows[0]
    assert pid == "P01"
    assert seg["id"] == "P01:0"
    assert [m["id"] for m in marks] == ["m1"]


def test_filter_transcript_segments_by_text_substr():
    manifest = {
        "source_transcripts": {
            "P01": {
                "segments": [
                    {"id": "P01:0", "start": 0.0, "end": 5.0, "text": "Checkout flow"},
                    {"id": "P01:1", "start": 5.0, "end": 10.0, "text": "Other thing"},
                ]
            }
        },
        "marks": [],
    }
    rows = cli._filter_transcript_segments(
        manifest,
        participants=None,
        mark_categories=None,
        text_substr="checkout",
    )
    assert [r[1]["id"] for r in rows] == ["P01:0"]


# ---- Cluster builders ----


def test_build_ss_clusters_groups_by_participant_and_detector():
    events = [
        _ev(id="a", participant="P01", detector="change", time_in=1.0, time_out=1.0),
        _ev(id="b", participant="P01", detector="change", time_in=2.0, time_out=2.0),
        _ev(id="c", participant="P02", detector="change", time_in=1.5, time_out=1.5),
        _ev(id="d", participant="P01", detector="color", time_in=1.5, time_out=1.5),
    ]
    clusters = cli._build_clusters_from_ss_events(
        events, gap=5.0, pad_pre=0.0, pad_post=0.0, max_duration=0.0
    )
    # P01/change merges (a, b); P01/color stays alone; P02/change stays alone.
    keys = sorted((c["participant"], c["detector"]) for c in clusters)
    assert keys == [("P01", "change"), ("P01", "color"), ("P02", "change")]
    p01_change = next(
        c for c in clusters if c["participant"] == "P01" and c["detector"] == "change"
    )
    assert sorted(p01_change["member_event_ids"]) == ["a", "b"]


def test_build_ss_clusters_applies_padding():
    events = [_ev(id="a", time_in=10.0, time_out=10.0)]
    clusters = cli._build_clusters_from_ss_events(
        events, gap=0.0, pad_pre=2.0, pad_post=3.0, max_duration=0.0
    )
    assert len(clusters) == 1
    assert clusters[0]["start"] == 8.0
    assert clusters[0]["end"] == 13.0


def test_build_transcript_clusters_groups_by_participant():
    manifest = {
        "source_transcripts": {
            "P01": {
                "segments": [
                    {"id": "P01:0", "start": 0.0, "end": 2.0, "text": "alpha"},
                    {"id": "P01:1", "start": 3.0, "end": 5.0, "text": "beta"},
                ],
                "source_file": "/data/study_P01.mp4",
            },
            "P02": {
                "segments": [
                    {"id": "P02:0", "start": 0.0, "end": 2.0, "text": "gamma"},
                ],
                "source_file": "/data/study_P02.mp4",
            },
        },
        "marks": [],
    }
    rows = cli._filter_transcript_segments(
        manifest, participants=None, mark_categories=None, text_substr=None
    )
    clusters = cli._build_clusters_from_transcript_segments(
        rows, manifest, gap=2.0, pad_pre=0.0, pad_post=0.0, max_duration=0.0
    )
    by_pid = {c["participant"]: c for c in clusters}
    assert set(by_pid.keys()) == {"P01", "P02"}
    # P01's two segments merge with gap=2.0 (gap between them is 1.0).
    assert by_pid["P01"]["start"] == 0.0
    assert by_pid["P01"]["end"] == 5.0
    assert "alpha" in by_pid["P01"]["text"]
    assert "beta" in by_pid["P01"]["text"]
    assert by_pid["P01"]["source_video"] == "study_P01.mp4"


# ---- Synthetic clip builder ----


def test_build_clip_records_uses_negative_row_and_padded_times():
    rec = files.build_clip_records(
        participant="P01",
        source_filename="mystudy_P01.mp4",
        time_ranges=[(10.0, 20.0)],
        description="change dialog",
        category="screenspace-change",
        study="mystudy",
        cell_col=1,
    )[0]
    assert rec["cell"].row == -1
    assert rec["cell"].col == 1
    # Pre-filled times trigger the prepare_clip fast path.
    assert rec["times"] == [("0:00:10", "0:00:20")]
    assert rec["source_filename"] == "mystudy_P01.mp4"


def test_build_clip_records_namespaces_by_cell_row_base_and_col():
    a = files.build_clip_records(
        participant="P01",
        source_filename="s_P01.mp4",
        time_ranges=[(0, 1)],
        description="d",
        category="c",
        study="s",
        cell_col=1,
        cell_row_base=0,
    )[0]
    b = files.build_clip_records(
        participant="P01",
        source_filename="s_P01.mp4",
        time_ranges=[(0, 1)],
        description="d",
        category="c",
        study="s",
        cell_col=2,
        cell_row_base=5,
    )[0]
    assert a["cell"].row == -1
    assert b["cell"].row == -6
    assert a["cell"].col == 1
    assert b["cell"].col == 2


# ---- Smoke dispatch tests ----


def test_run_ss_clips_smoke_dispatches_pipeline(monkeypatch):
    """Wires manifest -> filter -> cluster -> process_clips -> save_manifest."""
    import pipeline
    import screenspace
    import viewer

    fake_events = [
        {
            "id": "ev1",
            "source_video": "study_P01.mp4",
            "participant": "P01",
            "detector": "change",
            "event_type": "change: dialog",
            "time_in": 10.0,
            "time_out": 10.0,
            "confidence": 0.9,
            "excluded": False,
            "region": "dialog",
        }
    ]
    monkeypatch.setattr(
        screenspace,
        "load_screenspace_manifest",
        lambda: {"events": fake_events, "regions": {}, "tasks": [], "stashes": []},
    )
    captured: dict = {}

    def fake_process(clips_list, output_format="clip", include_severity=False, **kw):
        captured["clips"] = clips_list
        captured["output_format"] = output_format
        return (1, [{"id": "a-1c1s0", "file": "out.mp4"}])

    monkeypatch.setattr(pipeline, "process_clips", fake_process)
    saved: dict = {}

    def fake_save(artifacts, **kw):
        saved["artifacts"] = artifacts
        saved["mode"] = kw.get("mode")

    monkeypatch.setattr(viewer, "save_manifest", fake_save)

    args = _ss_args(ss_clips=True)
    cli._run_ss_clips(args)

    assert len(captured["clips"]) == 1
    clip = captured["clips"][0]
    assert clip["participant"] == "P01"
    assert clip["category"] == "screenspace-change"
    assert clip["source_filename"] == "study_P01.mp4"
    assert saved["mode"] == "ss-clips"
    assert saved["artifacts"][0]["id"] == "a-1c1s0"


def test_run_ss_clips_no_events_warns(monkeypatch, capsys):
    import screenspace

    monkeypatch.setattr(
        screenspace,
        "load_screenspace_manifest",
        lambda: {"events": [], "regions": {}, "tasks": [], "stashes": []},
    )
    args = _ss_args(ss_clips=True)
    cli._run_ss_clips(args)
    out = capsys.readouterr().out
    assert "No Screenspace events" in out


def test_run_transcript_clips_smoke_dispatches_pipeline(monkeypatch):
    import pipeline
    import transcripts
    import viewer

    fake_manifest = {
        "source_transcripts": {
            "P01": {
                "segments": [
                    {"id": "P01:0", "start": 0.0, "end": 2.0, "text": "alpha beta"},
                ],
                "source_file": "/data/study_P01.mp4",
            }
        },
        "corrections": [],
        "marks": [],
    }
    monkeypatch.setattr(transcripts, "load_transcripts_manifest", lambda: fake_manifest)
    captured: dict = {}

    def fake_process(clips_list, output_format="clip", include_severity=False, **kw):
        captured["clips"] = clips_list
        return (1, [{"id": "a-1c2s0", "file": "out.mp4"}])

    monkeypatch.setattr(pipeline, "process_clips", fake_process)
    saved: dict = {}

    def fake_save(artifacts, **kw):
        saved["artifacts"] = artifacts
        saved["mode"] = kw.get("mode")

    monkeypatch.setattr(viewer, "save_manifest", fake_save)

    args = _ss_args(transcript_clips=True)
    cli._run_transcript_clips(args)

    assert len(captured["clips"]) == 1
    clip = captured["clips"][0]
    assert clip["participant"] == "P01"
    assert clip["category"] == "transcript"
    assert clip["source_filename"] == "study_P01.mp4"
    assert saved["mode"] == "transcript-clips"


def test_run_transcript_clips_with_mark_filter_uses_mark_category(monkeypatch):
    import pipeline
    import transcripts
    import viewer

    fake_manifest = {
        "source_transcripts": {
            "P01": {
                "segments": [
                    {"id": "P01:0", "start": 0.0, "end": 2.0, "text": "interesting"},
                ],
                "source_file": "/data/study_P01.mp4",
            }
        },
        "marks": [{"id": "m1", "segment_id": "P01:0", "category": "insight"}],
    }
    monkeypatch.setattr(transcripts, "load_transcripts_manifest", lambda: fake_manifest)
    captured: dict = {}

    def fake_process(clips_list, output_format="clip", include_severity=False, **kw):
        captured["clips"] = clips_list
        return (1, [{"id": "a-1c2s0", "file": "out.mp4"}])

    monkeypatch.setattr(pipeline, "process_clips", fake_process)
    monkeypatch.setattr(viewer, "save_manifest", lambda *_a, **_k: None)

    args = _ss_args(transcript_clips=True, transcript_clips_mark="insight")
    cli._run_transcript_clips(args)

    assert captured["clips"][0]["category"] == "mark-insight"


# ---- --transcript-mark argparse + conflict ----


@pytest.mark.parametrize(
    "argv_extra,expected_attrs",
    [
        (
            [
                "--transcript-mark",
                "checkout",
                "--transcript-mark-category",
                "insight",
            ],
            {
                "transcript_mark": "checkout",
                "transcript_mark_category": "insight",
                "transcript_mark_participant": None,
                "transcript_mark_label": None,
            },
        ),
        (
            [
                "--transcript-mark",
                "checkout flow",
                "--transcript-mark-category",
                "pain_point",
                "--transcript-mark-participant",
                "P01,P02",
                "--transcript-mark-label",
                "follow up",
            ],
            {
                "transcript_mark": "checkout flow",
                "transcript_mark_category": "pain_point",
                "transcript_mark_participant": "P01,P02",
                "transcript_mark_label": "follow up",
            },
        ),
    ],
    ids=["minimal", "with_filters"],
)
def test_parse_transcript_mark(monkeypatch, argv_extra, expected_attrs):
    monkeypatch.setattr("sys.argv", ["clipgen.py", *argv_extra])
    args = cli.parse_arguments()
    for attr, expected in expected_attrs.items():
        assert getattr(args, attr) == expected


def test_transcript_mark_conflicts_with_transcript_clips():
    args = _ss_args(transcript_mark="x", transcript_clips=True)
    with pytest.raises(SystemExit):
        cli._validate_mode_conflicts(args)


def test_transcript_mark_conflicts_with_studio():
    args = _ss_args(transcript_mark="x", studio=True)
    with pytest.raises(SystemExit):
        cli._validate_mode_conflicts(args)


def test_transcript_mark_alone_validates():
    args = _ss_args(transcript_mark="x")
    modes = cli._validate_mode_conflicts(args)
    assert modes["transcript_mark"] is True
    assert modes["transcript_clips"] is False


# ---- _run_transcript_mark behavior ----


@pytest.fixture
def no_running_server(monkeypatch):
    """Force _run_transcript_mark to take the direct-disk-write path."""
    monkeypatch.setattr(cli, "_post_marks_to_running_server", lambda *_a, **_k: None)


def _mark_manifest():
    return {
        "source_transcripts": {
            "P01": {
                "segments": [
                    {
                        "id": "P01:0",
                        "start": 0.0,
                        "end": 2.0,
                        "text": "Checkout flow is slow",
                    },
                    {"id": "P01:1", "start": 2.0, "end": 4.0, "text": "Login is fine"},
                ],
                "source_file": "/data/study_P01.mp4",
            },
            "P02": {
                "segments": [
                    {
                        "id": "P02:0",
                        "start": 0.0,
                        "end": 2.0,
                        "text": "Talked about CHECKOUT",
                    },
                ],
                "source_file": "/data/study_P02.mp4",
            },
        },
        "corrections": [],
        "marks": [],
    }


@pytest.mark.parametrize(
    "args_overrides,expected_seg_ids",
    [
        (
            dict(transcript_mark="checkout", transcript_mark_category="insight"),
            ["P01:0", "P02:0"],
        ),
        (
            dict(
                transcript_mark="checkout",
                transcript_mark_category="insight",
                transcript_mark_participant="P01",
            ),
            ["P01:0"],
        ),
    ],
    ids=["all_participants", "participant_filter"],
)
def test_run_transcript_mark_creates_marks(
    monkeypatch, no_running_server, args_overrides, expected_seg_ids
):
    import transcripts

    manifest = _mark_manifest()
    monkeypatch.setattr(transcripts, "load_transcripts_manifest", lambda: manifest)
    saved: dict = {}
    monkeypatch.setattr(
        transcripts,
        "save_transcripts_manifest",
        lambda src, corr, marks=None: saved.update({"marks": marks}),
    )

    args = _ss_args(**args_overrides)
    cli._run_transcript_mark(args)

    assert sorted(m["segment_id"] for m in saved["marks"]) == expected_seg_ids
    assert all(m["category"] == "insight" for m in saved["marks"])
    assert all(m["label"] is None for m in saved["marks"])


def test_run_transcript_mark_updates_existing_in_place(monkeypatch, no_running_server):
    import transcripts

    manifest = _mark_manifest()
    manifest["marks"] = [
        {
            "id": "m_old1",
            "segment_id": "P01:0",
            "category": "bookmark",
            "label": "old",
            "created": "2025-01-01T00:00:00+00:00",
        },
        {
            "id": "m_old2",
            "segment_id": "P01:1",
            "category": "delight",
            "label": None,
            "created": "2025-01-01T00:00:00+00:00",
        },
    ]
    monkeypatch.setattr(transcripts, "load_transcripts_manifest", lambda: manifest)
    saved: dict = {}
    monkeypatch.setattr(
        transcripts,
        "save_transcripts_manifest",
        lambda src, corr, marks=None: saved.update({"marks": marks}),
    )

    args = _ss_args(
        transcript_mark="checkout",
        transcript_mark_category="insight",
        transcript_mark_label="new label",
    )
    cli._run_transcript_mark(args)

    by_seg = {m["segment_id"]: m for m in saved["marks"]}
    # P01:0 matched: updated in place (id preserved), category + label changed.
    assert by_seg["P01:0"]["id"] == "m_old1"
    assert by_seg["P01:0"]["category"] == "insight"
    assert by_seg["P01:0"]["label"] == "new label"
    # P01:1 did NOT match: untouched.
    assert by_seg["P01:1"]["id"] == "m_old2"
    assert by_seg["P01:1"]["category"] == "delight"
    assert by_seg["P01:1"]["label"] is None
    # P02:0 matched and was new: appended.
    assert by_seg["P02:0"]["category"] == "insight"
    # No duplicates.
    assert len(saved["marks"]) == 3


def test_run_transcript_mark_invalid_category_does_not_save(
    monkeypatch, capsys, no_running_server
):
    import transcripts

    manifest = _mark_manifest()
    monkeypatch.setattr(transcripts, "load_transcripts_manifest", lambda: manifest)
    called: dict = {"saved": False}

    def fake_save(*_a, **_k):
        called["saved"] = True

    monkeypatch.setattr(transcripts, "save_transcripts_manifest", fake_save)

    args = _ss_args(
        transcript_mark="checkout", transcript_mark_category="not_a_category"
    )
    cli._run_transcript_mark(args)

    assert called["saved"] is False
    out = capsys.readouterr().out
    assert "transcript-mark-category" in out


def test_run_transcript_mark_no_matches_warns(monkeypatch, capsys, no_running_server):
    import transcripts

    manifest = _mark_manifest()
    monkeypatch.setattr(transcripts, "load_transcripts_manifest", lambda: manifest)
    called: dict = {"saved": False}
    monkeypatch.setattr(
        transcripts,
        "save_transcripts_manifest",
        lambda *_a, **_k: called.update({"saved": True}),
    )

    args = _ss_args(transcript_mark="zzz", transcript_mark_category="insight")
    cli._run_transcript_mark(args)

    assert called["saved"] is False
    out = capsys.readouterr().out
    assert "No transcript segments contain" in out


def test_run_transcript_mark_routes_through_running_server(monkeypatch, capsys):
    """When the Transcripts server is running, posts via API and skips disk write."""
    import transcripts

    manifest = _mark_manifest()
    monkeypatch.setattr(transcripts, "load_transcripts_manifest", lambda: manifest)

    posted: dict = {}

    def fake_post(seg_ids, category, label):
        posted["seg_ids"] = list(seg_ids)
        posted["category"] = category
        posted["label"] = label
        return {"ok": True, "marks": [{"segment_id": s} for s in seg_ids]}

    monkeypatch.setattr(cli, "_post_marks_to_running_server", fake_post)

    saved: dict = {"called": False}
    monkeypatch.setattr(
        transcripts,
        "save_transcripts_manifest",
        lambda *_a, **_k: saved.update({"called": True}),
    )

    args = _ss_args(transcript_mark="checkout", transcript_mark_category="insight")
    cli._run_transcript_mark(args)

    assert sorted(posted["seg_ids"]) == ["P01:0", "P02:0"]
    assert posted["category"] == "insight"
    assert saved["called"] is False
    out = capsys.readouterr().out
    assert "via running Transcripts server" in out
