"""Tests for data_export module: builders, CSV writer, and bundle writer."""

import csv
import io
import json
from pathlib import Path

import pytest

import config
import data_export
import utils


# ---- Fixtures -----------------------------------------------------------


def _ss_event(detector, **overrides):
    """Build a Screenspace event dict matching screenspace.create_event() shape."""
    base = {
        "id": f"ev_{detector}_1",
        "source_video": "study_P01.mp4",
        "participant": "P01",
        "detector": detector,
        "event_type": detector + ": region1",
        "time_in": 12.5,
        "time_out": 12.5,
        "confidence": 0.85,
        "metadata": {},
        "excluded": False,
        "task_id": "ss_abc12345",
        "region": "region1",
    }
    base.update(overrides)
    return base


@pytest.fixture
def screenspace_manifest():
    return {
        "regions": {},
        "tasks": [],
        "events": [
            _ss_event("color"),
            _ss_event("change", metadata={"magnitude": 0.42}),
            _ss_event("similarity", metadata={"score": 0.91}),
            _ss_event("text", metadata={"text_found": "Game Over"}),
            _ss_event("numbers", metadata={"value": 250}),
            _ss_event("template", metadata={"match_count": 3, "best_score": 0.78}),
            _ss_event("flow", metadata={"magnitude": 0.55, "angle": 90.0}),
            _ss_event("scene", metadata={"scene_name": "menu", "score": 0.93}),
            _ss_event(
                "inactivity",
                time_in=20.0,
                time_out=35.5,
                metadata={"duration": 15.5, "avg_distance": 0.02},
            ),
            _ss_event(
                "color",
                id="ev_excluded",
                participant="P02",
                excluded=True,
            ),
        ],
        "stashes": [],
    }


@pytest.fixture
def insights_manifest():
    return {
        "meta": {"version": "0.0.0"},
        "insights": [
            {
                "id": "ins_abc12345",
                "title": "Players miss the save button",
                "summary": "Three of five players failed to find Save.",
                "severity": "High",
                "status": "final",
                "createdAt": "2026-04-01T10:00:00+00:00",
                "updatedAt": "2026-04-02T10:00:00+00:00",
                "timelineContext": "early-game tutorial",
                "causes": {
                    "narrative": "Save icon is below the fold.",
                    "artifacts": ["a1", "a2"],
                },
                "behaviors": {
                    "narrative": "Users scroll up looking for it.",
                    "artifacts": [],
                },
                "impacts": {
                    "narrative": "Lost progress on quit.",
                    "artifacts": ["a3"],
                },
            }
        ],
    }


@pytest.fixture
def transcripts_manifest():
    return {
        "source_transcripts": {
            "P01": {
                "language": "en",
                "model": "base",
                "source_file": "study_P01.mp4",
                "transcribed_at": "2026-04-01T10:00:00+00:00",
                "segments": [
                    {"id": "P01:0", "start": 0.0, "end": 5.0, "text": "Hello there."},
                    {
                        "id": "P01:1",
                        "start": 5.0,
                        "end": 10.0,
                        "text": "What's this button?",
                    },
                ],
            },
            "P02": {
                "language": "en",
                "model": "base",
                "source_file": "study_P02.mp4",
                "transcribed_at": "2026-04-01T11:00:00+00:00",
                "segments": [
                    {"id": "P02:0", "start": 0.0, "end": 4.5, "text": "Is this it?"},
                ],
            },
        },
        "corrections": [],
        "marks": [
            {
                "id": "m_x",
                "segment_id": "P01:1",
                "category": "pain_point",
                "label": "confused",
                "created": "2026-04-01T12:00:00+00:00",
            }
        ],
    }


# ---- Screenspace events builder -----------------------------------------


_ALL_DETECTORS = {
    "color",
    "change",
    "similarity",
    "text",
    "numbers",
    "template",
    "flow",
    "scene",
    "inactivity",
}


@pytest.mark.parametrize(
    "kwargs,expected_count,expected_detectors,expected_participants,expected_excluded_count",
    [
        ({}, 9, _ALL_DETECTORS, None, 0),
        ({"include_excluded": True}, 10, None, None, 1),
        ({"include_excluded": True, "participants": ["P02"]}, 1, None, {"P02"}, None),
        ({"detectors": ["change", "flow"]}, 2, {"change", "flow"}, None, None),
    ],
    ids=["default", "include_excluded", "participant_filter", "detector_filter"],
)
def test_screenspace_events_filters(
    screenspace_manifest,
    kwargs,
    expected_count,
    expected_detectors,
    expected_participants,
    expected_excluded_count,
):
    rows = data_export.build_screenspace_events(screenspace_manifest, **kwargs)
    assert len(rows) == expected_count
    if expected_detectors is not None:
        assert {r["detector"] for r in rows} == expected_detectors
    if expected_participants is not None:
        assert {r["participant"] for r in rows} == expected_participants
    if expected_excluded_count is not None:
        assert sum(1 for r in rows if r["excluded"]) == expected_excluded_count


def test_screenspace_events_metadata_hoisted(screenspace_manifest):
    rows = data_export.build_screenspace_events(screenspace_manifest)
    by_detector = {r["detector"]: r for r in rows}

    assert by_detector["change"]["magnitude"] == 0.42
    assert by_detector["similarity"]["score"] == 0.91
    assert by_detector["text"]["text_found"] == "Game Over"
    assert by_detector["numbers"]["value"] == 250
    assert by_detector["template"]["match_count"] == 3
    assert by_detector["template"]["best_score"] == 0.78
    assert by_detector["flow"]["magnitude"] == 0.55
    assert by_detector["flow"]["angle"] == 90.0
    assert by_detector["scene"]["scene_name"] == "menu"
    assert by_detector["inactivity"]["duration"] == 15.5
    assert by_detector["inactivity"]["avg_distance"] == 0.02


def test_screenspace_events_metadata_does_not_overwrite_canonical(
    screenspace_manifest,
):
    # Sneak a metadata key that collides with a canonical column;
    # the canonical column must win.
    screenspace_manifest["events"][0]["metadata"] = {"participant": "P99"}
    rows = data_export.build_screenspace_events(screenspace_manifest)
    color_row = next(r for r in rows if r["detector"] == "color")
    assert color_row["participant"] == "P01"


def test_screenspace_events_duration_computed(screenspace_manifest):
    rows = data_export.build_screenspace_events(screenspace_manifest)
    inactivity = next(r for r in rows if r["detector"] == "inactivity")
    assert inactivity["duration"] == pytest.approx(15.5, abs=1e-6)


def test_screenspace_events_handles_nonfinite_floats():
    manifest = {
        "events": [
            {
                "id": "ev_nan",
                "participant": "P01",
                "detector": "change",
                "event_type": "change",
                "time_in": 1.0,
                "time_out": 1.0,
                "confidence": float("nan"),
                "metadata": {"magnitude": float("inf")},
                "excluded": False,
                "task_id": "t",
                "region": "r",
            }
        ]
    }
    rows = data_export.build_screenspace_events(manifest)
    assert rows[0]["confidence"] is None
    assert rows[0]["magnitude"] is None


# ---- Insights builder ---------------------------------------------------


def test_insights_builder_flattens_narratives(insights_manifest):
    rows = data_export.build_insights_records(insights_manifest)
    assert len(rows) == 1
    r = rows[0]
    assert r["title"] == "Players miss the save button"
    assert r["causes_narrative"] == "Save icon is below the fold."
    assert r["behaviors_narrative"] == "Users scroll up looking for it."
    assert r["impacts_narrative"] == "Lost progress on quit."
    assert r["causes_artifact_ids"] == ["a1", "a2"]
    assert r["behaviors_artifact_ids"] == []
    assert r["impacts_artifact_ids"] == ["a3"]


def test_insights_builder_handles_missing_subsections():
    manifest = {"insights": [{"id": "ins_x", "title": "T"}]}
    rows = data_export.build_insights_records(manifest)
    assert rows[0]["causes_narrative"] == ""
    assert rows[0]["behaviors_artifact_ids"] == []


# ---- Transcript segments builder ----------------------------------------


def test_build_transcript_segments(transcripts_manifest):
    """One row per segment, with participant metadata, mark joins, and duration."""
    rows = data_export.build_transcript_segments(transcripts_manifest)
    by_id = {r["segment_id"]: r for r in rows}

    assert len(rows) == 3
    assert set(by_id) == {"P01:0", "P01:1", "P02:0"}

    for r in rows:
        assert r["language"] == "en"
        assert r["model"] == "base"
    assert by_id["P01:0"]["source_file"] == "study_P01.mp4"

    assert by_id["P01:1"]["mark_categories"] == ["pain_point"]
    assert by_id["P01:1"]["mark_labels"] == ["confused"]
    assert by_id["P01:0"]["mark_categories"] == []

    assert by_id["P01:0"]["duration"] == pytest.approx(5.0, abs=1e-6)
    assert by_id["P02:0"]["duration"] == pytest.approx(4.5, abs=1e-6)


# ---- CSV serialization --------------------------------------------------


def test_to_csv_uses_preferred_order_then_alphabetical(screenspace_manifest):
    rows = data_export.build_screenspace_events(screenspace_manifest)
    csv_text = data_export.to_csv(
        rows, preferred_column_order=data_export.SCREENSPACE_EVENT_COLUMNS
    )
    reader = csv.DictReader(io.StringIO(csv_text))
    headers = reader.fieldnames
    assert headers is not None
    canonical = list(data_export.SCREENSPACE_EVENT_COLUMNS)
    # Canonical columns should appear first, in order
    assert headers[: len(canonical)] == canonical
    # Tail should be alphabetical
    tail = headers[len(canonical) :]
    assert tail == sorted(tail)


def test_to_csv_flattens_lists_with_semicolon():
    rows = [{"a": [1, 2, 3], "b": "x"}]
    csv_text = data_export.to_csv(rows)
    reader = csv.DictReader(io.StringIO(csv_text))
    row = next(reader)
    assert row["a"] == "1;2;3"


def test_to_csv_empty_records_returns_empty_string():
    assert data_export.to_csv([]) == ""


def test_to_json_envelope_shape(screenspace_manifest):
    rows = data_export.build_screenspace_events(screenspace_manifest)
    parsed = json.loads(data_export.to_json(rows))
    assert set(parsed.keys()) == {"exported_at", "version", "records"}
    assert parsed["version"] == utils.get_version()
    assert len(parsed["records"]) == len(rows)


# ---- Bundle writer ------------------------------------------------------


def _write_manifest(tmp_path: Path, filename: str, content: dict) -> None:
    (tmp_path / filename).write_text(json.dumps(content), encoding="utf-8")


_EMPTY_SS_MANIFEST = {"regions": {}, "tasks": [], "events": [], "stashes": []}


@pytest.mark.parametrize(
    "manifest_specs,expected_count,expected_name_substrs",
    [
        (
            ["screenspace", "insights", "transcripts"],
            6,
            {"screenspace_events", "insights", "transcripts"},
        ),
        (["screenspace"], 2, {"screenspace_events"}),
        ([], 0, set()),
        (["screenspace_empty"], 0, set()),
    ],
    ids=["all_three", "skips_missing", "no_manifests", "skips_empty"],
)
def test_bundle_writer(
    tmp_path,
    screenspace_manifest,
    insights_manifest,
    transcripts_manifest,
    manifest_specs,
    expected_count,
    expected_name_substrs,
):
    fixture_map = {
        "screenspace": (config.SCREENSPACE_MANIFEST_FILENAME, screenspace_manifest),
        "screenspace_empty": (config.SCREENSPACE_MANIFEST_FILENAME, _EMPTY_SS_MANIFEST),
        "insights": (config.INSIGHTS_MANIFEST_FILENAME, insights_manifest),
        "transcripts": (config.TRANSCRIPTS_MANIFEST_FILENAME, transcripts_manifest),
    }
    for spec in manifest_specs:
        filename, content = fixture_map[spec]
        _write_manifest(tmp_path, filename, content)

    written = data_export.write_export_bundle(tmp_path)
    assert len(written) == expected_count
    names = {p.name for p in written}
    for substr in expected_name_substrs:
        assert any(substr in n for n in names), f"missing {substr!r} in {names}"


def test_bundle_csv_has_data_rows(tmp_path, screenspace_manifest):
    _write_manifest(
        tmp_path, config.SCREENSPACE_MANIFEST_FILENAME, screenspace_manifest
    )
    data_export.write_export_bundle(tmp_path)
    csv_path = tmp_path / "clipgen_export_screenspace_events.csv"
    text = csv_path.read_text(encoding="utf-8")
    reader = csv.DictReader(io.StringIO(text))
    rows = list(reader)
    # 9 visible events (excluded one filtered out)
    assert len(rows) == 9
    fieldnames = reader.fieldnames
    assert fieldnames is not None
    assert "magnitude" in fieldnames
