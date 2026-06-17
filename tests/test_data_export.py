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


# ---- Friction builders --------------------------------------------------


def _with_friction(manifest):
    manifest["source_transcripts"]["P01"]["friction"] = {
        "computed_at": "2026-04-01T13:00:00+00:00",
        "model": "qwen3.5:9b",
        "stale": False,
        "moments": [
            {
                "segment_ids": ["P01:1"],
                "category": "confusion",
                "rationale": "Unsure what the button does",
                "score": 0.7,
            }
        ],
        "segments": [
            {
                "id": "P01:0",
                "score": 0.0,
                "categories": [],
                "markers": [],
                "counts": {},
            },
            {
                "id": "P01:1",
                "score": 0.5,
                "categories": ["confusion"],
                "markers": ["what's this"],
                "counts": {"confusion": 1},
            },
        ],
    }
    return manifest


def test_build_friction_moments(transcripts_manifest):
    rows = data_export.build_friction_moments(_with_friction(transcripts_manifest))
    assert len(rows) == 1
    row = rows[0]
    assert row["participant"] == "P01"
    assert row["segment_ids"] == ["P01:1"]
    assert row["category"] == "confusion"
    assert row["score"] == 0.7
    assert row["model"] == "qwen3.5:9b"
    assert row["computed_at"] == "2026-04-01T13:00:00+00:00"


def test_build_friction_segments_skips_zero_score(transcripts_manifest):
    rows = data_export.build_friction_segments(_with_friction(transcripts_manifest))
    # Only the non-zero-score segment is exported.
    assert len(rows) == 1
    assert rows[0]["segment_id"] == "P01:1"
    assert rows[0]["score"] == 0.5
    assert rows[0]["categories"] == ["confusion"]


def test_build_friction_moments_empty_without_friction(transcripts_manifest):
    assert data_export.build_friction_moments(transcripts_manifest) == []
    assert data_export.build_friction_segments(transcripts_manifest) == []


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


def test_to_csv_scrubs_nonfinite_floats_inside_lists():
    rows = [{"a": [1.0, float("nan"), float("inf"), 2.0]}]
    csv_text = data_export.to_csv(rows)
    reader = csv.DictReader(io.StringIO(csv_text))
    row = next(reader)
    assert row["a"] == "1.0;;;2.0"


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
            ["screenspace", "transcripts"],
            4,
            {"screenspace_events", "transcripts"},
        ),
        (["screenspace"], 2, {"screenspace_events"}),
        ([], 0, set()),
        (["screenspace_empty"], 0, set()),
    ],
    ids=["all_surfaces", "skips_missing", "no_manifests", "skips_empty"],
)
def test_bundle_writer(
    tmp_path,
    screenspace_manifest,
    transcripts_manifest,
    manifest_specs,
    expected_count,
    expected_name_substrs,
):
    fixture_map = {
        "screenspace": (config.SCREENSPACE_MANIFEST_FILENAME, screenspace_manifest),
        "screenspace_empty": (config.SCREENSPACE_MANIFEST_FILENAME, _EMPTY_SS_MANIFEST),
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


# ---- Pins export --------------------------------------------------------


def _ss_manifest_with_pins() -> dict:
    """Minimal screenspace manifest carrying calibration pins for two people."""
    return {
        "regions": {},
        "tasks": [],
        "events": [],
        "stashes": [],
        "pins": {
            "P01": [
                {
                    "id": "pin_aaaaaaaa",
                    "timestamp": 12.5,
                    "polarity": "positive",
                    "label": "health bar red",
                    "created_at": "2026-06-16T00:00:00+00:00",
                },
                {
                    "id": "pin_bbbbbbbb",
                    "timestamp": 30.0,
                    "polarity": "negative",
                    "label": "",
                    "created_at": "2026-06-16T00:01:00+00:00",
                },
            ],
            "P02": [
                {
                    "id": "pin_cccccccc",
                    "timestamp": 5.25,
                    "polarity": "positive",
                    "label": "menu open",
                    "created_at": "2026-06-16T00:02:00+00:00",
                },
            ],
        },
    }


def test_screenspace_pins_builder():
    rows = data_export.build_screenspace_pins(_ss_manifest_with_pins())
    assert len(rows) == 3
    by_id = {r["id"]: r for r in rows}
    assert by_id["pin_aaaaaaaa"]["participant"] == "P01"
    assert by_id["pin_aaaaaaaa"]["polarity"] == "positive"
    assert by_id["pin_aaaaaaaa"]["timestamp"] == 12.5
    assert by_id["pin_bbbbbbbb"]["polarity"] == "negative"
    assert by_id["pin_cccccccc"]["participant"] == "P02"


def test_screenspace_pins_builder_no_pins(screenspace_manifest):
    # The shared fixture has no "pins" key — builder must return an empty list,
    # not raise, so the bundle writer simply skips the surface.
    assert data_export.build_screenspace_pins(screenspace_manifest) == []


@pytest.mark.parametrize(
    "pins_value",
    [
        None,
        [],
        {"P01": {"id": "pin_bad"}},
        {"P01": [None, "bad", {"id": "pin_ok", "timestamp": 1.0}]},
    ],
)
def test_screenspace_pins_builder_handles_malformed_shapes(pins_value):
    rows = data_export.build_screenspace_pins({"pins": pins_value})
    if isinstance(pins_value, dict) and isinstance(pins_value.get("P01"), list):
        assert rows == [
            {
                "participant": "P01",
                "id": "pin_ok",
                "timestamp": 1.0,
                "polarity": "",
                "label": "",
                "created_at": "",
            }
        ]
    else:
        assert rows == []


def test_bundle_writer_includes_pins(tmp_path):
    _write_manifest(
        tmp_path, config.SCREENSPACE_MANIFEST_FILENAME, _ss_manifest_with_pins()
    )
    written = data_export.write_export_bundle(tmp_path)
    names = {p.name for p in written}
    assert "clipgen_export_screenspace_pins.json" in names
    assert "clipgen_export_screenspace_pins.csv" in names

    csv_path = tmp_path / "clipgen_export_screenspace_pins.csv"
    reader = csv.DictReader(io.StringIO(csv_path.read_text(encoding="utf-8")))
    rows = list(reader)
    assert len(rows) == 3
    assert reader.fieldnames is not None
    # Preferred column order leads the header.
    assert reader.fieldnames[:6] == list(data_export.SCREENSPACE_PIN_COLUMNS)


def test_bundle_writer_events_and_pins_coexist(tmp_path, screenspace_manifest):
    # Events and pins are independent surfaces off the same manifest: a manifest
    # carrying both must emit four files (events + pins, each JSON & CSV).
    manifest = dict(screenspace_manifest)
    manifest["pins"] = _ss_manifest_with_pins()["pins"]
    _write_manifest(tmp_path, config.SCREENSPACE_MANIFEST_FILENAME, manifest)
    written = data_export.write_export_bundle(tmp_path)
    names = {p.name for p in written}
    assert names == {
        "clipgen_export_screenspace_events.json",
        "clipgen_export_screenspace_events.csv",
        "clipgen_export_screenspace_pins.json",
        "clipgen_export_screenspace_pins.csv",
    }
