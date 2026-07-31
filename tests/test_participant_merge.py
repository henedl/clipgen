"""Tests for the sheet ∪ disk participant union.

With a spreadsheet loaded, Screenspace and Transcripts used to list *only* the
sheet's participant columns, so a source video on disk that the sheet didn't
mention was invisible while Composer/Workflows listed it. ``files.resolve_participant_videos``
merges both sources; ``_refresh_participants`` in the two blueprints rebuilds
that merge when the input directory changes.
"""

from pathlib import Path
from types import SimpleNamespace

import pytest
from flask import Flask

import config
import files
import utils
from conftest import make_sheet_context


def _names(entry):
    return [Path(p).name for p in entry["video_paths"]]


def _ctx(participants, study="clipgen-test", filename_row=None):
    """Sheet context whose header row carries *participants* starting at col 2.

    ``filename_row`` (a list aligned to the header) adds an optional ``Filename``
    override row directly under the header.
    """
    header = ["ID"] + list(participants)
    sheet_data = [["Study", study], header]
    filename_row_idx = None
    if filename_row is not None:
        sheet_data.append(["Filename"] + list(filename_row))
        filename_row_idx = 2
    sheet_data.append(["1", *([""] * len(participants))])
    return make_sheet_context(
        sheet_data=sheet_data,
        id_cell=SimpleNamespace(row=2, col=1),
        observation_cell=SimpleNamespace(row=2, col=3),
        num_participants=len(participants),
        study_name=study,
        filename_row_idx=filename_row_idx,
    )


@pytest.fixture
def input_dir(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    return tmp_path


# ---- files.resolve_participant_videos ----


def test_appends_disk_only_participants(input_dir):
    (input_dir / "clipgen-test_P01.mp4").write_text("v")
    (input_dir / "study_P13.mp4").write_text("v")

    found = files.resolve_participant_videos(_ctx(["P01", "P02"]))

    assert [p["id"] for p in found] == ["P01", "P02", "P13"]
    assert [p["in_sheet"] for p in found] == [True, True, False]
    # A sheet column with no file stays visible rather than vanishing silently.
    assert found[1]["has_video"] is False
    # The core regression: a disk-only participant keeps its own discovered path.
    # Re-resolving P13 against the sheet's study name would invent a
    # "clipgen-test_P13.mp4" that is not on disk.
    assert _names(found[2]) == ["study_P13.mp4"]
    assert found[2]["has_video"] is True


def test_preserves_sheet_order_and_sorts_disk_only(input_dir):
    for name in (
        "clipgen-test_P05.mp4",
        "clipgen-test_P01.mp4",
        "x_P20.mp4",
        "x_P03.mp4",
    ):
        (input_dir / name).write_text("v")

    found = files.resolve_participant_videos(_ctx(["P05", "P01"]))

    assert [p["id"] for p in found] == ["P05", "P01", "P03", "P20"]


def test_sheet_wins_on_dedup(input_dir):
    # Same participant reachable two ways; the sheet's study-scoped file wins and
    # the id is not listed twice.
    (input_dir / "clipgen-test_P01.mp4").write_text("v")
    (input_dir / "other_P01.mp4").write_text("v")

    found = files.resolve_participant_videos(_ctx(["P01"]))

    assert [p["id"] for p in found] == ["P01"]
    assert _names(found[0]) == ["clipgen-test_P01.mp4"]
    assert found[0]["in_sheet"] is True


def test_honours_filename_override(input_dir):
    (input_dir / "morning.mp4").write_text("v")
    (input_dir / "afternoon.mp4").write_text("v")

    found = files.resolve_participant_videos(
        _ctx(["P01"], filename_row=["morning.mp4 + afternoon.mp4"])
    )

    assert _names(found[0]) == ["morning.mp4", "afternoon.mp4"]
    assert found[0]["has_video"] is True


def test_strict_sheet_resolution_does_not_borrow_another_study(input_dir):
    # Deliberate: a sheet participant whose expected file is missing stays
    # has_video False even though a same-id file from another study sits right
    # there. Borrowing it would let a typo'd Filename override silently resolve
    # to the wrong recording.
    (input_dir / "otherstudy_P05.mp4").write_text("v")

    found = files.resolve_participant_videos(_ctx(["P05"]))

    assert [p["id"] for p in found] == ["P05"]
    assert found[0]["has_video"] is False
    assert _names(found[0]) == ["clipgen-test_P05.mp4"]


def test_no_sheet_is_a_plain_scan(input_dir):
    (input_dir / "study_P01.mp4").write_text("v")
    (input_dir / "study_P02.mp4").write_text("v")

    found = files.resolve_participant_videos(None)

    assert [p["id"] for p in found] == ["P01", "P02"]
    assert all(p["in_sheet"] is False for p in found)


def test_does_not_mutate_the_discovery_cache(input_dir):
    (input_dir / "study_P01.mp4").write_text("v")

    files.resolve_participant_videos(None)

    # discover_participant_videos memoizes its result dicts and shares them with
    # the Workflows blueprint; the merge must build fresh ones, never stamp on.
    assert all("in_sheet" not in p for p in utils.discover_participant_videos())


# ---- Blueprint _refresh_participants ----


@pytest.mark.parametrize("module_name", ["screenspace_server", "transcripts_server"])
def test_refresh_is_a_noop_before_init(monkeypatch, module_name):
    module = __import__(module_name)
    pinned = [{"id": "P99", "video_paths": ["/tmp/x.mp4"], "has_video": False}]
    monkeypatch.setattr(module, "_participant_source", None)
    monkeypatch.setattr(module, "_participants", pinned)

    module._refresh_participants()

    assert module._participants is pinned


@pytest.mark.parametrize(
    ("module_name", "bp_name"),
    [
        ("screenspace_server", "screenspace_bp"),
        ("transcripts_server", "transcripts_bp"),
    ],
)
def test_refresh_picks_up_a_new_file(monkeypatch, input_dir, module_name, bp_name):
    import viewer

    module = __import__(module_name)
    (input_dir / "study_P01.mp4").write_text("v")

    monkeypatch.setattr(viewer, "load_manifest_artifacts", list)
    monkeypatch.setattr(module, "_participants", [])
    monkeypatch.setattr(
        module,
        "_participant_source",
        {"sheet_context": None, "dir": "", "mtime": None},
    )
    module._refresh_participants()
    assert [p["id"] for p in module._participants] == ["P01"]

    # A video dropped into the input dir mid-session must show up without a
    # server restart — the dir's mtime_ns advances and the merge is rebuilt.
    (input_dir / "study_P99.mp4").write_text("v")

    app = Flask(__name__)
    app.register_blueprint(getattr(module, bp_name), url_prefix="/x")
    with app.test_client() as client:
        data = client.get("/x/api/participants").get_json()

    assert [p["id"] for p in data["participants"]] == ["P01", "P99"]
    assert all(p["in_sheet"] is False for p in data["participants"])
    assert data["has_sheet"] is False
