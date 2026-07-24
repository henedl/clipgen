import sys
from pathlib import Path
from types import SimpleNamespace

import gspread
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


@pytest.fixture(autouse=True)
def _sandbox_cwd(tmp_path, monkeypatch):
    """Run every test from a throwaway working directory.

    Many code paths fall back to ``Path.cwd()`` when ``config.OUTPUT_DIR`` is
    unset — notably ``files.get_unique_filename`` reserving a 0-byte placeholder.
    A test that exercises those without sandboxing the output dir would otherwise
    drop artifacts into the repo root (``.mp4`` is gitignored, so they accumulate
    invisibly). Chdir-ing into ``tmp_path`` keeps any such fallback inside
    pytest's auto-cleaned tmp tree. Tests that set ``config.OUTPUT_DIR``
    explicitly are unaffected; this is purely a safety net for the rest.
    """
    monkeypatch.chdir(tmp_path)


@pytest.fixture(autouse=True)
def _reset_no_input_mode():
    """Restore ``utils.NO_INPUT_MODE`` after every test.

    ``server.create_combined_app`` sets the module global ``NO_INPUT_MODE = True``
    and never resets it, so under ``pytest-randomly`` a server test ordered before
    an interactive-prompt test (e.g. the Excel fallback) leaks the flag and makes
    the prompt refuse to read input. Snapshot + restore keeps it isolated per test.
    """
    import utils

    original = utils.NO_INPUT_MODE
    try:
        yield
    finally:
        utils.NO_INPUT_MODE = original


@pytest.fixture(autouse=True)
def _reset_overview_observation_getter():
    """Restore ``overview._observation_rows_getter`` after every test.

    ``server.build_combined_app`` injects the live-sheet getter through
    ``overview.configure()`` and never unsets it, so any test that builds a
    combined app leaves it wired for the rest of the worker process. A later
    ``overview`` test then reads whatever sheet another test left in
    ``server._sheet_context`` and reports phantom participants — four
    ``test_overview.py`` assertions failed exactly that way under CI's
    ``-n auto --randomly-seed`` ordering, while a serial local run happened to
    order them safely. Snapshot + restore keeps it isolated per test.
    """
    import overview

    original = overview._observation_rows_getter
    try:
        yield
    finally:
        overview._observation_rows_getter = original


@pytest.fixture
def make_clip():
    def _make_clip(
        *,
        row=3,
        col=2,
        value="00:10-00:20",
        study="study",
        participant="P01",
        desc="desc",
        category="cat",
    ):
        return {
            "cell": gspread.cell.Cell(row, col, value),
            "study": study,
            "participant": participant,
            "desc": desc,
            "category": category,
        }

    return _make_clip


@pytest.fixture
def fake_sheet_meta():
    # Header row is spreadsheet row 2 in this program.
    return SimpleNamespace(
        id_cell=SimpleNamespace(row=2, col=1),
        observation_cell=SimpleNamespace(row=2, col=4),
        category_cell=SimpleNamespace(row=2, col=5),
    )


def make_sheet_context(
    sheet_data,
    id_cell,
    observation_cell,
    category_cell=None,
    num_participants=2,
    study_name="study",
    baseline_row_idx=None,
    filename_row_idx=None,
):
    """Build a SheetContext with test-friendly defaults.

    Plain helper, not a fixture — import with
    ``from conftest import make_sheet_context`` (call sites pass everything by
    keyword). ``category_cell=None`` defaults to a header on the ID row, col 5.
    """
    from spreadsheet import SheetContext

    if category_cell is None:
        category_cell = SimpleNamespace(row=id_cell.row, col=5)
    return SheetContext(
        sheet_data=sheet_data,
        id_cell=id_cell,
        observation_cell=observation_cell,
        category_cell=category_cell,
        num_participants=num_participants,
        study_name=study_name,
        baseline_row_idx=baseline_row_idx,
        filename_row_idx=filename_row_idx,
    )
