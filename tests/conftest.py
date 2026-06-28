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
