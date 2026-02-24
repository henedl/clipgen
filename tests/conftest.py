import sys
from pathlib import Path
from types import SimpleNamespace

import gspread
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


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
