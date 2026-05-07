"""Tests for --no-input flag parsing and the targeted blocking-prompt guards."""

from unittest.mock import MagicMock

import pytest

import cli
import clipgen
import excel_io
import utils


@pytest.fixture(autouse=True)
def _reset_no_input_mode():
    """Ensure NO_INPUT_MODE doesn't leak between tests."""
    utils.NO_INPUT_MODE = False
    yield
    utils.NO_INPUT_MODE = False


# ---- Argparse rename ----


def test_no_input_long_flag_parses(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--no-input"])
    args = cli.parse_arguments()
    assert args.no_input is True


def test_no_input_default_false(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py"])
    args = cli.parse_arguments()
    assert args.no_input is False


def test_short_y_flag_is_rejected(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "-y"])
    with pytest.raises(SystemExit) as exc:
        cli.parse_arguments()
    assert exc.value.code == 2


def test_long_yes_flag_is_rejected(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--yes"])
    with pytest.raises(SystemExit) as exc:
        cli.parse_arguments()
    assert exc.value.code == 2


# ---- utils.suggest_close_match guard ----


def test_suggest_close_match_returns_none_under_no_input(monkeypatch):
    # Sanity: without NO_INPUT_MODE, the prompt is reached.
    monkeypatch.setattr(utils, "read_user_input", lambda _prompt: "n")
    assert utils.suggest_close_match("foobar", ["foobaz"]) is None

    utils.NO_INPUT_MODE = True

    def _fail(_prompt):
        raise AssertionError("read_user_input must not be called under NO_INPUT_MODE")

    monkeypatch.setattr(utils, "read_user_input", _fail)
    assert utils.suggest_close_match("foobar", ["foobaz"]) is None


# ---- clipgen.select_spreadsheet guard ----


def test_select_spreadsheet_exits_under_no_input(monkeypatch):
    utils.NO_INPUT_MODE = True

    def _fail(_prompt):
        raise AssertionError("read_user_input must not be called under NO_INPUT_MODE")

    monkeypatch.setattr(utils, "read_user_input", _fail)
    with pytest.raises(SystemExit) as exc:
        clipgen.select_spreadsheet(MagicMock(), ["doc1", "doc2"])
    assert exc.value.code == 2


# ---- clipgen.run_interactive_mode guard ----


def test_run_interactive_mode_exits_under_no_input(monkeypatch):
    utils.NO_INPUT_MODE = True

    def _fail(_prompt):
        raise AssertionError("read_user_input must not be called under NO_INPUT_MODE")

    monkeypatch.setattr(utils, "read_user_input", _fail)
    with pytest.raises(SystemExit) as exc:
        clipgen.run_interactive_mode(MagicMock())
    assert exc.value.code == 2


# ---- excel_io guards ----


def test_prompt_for_excel_fallback_returns_none_under_no_input(monkeypatch):
    utils.NO_INPUT_MODE = True

    def _fail(_prompt):
        raise AssertionError("read_user_input must not be called under NO_INPUT_MODE")

    monkeypatch.setattr(utils, "read_user_input", _fail)
    assert excel_io.prompt_for_excel_fallback() is None


def test_select_excel_file_returns_none_under_no_input_with_multiple(monkeypatch):
    utils.NO_INPUT_MODE = True
    monkeypatch.setattr(excel_io, "list_excel_in_cwd", lambda: ["./a.xlsx", "./b.xlsx"])

    def _fail(_prompt):
        raise AssertionError("read_user_input must not be called under NO_INPUT_MODE")

    monkeypatch.setattr(utils, "read_user_input", _fail)
    assert excel_io.select_excel_file() is None


def test_select_excel_file_opens_single_match_under_no_input(monkeypatch):
    """Single-file branch should still work in non-interactive mode."""
    utils.NO_INPUT_MODE = True
    monkeypatch.setattr(excel_io, "list_excel_in_cwd", lambda: ["./only.xlsx"])
    sentinel = object()
    monkeypatch.setattr(excel_io, "open_excel_workbook", lambda _path: sentinel)
    assert excel_io.select_excel_file() is sentinel
