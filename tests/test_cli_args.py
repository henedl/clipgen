from argparse import Namespace
from unittest.mock import Mock

import pytest

import cli
import clipgen


def _base_args(**overrides):
    args = {
        "batch": False,
        "lines": None,
        "range": None,
        "category": None,
        "cell": None,
        "participant": None,
        "keyword": False,
        "severity": None,
        "mixed": None,
        "reel": None,
        "chronologic": None,
        "highlights": None,
        "screen": False,
        "gif": False,
        "yes": False,
        "verbose": False,
        "spreadsheet": None,
        "viewer": False,
        "titlecards": None,
        "filmstrip": None,
        "screenspace": False,
        "transcripts": False,
        "pre_transcribe": None,
    }
    args.update(overrides)
    return Namespace(**args)


def test_parse_arguments_rejects_conflicting_mode_flags(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "-b", "-l", "5"])
    with pytest.raises(SystemExit) as exc:
        cli.parse_arguments()
    assert exc.value.code == 2


def test_parse_arguments_rejects_conflicting_output_flags(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--screen", "--gif"])
    with pytest.raises(SystemExit) as exc:
        cli.parse_arguments()
    assert exc.value.code == 2


def test_parse_arguments_titlecards_flags(monkeypatch):
    # Default: no flag → None (use config default)
    monkeypatch.setattr("sys.argv", ["clipgen.py"])
    args = cli.parse_arguments()
    assert getattr(args, "titlecards", None) is None

    # --titlecards → True
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--titlecards"])
    args = cli.parse_arguments()
    assert args.titlecards is True

    # --no-titlecards → False
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--no-titlecards"])
    args = cli.parse_arguments()
    assert args.titlecards is False


def test_parse_arguments_filmstrip_flags(monkeypatch):
    # Default: no flag → None (use config default)
    monkeypatch.setattr("sys.argv", ["clipgen.py"])
    args = cli.parse_arguments()
    assert getattr(args, "filmstrip", None) is None

    # --filmstrip → True
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--filmstrip"])
    args = cli.parse_arguments()
    assert args.filmstrip is True

    # --no-filmstrip → False
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--no-filmstrip"])
    args = cli.parse_arguments()
    assert args.filmstrip is False


def test_parse_cli_mode_args_parses_mixed_line_separators():
    args = _base_args(lines="1, 4+5")
    parsed = cli.parse_cli_mode_args(args)
    assert parsed.line_numbers == [1, 4, 5]


def test_parse_cli_mode_args_rejects_reversed_range():
    args = _base_args(range="10-1")
    with pytest.raises(SystemExit) as exc:
        cli.parse_cli_mode_args(args)
    assert exc.value.code == 1


def test_parse_cli_mode_args_accepts_valid_range():
    args = _base_args(range="5-10")
    parsed = cli.parse_cli_mode_args(args)
    assert parsed.range_start == 5
    assert parsed.range_end == 10


def test_parse_cli_mode_args_delegates_cell_parsing(monkeypatch):
    args = _base_args(cell="P01.11, P02.12")
    expected = [("P01", 11), ("P02", 12)]
    monkeypatch.setattr(
        cli.spreadsheet,
        "parse_cell_specifications",
        lambda value: expected if value == "P01.11, P02.12" else [],
    )
    parsed = cli.parse_cli_mode_args(args)
    assert parsed.cell_specs == expected


def test_run_cli_mode_rejects_reel_with_gif():
    args = _base_args(reel="11, P01", gif=True)
    with pytest.raises(SystemExit) as exc:
        cli.run_cli_mode(None, args, cli.CliModeArgs(None, None, None, None))
    assert exc.value.code == 1


def test_run_cli_mode_batch_happy_path_dispatch(monkeypatch, make_clip):
    args = _base_args(batch=True, yes=True)
    clips = [make_clip()]
    generate_list = Mock(return_value=clips)
    process_clips = Mock(return_value=(1, []))
    completion = Mock()

    monkeypatch.setattr(cli.spreadsheet, "generate_list", generate_list)
    monkeypatch.setattr(clipgen, "process_clips", process_clips)
    monkeypatch.setattr(clipgen, "_print_completion_message", completion)

    cli.run_cli_mode(None, args, cli.CliModeArgs(None, None, None, None))

    generate_list.assert_called_once_with(None, "batch", skip_prompts=True)
    process_clips.assert_called_once_with(
        clips, output_format="clip", include_severity=False
    )
    completion.assert_called_once()


def test_run_cli_mode_line_with_screen_output(monkeypatch, make_clip):
    args = _base_args(lines="5", screen=True, yes=True)
    clips = [make_clip()]
    generate_list = Mock(return_value=clips)
    process_clips = Mock(return_value=(1, []))
    completion = Mock()

    monkeypatch.setattr(cli.spreadsheet, "generate_list", generate_list)
    monkeypatch.setattr(clipgen, "process_clips", process_clips)
    monkeypatch.setattr(clipgen, "_print_completion_message", completion)

    parsed = cli.CliModeArgs(
        line_numbers=[5], range_start=None, range_end=None, cell_specs=None
    )
    cli.run_cli_mode(None, args, parsed)

    generate_list.assert_called_once_with(
        None,
        "line",
        line_numbers=[5],
        skip_prompts=True,
    )
    process_clips.assert_called_once_with(
        clips, output_format="screen", include_severity=False
    )
    completion.assert_called_once()


def test_run_cli_mode_category_cli_dispatch(monkeypatch, make_clip):
    args = _base_args(category="Observations,Onboarding", yes=True)
    clips = [make_clip()]
    generate_list = Mock(return_value=clips)
    process_clips = Mock(return_value=(1, []))
    completion = Mock()

    monkeypatch.setattr(cli.spreadsheet, "generate_list", generate_list)
    monkeypatch.setattr(clipgen, "process_clips", process_clips)
    monkeypatch.setattr(clipgen, "_print_completion_message", completion)

    parsed = cli.CliModeArgs(None, None, None, None)
    cli.run_cli_mode(None, args, parsed)

    generate_list.assert_called_once_with(
        None,
        "category",
        skip_prompts=True,
        categories=["Observations", "Onboarding"],
    )
    process_clips.assert_called_once_with(
        clips, output_format="clip", include_severity=False
    )
    completion.assert_called_once()


def test_run_cli_mode_reel_cli_dispatch(monkeypatch, make_clip):
    args = _base_args(reel="11, P01.5", yes=True)
    clips = [make_clip()]
    generate_list = Mock(return_value=clips)
    process_reel = Mock(return_value=(1, []))
    completion = Mock()

    monkeypatch.setattr(cli.spreadsheet, "generate_list", generate_list)
    monkeypatch.setattr(clipgen, "process_reel", process_reel)
    monkeypatch.setattr(clipgen, "_print_completion_message", completion)

    parsed = cli.CliModeArgs(None, None, None, None)
    cli.run_cli_mode(None, args, parsed)

    generate_list.assert_called_once_with(
        None, "reel", reel_input="11, P01.5", skip_prompts=True
    )
    assert process_reel.call_count == 1
    # First positional argument should be the clips list returned from generate_list.
    assert process_reel.call_args[0][0] is clips
    completion.assert_called_once()


# ---- Excel auth-skipping ----


@pytest.mark.parametrize(
    "value,expected",
    [
        ("excel", True),
        ("EXCEL", True),
        ("  excel  ", True),
        ("notes.xlsx", True),
        ("path/to/file.XLSX", True),
        (None, False),
        ("", False),
        ("my-spreadsheet", False),
        ("https://docs.google.com/spreadsheets/d/abc", False),
        ("42", False),
    ],
)
def test_is_excel_spreadsheet_arg(value, expected):
    assert cli._is_excel_spreadsheet_arg(value) is expected


# ---- --pre-transcribe flag parsing ----


def test_pre_transcribe_absent(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py"])
    args = cli.parse_arguments()
    assert args.pre_transcribe is None


def test_pre_transcribe_no_ids(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--pre-transcribe"])
    args = cli.parse_arguments()
    assert args.pre_transcribe == []


def test_pre_transcribe_with_ids(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--pre-transcribe", "P01", "P03"])
    args = cli.parse_arguments()
    assert args.pre_transcribe == ["P01", "P03"]
