from argparse import Namespace
from unittest.mock import Mock

import pytest

import clipgen
import utils


def _base_args(**overrides):
    args = {
        "batch": False,
        "lines": None,
        "range": None,
        "cell": None,
        "participant": None,
        "filter": False,
        "reel": None,
        "timeline": None,
        "screen": False,
        "gif": False,
        "yes": False,
        "verbose": False,
        "spreadsheet": None,
    }
    args.update(overrides)
    return Namespace(**args)


def test_parse_arguments_rejects_conflicting_mode_flags(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "-b", "-l", "5"])
    with pytest.raises(SystemExit) as exc:
        utils.parse_arguments()
    assert exc.value.code == 2


def test_parse_arguments_rejects_conflicting_output_flags(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--screen", "--gif"])
    with pytest.raises(SystemExit) as exc:
        utils.parse_arguments()
    assert exc.value.code == 2


def test_parse_cli_mode_args_parses_mixed_line_separators():
    args = _base_args(lines="1, 4+5")
    parsed = clipgen.parse_cli_mode_args(args)
    assert parsed.line_numbers == [1, 4, 5]


def test_parse_cli_mode_args_rejects_reversed_range():
    args = _base_args(range="10-1")
    with pytest.raises(SystemExit) as exc:
        clipgen.parse_cli_mode_args(args)
    assert exc.value.code == 1


def test_run_cli_mode_rejects_reel_with_gif():
    args = _base_args(reel="11, P01", gif=True)
    with pytest.raises(SystemExit) as exc:
        clipgen.run_cli_mode(None, args, clipgen.CliModeArgs(None, None, None, None))
    assert exc.value.code == 1


def test_run_cli_mode_batch_happy_path_dispatch(monkeypatch, make_clip):
    args = _base_args(batch=True, yes=True)
    clips = [make_clip()]
    generate_list = Mock(return_value=clips)
    process_clips = Mock(return_value=1)
    completion = Mock()

    monkeypatch.setattr(clipgen.spreadsheet, "generate_list", generate_list)
    monkeypatch.setattr(clipgen, "process_clips", process_clips)
    monkeypatch.setattr(clipgen, "_print_completion_message", completion)

    clipgen.run_cli_mode(None, args, clipgen.CliModeArgs(None, None, None, None))

    generate_list.assert_called_once_with(None, "batch", skip_prompts=True)
    process_clips.assert_called_once_with(clips, output_format="clip")
    completion.assert_called_once()
