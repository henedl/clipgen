from argparse import Namespace
from unittest.mock import Mock

import pytest

import cli
import clipgen
import viewer as viewer_mod


def _args(**overrides):
    base = dict(
        batch=False,
        lines=None,
        range=None,
        category=None,
        cell=None,
        participant=None,
        keyword=False,
        severity=None,
        mixed=None,
        reel=None,
        chronologic=None,
        highlights=None,
        screen=False,
        gif=False,
        yes=False,
        verbose=False,
        spreadsheet=None,
        viewer=False,
        screenspace=False,
    )
    base.update(overrides)
    return Namespace(**base)


def test_run_cli_mode_rejects_chronologic_in_mixed_mode(monkeypatch):
    args = _args(mixed="chronologic, 11")

    monkeypatch.setattr(
        cli.spreadsheet, "parse_reel_input", lambda value: {"chronologic": True}
    )

    with pytest.raises(SystemExit) as exc:
        cli.run_cli_mode(None, args, cli.CliModeArgs(None, None, None, None))

    assert exc.value.code == 1


def test_generate_cli_clips_prefers_mixed_over_batch(monkeypatch):
    worksheet = object()
    args = _args(mixed="11, P01.5")
    parsed = cli.CliModeArgs(None, None, None, None)

    generate_list = Mock(return_value=["dummy"])
    monkeypatch.setattr(cli.spreadsheet, "generate_list", generate_list)

    cli._generate_cli_clips(worksheet, args, parsed)

    generate_list.assert_called_once()
    call_args = generate_list.call_args
    assert call_args.args[0] is worksheet
    assert call_args.args[1] == "reel"
    assert call_args.kwargs.get("reel_input") == "11, P01.5"


def test_run_cli_mode_chronologic_reel_and_viewer(monkeypatch):
    worksheet = Namespace(title="Sheet", spreadsheet=Namespace(url="http://example"))
    args = _args(chronologic="P01", viewer=True)
    parsed = cli.CliModeArgs(None, None, None, None)

    clips = [{"study": "study", "participant": "P01"}]
    monkeypatch.setattr(cli, "_generate_cli_clips", lambda _ws, _a, _p: clips)

    process_reel = Mock(return_value=(1, [{"study": "study", "participant": "P01"}]))
    monkeypatch.setattr(clipgen, "process_reel", process_reel)

    viewer_path = "clips_viewer.html"
    monkeypatch.setattr(
        viewer_mod, "generate_timeline_viewer", lambda _data: viewer_path
    )
    monkeypatch.setattr(cli.utils, "info_print", lambda _msg: None)

    cli.run_cli_mode(worksheet, args, parsed)

    process_reel.assert_called_once()
