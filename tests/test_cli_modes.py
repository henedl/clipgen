from argparse import Namespace
from unittest.mock import Mock

import pytest

import cli
import clipgen
import config
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
        no_input=False,
        verbose=False,
        spreadsheet=None,
        viewer=False,
        screenspace=False,
        transcripts=False,
        workflows=False,
    )
    base.update(overrides)
    return Namespace(**base)


def test_workflows_mode_accepted_alone():
    """--workflows alone validates and is reported active by mode detection."""
    modes = cli._validate_mode_conflicts(_args(workflows=True))
    assert modes["workflows"] is True


def test_workflows_mode_conflicts_with_other_web_modes():
    """--workflows is mutually exclusive with the other web frontends."""
    for other in ("studio", "screenspace", "transcripts"):
        with pytest.raises(SystemExit):
            cli._validate_mode_conflicts(_args(workflows=True, **{other: True}))


def test_workflows_mode_conflicts_with_selection_flags():
    """--workflows accepts only -s/-i/-o/-v, not selection/format modes."""
    with pytest.raises(SystemExit):
        cli._validate_mode_conflicts(_args(workflows=True, batch=True))


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

    reel_records = [
        {
            "id": "reel_abc",
            "file": "study_reel.mp4",
            "study": "study",
            "components": [],
        }
    ]
    process_reel = Mock(return_value=(1, reel_records))
    monkeypatch.setattr(clipgen, "process_reel", process_reel)

    viewer_calls = []
    viewer_path = "clips_viewer.html"

    def capture_viewer(data):
        viewer_calls.append(data)
        return viewer_path

    monkeypatch.setattr(viewer_mod, "generate_timeline_viewer", capture_viewer)
    monkeypatch.setattr(cli.utils, "info_print", lambda _msg: None)

    cli.run_cli_mode(worksheet, args, parsed)

    process_reel.assert_called_once()
    assert len(viewer_calls) == 1
    assert viewer_calls[0]["artifacts"] == []
    assert viewer_calls[0]["reels"] == reel_records


def test_standalone_viewer_from_reel_only_manifest(monkeypatch, tmp_path):
    import viewer

    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    reel = {
        "id": "reel_only",
        "file": "only_reel.mp4",
        "study": "study",
        "components": [],
    }
    viewer.save_manifest([], new_reels=[reel], study="study", mode="reel")

    viewer_calls = []
    monkeypatch.setattr(
        viewer_mod,
        "generate_timeline_viewer",
        lambda data: viewer_calls.append(data) or "viewer.html",
    )
    monkeypatch.setattr(cli.utils, "info_print", lambda _msg: None)

    args = _args(viewer=True)
    assert cli._dispatch_standalone_mode(args, cli_mode=False, gallery_arg=None) is True
    assert len(viewer_calls) == 1
    assert viewer_calls[0]["artifacts"] == []
    assert len(viewer_calls[0]["reels"]) == 1
    assert viewer_calls[0]["reels"][0]["id"] == "reel_only"
