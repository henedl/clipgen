from argparse import Namespace
from unittest.mock import Mock

import pytest

import clipgen
import spreadsheet


def _args(**overrides):
    base = dict(
        batch=False,
        lines=None,
        range=None,
        cell=None,
        participant=None,
        filter=False,
        mixed=None,
        reel=None,
        timeline=None,
        screen=False,
        gif=False,
        yes=False,
        verbose=False,
        spreadsheet=None,
        viewer=False,
    )
    base.update(overrides)
    return Namespace(**base)


def test_run_cli_mode_rejects_timeline_in_mixed_mode(monkeypatch):
    args = _args(mixed="timeline, 11")

    monkeypatch.setattr(clipgen.spreadsheet, "parse_reel_input", lambda value: {"timeline": True})

    with pytest.raises(SystemExit) as exc:
        clipgen.run_cli_mode(None, args, clipgen.CliModeArgs(None, None, None, None))

    assert exc.value.code == 1


def test_generate_cli_clips_prefers_mixed_over_batch(monkeypatch):
    worksheet = object()
    args = _args(mixed="11, P01.5")
    parsed = clipgen.CliModeArgs(None, None, None, None)

    def fake_generate_list(sheet, mode, **kwargs):
        return [mode, kwargs]

    monkeypatch.setattr(clipgen.spreadsheet, "generate_list", fake_generate_list)

    clips = clipgen._generate_cli_clips(worksheet, args, parsed)
    # Expect reel mode with the mixed selectors as reel_input.
    assert clips[0] == "reel"
    assert clips[1]["reel_input"] == "11, P01.5"


def test_run_cli_mode_timeline_reel_and_viewer(monkeypatch):
    worksheet = Namespace(title="Sheet", spreadsheet=Namespace(url="http://example"))
    args = _args(timeline="P01", viewer=True)
    parsed = clipgen.CliModeArgs(None, None, None, None)

    clips = [{"study": "study", "participant": "P01"}]
    monkeypatch.setattr(clipgen, "_generate_cli_clips", lambda _ws, _a, _p: clips)

    process_reel = Mock(return_value=(1, [{"study": "study", "participant": "P01"}]))
    monkeypatch.setattr(clipgen, "process_reel", process_reel)

    viewer_path = "clips_viewer.html"
    monkeypatch.setattr(clipgen, "generate_timeline_viewer", lambda _data: viewer_path)
    monkeypatch.setattr(clipgen.utils, "info_print", lambda _msg: None)

    clipgen.run_cli_mode(worksheet, args, parsed)

    process_reel.assert_called_once()
    assert process_reel.call_args.kwargs["collect_artifacts"] is True

