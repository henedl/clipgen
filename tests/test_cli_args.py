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
        "no_input": False,
        "verbose": False,
        "spreadsheet": None,
        "viewer": False,
        "titlecards": None,
        "filmstrip": None,
        "screenspace": False,
        "transcripts": False,
        "pre_transcribe": None,
        "ss_task": None,
        "ss_list_regions": False,
        "ss_list_stashes": False,
        "ss_list_tasks": None,
        "ss_target_color": None,
        "ss_tolerance": None,
        "ss_threshold": None,
        "ss_reference_timestamp": None,
        "ss_text": None,
        "ss_fuzzy_threshold": None,
        "ss_operator": None,
        "ss_target_value": None,
        "ss_range_min": None,
        "ss_range_max": None,
        "ss_speedup": None,
        "ss_output_format": None,
        "ss_start": None,
        "ss_end": None,
        "ss_interval": None,
        "ss_event_label": None,
        "summarize": None,
        "citations": None,
        "export": False,
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


def test_parse_arguments_export_flag(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--export"])
    args = cli.parse_arguments()
    assert getattr(args, "export", False) is True


def test_export_dispatch_calls_run_cli_export(monkeypatch):
    """--export must dispatch to data_export.run_cli_export and exit cleanly."""
    import data_export

    args = _base_args(export=True)
    called = {}

    def fake_run():
        called["yes"] = True
        return 0

    monkeypatch.setattr(data_export, "run_cli_export", fake_run)
    with pytest.raises(SystemExit) as exc:
        cli._dispatch_standalone_mode(args, cli_mode=True, gallery_arg=None)
    assert exc.value.code == 0
    assert called.get("yes") is True


def test_export_conflicts_with_studio(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--export", "--studio"])
    args = cli.parse_arguments()
    with pytest.raises(SystemExit) as exc:
        cli._validate_mode_conflicts(args)
    assert exc.value.code == 1


@pytest.mark.parametrize(
    "flag,expected_default_page",
    [
        ("studio", "studio"),
        ("screenspace", "screenspace"),
        ("transcripts", "transcripts"),
    ],
)
def test_web_mode_without_spreadsheet_dispatches_standalone(
    monkeypatch, flag, expected_default_page
):
    """--studio / --screenspace / --transcripts without -s launches the combined
    server with worksheet=None and the right default page."""
    import server

    captured = {}

    def fake_start(
        *, worksheet=None, port=None, default_page="studio", gspread_client=None
    ):
        captured["worksheet"] = worksheet
        captured["default_page"] = default_page
        captured["gspread_client"] = gspread_client

    monkeypatch.setattr(server, "start_combined_server", fake_start)
    # Skip persisted-dir application in this isolated test
    monkeypatch.setattr(cli, "_maybe_apply_persisted_dirs", lambda _args: None)
    # Force the silent-auth helper to return None (no cached token to reuse).
    monkeypatch.setattr(cli, "_try_silent_google_auth", lambda: None)

    args = _base_args(**{flag: True})
    result = cli._dispatch_standalone_mode(args, cli_mode=False, gallery_arg=None)

    assert result is True
    assert captured["worksheet"] is None
    assert captured["default_page"] == expected_default_page
    assert captured["gspread_client"] is None


def test_studio_with_spreadsheet_does_not_short_circuit(monkeypatch):
    """If -s is provided, the standalone short-circuit must not fire."""
    import server

    monkeypatch.setattr(
        server,
        "start_combined_server",
        lambda **_: pytest.fail("standalone path should not be taken"),
    )
    monkeypatch.setattr(cli, "_maybe_apply_persisted_dirs", lambda _args: None)

    args = _base_args(studio=True, spreadsheet="mystudy")
    result = cli._dispatch_standalone_mode(args, cli_mode=False, gallery_arg=None)
    assert result is False


def test_maybe_apply_persisted_dirs_uses_last_known(monkeypatch, tmp_path):
    """Persisted dirs are applied when CLI didn't set them and the path still exists."""
    import config
    import start_settings

    last_in = tmp_path / "in"
    last_out = tmp_path / "out"
    last_in.mkdir()
    last_out.mkdir()

    monkeypatch.setattr(
        start_settings,
        "load_start_settings",
        lambda: {
            "persist_enabled": True,
            "last_input": str(last_in),
            "last_output": str(last_out),
        },
    )
    monkeypatch.setattr(config, "INPUT_DIR", "")
    monkeypatch.setattr(config, "OUTPUT_DIR", "")

    args = _base_args(studio=True)
    args.input = None
    args.output = None
    cli._maybe_apply_persisted_dirs(args)

    assert config.INPUT_DIR == str(last_in)
    assert config.OUTPUT_DIR == str(last_out)


def test_maybe_apply_persisted_dirs_skips_when_disabled(monkeypatch):
    import config
    import start_settings

    monkeypatch.setattr(
        start_settings,
        "load_start_settings",
        lambda: {
            "persist_enabled": False,
            "last_input": "/anything",
            "last_output": "/anything",
        },
    )
    monkeypatch.setattr(config, "INPUT_DIR", "")
    monkeypatch.setattr(config, "OUTPUT_DIR", "")

    args = _base_args()
    args.input = None
    args.output = None
    cli._maybe_apply_persisted_dirs(args)

    assert config.INPUT_DIR == ""
    assert config.OUTPUT_DIR == ""


@pytest.mark.parametrize(
    "argv_extra,attr,expected",
    [
        ([], "titlecards", None),
        (["--titlecards"], "titlecards", True),
        (["--no-titlecards"], "titlecards", False),
        ([], "filmstrip", None),
        (["--filmstrip"], "filmstrip", True),
        (["--no-filmstrip"], "filmstrip", False),
    ],
)
def test_three_state_bool_flags(monkeypatch, argv_extra, attr, expected):
    monkeypatch.setattr("sys.argv", ["clipgen.py", *argv_extra])
    args = cli.parse_arguments()
    assert getattr(args, attr, None) is expected


def test_parse_cli_mode_args_parses_mixed_line_separators():
    args = _base_args(lines="1, 4+5")
    parsed = cli.parse_cli_mode_args(args)
    assert parsed.line_numbers == [1, 4, 5]


def _mock_main_side_effects(monkeypatch, tmp_path):
    """Silence the I/O-heavy parts of cli.main() so tests can drive dispatch."""
    import os as _os
    import server
    import utils
    import video

    monkeypatch.setattr(cli, "setup_encoding", lambda: None)
    monkeypatch.setattr(cli, "get_runtime_working_dir", lambda: str(tmp_path))
    monkeypatch.setattr(_os, "chdir", lambda _p: None)
    monkeypatch.setattr(utils, "validate_runtime_directories", lambda: None)
    monkeypatch.setattr(video, "check_ffmpeg_tools_available", lambda: True)
    monkeypatch.setattr(video, "check_webp_support", lambda: True)
    monkeypatch.setattr(cli, "_maybe_apply_persisted_dirs", lambda _args: None)

    captured: dict[str, object] = {}

    def fake_start(*, worksheet=None, port=None, default_page="studio", **_kw):
        captured["worksheet"] = worksheet
        captured["default_page"] = default_page

    monkeypatch.setattr(server, "start_combined_server", fake_start)
    return captured


def test_frozen_no_args_launches_studio(monkeypatch, tmp_path):
    """Double-clicking the .app/.exe (frozen, no argv) should land in Studio."""
    captured = _mock_main_side_effects(monkeypatch, tmp_path)
    monkeypatch.setattr("sys.argv", ["clipgen"])
    monkeypatch.setattr("sys.frozen", True, raising=False)

    with pytest.raises(SystemExit) as exc:
        cli.main()

    assert exc.value.code == 0
    assert captured.get("default_page") == "studio"
    assert captured.get("worksheet") is None


def test_frozen_with_explicit_flag_is_respected(monkeypatch, tmp_path):
    """Frozen binary invoked with an explicit flag must not be overridden."""
    captured = _mock_main_side_effects(monkeypatch, tmp_path)
    monkeypatch.setattr("sys.argv", ["clipgen", "--screenspace"])
    monkeypatch.setattr("sys.frozen", True, raising=False)

    with pytest.raises(SystemExit) as exc:
        cli.main()

    assert exc.value.code == 0
    assert captured.get("default_page") == "screenspace"


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


@pytest.mark.parametrize(
    "args_overrides,parsed_kwargs,mode,gen_kwargs,process_attr,process_kwargs",
    [
        (
            {"batch": True, "no_input": True},
            {},
            "batch",
            {"skip_prompts": True},
            "process_clips",
            {"output_format": "clip", "include_severity": False},
        ),
        (
            {"lines": "5", "screen": True, "no_input": True},
            {"line_numbers": [5]},
            "line",
            {"line_numbers": [5], "skip_prompts": True},
            "process_clips",
            {"output_format": "screen", "include_severity": False},
        ),
        (
            {"category": "Observations,Onboarding", "no_input": True},
            {},
            "category",
            {"skip_prompts": True, "categories": ["Observations", "Onboarding"]},
            "process_clips",
            {"output_format": "clip", "include_severity": False},
        ),
        (
            {"reel": "11, P01.5", "no_input": True},
            {},
            "reel",
            {"reel_input": "11, P01.5", "skip_prompts": True},
            "process_reel",
            None,
        ),
    ],
    ids=["batch", "line_screen", "category", "reel"],
)
def test_run_cli_mode_dispatch(
    monkeypatch,
    make_clip,
    args_overrides,
    parsed_kwargs,
    mode,
    gen_kwargs,
    process_attr,
    process_kwargs,
):
    args = _base_args(**args_overrides)
    clips = [make_clip()]
    generate_list = Mock(return_value=clips)
    process_fn = Mock(return_value=(1, []))
    completion = Mock()

    monkeypatch.setattr(cli.spreadsheet, "generate_list", generate_list)
    monkeypatch.setattr(clipgen, process_attr, process_fn)
    monkeypatch.setattr(clipgen, "_print_completion_message", completion)

    parsed_defaults = dict(
        line_numbers=None, range_start=None, range_end=None, cell_specs=None
    )
    parsed_defaults.update(parsed_kwargs)

    cli.run_cli_mode(None, args, cli.CliModeArgs(**parsed_defaults))

    generate_list.assert_called_once_with(None, mode, **gen_kwargs)
    if process_kwargs is None:
        assert process_fn.call_count == 1
        assert process_fn.call_args[0][0] is clips
    else:
        process_fn.assert_called_once_with(clips, **process_kwargs)
    completion.assert_called_once()


# ---- Excel auth-skipping ----


@pytest.mark.parametrize(
    "value,expected",
    [
        ("excel", True),
        ("  excel  ", True),
        ("notes.xlsx", True),
        ("path/to/file.XLSX", True),
        (None, False),
        ("my-spreadsheet", False),
    ],
)
def test_is_excel_spreadsheet_arg(value, expected):
    assert cli._is_excel_spreadsheet_arg(value) is expected


# ---- --pre-transcribe flag parsing ----


@pytest.mark.parametrize(
    "argv_extra,expected",
    [
        ([], None),
        (["--pre-transcribe"], []),
        (["--pre-transcribe", "P01", "P03"], ["P01", "P03"]),
    ],
)
def test_pre_transcribe_flag(monkeypatch, argv_extra, expected):
    monkeypatch.setattr("sys.argv", ["clipgen.py", *argv_extra])
    args = cli.parse_arguments()
    assert args.pre_transcribe == expected


@pytest.mark.parametrize(
    "flag,attr,value",
    [
        ("--whisper-model", "whisper_model", "medium"),
        ("--ollama-model", "ollama_model", "gemma3:4b"),
    ],
)
def test_model_flag_parses(monkeypatch, flag, attr, value):
    monkeypatch.setattr("sys.argv", ["clipgen.py", flag, value])
    args = cli.parse_arguments()
    assert getattr(args, attr) == value


def test_whisper_model_flag_rejects_invalid(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--whisper-model", "huge"])
    with pytest.raises(SystemExit):
        cli.parse_arguments()


@pytest.mark.parametrize(
    "flag,value,config_attr",
    [
        ("--whisper-model", "small", "TRANSCRIBE_MODEL"),
        ("--ollama-model", "gemma3:4b", "OLLAMA_SUMMARY_MODEL"),
    ],
)
def test_model_flag_applies_to_config(monkeypatch, flag, value, config_attr):
    import config

    original = getattr(config, config_attr)
    try:
        monkeypatch.setattr("sys.argv", ["clipgen.py", flag, value])
        args = cli.parse_arguments()
        cli._apply_config_overrides(args, cli_mode=True)
        assert getattr(config, config_attr) == value
    finally:
        setattr(config, config_attr, original)
