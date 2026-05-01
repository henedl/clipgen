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


@pytest.mark.parametrize(
    "args_overrides,parsed_kwargs,mode,gen_kwargs,process_attr,process_kwargs",
    [
        (
            {"batch": True, "yes": True},
            {},
            "batch",
            {"skip_prompts": True},
            "process_clips",
            {"output_format": "clip", "include_severity": False},
        ),
        (
            {"lines": "5", "screen": True, "yes": True},
            {"line_numbers": [5]},
            "line",
            {"line_numbers": [5], "skip_prompts": True},
            "process_clips",
            {"output_format": "screen", "include_severity": False},
        ),
        (
            {"category": "Observations,Onboarding", "yes": True},
            {},
            "category",
            {"skip_prompts": True, "categories": ["Observations", "Onboarding"]},
            "process_clips",
            {"output_format": "clip", "include_severity": False},
        ),
        (
            {"reel": "11, P01.5", "yes": True},
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


def test_whisper_model_flag_parses(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--whisper-model", "medium"])
    args = cli.parse_arguments()
    assert args.whisper_model == "medium"


def test_whisper_model_flag_rejects_invalid(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--whisper-model", "huge"])
    with pytest.raises(SystemExit):
        cli.parse_arguments()


def test_ollama_model_flag_parses(monkeypatch):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--ollama-model", "gemma3:4b"])
    args = cli.parse_arguments()
    assert args.ollama_model == "gemma3:4b"


def test_whisper_model_applies_to_config(monkeypatch):
    import config

    original = config.TRANSCRIBE_MODEL
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--whisper-model", "small"])
    args = cli.parse_arguments()
    cli._apply_config_overrides(args, cli_mode=True)
    assert config.TRANSCRIBE_MODEL == "small"
    config.TRANSCRIBE_MODEL = original


def test_ollama_model_applies_to_config(monkeypatch):
    import config

    original = config.OLLAMA_SUMMARY_MODEL
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--ollama-model", "gemma3:4b"])
    args = cli.parse_arguments()
    cli._apply_config_overrides(args, cli_mode=True)
    assert config.OLLAMA_SUMMARY_MODEL == "gemma3:4b"
    config.OLLAMA_SUMMARY_MODEL = original
