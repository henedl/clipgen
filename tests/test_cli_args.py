import os
from argparse import Namespace
from unittest.mock import Mock

import pytest

import cli
import app


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
    """Silence the I/O-heavy parts of cli.main() so tests can drive dispatch.

    Both launch paths must be stubbed. ``desktop.launch`` is the one that bites:
    the frozen-no-argv tests below reproduce a Finder double-click exactly, which
    is the real trigger for a native window, so leaving it live blocks the whole
    run inside ``webview.start()`` — a visible window and no test output.
    ``captured["launcher"]`` records which path dispatch actually chose.
    """
    import os as _os

    import desktop
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

    def _record(launcher):
        def fake(*, worksheet=None, port=None, default_page="studio", **_kw):
            captured["launcher"] = launcher
            captured["worksheet"] = worksheet
            captured["default_page"] = default_page

        return fake

    monkeypatch.setattr(server, "start_combined_server", _record("browser"))
    monkeypatch.setattr(desktop, "launch", _record("desktop"))
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
    # A double-click is the one case that opens a native window.
    assert captured.get("launcher") == "desktop"


def test_frozen_with_explicit_flag_is_respected(monkeypatch, tmp_path):
    """Frozen binary invoked with an explicit flag must not be overridden."""
    captured = _mock_main_side_effects(monkeypatch, tmp_path)
    monkeypatch.setattr("sys.argv", ["clipgen", "--screenspace"])
    monkeypatch.setattr("sys.frozen", True, raising=False)

    with pytest.raises(SystemExit) as exc:
        cli.main()

    assert exc.value.code == 0
    assert captured.get("default_page") == "screenspace"
    # Explicit CLI use of a frozen binary stays on the browser path — the window
    # is reserved for the argument-less double-click.
    assert captured.get("launcher") == "browser"


def test_frozen_no_args_with_browser_flag_uses_browser(monkeypatch, tmp_path):
    """--browser is the escape hatch out of the desktop window."""
    captured = _mock_main_side_effects(monkeypatch, tmp_path)
    monkeypatch.setattr("sys.argv", ["clipgen", "--browser"])
    monkeypatch.setattr("sys.frozen", True, raising=False)

    with pytest.raises(SystemExit) as exc:
        cli.main()

    assert exc.value.code == 0
    assert captured.get("launcher") == "browser"


def test_desktop_flag_from_source_opens_window(monkeypatch, tmp_path):
    """--desktop alone means Studio in a window, mirroring a double-click."""
    captured = _mock_main_side_effects(monkeypatch, tmp_path)
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--desktop"])
    monkeypatch.delattr("sys.frozen", raising=False)

    with pytest.raises(SystemExit) as exc:
        cli.main()

    assert exc.value.code == 0
    assert captured.get("launcher") == "desktop"
    assert captured.get("default_page") == "studio"


def test_web_frontend_from_source_defaults_to_browser(monkeypatch, tmp_path):
    """A plain source-checkout launch keeps the pre-existing browser behaviour."""
    captured = _mock_main_side_effects(monkeypatch, tmp_path)
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--composer"])
    monkeypatch.delattr("sys.frozen", raising=False)

    with pytest.raises(SystemExit) as exc:
        cli.main()

    assert exc.value.code == 0
    assert captured.get("launcher") == "browser"
    assert captured.get("default_page") == "composer"


def test_frozen_no_args_threads_into_standalone_branch(monkeypatch, tmp_path):
    """Frozen-no-argv must dispatch through the no-spreadsheet branch in
    _dispatch_standalone_mode, not the later --studio + -s branch that goes
    through worksheet selection."""
    import excel_io

    captured = _mock_main_side_effects(monkeypatch, tmp_path)
    monkeypatch.setattr(
        cli,
        "select_worksheet",
        lambda *_a, **_kw: pytest.fail(
            "frozen-no-args path must short-circuit before worksheet selection"
        ),
    )
    monkeypatch.setattr(
        excel_io,
        "prompt_for_excel_fallback",
        lambda *_a, **_kw: pytest.fail(
            "frozen-no-args path must not reach Excel fallback prompt"
        ),
    )
    monkeypatch.setattr(cli, "_try_silent_google_auth", lambda: None)
    monkeypatch.setattr("sys.argv", ["clipgen"])
    monkeypatch.setattr("sys.frozen", True, raising=False)

    with pytest.raises(SystemExit) as exc:
        cli.main()

    assert exc.value.code == 0
    assert captured.get("default_page") == "studio"
    assert captured.get("worksheet") is None


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
    monkeypatch.setattr(app, process_attr, process_fn)
    monkeypatch.setattr(app, "_print_completion_message", completion)

    parsed_defaults = {
        "line_numbers": None,
        "range_start": None,
        "range_end": None,
        "cell_specs": None,
    }
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


def test_select_worksheet_url_skips_drive_listing(monkeypatch):
    """A `-s <url>` open must not pay the rate-limited Drive listing round-trip:
    get_all_spreadsheets is deferred and only fetched by paths that need it."""
    import google_api

    monkeypatch.setattr(
        google_api,
        "get_all_spreadsheets",
        lambda _c: pytest.fail("Drive listing must be skipped for URL opens"),
    )
    sentinel = object()
    seen = {}

    def fake_open_by_url(_client, url, **_kw):
        seen["url"] = url
        return sentinel

    monkeypatch.setattr(app, "open_spreadsheet_by_url", fake_open_by_url)
    monkeypatch.setattr(app, "_is_excel_worksheet", lambda _w: False)

    args = _base_args(spreadsheet="http://example.com/sheet")
    result = cli.select_worksheet(Mock(), args, cli_mode=True)

    assert result is sentinel
    assert seen["url"] == "http://example.com/sheet"


# ---- Single-xlsx fallback in CLI mode ----


def test_single_xlsx_fallback_path_requires_exactly_one(monkeypatch):
    import excel_io

    monkeypatch.setattr(excel_io, "list_excel_in_cwd", list)
    assert cli._single_xlsx_fallback_path("reason") is None

    monkeypatch.setattr(excel_io, "list_excel_in_cwd", lambda: ["a.xlsx", "b.xlsx"])
    assert cli._single_xlsx_fallback_path("reason") is None

    monkeypatch.setattr(excel_io, "list_excel_in_cwd", lambda: ["only.xlsx"])
    assert cli._single_xlsx_fallback_path("reason") == "only.xlsx"


def test_select_worksheet_cli_falls_back_to_single_xlsx(monkeypatch):
    """No -s and no Drive match in CLI mode uses the sole local .xlsx."""
    import excel_io
    import google_api

    monkeypatch.setattr(google_api, "get_all_spreadsheets", lambda _c: [])
    monkeypatch.setattr(app, "open_spreadsheet_by_name", lambda *_a, **_kw: None)
    sentinel = object()
    monkeypatch.setattr(excel_io, "list_excel_in_cwd", lambda: ["only.xlsx"])
    monkeypatch.setattr(excel_io, "open_excel_workbook", lambda _path: sentinel)
    monkeypatch.setattr(app, "_is_excel_worksheet", lambda _w: True)

    result = cli.select_worksheet(Mock(), _base_args(), cli_mode=True)
    assert result is sentinel


def test_select_worksheet_cli_ambiguous_xlsx_still_exits(monkeypatch, capsys):
    """Zero or several local .xlsx files keep the hard error (now mentioning -s excel)."""
    import excel_io
    import google_api

    monkeypatch.setattr(google_api, "get_all_spreadsheets", lambda _c: [])
    monkeypatch.setattr(app, "open_spreadsheet_by_name", lambda *_a, **_kw: None)
    monkeypatch.setattr(excel_io, "list_excel_in_cwd", lambda: ["a.xlsx", "b.xlsx"])
    monkeypatch.setattr(
        excel_io,
        "open_excel_workbook",
        lambda _path: pytest.fail("must not open an ambiguous Excel file"),
    )

    with pytest.raises(SystemExit):
        cli.select_worksheet(Mock(), _base_args(), cli_mode=True)
    assert "-s excel" in capsys.readouterr().out


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


# ---- --settings flag ----


@pytest.mark.parametrize(
    "argv_extra,expected",
    [
        ([], False),
        (["--settings"], True),
    ],
)
def test_settings_flag_parses(monkeypatch, argv_extra, expected):
    monkeypatch.setattr("sys.argv", ["clipgen.py", *argv_extra])
    args = cli.parse_arguments()
    assert args.settings is expected


def test_settings_rejected_with_no_input(monkeypatch, capsys):
    import os

    import utils

    monkeypatch.setattr("sys.argv", ["clipgen.py", "--settings", "--no-input"])
    monkeypatch.setattr(os, "chdir", lambda _p: None)
    monkeypatch.setattr(utils, "validate_runtime_directories", lambda: None)
    monkeypatch.setattr(
        utils,
        "set_program_settings",
        lambda: pytest.fail("settings editor must not open under --no-input"),
    )

    with pytest.raises(SystemExit) as exc:
        cli.main()
    assert exc.value.code == 1
    assert "--settings" in capsys.readouterr().out


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


def test_no_whisper_vad_flag_disables_vad(monkeypatch):
    import config

    original = config.TRANSCRIBE_VAD_FILTER
    try:
        monkeypatch.setattr("sys.argv", ["clipgen.py", "--no-whisper-vad"])
        args = cli.parse_arguments()
        cli._apply_config_overrides(args, cli_mode=True)
        assert config.TRANSCRIBE_VAD_FILTER is False
    finally:
        config.TRANSCRIBE_VAD_FILTER = original


def test_whisper_hallucination_silence_flag_applies_to_config(monkeypatch):
    import config

    original = config.TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD
    try:
        monkeypatch.setattr(
            "sys.argv", ["clipgen.py", "--whisper-hallucination-silence", "2.5"]
        )
        args = cli.parse_arguments()
        cli._apply_config_overrides(args, cli_mode=True)
        assert config.TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD == 2.5
    finally:
        config.TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD = original


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


def test_run_gallery_cli_negative_interval_falls_back_to_default(tmp_path, monkeypatch):
    """A negative --interval must fall back to the default, mirroring the
    interactive gallery guard (a bare `or` would pass the truthy negative)."""
    import config
    import video

    video_file = tmp_path / "clip.mp4"
    video_file.write_bytes(b"x")

    captured = {}

    def fake_captures(_path, *, interval_seconds, output_format, gif_duration_seconds):
        captured["interval"] = interval_seconds
        return []  # return early before viewer build

    monkeypatch.setattr(video, "generate_interval_captures", fake_captures)

    args = Namespace(gallery=str(video_file), interval=-5, gif=False, bundle=False)
    cli._run_gallery_cli(args)
    assert captured["interval"] == config.GALLERY_INTERVAL_SECONDS


# ---- credentials.json resolution ----


def test_credentials_search_path_order(monkeypatch, tmp_path):
    """CWD (beside the app) wins, then gspread's config dir, then clipgen's."""
    import start_settings

    cwd = tmp_path / "app"
    gspread_dir = tmp_path / "gspread"
    clipgen_dir = tmp_path / "clipgen"
    for d in (cwd, gspread_dir, clipgen_dir):
        d.mkdir()

    monkeypatch.chdir(cwd)
    monkeypatch.setattr("gspread.auth.DEFAULT_CONFIG_DIR", str(gspread_dir))
    monkeypatch.setattr(start_settings, "config_dir", lambda: clipgen_dir)

    assert cli.resolve_credentials_path() is None

    # Lowest priority first, then check each higher one takes over.
    (clipgen_dir / "credentials.json").write_text("{}")
    assert cli.resolve_credentials_path() == clipgen_dir / "credentials.json"

    (gspread_dir / "credentials.json").write_text("{}")
    assert cli.resolve_credentials_path() == gspread_dir / "credentials.json"

    (cwd / "credentials.json").write_text("{}")
    assert cli.resolve_credentials_path() == cwd / "credentials.json"


def test_silent_auth_needs_only_a_cached_token(monkeypatch, tmp_path):
    """A valid cached token is sufficient; credentials.json need not be present.

    Requiring both is what pushed users with a good token back through
    "Connect Google" whenever their credentials file lived elsewhere.
    """
    token = tmp_path / "authorized_user.json"
    token.write_text("{}")
    monkeypatch.setattr("gspread.auth.DEFAULT_AUTHORIZED_USER_FILENAME", str(token))
    monkeypatch.setattr(cli, "resolve_credentials_path", lambda: None)

    called = {}

    def fake_oauth(**kwargs):
        called.update(kwargs)
        return "client"

    monkeypatch.setattr("gspread.oauth", fake_oauth)
    assert cli._try_silent_google_auth() == "client"

    # No token -> no attempt at all.
    monkeypatch.setattr(
        "gspread.auth.DEFAULT_AUTHORIZED_USER_FILENAME", str(tmp_path / "missing.json")
    )
    assert cli._try_silent_google_auth() is None


# ---- frozen working directory ----


def test_frozen_macos_working_dir_is_beside_the_bundle(monkeypatch):
    """Writing inside .app/Contents/MacOS would break the signature and hide files."""
    monkeypatch.setattr("sys.frozen", True, raising=False)
    monkeypatch.setattr(
        "sys.executable", "/Applications/clipgen.app/Contents/MacOS/clipgen"
    )
    assert cli.get_runtime_working_dir() == "/Applications"


def test_frozen_plain_binary_uses_its_own_directory(monkeypatch):
    monkeypatch.setattr("sys.frozen", True, raising=False)
    monkeypatch.setattr("sys.executable", "/opt/tools/clipgen")
    monkeypatch.delattr("sys._MEIPASS", raising=False)
    assert cli.get_runtime_working_dir() == "/opt/tools"


def test_frozen_onefile_temp_meipass_is_ignored(monkeypatch):
    """One-file's _MEIPASS is an unrelated temp dir and must not shift the CWD."""
    monkeypatch.setattr("sys.frozen", True, raising=False)
    monkeypatch.setattr("sys.executable", "/opt/tools/clipgen")
    monkeypatch.setattr("sys._MEIPASS", "/var/folders/xy/_MEI123456", raising=False)
    assert cli.get_runtime_working_dir() == "/opt/tools"


def test_frozen_onedir_resolves_beside_the_payload_folder(monkeypatch):
    """One-dir puts the exe inside clipgen/ next to _internal/.

    "Next to the application" is then the folder *containing* clipgen/, which is
    what the user dragged out of the archive and where they will drop
    credentials.json — not the folder holding _internal.
    """
    monkeypatch.setattr("sys.frozen", True, raising=False)
    monkeypatch.setattr("sys.executable", "/Users/me/Apps/clipgen/clipgen.exe")
    monkeypatch.setattr(
        "sys._MEIPASS", "/Users/me/Apps/clipgen/_internal", raising=False
    )
    assert cli.get_runtime_working_dir() == "/Users/me/Apps"


def test_frozen_macos_app_wins_over_onedir_branch(monkeypatch):
    """A one-dir .app has its payload in Contents/Frameworks, not beside the exe."""
    monkeypatch.setattr("sys.frozen", True, raising=False)
    monkeypatch.setattr(
        "sys.executable", "/Applications/clipgen.app/Contents/MacOS/clipgen"
    )
    monkeypatch.setattr(
        "sys._MEIPASS", "/Applications/clipgen.app/Contents/Frameworks", raising=False
    )
    assert cli.get_runtime_working_dir() == "/Applications"


# ---- Finder-launched .app startup (regression: silent exit 1) ----


def test_path_augmentation_only_when_frozen(monkeypatch):
    """A source run already carries the developer's real PATH; leave it alone."""
    import utils

    monkeypatch.delattr("sys.frozen", raising=False)
    monkeypatch.setenv("PATH", "/usr/bin:/bin")
    assert utils.augment_path_for_gui_launch() == []
    assert os.environ["PATH"] == "/usr/bin:/bin"


def test_path_augmentation_appends_existing_dirs_only(monkeypatch, tmp_path):
    """Finder hands a GUI app a PATH without Homebrew, so ffmpeg looks missing.

    Entries are appended rather than prepended: this is about making binaries
    findable, not about overriding a resolution order the user already has.
    """
    import utils

    real = tmp_path / "brew-bin"
    real.mkdir()
    missing = tmp_path / "not-installed"

    monkeypatch.setattr("sys.frozen", True, raising=False)
    monkeypatch.setattr("sys.platform", "darwin")
    monkeypatch.setattr(utils, "_GUI_PATH_DIRS", (str(real), str(missing)))
    monkeypatch.setenv("PATH", "/usr/bin:/bin")

    assert utils.augment_path_for_gui_launch() == [str(real)]
    assert os.environ["PATH"] == f"/usr/bin:/bin{os.pathsep}{real}"
    # Idempotent: relaunching inside one process must not grow PATH forever.
    assert utils.augment_path_for_gui_launch() == []


def test_path_augmentation_skipped_off_macos(monkeypatch, tmp_path):
    """Windows GUI processes inherit the machine PATH, so there is nothing to fix."""
    import utils

    monkeypatch.setattr("sys.frozen", True, raising=False)
    monkeypatch.setattr("sys.platform", "win32")
    monkeypatch.setenv("PATH", "C:\\Windows")
    assert utils.augment_path_for_gui_launch() == []


def test_fatal_startup_error_is_silent_without_gui(monkeypatch):
    """Console runs keep printing only — no dialog subprocess."""
    import utils

    called = []
    monkeypatch.setattr(utils, "_show_native_alert", lambda *a: called.append(a))
    monkeypatch.setattr(utils, "GUI_LAUNCH", False)
    utils.fatal_startup_error("boom", ["detail"])
    assert called == []


def test_fatal_startup_error_surfaces_natively_in_gui_launch(monkeypatch):
    """The actual regression: a windowed launch must not die with a blank screen.

    ffmpeg missing from a Finder-inherited PATH exited 1 with the guidance
    printed to a stdout nobody could see.
    """
    import utils

    called = []
    monkeypatch.setattr(utils, "_show_native_alert", lambda *a: called.append(a))
    monkeypatch.setattr(utils, "GUI_LAUNCH", True)
    utils.fatal_startup_error(
        "Required video tools are missing from PATH.", ["install ffmpeg"]
    )
    assert len(called) == 1
    assert "install ffmpeg" in called[0][1]


def test_missing_ffmpeg_routes_through_fatal_startup_error(monkeypatch):
    """video.check_ffmpeg_tools_available must use the surfacing path, not error_print."""
    import video

    monkeypatch.setattr(video.shutil, "which", lambda _tool: None)
    seen = []
    monkeypatch.setattr(
        video.utils, "fatal_startup_error", lambda m, d=None: seen.append((m, d))
    )
    assert video.check_ffmpeg_tools_available() is False
    assert len(seen) == 1
    assert "ffmpeg" in seen[0][1][0]


class TestInstallGuidance:
    """The guidance in the fatal alert is the only thing a windowed launch
    shows, so it has to work for a Mac without Homebrew too."""

    def _lines(self, monkeypatch, *, platform, has_brew):
        import utils as u

        monkeypatch.setattr(u.sys, "platform", platform)
        monkeypatch.setattr(
            u.shutil,
            "which",
            lambda tool: "/opt/homebrew/bin/brew" if has_brew else None,
        )
        return u.install_guidance_lines(
            brew_command="brew install ffmpeg",
            linux=["Linux (Debian/Ubuntu): sudo apt install ffmpeg"],
            windows=["Windows (winget): winget install Gyan.FFmpeg"],
            download_url="https://www.ffmpeg.org/download.html",
            verify_commands=["ffmpeg -version"],
        )

    def test_macos_with_brew_leads_with_brew(self, monkeypatch):
        lines = self._lines(monkeypatch, platform="darwin", has_brew=True)
        assert any("brew install ffmpeg" in line for line in lines)
        assert not any("brew.sh" in line for line in lines)

    def test_macos_without_brew_leads_with_the_download(self, monkeypatch):
        """`brew install …` is an instruction a brewless Mac cannot follow."""
        lines = self._lines(monkeypatch, platform="darwin", has_brew=False)
        blob = "\n".join(lines)
        assert "Homebrew is not installed" in blob
        assert "https://www.ffmpeg.org/download.html" in blob
        assert "https://brew.sh" in blob
        # The brew command survives as the second route, not the only one.
        assert "brew install ffmpeg" in blob

    def test_first_line_is_actionable_on_its_own(self, monkeypatch):
        """The browser surfaces (Overview gate, Settings note, summary hint) are
        one-liners that render only ``install_hint[0]``. A first line that is a
        bare header — "macOS: Homebrew is not installed." — leaves those users
        with nothing to act on, which is exactly what it did once."""
        for platform, has_brew in (
            ("darwin", True),
            ("darwin", False),
            ("linux", False),
            ("win32", False),
            ("sunos5", False),
        ):
            first = self._lines(monkeypatch, platform=platform, has_brew=has_brew)[0]
            assert (
                "ffmpeg.org/download.html" in first
                or "install ffmpeg" in first
                or "install Gyan.FFmpeg" in first
            ), f"{platform} brew={has_brew}: {first!r}"
            # A header line ends in a period and carries no command or URL.
            assert not first.endswith("is not installed.")

    def test_download_url_reachable_on_every_platform(self, monkeypatch):
        for platform in ("darwin", "linux", "win32", "sunos5"):
            lines = self._lines(monkeypatch, platform=platform, has_brew=True)
            assert any(
                "https://www.ffmpeg.org/download.html" in line for line in lines
            ), platform

    def test_verify_commands_come_last(self, monkeypatch):
        lines = self._lines(monkeypatch, platform="linux", has_brew=False)
        assert lines[-1] == "  ffmpeg -version"
        assert lines[-2] == "Then verify in a new terminal:"
