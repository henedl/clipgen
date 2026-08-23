"""Tests for the thinking-agent CLI flags (--summarize, --citations, --friction)."""

from argparse import Namespace

import pytest

import cli


def _agent_args(**overrides):
    defaults = {
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
        "manifest": False,
        "regenerate": False,
        "studio": False,
        "screenspace": False,
        "transcripts": False,
        "timeline_viewer": False,
        "gallery": None,
        "interval": None,
        "bundle": False,
        "input": None,
        "output": None,
        "titlecards": None,
        "filmstrip": None,
        "transcribe": False,
        "transcript_format": None,
        "pre_transcribe": None,
        "whisper_model": None,
        "llm_model": None,
        "summarize": None,
        "citations": None,
        "friction": None,
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
    }
    defaults.update(overrides)
    return Namespace(**defaults)


def _make_manifest(**entries):
    """Build a fake transcripts manifest dict with the given entries."""
    return {
        "source_transcripts": dict(entries),
        "corrections": {},
        "marks": {},
    }


# ---- Argparse parsing ----


@pytest.mark.parametrize(
    "argv_extra,expected",
    [
        ([], []),
        (["P01", "P03"], ["P01", "P03"]),
    ],
)
def test_parse_summarize(monkeypatch, argv_extra, expected):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--summarize", *argv_extra])
    args = cli.parse_arguments()
    assert args.summarize == expected


def test_parse_citations_with_llm_model(monkeypatch):
    monkeypatch.setattr(
        "sys.argv", ["clipgen.py", "--citations", "P01", "--llm-model", "gemma3:4b"]
    )
    args = cli.parse_arguments()
    assert args.citations == ["P01"]
    assert args.llm_model == "gemma3:4b"


# ---- Conflict validation ----


def test_summarize_conflicts_with_screenspace_ui():
    args = _agent_args(summarize=[], screenspace=True)
    with pytest.raises(SystemExit):
        cli._validate_mode_conflicts(args)


def test_summarize_and_citations_conflict():
    args = _agent_args(summarize=[], citations=[])
    with pytest.raises(SystemExit):
        cli._validate_mode_conflicts(args)


# ---- _run_summarize ----


def test_summarize_default_runs_all(monkeypatch):
    saved = {}

    def fake_save(source_transcripts, corrections, marks=None, known_terms=None):
        saved["source_transcripts"] = {
            k: dict(v) for k, v in source_transcripts.items()
        }

    manifest = _make_manifest(
        P01={"segments": [{"start": 0, "end": 1, "text": "hello"}]},
        P02={"segments": [{"start": 0, "end": 1, "text": "world"}]},
    )
    monkeypatch.setattr(cli.transcripts, "load_transcripts_manifest", lambda: manifest)
    monkeypatch.setattr(cli.transcripts, "save_transcripts_manifest", fake_save)

    import thinking_agents

    monkeypatch.setattr(thinking_agents, "summarize_transcript", lambda segs: "S")

    args = _agent_args(summarize=[])
    cli._run_summarize(args)

    assert saved["source_transcripts"]["P01"]["summary"] == "S"
    assert saved["source_transcripts"]["P02"]["summary"] == "S"


def test_summarize_specific_ids_only(monkeypatch):
    saved = []

    def fake_save(source_transcripts, corrections, marks=None, known_terms=None):
        saved.append({k: dict(v) for k, v in source_transcripts.items()})

    manifest = _make_manifest(
        P01={"segments": [{"start": 0, "end": 1, "text": "hello"}]},
        P02={"segments": [{"start": 0, "end": 1, "text": "world"}]},
    )
    monkeypatch.setattr(cli.transcripts, "load_transcripts_manifest", lambda: manifest)
    monkeypatch.setattr(cli.transcripts, "save_transcripts_manifest", fake_save)

    import thinking_agents

    monkeypatch.setattr(thinking_agents, "summarize_transcript", lambda segs: "S")

    args = _agent_args(summarize=["P02"])
    cli._run_summarize(args)

    # Only P02 should have summary; P01 untouched.
    final = saved[-1]
    assert final["P02"]["summary"] == "S"
    assert "summary" not in final.get("P01", {})


def test_summarize_skips_existing_without_no_input(monkeypatch, capsys):
    saved = []

    def fake_save(source_transcripts, corrections, marks=None, known_terms=None):
        saved.append(True)

    manifest = _make_manifest(
        P01={"segments": [{"start": 0, "end": 1, "text": "hello"}], "summary": "old"},
    )
    monkeypatch.setattr(cli.transcripts, "load_transcripts_manifest", lambda: manifest)
    monkeypatch.setattr(cli.transcripts, "save_transcripts_manifest", fake_save)

    import thinking_agents

    monkeypatch.setattr(thinking_agents, "summarize_transcript", lambda segs: "NEW")

    args = _agent_args(summarize=["P01"], no_input=False)
    cli._run_summarize(args)

    out = capsys.readouterr().out
    assert "already present" in out
    # Manifest must NOT have been overwritten — summarize should not have been called
    assert manifest["source_transcripts"]["P01"]["summary"] == "old"


def test_summarize_overwrites_with_no_input(monkeypatch):
    saved = []

    def fake_save(source_transcripts, corrections, marks=None, known_terms=None):
        saved.append({k: dict(v) for k, v in source_transcripts.items()})

    manifest = _make_manifest(
        P01={"segments": [{"start": 0, "end": 1, "text": "hello"}], "summary": "old"},
    )
    monkeypatch.setattr(cli.transcripts, "load_transcripts_manifest", lambda: manifest)
    monkeypatch.setattr(cli.transcripts, "save_transcripts_manifest", fake_save)

    import thinking_agents

    monkeypatch.setattr(thinking_agents, "summarize_transcript", lambda segs: "NEW")

    args = _agent_args(summarize=["P01"], no_input=True)
    cli._run_summarize(args)

    assert saved[-1]["P01"]["summary"] == "NEW"


def test_summarize_handles_none_result(monkeypatch, capsys):
    manifest = _make_manifest(
        P01={"segments": [{"start": 0, "end": 1, "text": "hi"}]},
    )
    monkeypatch.setattr(cli.transcripts, "load_transcripts_manifest", lambda: manifest)
    monkeypatch.setattr(
        cli.transcripts, "save_transcripts_manifest", lambda *a, **kw: None
    )

    import thinking_agents

    monkeypatch.setattr(thinking_agents, "summarize_transcript", lambda segs: None)

    args = _agent_args(summarize=["P01"])
    cli._run_summarize(args)

    err = capsys.readouterr().out
    assert "P01" in err
    assert "summary" not in manifest["source_transcripts"]["P01"]


# ---- _run_citations ----


def test_citations_requires_summary(monkeypatch, capsys):
    manifest = _make_manifest(
        P01={"segments": [{"start": 0, "end": 1, "text": "hi"}]},
    )
    monkeypatch.setattr(cli.transcripts, "load_transcripts_manifest", lambda: manifest)
    monkeypatch.setattr(
        cli.transcripts, "save_transcripts_manifest", lambda *a, **kw: None
    )

    import thinking_agents

    monkeypatch.setattr(
        thinking_agents,
        "find_citations",
        lambda s, segs: [{"sentence": "x", "refs": []}],
    )

    args = _agent_args(citations=["P01"])
    cli._run_citations(args)

    err = capsys.readouterr().out
    assert "P01" in err and "summary" in err
    assert "citations" not in manifest["source_transcripts"]["P01"]


def test_citations_writes_refs(monkeypatch):
    saved = []

    def fake_save(source_transcripts, corrections, marks=None, known_terms=None):
        saved.append({k: dict(v) for k, v in source_transcripts.items()})

    manifest = _make_manifest(
        P01={
            "segments": [{"start": 0, "end": 1, "text": "hi"}],
            "summary": "S",
        },
    )
    monkeypatch.setattr(cli.transcripts, "load_transcripts_manifest", lambda: manifest)
    monkeypatch.setattr(cli.transcripts, "save_transcripts_manifest", fake_save)

    import thinking_agents

    fake_citations = [{"sentence": "claim", "refs": [{"start": 0, "end": 1}]}]
    monkeypatch.setattr(
        thinking_agents, "find_citations", lambda s, segs: fake_citations
    )

    args = _agent_args(citations=["P01"])
    cli._run_citations(args)

    assert saved[-1]["P01"]["citations"] == fake_citations


# ---- _run_friction_agent ----


def _fake_friction_agent(monkeypatch, result):
    """Stub the friction registry entry with a run() returning `result`."""
    import thinking_agents

    calls = []

    def fake_run(entry, cancel_event, on_token=None):
        calls.append(entry)
        return result

    monkeypatch.setattr(
        thinking_agents, "get_agent", lambda key: {"key": key, "run": fake_run}
    )
    return calls


@pytest.mark.parametrize(
    "argv_extra,expected",
    [
        ([], []),
        (["P01", "P03"], ["P01", "P03"]),
    ],
)
def test_parse_friction(monkeypatch, argv_extra, expected):
    monkeypatch.setattr("sys.argv", ["clipgen.py", "--friction", *argv_extra])
    args = cli.parse_arguments()
    assert args.friction == expected


def test_friction_and_summarize_conflict():
    args = _agent_args(friction=[], summarize=[])
    with pytest.raises(SystemExit):
        cli._validate_mode_conflicts(args)


def test_friction_requires_summary(monkeypatch, capsys):
    manifest = _make_manifest(
        P01={"segments": [{"start": 0, "end": 1, "text": "hi"}]},
    )
    monkeypatch.setattr(cli.transcripts, "load_transcripts_manifest", lambda: manifest)
    monkeypatch.setattr(
        cli.transcripts, "save_transcripts_manifest", lambda *a, **kw: None
    )
    calls = _fake_friction_agent(monkeypatch, {"moments": [], "llm_ok": True})

    args = _agent_args(friction=["P01"])
    cli._run_friction_agent(args)

    out = capsys.readouterr().out
    assert "P01" in out and "summary" in out
    assert calls == []
    assert "friction" not in manifest["source_transcripts"]["P01"]


def test_friction_skips_existing_without_no_input(monkeypatch, capsys):
    manifest = _make_manifest(
        P01={
            "segments": [{"start": 0, "end": 1, "text": "hi"}],
            "summary": "S",
            "friction": {"moments": [], "llm_ok": True},
        },
    )
    monkeypatch.setattr(cli.transcripts, "load_transcripts_manifest", lambda: manifest)
    monkeypatch.setattr(
        cli.transcripts, "save_transcripts_manifest", lambda *a, **kw: None
    )
    calls = _fake_friction_agent(monkeypatch, {"moments": [], "llm_ok": True})

    args = _agent_args(friction=["P01"], no_input=False)
    cli._run_friction_agent(args)

    assert "already present" in capsys.readouterr().out
    assert calls == []


def test_friction_writes_result_with_no_input(monkeypatch):
    saved = []

    def fake_save(source_transcripts, corrections, marks=None, known_terms=None):
        saved.append({k: dict(v) for k, v in source_transcripts.items()})

    manifest = _make_manifest(
        P01={
            "segments": [{"start": 0, "end": 1, "text": "hi"}],
            "summary": "S",
            "friction": {"moments": [], "llm_ok": False},
        },
    )
    monkeypatch.setattr(cli.transcripts, "load_transcripts_manifest", lambda: manifest)
    monkeypatch.setattr(cli.transcripts, "save_transcripts_manifest", fake_save)
    fake_result = {"moments": [{"reason": "r"}], "llm_ok": True}
    _fake_friction_agent(monkeypatch, fake_result)

    args = _agent_args(friction=["P01"], no_input=True)
    cli._run_friction_agent(args)

    assert saved[-1]["P01"]["friction"] == fake_result


def test_friction_warns_when_llm_failed(monkeypatch, capsys):
    manifest = _make_manifest(
        P01={"segments": [{"start": 0, "end": 1, "text": "hi"}], "summary": "S"},
    )
    monkeypatch.setattr(cli.transcripts, "load_transcripts_manifest", lambda: manifest)
    monkeypatch.setattr(
        cli.transcripts, "save_transcripts_manifest", lambda *a, **kw: None
    )
    _fake_friction_agent(monkeypatch, {"moments": [], "llm_ok": False})

    args = _agent_args(friction=["P01"])
    cli._run_friction_agent(args)

    out = capsys.readouterr().out
    assert "programmatic scores stored" in out
    assert manifest["source_transcripts"]["P01"]["friction"]["llm_ok"] is False
