"""The two AI switches: ``AI_ENABLED`` and ``AI_TRANSCRIPTION_ENABLED``.

Off must hold at every layer: the model chokepoints, the routes, the agent
chain, the CLI and the Workflows runner. Existing results must stay readable.
"""

import argparse
import re
import threading
from pathlib import Path

import pytest
from flask import Flask

import cli
import config
import llm_client
import redact
import screenspace
import screenspace_ocr
import server
import server_utils
import speakers
import start_settings
import transcripts
import transcripts_server
import utils
import workflows
import workflows_catalog

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "source"
WEB = ROOT / "assets" / "web"


@pytest.fixture
def ai_off(monkeypatch):
    monkeypatch.setattr(config, "AI_ENABLED", False)


@pytest.fixture
def transcription_off(monkeypatch):
    monkeypatch.setattr(config, "AI_TRANSCRIPTION_ENABLED", False)


@pytest.fixture
def no_network(monkeypatch):
    """Any socket or subprocess the gated code opens fails the test."""

    def _boom(*_args, **_kwargs):
        raise AssertionError("gated code reached the network or a subprocess")

    monkeypatch.setattr(llm_client.urllib.request, "urlopen", _boom)
    monkeypatch.setattr(llm_client.subprocess, "Popen", _boom)


# ---- The switches are independent ----


def test_switches_block_only_their_own_kind(monkeypatch):
    assert not utils.ai_blocked("ai")
    assert not utils.ai_blocked("transcription")
    monkeypatch.setattr(config, "AI_ENABLED", False)
    assert utils.ai_blocked("ai")
    assert not utils.ai_blocked("transcription")
    monkeypatch.setattr(config, "AI_ENABLED", True)
    monkeypatch.setattr(config, "AI_TRANSCRIPTION_ENABLED", False)
    assert utils.ai_blocked("transcription")
    assert not utils.ai_blocked("ai")


def test_every_kind_has_a_message():
    assert set(config.AI_OFF_MESSAGES) == {"ai", "transcription"}
    for kind in config.AI_OFF_MESSAGES:
        assert utils.ai_off_message(kind)


# ---- Chokepoints ----


def test_llm_refuses_without_touching_the_server(ai_off, no_network, monkeypatch):
    # A server someone else started must not be used either.
    monkeypatch.setattr(llm_client, "is_available", lambda: True)
    assert llm_client.generate("hello") is None
    assert llm_client.take_last_error() == utils.ai_off_message()
    assert llm_client.ensure_server() is False
    assert llm_client.start_server() is False
    assert llm_client.download_model("org/repo:Q4_K_M") is False


def test_redact_refuses_detection_and_download(ai_off, no_network, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True)
    segments = [{"text": "mail me at a@b.se"}]
    assert redact.redact_entry(segments, min_score=0.5, org=False) is False
    assert "pii" not in segments[0]
    assert redact.download() is False


def test_speakers_refuse_even_in_debug_mode(ai_off, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True)
    segments = [{"start": 0.0, "end": 1.0, "text": "hi"}]
    block = speakers.diarize_entry(["x.mp4"], segments, 0, None, max_speakers=2)
    assert block is None
    assert "speaker" not in segments[0]


def test_ocr_refuses_before_building_an_engine(ai_off, monkeypatch):
    def _boom(_model):
        raise AssertionError("OCR engine built with AI off")

    monkeypatch.setattr(screenspace_ocr, "_build_ocr_reader", _boom)
    with (
        pytest.raises(RuntimeError, match="AI features are off"),
        screenspace_ocr._checkout_ocr_reader(["en"]),
    ):
        pass


def test_transcription_refuses_even_in_debug_mode(transcription_off, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True)
    assert transcripts.transcribe_video("x.mp4") is None
    assert transcripts.warmup_transcription_model() is False
    monkeypatch.setattr(config, "DEBUGGING", False)
    assert transcripts._load_model("tiny") is None


def test_transcription_survives_the_ai_switch(ai_off, monkeypatch):
    monkeypatch.setattr(config, "DEBUGGING", True)
    assert transcripts.warmup_transcription_model() is True


def test_download_aborts_when_the_switch_flips_mid_file(monkeypatch, tmp_path):
    class _Response:
        headers = {"Content-Length": "8"}

        def __enter__(self):
            return self

        def __exit__(self, *_exc):
            return False

        def read(self, _size):
            # The first chunk lands, then the user turns AI off.
            config.AI_ENABLED = False
            return b"abcd"

    monkeypatch.setattr(config, "AI_ENABLED", True)
    monkeypatch.setattr(
        llm_client.urllib.request, "urlopen", lambda *_a, **_k: _Response()
    )
    target = tmp_path / "model.gguf"
    assert not llm_client.stream_download("http://x", target, sha256="", size=8)
    assert not target.exists()
    assert not list(tmp_path.glob("*.part"))


# ---- Existing results stay readable ----


def test_stored_redaction_still_renders_with_ai_off(ai_off):
    text = "Anna said hi"
    span = {"start": 0, "end": 4, "label": "GIVEN_NAME", "text": "Anna", "score": 0.9}
    segments = [{"text": text, "pii": [span], "pii_crc": redact.text_crc(text)}]
    spans = redact.entry_spans(segments, [text])
    assert spans[0] and spans[0][0]["label"] == "GIVEN_NAME"


def test_manual_terms_still_match_with_ai_off(ai_off):
    spans = redact.entry_spans(
        [{"text": "Call Acme now"}], ["Call Acme now"], terms=["Acme"]
    )
    assert [s["label"] for s in spans[0]] == [redact.TERM_LABEL]


# ---- Routes ----


@pytest.fixture(scope="module")
def tr_app():
    app = Flask(__name__)
    app.register_blueprint(transcripts_server.transcripts_bp, url_prefix="/t")
    return app


@pytest.fixture
def tr_client(tr_app, monkeypatch):
    entry = {"segments": [{"id": "s1", "start": 0.0, "end": 1.0, "text": "hi"}]}
    monkeypatch.setattr(
        transcripts_server,
        "_manifest",
        {"source_transcripts": {"P01": entry}, "corrections": [], "marks": []},
    )
    monkeypatch.setattr(transcripts_server, "_participants", [])
    monkeypatch.setattr(transcripts_server, "_worker", None)
    monkeypatch.setattr(transcripts_server, "_persist_manifest", lambda: None)
    with tr_app.test_client() as client:
        yield client


AI_ROUTES = [
    ("post", "/t/api/speakers/P01/regenerate", None),
    ("put", "/t/api/speakers/P01", {"enabled": True}),
    ("post", "/t/api/redact/P01/regenerate", None),
    ("put", "/t/api/redact/P01", {"enabled": True}),
    ("post", "/t/api/agent/summary/P01/regenerate", None),
    ("post", "/t/api/models/llm/download", {"model": "org/repo:Q4_K_M"}),
    ("post", "/t/api/models/redact/download", None),
    ("post", "/t/api/models/llm/start", None),
]


@pytest.mark.parametrize("method,path,body", AI_ROUTES)
def test_ai_routes_answer_403(tr_client, ai_off, method, path, body):
    resp = getattr(tr_client, method)(path, json=body)
    assert resp.status_code == 403
    data = resp.get_json()
    assert data["ai_disabled"] == "ai"
    assert data["error"] == utils.ai_off_message("ai")


def test_transcribe_route_answers_403(tr_client, transcription_off):
    resp = tr_client.post("/t/api/transcribe", json={"participants": ["P01"]})
    assert resp.status_code == 403
    assert resp.get_json()["ai_disabled"] == "transcription"


def test_warmup_skips_quietly(tr_client, transcription_off):
    # Pages post this unprompted, so it must not read as an error.
    resp = tr_client.post("/t/api/transcribe/warmup", json={})
    assert resp.status_code == 200
    assert resp.get_json()["reason"] == "ai_disabled"


def test_stop_and_off_routes_stay_open(tr_client, ai_off):
    assert tr_client.post("/t/api/speakers/P01/stop").status_code == 200
    assert tr_client.post("/t/api/redact/P01/stop").status_code == 200
    assert tr_client.post("/t/api/agent/summary/P01/stop").status_code == 200
    assert (
        tr_client.put("/t/api/speakers/P01", json={"enabled": False}).status_code == 200
    )


def test_redact_switch_works_with_terms_and_ai_off(tr_client, ai_off):
    transcripts_server._manifest["redact_terms"] = ["Acme"]
    resp = tr_client.put("/t/api/redact/P01", json={"enabled": True})
    assert resp.status_code == 200
    assert resp.get_json()["task"] is None


# ---- Agent chain ----


def test_agents_never_start_with_ai_off(tr_client, ai_off, monkeypatch):
    orchestrator = transcripts_server._orchestrator
    assert orchestrator.next_eligible("P01") is None
    started: list[threading.Thread] = []
    monkeypatch.setattr(threading.Thread, "start", lambda self: started.append(self))
    orchestrator.run_agent("summary", "P01", force=True)
    assert not started


# ---- Settings: the off-hook, reset, records ----


def test_turning_a_switch_off_stops_its_work(monkeypatch):
    stopped: list[str] = []
    monkeypatch.setattr(transcripts_server, "stop_ai_work", stopped.append)
    monkeypatch.setattr(server, "_save_studio_settings", lambda _o: None)
    monkeypatch.setattr(config, "AI_ENABLED", True)
    monkeypatch.setattr(config, "AI_TRANSCRIPTION_ENABLED", True)

    server._apply_settings_payload({"settings": {"AI_ENABLED": False}})
    assert stopped == ["ai"]
    # Saving again while already off stops nothing twice.
    server._apply_settings_payload({"settings": {"AI_ENABLED": False}})
    assert stopped == ["ai"]
    server._apply_settings_payload({"settings": {"AI_TRANSCRIPTION_ENABLED": False}})
    assert stopped == ["ai", "transcription"]


def test_reset_never_turns_a_switch_back_on(monkeypatch):
    monkeypatch.setattr(server, "_save_studio_settings", lambda _o: None)
    monkeypatch.setattr(config, "AI_ENABLED", False)
    monkeypatch.setattr(config, "AI_TRANSCRIPTION_ENABLED", False)
    monkeypatch.setattr(config, "AI_CHOICE_MADE", True)
    for directive in ("all", "tab:General"):
        applied, error = server._apply_settings_payload({"reset": directive})
        assert error is None
        assert not any(name.startswith("AI_") for name in applied)
    assert config.AI_ENABLED is False
    assert config.AI_TRANSCRIPTION_ENABLED is False
    assert config.AI_CHOICE_MADE is True


def test_switches_lead_the_general_tab():
    names = list(config.STUDIO_SETTINGS)
    assert names[:2] == ["AI_ENABLED", "AI_TRANSCRIPTION_ENABLED"]
    for name in names[:2]:
        meta = config.STUDIO_SETTINGS[name]
        assert meta["tab"] == "General"
        assert meta["type"] == "ai_switch"
        assert meta["kind"] in config.AI_OFF_MESSAGES


def test_settings_records_carry_the_gate():
    records = {r["name"]: r for r in server._settings_records()}
    assert records["LLM_SUMMARY_MODEL"]["ai"] == "ai"
    assert records["TRANSCRIBE_SPEAKERS"]["ai"] == "ai"
    assert records["TRANSCRIBE_MODEL"]["ai"] == "transcription"
    assert records["MANIFEST_ENABLED"]["ai"] is None
    # Redaction settings also govern stored spans and manual terms.
    assert records["TRANSCRIBE_REDACT"]["ai"] is None


def test_every_summaries_setting_is_gated():
    for name, meta in config.STUDIO_SETTINGS.items():
        if meta["tab"] == "Summaries":
            assert meta.get("ai") == "ai", name
        if meta.get("ai"):
            assert meta["ai"] in config.AI_OFF_MESSAGES, name


# ---- Page flags ----


def test_index_page_carries_the_off_flags(monkeypatch):
    on = server_utils.render_index_html(WEB, "studio.html")
    assert "dataset.aiOff" not in on
    monkeypatch.setattr(config, "AI_ENABLED", False)
    ai_only = server_utils.render_index_html(WEB, "studio.html")
    assert "dataset.aiOff" in ai_only
    assert "dataset.transcriptionOff" not in ai_only
    monkeypatch.setattr(config, "AI_TRANSCRIPTION_ENABLED", False)
    both = server_utils.render_index_html(WEB, "studio.html")
    assert "dataset.transcriptionOff" in both
    # Turning the switches back on must not serve the cached off page.
    monkeypatch.setattr(config, "AI_ENABLED", True)
    monkeypatch.setattr(config, "AI_TRANSCRIPTION_ENABLED", True)
    assert "dataset.aiOff" not in server_utils.render_index_html(WEB, "studio.html")


# ---- CLI ----


def _save_switches(**values):
    start_settings.save_config_json(config.STUDIO_SETTINGS_FILENAME, values)


def test_cli_reads_the_switches_from_studio_settings(monkeypatch):
    monkeypatch.setattr(config, "AI_ENABLED", True)
    monkeypatch.setattr(config, "AI_TRANSCRIPTION_ENABLED", True)
    _save_switches(AI_ENABLED=False)
    cli._apply_ai_switches(argparse.Namespace())
    assert config.AI_ENABLED is False
    assert config.AI_TRANSCRIPTION_ENABLED is True


@pytest.mark.parametrize(
    "attr,value,saved,flag",
    [
        ("summarize", [], {"AI_ENABLED": False}, "--summarize"),
        ("citations", ["P01"], {"AI_ENABLED": False}, "--citations"),
        ("friction", [], {"AI_ENABLED": False}, "--friction"),
        ("pre_transcribe", [], {"AI_TRANSCRIPTION_ENABLED": False}, "--pre-transcribe"),
    ],
)
def test_cli_refuses_blocked_modes(monkeypatch, capsys, attr, value, saved, flag):
    monkeypatch.setattr(config, "AI_ENABLED", True)
    monkeypatch.setattr(config, "AI_TRANSCRIPTION_ENABLED", True)
    _save_switches(**saved)
    with pytest.raises(SystemExit) as exit_info:
        cli._apply_ai_switches(argparse.Namespace(**{attr: value}))
    assert exit_info.value.code == 1
    assert flag in capsys.readouterr().out


def test_cli_lets_unblocked_modes_through(monkeypatch):
    monkeypatch.setattr(config, "AI_ENABLED", True)
    monkeypatch.setattr(config, "AI_TRANSCRIPTION_ENABLED", True)
    _save_switches(AI_ENABLED=False)
    # Transcription is its own switch; an absent flag parses to None.
    cli._apply_ai_switches(argparse.Namespace(pre_transcribe=[], summarize=None))


# ---- Screenspace ----


def test_ocr_tools_are_the_only_gated_tools():
    gated = {name for name, facts in screenspace.tool_catalog().items() if facts["ai"]}
    assert gated == {"text", "numbers"}


def test_multitool_is_gated_by_its_steps():
    assert screenspace.task_uses_ai("text", {})
    assert not screenspace.task_uses_ai("color", {})
    assert screenspace.task_uses_ai(
        "multitool", {"steps": [{"type": "color"}, {"type": "numbers"}]}
    )
    assert not screenspace.task_uses_ai(
        "multitool", {"steps": [{"type": "color"}, {"type": "change"}]}
    )


def test_screenspace_task_validation_refuses_ocr(ai_off):
    import screenspace_server

    with pytest.raises(server_utils.ApiError) as raised:
        screenspace_server._validate_task_request(
            {"type": "text", "participant": "P01", "region": "r", "parameters": {}}
        )
    assert raised.value.code == 403
    assert raised.value.fields == {"ai_disabled": "ai"}


# ---- Workflows ----


def test_workflow_nodes_name_their_switch():
    kind = workflows_catalog.node_ai_kind
    assert kind("transcribe", {}) == "transcription"
    for node in (
        "summarize",
        "citations",
        "friction",
        "report",
        "ss_text",
        "ss_numbers",
    ):
        assert kind(node, {}) == "ai", node
    assert kind("make_clips", {}) is None
    # Detect follows its detector; the default is an OCR one.
    assert kind("detect", {}) == "ai"
    assert kind("detect", {"detector": "numbers"}) == "ai"
    assert kind("detect", {"detector": "color"}) is None


def test_every_thinking_node_is_gated():
    for node_id, spec in workflows_catalog.NODE_TYPES.items():
        if spec["domain"] == "thinking":
            assert spec.get("ai") == "ai", node_id


def test_runner_fails_a_gated_node_without_executing_it(tmp_path, ai_off, monkeypatch):
    def _boom(_ctx, _inputs, _params):
        raise AssertionError("gated node executed")

    monkeypatch.setitem(workflows.NODE_TYPES["summarize"], "execute", _boom)
    runner = workflows.WorkflowRunner(
        "run_test",
        {
            "id": "bp",
            "nodes": [{"id": "s", "type": "summarize", "params": {}}],
            "edges": [],
        },
        workflows.NodeContext(
            input_dir=tmp_path, output_dir=tmp_path, cancel_event=threading.Event()
        ),
    )
    runner.run()
    assert runner.node_states["s"]["status"] == "failed"
    assert runner.node_states["s"]["error"] == utils.ai_off_message("ai")


# ---- Downloaded models: usage and removal ----


@pytest.fixture
def models_client(tr_app, monkeypatch, tmp_path):
    whisper = tmp_path / "hub" / "models--Systran--faster-whisper-tiny"
    (whisper / "blobs").mkdir(parents=True)
    (whisper / "blobs" / "model.bin").write_bytes(b"w" * 10)
    monkeypatch.setattr(transcripts, "whisper_cache_dirs", lambda: [whisper])
    llm_dir = llm_client.models_dir()
    llm_dir.mkdir(parents=True, exist_ok=True)
    (llm_dir / "tiny.gguf").write_bytes(b"g" * 7)
    redact_dir = redact.models_dir()
    redact_dir.mkdir(parents=True, exist_ok=True)
    for asset in redact.ASSETS:
        (redact_dir / asset["filename"]).write_bytes(b"r")
    monkeypatch.setattr(transcripts_server, "_worker", None)
    monkeypatch.setattr(llm_client, "stop_server", lambda: None)
    with tr_app.test_client() as client:
        yield client, whisper, llm_dir, redact_dir


def test_downloaded_models_are_counted_per_switch(models_client):
    client, _whisper, _llm_dir, _redact_dir = models_client
    data = client.get("/t/api/models/downloaded").get_json()
    assert data["transcription"] == {"count": 1, "bytes": 10}
    assert data["ai"] == {
        "count": 1 + len(redact.ASSETS),
        "bytes": 7 + len(redact.ASSETS),
    }


def test_removing_ai_models_leaves_whisper_alone(models_client):
    client, whisper, llm_dir, redact_dir = models_client
    resp = client.delete("/t/api/models/downloaded/ai")
    assert resp.status_code == 200
    assert not list(llm_dir.glob("*.gguf"))
    assert not redact.is_redact_model_available()
    assert not any(redact_dir.iterdir())
    assert whisper.is_dir()


def test_removing_whisper_models_leaves_ai_alone(models_client):
    client, whisper, llm_dir, _redact_dir = models_client
    resp = client.delete("/t/api/models/downloaded/transcription")
    assert resp.status_code == 200
    assert not whisper.exists()
    assert (llm_dir / "tiny.gguf").is_file()


def test_removal_only_unlinks_a_symlinked_model(models_client, tmp_path):
    client, _whisper, llm_dir, _redact_dir = models_client
    shared = tmp_path / "other-tool-cache.gguf"
    shared.write_bytes(b"s" * 5)
    (llm_dir / "linked.gguf").symlink_to(shared)
    assert client.delete("/t/api/models/downloaded/ai").status_code == 200
    assert not (llm_dir / "linked.gguf").exists()
    assert shared.is_file()


def test_removal_is_refused_while_models_are_in_use(models_client, monkeypatch):
    client, whisper, _llm_dir, _redact_dir = models_client
    monkeypatch.setattr(transcripts_server, "_models_in_use", lambda _kind: True)
    resp = client.delete("/t/api/models/downloaded/transcription")
    assert resp.status_code == 400
    assert whisper.is_dir()


def test_removal_rejects_an_unknown_kind(models_client):
    client = models_client[0]
    assert client.delete("/t/api/models/downloaded/everything").status_code == 404


# ---- Source guards ----

# Each model is built in one file, behind that file's gate.
MODEL_SITES = {
    "WhisperModel(": "transcripts.py",
    "RapidOCR(": "screenspace_ocr.py",
    "InferenceSession(": "speakers.py",
    "tflite_numpy.load_graph(": "redact.py",
}


def test_models_are_only_built_in_gated_files():
    for call, home in MODEL_SITES.items():
        users = sorted(
            path.name
            for path in SOURCE.glob("*.py")
            if call in path.read_text(encoding="utf-8")
        )
        assert users == [home], f"{call} moved or gained a caller: {users}"
        assert "ai_blocked(" in (SOURCE / home).read_text(encoding="utf-8"), home


def test_frontend_gate_tags_name_a_real_switch():
    values = set()
    for path in [*WEB.glob("*.js"), *WEB.glob("*.html")]:
        text = path.read_text(encoding="utf-8")
        values.update(re.findall(r'data-ai="([a-z]+)"', text))
        values.update(re.findall(r'setAttribute\("data-ai", "([a-z]+)"\)', text))
    assert values == set(config.AI_OFF_MESSAGES)


def test_gate_intercepts_clicks_in_the_capture_phase():
    utils_js = (WEB / "utils.js").read_text(encoding="utf-8")
    listener = utils_js[utils_js.index("// Capture phase:") :]
    listener = listener[: listener.index("}, true);") + len("}, true);")]
    assert "clipgenAiGate(e.target)" in listener
    assert "stopImmediatePropagation" in listener


def test_settings_modal_reloads_after_a_switch_toggle():
    modal = (WEB / "settings-modal.js").read_text(encoding="utf-8")
    switch = modal[modal.index("function _buildAiSwitchRow") :]
    switch = switch[: switch.index("\n  }\n") + 1]
    assert "REOPEN_KEY" in switch
    # Only a toggle may send the switches; other saves leave them out.
    assert "_aiDirty = true" in switch
    assert 's.group === "AI" && !_aiDirty' in modal
    assert "window.location.reload()" in switch
    assert "focusSetting" in modal


def test_start_overlay_asks_once():
    overlay = (WEB / "start-overlay.js").read_text(encoding="utf-8")
    assert "AI_CHOICE_MADE: true" in overlay
    assert 'data-role="ai-choice"' in (WEB / "start-overlay.html").read_text(
        encoding="utf-8"
    )
