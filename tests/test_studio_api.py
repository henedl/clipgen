import io
import json
from unittest.mock import Mock

import pytest

Flask = pytest.importorskip("flask").Flask
import config
import server
import itertools


def _set_artifacts(monkeypatch, artifacts):
    """Patch _generated_artifacts and rebuild the lookup index in lockstep."""
    monkeypatch.setattr(server, "_generated_artifacts", list(artifacts))
    monkeypatch.setattr(server, "_generated_artifacts_index", {})
    server._rebuild_artifact_index()


@pytest.fixture(scope="module")
def studio_app():
    """The Flask app, built once for the module.

    Registering the blueprint compiles ~40 Werkzeug URL rules, which dominates
    this fixture's cost — and the app object holds no per-test state: everything
    these tests touch lives in ``server`` module globals, re-pinned per test by
    the function-scoped ``client`` below.
    """
    app = Flask(__name__)
    app.register_blueprint(server.studio_bp, url_prefix="/studio")
    return app


@pytest.fixture
def client(studio_app, monkeypatch):
    # Default: no worksheet/context loaded (error state)
    monkeypatch.setattr(server, "_worksheet", None)
    monkeypatch.setattr(server, "_sheet_context", None)
    monkeypatch.setattr(server, "_sheet_payload_cache", None)
    _set_artifacts(monkeypatch, [])
    monkeypatch.setattr(server, "_generated_reels", [])
    server._release_busy("generate")
    server._release_busy("reel")

    with studio_app.test_client() as c:
        yield c

    server._release_busy("generate")
    server._release_busy("reel")


@pytest.mark.parametrize(
    "method,path,payload",
    [
        ("post", "/studio/api/sheet/refresh", None),
        ("post", "/studio/api/generate", {"cells": ["P01.3"]}),
        ("post", "/studio/api/gallery", {"participant": "P01"}),
        ("post", "/studio/api/timeline-viewer", {}),
        ("post", "/studio/api/highlights-preview", {}),
        ("post", "/studio/api/reel", {"cells": ["P01.3"]}),
    ],
    ids=[
        "sheet_refresh",
        "generate",
        "gallery",
        "timeline_viewer",
        "highlights_preview",
        "reel",
    ],
)
def test_mutating_routes_return_400_when_no_sheet(client, method, path, payload):
    """Mutating routes report a precondition failure when no sheet is loaded."""
    fn = getattr(client, method)
    resp = fn(path, json=payload) if payload is not None else fn(path)
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False
    assert "No spreadsheet loaded" in data["error"]


def test_api_sheet_returns_empty_placeholder_when_no_sheet(client):
    """/api/sheet returns ok with sheet_loaded=False when no spreadsheet is loaded."""
    resp = client.get("/studio/api/sheet")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["sheet_loaded"] is False
    assert data["rows"] == []
    assert data["participants"] == []


# ── Titlecard / endcard background picker endpoints ──────────────────────


def test_api_titlecards_list_synthetic_items(client, tmp_path, monkeypatch):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    data = client.get("/studio/api/titlecards").get_json()
    assert data["ok"] is True
    title_kinds = [it["kind"] for it in data["title"]["items"]]
    assert "default" in title_kinds
    assert "color" in title_kinds
    assert "none" not in title_kinds  # titlecards always render text
    end_kinds = [it["kind"] for it in data["end"]["items"]]
    assert "default" in end_kinds
    assert "none" in end_kinds
    assert "color" in end_kinds


def test_api_titlecard_upload_rejects_bad_extension(client, tmp_path, monkeypatch):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    resp = client.post(
        "/studio/api/titlecards/upload",
        data={"file": (io.BytesIO(b"hello"), "notes.txt")},
        content_type="multipart/form-data",
    )
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


def test_api_titlecard_upload_list_and_serve(client, tmp_path, monkeypatch):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    resp = client.post(
        "/studio/api/titlecards/upload",
        data={"file": (io.BytesIO(b"\x89PNG fake"), "card.png")},
        content_type="multipart/form-data",
    )
    assert resp.status_code == 200
    item = resp.get_json()["item"]
    assert item["kind"] == "upload"
    name = item["id"]
    assert name == "card.png"

    listing = client.get("/studio/api/titlecards").get_json()
    upload_ids = [
        it["id"] for it in listing["title"]["items"] if it["kind"] == "upload"
    ]
    assert name in upload_ids

    served = client.get("/studio/api/titlecards/image/" + name)
    assert served.status_code == 200
    assert served.data == b"\x89PNG fake"


def test_api_titlecard_upload_makes_filename_url_safe(client, tmp_path, monkeypatch):
    """A filename with URL-reserved chars is sanitized so its served URL works."""
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    resp = client.post(
        "/studio/api/titlecards/upload",
        data={"file": (io.BytesIO(b"\x89PNG fake"), "my #1.png")},
        content_type="multipart/form-data",
    )
    assert resp.status_code == 200
    item = resp.get_json()["item"]
    name = item["id"]
    # No URL-reserved character survives into the stored name, id, or URL.
    for bad in ("#", " ", "%", "&"):
        assert bad not in name
        assert bad not in item["url"]
    assert item["url"] == "/api/titlecards/image/" + name

    served = client.get("/studio/api/titlecards/image/" + name)
    assert served.status_code == 200
    assert served.data == b"\x89PNG fake"


def test_api_titlecard_delete_resets_selection(client, tmp_path, monkeypatch):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    images = tmp_path / server.config.TITLECARD_IMAGES_DIRNAME
    images.mkdir()
    (images / "card.png").write_bytes(b"data")
    monkeypatch.setattr(server.config, "TITLECARD_IMAGE", "card.png")

    body = client.delete("/studio/api/titlecards/image/card.png").get_json()
    assert body["ok"] is True
    assert body["reset"].get("TITLECARD_IMAGE") == ""
    assert not (images / "card.png").exists()
    assert server.config.TITLECARD_IMAGE == ""


def test_process_intake_item_releases_reservation_on_cut_exception(
    tmp_path, monkeypatch
):
    """An exception from cut_global_range must not leave a 0-byte intake placeholder."""
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda *a, **k: ["v.mp4"]
    )
    monkeypatch.setattr(server.video, "timeline_or_none", lambda *a, **k: None)

    def boom(*_a, **_k):
        raise RuntimeError("ffmpeg blew up")

    monkeypatch.setattr(server.pipeline, "cut_global_range", boom)

    item = {"participant": "P01", "start": 0.0, "end": 5.0}
    with pytest.raises(RuntimeError):
        server._process_intake_item(item, "clip", "study")
    assert list(tmp_path.glob("*.mp4")) == []


def test_process_intake_item_enforces_size_cap_for_clip_only(tmp_path, monkeypatch):
    """A clip intake compresses to the cap (no wrap here); screenshots/GIFs do not."""
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server.config, "MAX_FILESIZE_MB", 50)
    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda *a, **k: ["v.mp4"]
    )
    monkeypatch.setattr(server.video, "timeline_or_none", lambda *a, **k: None)
    monkeypatch.setattr(
        server.pipeline,
        "cut_global_range",
        lambda *a, **k: {"sourceVideo": "v.mp4", "localStart": 0.0, "localEnd": 5.0},
    )

    enforce = Mock()
    monkeypatch.setattr(server.video, "enforce_filesize_limit", enforce)

    item = {"participant": "P01", "start": 0.0, "end": 5.0}

    server._process_intake_item(item, "clip", "study")
    assert enforce.call_count == 1

    server._process_intake_item(item, "screen", "study")
    server._process_intake_item(item, "gif", "study")
    assert enforce.call_count == 1  # unchanged — only the clip was compressed


def test_settings_records_include_card_pickers(client):
    data = client.get("/studio/api/settings").get_json()
    by_name = {s["name"]: s for s in data["settings"]}
    assert by_name["TITLECARD_IMAGE"]["type"] == "card_picker"
    assert by_name["TITLECARD_IMAGE"]["kind"] == "title"
    assert by_name["ENDCARD_IMAGE"]["kind"] == "end"
    # Solid-color fills are hidden settings (edited via the picker's swatch).
    assert by_name["TITLECARD_COLOR"]["type"] == "hidden"
    assert by_name["ENDCARD_COLOR"]["type"] == "hidden"


def test_settings_records_include_llm_model_pickers(client):
    data = client.get("/studio/api/settings").get_json()
    by_name = {s["name"]: s for s in data["settings"]}
    # Summary and friction models are dynamic pickers populated from installed
    # downloaded GGUF models, not a fixed list.
    assert by_name["LLM_SUMMARY_MODEL"]["type"] == "model_select"
    assert by_name["LLM_SUMMARY_MODEL"]["provider"] == "llm"
    assert by_name["LLM_FRICTION_MODEL"]["type"] == "model_select"
    assert by_name["LLM_FRICTION_MODEL"]["provider"] == "llm"
    # Friction and report inherit the summary model via a blank value, shown as
    # a label.
    assert by_name["LLM_FRICTION_MODEL"]["emptyLabel"]
    assert by_name["LLM_REPORT_MODEL"]["type"] == "model_select"
    assert by_name["LLM_REPORT_MODEL"]["provider"] == "llm"
    assert by_name["LLM_REPORT_MODEL"]["emptyLabel"]


def test_settings_records_include_update_toggle(client):
    data = client.get("/studio/api/settings").get_json()
    by_name = {s["name"]: s for s in data["settings"]}
    assert by_name["UPDATE_CHECK_ON_LAUNCH"]["type"] == "bool"
    assert by_name["UPDATE_CHECK_ON_LAUNCH"]["group"] == "Updates"


def test_settings_records_include_agent_prompts(client):
    data = client.get("/studio/api/settings").get_json()
    by_name = {s["name"]: s for s in data["settings"]}
    for name in (
        "LLM_SUMMARY_PROMPT",
        "LLM_CITATIONS_SYSTEM",
        "LLM_CITATIONS_PROMPT",
        "LLM_FRICTION_SYSTEM",
        "LLM_FRICTION_PROMPT",
        "LLM_REPORT_SYSTEM",
        "LLM_REPORT_PROMPT",
    ):
        assert by_name[name]["type"] == "prompt"
        assert by_name[name]["tab"] == "Summaries"
        assert by_name[name]["group"] == "Agent prompts"
    # User prompts are .format()-ed; their placeholders drive validation.
    assert by_name["LLM_SUMMARY_PROMPT"]["placeholders"] == ["text"]
    assert by_name["LLM_CITATIONS_PROMPT"]["placeholders"] == [
        "claims",
        "transcript",
    ]
    assert by_name["LLM_FRICTION_PROMPT"]["placeholders"] == [
        "summary",
        "segments",
        "limit",
    ]
    assert by_name["LLM_REPORT_PROMPT"]["placeholders"] == [
        "participant",
        "summary",
        "observations",
        "bookmarks",
    ]
    # System prompts are sent verbatim — no placeholders.
    assert by_name["LLM_CITATIONS_SYSTEM"]["placeholders"] == []
    assert by_name["LLM_FRICTION_SYSTEM"]["placeholders"] == []
    assert by_name["LLM_REPORT_SYSTEM"]["placeholders"] == []


def test_settings_put_persists_custom_prompt(client, tmp_path, monkeypatch):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    # Baseline via monkeypatch so the PUT's direct setattr is restored on teardown.
    monkeypatch.setattr(
        server.config, "LLM_SUMMARY_PROMPT", server.config.LLM_SUMMARY_PROMPT
    )
    custom = "Custom summary instructions.\n\nTranscript:\n{text}"
    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"LLM_SUMMARY_PROMPT": custom}},
    )
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True
    assert server.config.LLM_SUMMARY_PROMPT == custom
    saved = server.start_settings.load_config_json(
        server.config.STUDIO_SETTINGS_FILENAME, default={}
    )
    assert saved.get("LLM_SUMMARY_PROMPT") == custom
    # GET reflects the new value.
    by_name = {
        s["name"]: s for s in client.get("/studio/api/settings").get_json()["settings"]
    }
    assert by_name["LLM_SUMMARY_PROMPT"]["value"] == custom


def test_settings_put_rejects_prompt_missing_placeholder(client, tmp_path, monkeypatch):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    default = server.config.LLM_SUMMARY_PROMPT
    monkeypatch.setattr(server.config, "LLM_SUMMARY_PROMPT", default)
    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"LLM_SUMMARY_PROMPT": "No placeholder here."}},
    )
    assert resp.status_code == 400
    body = resp.get_json()
    assert body["ok"] is False
    assert "{text}" in body["error"]
    assert server.config.LLM_SUMMARY_PROMPT == default  # unchanged


def test_settings_put_rejects_prompt_unknown_placeholder(client, tmp_path, monkeypatch):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    default = server.config.LLM_CITATIONS_PROMPT
    monkeypatch.setattr(server.config, "LLM_CITATIONS_PROMPT", default)
    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"LLM_CITATIONS_PROMPT": "{claims} {transcript} {bogus}"}},
    )
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False
    assert server.config.LLM_CITATIONS_PROMPT == default


def test_settings_put_accepts_system_prompt_verbatim(client, tmp_path, monkeypatch):
    """*_SYSTEM prompts are never .format()-ed, so braces are literal and any
    text is accepted."""
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server.config, "LLM_CITATIONS_SYSTEM", "baseline")
    weird = 'Reply with JSON like {"a": 1} and emphasise {clarity}.'
    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"LLM_CITATIONS_SYSTEM": weird}},
    )
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True
    assert server.config.LLM_CITATIONS_SYSTEM == weird


def test_settings_reset_summaries_restores_prompt(client, tmp_path, monkeypatch):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    default = server._settings_defaults["LLM_SUMMARY_PROMPT"]
    monkeypatch.setattr(server.config, "LLM_SUMMARY_PROMPT", "Edited {text}")
    resp = client.put("/studio/api/settings", json={"reset": "tab:Summaries"})
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True
    assert server.config.LLM_SUMMARY_PROMPT == default


def test_settings_records_include_boundary_post_processing(client):
    data = client.get("/studio/api/settings").get_json()
    by_name = {s["name"]: s for s in data["settings"]}
    assert by_name["SCREENSPACE_BOUNDARY_MERGE_THRESHOLD"]["type"] == "float"
    assert by_name["SCREENSPACE_BOUNDARY_MERGE_THRESHOLD"]["tab"] == "Screenspace"
    assert by_name["SCREENSPACE_BOUNDARY_TYPE_THRESHOLD"]["type"] == "float"
    assert by_name["SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_ENABLED"]["type"] == "bool"


def test_settings_put_persists_boundary_knobs(client, tmp_path, monkeypatch):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    resp = client.put(
        "/studio/api/settings",
        json={
            "settings": {
                "SCREENSPACE_BOUNDARY_MERGE_THRESHOLD": 0.3,
                "SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_ENABLED": False,
            }
        },
    )
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True
    assert server.config.SCREENSPACE_BOUNDARY_MERGE_THRESHOLD == 0.3
    assert server.config.SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_ENABLED is False


def test_settings_put_persists_friction_model(client, tmp_path, monkeypatch):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server.config, "LLM_FRICTION_MODEL", "")
    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"LLM_FRICTION_MODEL": "gemma4:latest"}},
    )
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True
    assert server.config.LLM_FRICTION_MODEL == "gemma4:latest"


def test_settings_put_persists_card_color(client, tmp_path, monkeypatch):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    # Baseline via monkeypatch so the PUT's direct setattr is restored on teardown.
    monkeypatch.setattr(server.config, "TITLECARD_COLOR", "#000000")
    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"TITLECARD_COLOR": "#ff8800"}},
    )
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True
    assert server.config.TITLECARD_COLOR == "#ff8800"


def test_settings_partial_put_preserves_other_settings(client, tmp_path, monkeypatch):
    """A partial PUT must not drop unrelated non-default settings already saved."""
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    # A previously-selected card image (non-default) is live on config.
    monkeypatch.setattr(server.config, "TITLECARD_IMAGE", "card.png")

    # Submit only an inline titlecard key (mirrors the studio.js partial PUT;
    # duration has no ffmpeg-support gate so the test stays dependency-free).
    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"TITLECARD_DURATION_SECONDS": 5}},
    )
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True

    saved = server.start_settings.load_config_json(
        server.config.STUDIO_SETTINGS_FILENAME, default={}
    )
    assert saved.get("TITLECARD_IMAGE") == "card.png"  # preserved, not dropped
    assert saved.get("TITLECARD_DURATION_SECONDS") == 5


def test_settings_put_rejects_card_image_traversal(client, tmp_path, monkeypatch):
    """A card image setting that escapes the upload pool is rejected, not stored."""
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server.config, "TITLECARD_IMAGE", "")
    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"TITLECARD_IMAGE": "../secret.png"}},
    )
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False
    assert server.config.TITLECARD_IMAGE == ""


def test_settings_rejects_non_hex_card_color(client, tmp_path, monkeypatch):
    """Card colors feed ffmpeg's lavfi color input — only #rrggbb is accepted."""
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server.config, "TITLECARD_COLOR", "#000000")
    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"TITLECARD_COLOR": "red; drawbox"}},
    )
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False
    assert server.config.TITLECARD_COLOR == "#000000"  # unchanged


def test_settings_put_endcard_accepts_none_title_rejects_none(
    client, tmp_path, monkeypatch
):
    """`__none__` is valid only for endcards; titlecards always render."""
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server.config, "ENDCARD_IMAGE", "")
    monkeypatch.setattr(server.config, "TITLECARD_IMAGE", "")

    ok = client.put(
        "/studio/api/settings",
        json={"settings": {"ENDCARD_IMAGE": server.config.CARD_IMAGE_NONE}},
    )
    assert ok.status_code == 200
    assert server.config.ENDCARD_IMAGE == server.config.CARD_IMAGE_NONE

    bad = client.put(
        "/studio/api/settings",
        json={"settings": {"TITLECARD_IMAGE": server.config.CARD_IMAGE_NONE}},
    )
    assert bad.status_code == 400
    assert server.config.TITLECARD_IMAGE == ""


def test_settings_put_card_image_validates_against_pool(client, tmp_path, monkeypatch):
    """An existing uploaded basename is accepted; a missing one is rejected."""
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server.config, "TITLECARD_IMAGE", "")
    images = tmp_path / server.config.TITLECARD_IMAGES_DIRNAME
    images.mkdir()
    (images / "card.png").write_bytes(b"x")

    ok = client.put(
        "/studio/api/settings",
        json={"settings": {"TITLECARD_IMAGE": "card.png"}},
    )
    assert ok.status_code == 200
    assert server.config.TITLECARD_IMAGE == "card.png"

    missing = client.put(
        "/studio/api/settings",
        json={"settings": {"TITLECARD_IMAGE": "ghost.png"}},
    )
    assert missing.status_code == 400
    # A rejected value must not overwrite the prior good selection.
    assert server.config.TITLECARD_IMAGE == "card.png"


def test_api_sheet_baseline_returns_empty_when_no_sheet(client):
    resp = client.get("/studio/api/sheet/baseline")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["sheet_loaded"] is False
    assert data["baselines"] == {}


def test_api_thumbnail_returns_404_when_no_sheet(client):
    resp = client.get("/studio/api/thumbnail/P01/0")
    assert resp.status_code == 404
    data = resp.get_json()
    assert data["ok"] is False


def test_api_thumbnail_accepts_fractional_timestamp(client, monkeypatch, tmp_path):
    """A fractional second must not 400 (bare int('12.5') would); it floors to
    the second-granular thumbnail."""
    import types

    import video

    monkeypatch.setattr(server, "_sheet_context", types.SimpleNamespace())
    source = tmp_path / "study_P01.mp4"
    source.write_bytes(b"x")
    monkeypatch.setattr(server, "_resolve_participant_sources", lambda _p: [source])

    captured = {}

    def fake_extract(path, cut_sec, *, width):
        captured["cut_sec"] = cut_sec
        return b"jpegbytes"

    monkeypatch.setattr(video, "extract_thumbnail_bytes", fake_extract)

    resp = client.get("/studio/api/thumbnail/P01/12.5")
    assert resp.status_code == 200
    assert resp.mimetype == "image/jpeg"
    assert captured["cut_sec"] == 12


@pytest.mark.parametrize(
    "path,payload,context_attr,expected_err",
    [
        ("/studio/api/generate", {"cells": []}, "_worksheet", "No cells"),
        (
            "/studio/api/generate",
            {"cells": ["P01.3"], "format": "pdf"},
            "_worksheet",
            "Invalid format",
        ),
        ("/studio/api/reel", {"cells": []}, "_worksheet", "No cells"),
        (
            "/studio/api/gallery",
            {"participant": "", "format": "screen"},
            "_sheet_context",
            "No participant",
        ),
        (
            "/studio/api/gallery",
            {"participant": "P01", "format": "clip"},
            "_sheet_context",
            "Invalid format",
        ),
    ],
    ids=[
        "generate_no_cells",
        "generate_bad_format",
        "reel_no_cells",
        "gallery_no_participant",
        "gallery_bad_format",
    ],
)
def test_api_returns_400_for_invalid_input(
    client, monkeypatch, path, payload, context_attr, expected_err
):
    monkeypatch.setattr(server, context_attr, object())
    resp = client.post(path, json=payload)
    assert resp.status_code == 400
    data = resp.get_json()
    assert expected_err in data["error"]


def test_api_sheet_refresh_success(client, monkeypatch):
    import spreadsheet

    fake_context = object()
    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr(spreadsheet, "build_sheet_context", lambda ws: fake_context)

    resp = client.post("/studio/api/sheet/refresh")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert server._sheet_context is fake_context


def test_api_sheet_refresh_500_when_build_fails(client, monkeypatch):
    import spreadsheet

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr(spreadsheet, "build_sheet_context", lambda ws: None)

    resp = client.post("/studio/api/sheet/refresh")
    assert resp.status_code == 500
    data = resp.get_json()
    assert data["ok"] is False
    assert "Failed to refresh" in data["error"]


def test_api_sheet_baseline_no_baseline_row(client, monkeypatch):
    import types

    ctx = types.SimpleNamespace(
        header_row=["ID", "P01", "P02"],
        id_cell=types.SimpleNamespace(col=1),
        num_participants=2,
        baseline_row_idx=None,
        sheet_data=[["study"], ["ID", "P01", "P02"]],
    )
    monkeypatch.setattr(server, "_sheet_context", ctx)

    resp = client.get("/studio/api/sheet/baseline")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["baselines"] == {}


def test_api_sheet_baseline_with_values(client, monkeypatch):
    import types

    ctx = types.SimpleNamespace(
        header_row=["ID", "P01", "P02"],
        id_cell=types.SimpleNamespace(col=1),
        num_participants=2,
        baseline_row_idx=2,
        sheet_data=[
            ["study"],
            ["ID", "P01", "P02"],
            ["Baseline time", "09:12:00", "09:15:30"],
        ],
    )
    monkeypatch.setattr(server, "_sheet_context", ctx)

    resp = client.get("/studio/api/sheet/baseline")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["baselines"] == {"P01": 33120, "P02": 33330}


def test_api_sheet_baseline_partial(client, monkeypatch):
    import types

    ctx = types.SimpleNamespace(
        header_row=["ID", "P01", "P02", "P03"],
        id_cell=types.SimpleNamespace(col=1),
        num_participants=3,
        baseline_row_idx=2,
        sheet_data=[
            ["study"],
            ["ID", "P01", "P02", "P03"],
            ["Baseline time", "09:12:00", "", "09:20:00"],
        ],
    )
    monkeypatch.setattr(server, "_sheet_context", ctx)

    resp = client.get("/studio/api/sheet/baseline")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["baselines"]["P01"] == 33120
    assert data["baselines"]["P02"] == 0
    assert data["baselines"]["P03"] == 33600


def test_api_reel_highlights_duration_override(client, monkeypatch):
    """highlights_duration temporarily overrides config and is restored after."""
    import config

    monkeypatch.setattr(server, "_worksheet", object())
    original = config.HIGHLIGHTS_REEL_DURATION_SECONDS
    captured = {}

    def fake_generate_list(ws, mode, *, ctx=None, reel_input, skip_prompts):
        captured["duration"] = config.HIGHLIGHTS_REEL_DURATION_SECONDS
        return []

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)

    resp = client.post(
        "/studio/api/reel",
        json={"cells": ["highlights", "batch"], "highlights_duration": 120},
    )
    # Reel route is now NDJSON-streaming, so empty-clips errors arrive as a
    # single line in the response body rather than an HTTP 400.
    assert resp.status_code == 200
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert lines[-1]["ok"] is False
    assert captured["duration"] == 120
    assert config.HIGHLIGHTS_REEL_DURATION_SECONDS == original


def test_api_reel_highlights_duration_restored_on_error(client, monkeypatch):
    """Config is restored even if generate_list raises."""
    import config

    monkeypatch.setattr(server, "_worksheet", object())
    original = config.HIGHLIGHTS_REEL_DURATION_SECONDS

    def raise_generate_list(ws, mode, *, ctx=None, reel_input, skip_prompts):
        raise RuntimeError("boom")

    monkeypatch.setattr("spreadsheet.generate_list", raise_generate_list)

    resp = client.post(
        "/studio/api/reel",
        json={"cells": ["highlights"], "highlights_duration": 999},
    )
    # Exceptions inside the streaming generator are caught and surfaced as a
    # single JSON line; status stays at 200 because headers were already sent.
    assert resp.status_code == 200
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert lines[-1]["ok"] is False
    assert "boom" in lines[-1].get("error", "")
    assert config.HIGHLIGHTS_REEL_DURATION_SECONDS == original


def test_api_thumbnail_returns_jpeg(client, monkeypatch, tmp_path):
    import types

    import video

    fake_jpeg = b"\xff\xd8\xff\xe0fake-jpeg-data"
    dummy_video = tmp_path / "study_P01.mp4"
    dummy_video.write_bytes(b"not a real video")

    ctx = types.SimpleNamespace(
        header_row=["ID", "P01"],
        id_cell=types.SimpleNamespace(col=1),
        num_participants=1,
        study_name="study",
        filename_row_idx=None,
        sheet_data=[],
    )
    monkeypatch.setattr(server, "_sheet_context", ctx)
    monkeypatch.setattr(
        server, "_thumbnail_cache", server._MediaCache(server._THUMBNAIL_CACHE_MAX)
    )
    monkeypatch.setattr("config.INPUT_DIR", str(tmp_path))
    monkeypatch.setattr(video, "extract_thumbnail_bytes", lambda *a, **kw: fake_jpeg)

    resp = client.get("/studio/api/thumbnail/P01/10")
    assert resp.status_code == 200
    assert resp.content_type == "image/jpeg"
    assert resp.data == fake_jpeg


def test_api_thumbnail_caches(client, monkeypatch, tmp_path):
    import types

    import video

    call_count = [0]
    fake_jpeg = b"\xff\xd8\xff\xe0cached"
    dummy_video = tmp_path / "study_P01.mp4"
    dummy_video.write_bytes(b"x")

    ctx = types.SimpleNamespace(
        header_row=["ID", "P01"],
        id_cell=types.SimpleNamespace(col=1),
        num_participants=1,
        study_name="study",
        filename_row_idx=None,
        sheet_data=[],
    )
    monkeypatch.setattr(server, "_sheet_context", ctx)
    monkeypatch.setattr(
        server, "_thumbnail_cache", server._MediaCache(server._THUMBNAIL_CACHE_MAX)
    )
    monkeypatch.setattr("config.INPUT_DIR", str(tmp_path))

    def counting_extract(*a, **kw):
        call_count[0] += 1
        return fake_jpeg

    monkeypatch.setattr(video, "extract_thumbnail_bytes", counting_extract)

    resp1 = client.get("/studio/api/thumbnail/P01/5")
    resp2 = client.get("/studio/api/thumbnail/P01/5")
    assert resp1.status_code == 200
    assert resp2.status_code == 200
    assert call_count[0] == 1


def test_api_manifest_get_returns_artifacts(client, monkeypatch):
    import viewer

    fake_artifacts = [
        {"id": "a5c2s0", "type": "clip", "participant": "P01", "cellRow": 5}
    ]
    monkeypatch.setattr(viewer, "load_manifest_both", lambda: (fake_artifacts, []))
    resp = client.get("/studio/api/manifest")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert len(data["artifacts"]) == 1
    assert data["artifacts"][0]["id"] == "a5c2s0"
    assert data["reels"] == []


def test_api_manifest_get_empty(client, monkeypatch):
    import viewer

    monkeypatch.setattr(viewer, "load_manifest_both", lambda: ([], []))
    resp = client.get("/studio/api/manifest")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["artifacts"] == []
    assert data["reels"] == []


def test_api_manifest_post_still_works(client, monkeypatch):
    _set_artifacts(monkeypatch, [])
    monkeypatch.setattr(server, "_generated_reels", [])
    resp = client.post("/studio/api/manifest")
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False
    assert "No artifacts" in data["error"]


def test_api_generate_skips_existing_artifacts(client, monkeypatch, tmp_path):
    """Already-generated artifacts are returned without re-running process_clips."""
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))

    # Create the artifact file on disk
    (tmp_path / "clip.mp4").write_bytes(b"video")

    existing = [
        {"id": "a5c2s0", "type": "clip", "file": "clip.mp4", "cellRow": 5, "cellCol": 2}
    ]
    _set_artifacts(monkeypatch, list(existing))

    cell = types.SimpleNamespace(row=5, col=2, value="1:00")

    def fake_generate_list(ws, mode, *, ctx=None, cell_specs, skip_prompts):
        return [{"participant": "P01", "cell": cell}]

    def fake_parse_cell_specs(text):
        return [("P01", 5)]

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)
    monkeypatch.setattr("spreadsheet.parse_cell_specifications", fake_parse_cell_specs)

    process_called = []
    monkeypatch.setattr(
        "pipeline.process_clips",
        lambda *a, **kw: process_called.append(1) or (1, []),
    )

    resp = client.post(
        "/studio/api/generate", json={"cells": ["P01.5"], "format": "clip"}
    )
    assert resp.status_code == 200
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert lines[0]["ok"] is True
    assert lines[0]["skipped"] is True
    assert lines[0]["artifacts"] == existing
    assert process_called == []


def test_api_generate_regenerates_when_file_missing(client, monkeypatch, tmp_path):
    """If artifact file is missing from disk, regeneration proceeds normally."""
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))

    # Artifact record exists but file does NOT
    stale = [
        {"id": "a5c2s0", "type": "clip", "file": "gone.mp4", "cellRow": 5, "cellCol": 2}
    ]
    _set_artifacts(monkeypatch, list(stale))

    cell = types.SimpleNamespace(row=5, col=2, value="1:00")

    def fake_generate_list(ws, mode, *, ctx=None, cell_specs, skip_prompts):
        return [{"participant": "P01", "cell": cell}]

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)
    monkeypatch.setattr("spreadsheet.parse_cell_specifications", lambda t: [("P01", 5)])

    new_artifact = {
        "id": "a5c2s0",
        "type": "clip",
        "file": "new.mp4",
        "cellRow": 5,
        "cellCol": 2,
    }
    monkeypatch.setattr(
        "pipeline.process_clips",
        lambda *a, **kw: (1, [new_artifact]),
    )

    resp = client.post(
        "/studio/api/generate", json={"cells": ["P01.5"], "format": "clip"}
    )
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert lines[0]["ok"] is True
    assert "skipped" not in lines[0]
    assert lines[0]["artifacts"] == [new_artifact]


def test_api_generate_regenerates_when_titlecards_toggled(
    client, monkeypatch, tmp_path
):
    """Toggling titlecards on regenerates a cached no-titlecard clip and discards
    the stale record + file."""
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))

    # Existing clip generated WITHOUT titlecards; file present on disk.
    (tmp_path / "clip.mp4").write_bytes(b"video")
    existing = [
        {
            "id": "a5c2s0",
            "type": "clip",
            "file": "clip.mp4",
            "cellRow": 5,
            "cellCol": 2,
            "titlecards": False,
            "titlecardDuration": 0,
        }
    ]
    _set_artifacts(monkeypatch, list(existing))

    cell = types.SimpleNamespace(row=5, col=2, value="1:00")
    monkeypatch.setattr(
        "spreadsheet.generate_list",
        lambda ws, mode, *, ctx=None, cell_specs, skip_prompts: [
            {"participant": "P01", "cell": cell}
        ],
    )
    monkeypatch.setattr("spreadsheet.parse_cell_specifications", lambda t: [("P01", 5)])

    process_called = []
    new_artifact = {
        "id": "a5c2s0",
        "type": "clip",
        "file": "clip.mp4",
        "cellRow": 5,
        "cellCol": 2,
        "titlecards": True,
        "titlecardDuration": 2,
    }
    monkeypatch.setattr(
        "pipeline.process_clips",
        lambda *a, **kw: process_called.append(kw) or (1, [new_artifact]),
    )

    resp = client.post(
        "/studio/api/generate",
        json={"cells": ["P01.5"], "format": "clip", "titlecards_enabled": True},
    )
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert process_called and process_called[0]["titlecards_enabled"] is True
    assert "skipped" not in lines[0]
    assert lines[0]["artifacts"] == [new_artifact]
    # The stale no-titlecard file was discarded.
    assert not (tmp_path / "clip.mp4").exists()


def test_api_generate_skips_when_titlecards_match(client, monkeypatch, tmp_path):
    """A cached clip whose titlecard state matches the request is reused."""
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr("config.TITLECARD_DURATION_SECONDS", 2)

    (tmp_path / "clip.mp4").write_bytes(b"video")
    existing = [
        {
            "id": "a5c2s0",
            "type": "clip",
            "file": "clip.mp4",
            "cellRow": 5,
            "cellCol": 2,
            "titlecards": True,
            "titlecardDuration": 2,
        }
    ]
    _set_artifacts(monkeypatch, list(existing))

    cell = types.SimpleNamespace(row=5, col=2, value="1:00")
    monkeypatch.setattr(
        "spreadsheet.generate_list",
        lambda ws, mode, *, ctx=None, cell_specs, skip_prompts: [
            {"participant": "P01", "cell": cell}
        ],
    )
    monkeypatch.setattr("spreadsheet.parse_cell_specifications", lambda t: [("P01", 5)])

    process_called = []
    monkeypatch.setattr(
        "pipeline.process_clips",
        lambda *a, **kw: process_called.append(1) or (1, []),
    )

    resp = client.post(
        "/studio/api/generate",
        json={"cells": ["P01.5"], "format": "clip", "titlecards_enabled": True},
    )
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert lines[0]["skipped"] is True
    assert lines[0]["artifacts"] == existing
    assert process_called == []
    # The matching cached file is left intact.
    assert (tmp_path / "clip.mp4").exists()


def test_api_reel_regenerates_when_titlecards_toggled(client, monkeypatch, tmp_path):
    """Toggling titlecards on must not reuse a cached reel built without them."""
    import types

    import pipeline

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))

    (tmp_path / "study_reel.mp4").write_bytes(b"old-reel")
    cell = types.SimpleNamespace(row=5, col=2, value="1:00-1:30")
    components = [{"cellRow": 5, "cellCol": 2, "start": 60.0, "end": 90.0}]
    expected_id = pipeline.compute_reel_id(components)

    existing_reel = {
        "id": expected_id,
        "file": "study_reel.mp4",
        "study": "study",
        "components": components,
        "titlecards": False,
        "titlecardDuration": 0,
    }
    monkeypatch.setattr(server, "_generated_reels", [existing_reel])

    def fake_generate_list(ws, mode, *, ctx=None, reel_input, skip_prompts):
        return [
            {
                "participant": "P01",
                "cell": cell,
                "desc": "test",
                "category": "cat",
                "study": "study",
                "severity": "",
                "times": [("1:00", "1:30")],
            }
        ]

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)
    monkeypatch.setattr("files.prepare_clip", lambda clip: clip)

    process_called = []

    def fake_stream(clips, cancel_flag, **kwargs):
        process_called.append(kwargs)
        yield (
            json.dumps({"ok": True, "generated": 1, "reels": [{"id": expected_id}]})
            + "\n"
        )

    monkeypatch.setattr(server, "_stream_process_reel", fake_stream)
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)

    resp = client.post(
        "/studio/api/reel",
        json={"cells": ["P01.5"], "titlecards_enabled": True, "titlecard_duration": 2},
    )
    assert resp.status_code == 200
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert lines[-1]["ok"] is True
    assert process_called
    assert process_called[0]["titlecards_enabled"] is True
    assert not (tmp_path / "study_reel.mp4").exists()


def test_api_reel_skips_existing_reel(client, monkeypatch, tmp_path):
    """An identical reel is returned without re-running process_reel."""
    import types

    import pipeline

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))

    # Create the reel file on disk
    (tmp_path / "study_reel.mp4").write_bytes(b"reel")

    cell = types.SimpleNamespace(row=5, col=2, value="1:00-1:30")

    # Compute expected reel ID using the same function the server will use
    components = [{"cellRow": 5, "cellCol": 2, "start": 60.0, "end": 90.0}]
    expected_id = pipeline.compute_reel_id(components)

    existing_reel = {
        "id": expected_id,
        "file": "study_reel.mp4",
        "study": "study",
        "components": components,
        "titlecards": False,
        "titlecardDuration": 0,
    }
    monkeypatch.setattr(server, "_generated_reels", [existing_reel])

    def fake_generate_list(ws, mode, *, ctx=None, reel_input, skip_prompts):
        return [
            {
                "participant": "P01",
                "cell": cell,
                "desc": "test",
                "category": "cat",
                "study": "study",
                "severity": "",
            }
        ]

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)

    process_called = []
    monkeypatch.setattr(
        "pipeline.process_reel",
        lambda *a, **kw: process_called.append(1) or (1, []),
    )

    resp = client.post(
        "/studio/api/reel",
        json={"cells": ["P01.5"], "titlecards_enabled": False},
    )
    assert resp.status_code == 200
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    final = lines[-1]
    assert final["ok"] is True
    assert final["skipped"] is True
    assert final["reels"] == [existing_reel]
    assert process_called == []


def test_apply_time_overrides_single_and_multi_segment():
    """Overrides replace a cell's whole time list; un-overridden clips untouched."""
    import types

    clips = [
        {"participant": "P01", "cell": types.SimpleNamespace(row=5)},
        {"participant": "P02", "cell": types.SimpleNamespace(row=3)},
        {"participant": "P09", "cell": types.SimpleNamespace(row=1)},
    ]
    server._apply_time_overrides(
        clips,
        {"P01.5": [[10, 70]], "P02.3": [[5, 15], [100, 160]]},
    )
    assert clips[0]["times"] == [("0:10", "1:10")]
    assert clips[1]["times"] == [("0:05", "0:15"), ("1:40", "2:40")]
    assert "times" not in clips[2]


def test_apply_time_overrides_matches_a_ref_case_insensitively():
    """The posted ref and the sheet header only ever match case-insensitively.

    find_participant_column resolves a ref against the header with .lower(), so
    an exact-match lookup here would silently drop the user's trimmed in/out
    points for any ref not spelled exactly like its column.
    """
    import types

    clips = [{"participant": "P01", "cell": types.SimpleNamespace(row=5)}]
    server._apply_time_overrides(clips, {"p01.5": [[10, 70]]})
    assert clips[0]["times"] == [("0:10", "1:10")]


def test_apply_time_overrides_forces_hours_across_hour_boundary():
    """When either endpoint crosses an hour, both render H:MM:SS (no mixed pair)."""
    import types

    clips = [{"participant": "P01", "cell": types.SimpleNamespace(row=7)}]
    server._apply_time_overrides(clips, {"P01.7": [[3590, 3660]]})
    assert clips[0]["times"] == [("0:59:50", "1:01:00")]


def test_apply_time_overrides_skips_invalid_and_empty():
    """Zero/negative-length pairs are dropped; empty overrides are a no-op."""
    import types

    clips = [{"participant": "P01", "cell": types.SimpleNamespace(row=2)}]
    server._apply_time_overrides(clips, {"P01.2": [[50, 50]]})
    assert "times" not in clips[0]
    server._apply_time_overrides(clips, {})
    assert "times" not in clips[0]


def test_api_generate_applies_time_overrides_and_forces_regen(
    client, monkeypatch, tmp_path
):
    """An override replaces the clip times AND forces regeneration even when a
    matching cached artifact exists (cache is keyed only by cell/format)."""
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))

    # A cached clip that would normally be reused (titlecards match the request).
    (tmp_path / "clip.mp4").write_bytes(b"video")
    existing = [
        {
            "id": "a5c2s0",
            "type": "clip",
            "file": "clip.mp4",
            "cellRow": 5,
            "cellCol": 2,
            "titlecards": False,
            "titlecardDuration": 0,
        }
    ]
    _set_artifacts(monkeypatch, list(existing))

    cell = types.SimpleNamespace(row=5, col=2, value="1:00")
    monkeypatch.setattr(
        "spreadsheet.generate_list",
        lambda ws, mode, *, ctx=None, cell_specs, skip_prompts: [
            {"participant": "P01", "cell": cell}
        ],
    )
    monkeypatch.setattr("spreadsheet.parse_cell_specifications", lambda t: [("P01", 5)])

    captured = {}

    def fake_process(clips, **kw):
        captured["times"] = clips[0].get("times")
        return (1, [{"id": "a5c2s0", "type": "clip", "file": "clip.mp4"}])

    monkeypatch.setattr("pipeline.process_clips", fake_process)

    resp = client.post(
        "/studio/api/generate",
        json={
            "cells": ["P01.5"],
            "format": "clip",
            "overrides": {"P01.5": [[10, 70]]},
        },
    )
    assert resp.status_code == 200
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert "skipped" not in lines[0]
    assert captured["times"] == [("0:10", "1:10")]


def test_api_reel_applies_time_overrides(client, monkeypatch, tmp_path):
    """The pure-spreadsheet reel path honours edited in/out points."""
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server, "_generated_reels", [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)

    cell = types.SimpleNamespace(row=5, col=2, value="1:00-1:30")
    monkeypatch.setattr(
        "spreadsheet.generate_list",
        lambda ws, mode, *, ctx=None, reel_input, skip_prompts: [
            {
                "participant": "P01",
                "cell": cell,
                "desc": "t",
                "category": "c",
                "study": "study",
                "severity": "",
            }
        ],
    )
    # Pass-through prepare_clip: the override already set clip["times"].
    monkeypatch.setattr("files.prepare_clip", lambda clip: clip)

    captured = {}

    def fake_stream(clips, cancel_flag, **kwargs):
        captured["times"] = clips[0].get("times")
        yield json.dumps({"ok": True, "generated": 1, "reels": [{"id": "x"}]}) + "\n"

    monkeypatch.setattr(server, "_stream_process_reel", fake_stream)

    resp = client.post(
        "/studio/api/reel",
        json={"cells": ["P01.5"], "overrides": {"P01.5": [[10, 70]]}},
    )
    assert resp.status_code == 200
    assert captured["times"] == [("0:10", "1:10")]


def test_api_gallery_404_when_video_not_found(client, monkeypatch):
    monkeypatch.setattr(server, "_sheet_context", object())
    monkeypatch.setattr(server, "_resolve_participant_sources", lambda p: [])
    resp = client.post(
        "/studio/api/gallery",
        json={"participant": "P01", "format": "screen", "interval": 10},
    )
    assert resp.status_code == 404
    data = resp.get_json()
    assert "not found" in data["error"]


def test_api_viewer_400_when_no_artifacts(client):
    resp = client.post("/studio/api/viewer")
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False
    assert "No artifacts" in data["error"]


def test_api_timeline_viewer_without_intake(client, monkeypatch):
    import pipeline
    import spreadsheet
    import viewer

    fake_ws = type("FakeWS", (), {"title": "Sheet1"})()
    monkeypatch.setattr(server, "_worksheet", fake_ws)

    fake_clips = [{"desc": "test", "participant": "P01"}]
    fake_artifacts = [
        {
            "id": "a1",
            "type": "clip",
            "study": "s",
            "participant": "P01",
            "start": 0,
            "end": 5,
        }
    ]
    monkeypatch.setattr(spreadsheet, "generate_list", lambda *a, **kw: fake_clips)
    monkeypatch.setattr(pipeline, "process_clips", lambda *a, **kw: (1, fake_artifacts))
    monkeypatch.setattr(pipeline, "is_excel_worksheet", lambda ws: False)
    monkeypatch.setattr(viewer, "load_screenspace_events_for_viewer", list)
    monkeypatch.setattr(viewer, "finalize_timeline_data", lambda *a, **kw: {"meta": {}})
    monkeypatch.setattr(
        viewer, "generate_timeline_viewer", lambda *a, **kw: "/out/viewer.html"
    )
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)

    resp = client.post("/studio/api/timeline-viewer", json={})
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["generated"] == 1


def test_api_timeline_viewer_with_intake(client, monkeypatch):
    import pipeline
    import spreadsheet
    import viewer

    fake_ws = type("FakeWS", (), {"title": "Sheet1"})()
    monkeypatch.setattr(server, "_worksheet", fake_ws)

    fake_clips = [{"desc": "test", "participant": "P01"}]
    sheet_artifacts = [
        {
            "id": "a1",
            "type": "clip",
            "study": "s",
            "participant": "P01",
            "start": 0,
            "end": 5,
        }
    ]
    intake_artifact = {
        "id": "intake_abc_s0",
        "type": "clip",
        "source": "screenspace",
        "participant": "P02",
        "start": 10.0,
        "end": 15.0,
        "_ok": True,
        "_error": "",
    }

    monkeypatch.setattr(spreadsheet, "generate_list", lambda *a, **kw: fake_clips)
    monkeypatch.setattr(
        pipeline, "process_clips", lambda *a, **kw: (1, sheet_artifacts)
    )
    monkeypatch.setattr(pipeline, "is_excel_worksheet", lambda ws: False)
    monkeypatch.setattr(viewer, "load_screenspace_events_for_viewer", list)
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)

    captured_artifacts = []

    def fake_finalize(artifacts, **kw):
        captured_artifacts.extend(artifacts)
        return {"meta": {}}

    monkeypatch.setattr(viewer, "finalize_timeline_data", fake_finalize)
    monkeypatch.setattr(
        viewer, "generate_timeline_viewer", lambda *a, **kw: "/out/viewer.html"
    )
    monkeypatch.setattr(
        server, "_generate_intake_clips", lambda items, **kw: [intake_artifact]
    )

    resp = client.post(
        "/studio/api/timeline-viewer",
        json={
            "include_intake": True,
            "intake_items": [
                {
                    "participant": "P02",
                    "start": 10.0,
                    "end": 15.0,
                    "event_type": "text",
                },
            ],
        },
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["generated"] == 2
    assert len(captured_artifacts) == 2
    assert captured_artifacts[1]["source"] == "screenspace"


def test_api_timeline_viewer_cancel_endpoint(client):
    """POST /api/timeline-viewer/cancel sets the cancel event and returns ok."""
    server._timeline_viewer_cancel_event.clear()
    resp = client.post("/studio/api/timeline-viewer/cancel")
    data = resp.get_json()
    assert resp.status_code == 200
    assert data["ok"] is True
    assert server._timeline_viewer_cancel_event.is_set()
    server._timeline_viewer_cancel_event.clear()


def test_api_gallery_cancel_endpoint(client):
    """POST /api/gallery/cancel sets the cancel event and returns ok."""
    server._gallery_cancel_event.clear()
    resp = client.post("/studio/api/gallery/cancel")
    data = resp.get_json()
    assert resp.status_code == 200
    assert data["ok"] is True
    assert server._gallery_cancel_event.is_set()
    server._gallery_cancel_event.clear()


def test_api_timeline_viewer_cancel_discards_clips(client, monkeypatch, tmp_path):
    """A cancel mid-build unlinks clips already written to disk, leaving no orphan
    media (the manifest is never published in that case)."""
    import pipeline
    import spreadsheet
    import utils

    monkeypatch.setattr(server, "_worksheet", type("FakeWS", (), {"title": "S"})())

    clip_file = tmp_path / "clip_a1.mp4"
    clip_file.write_text("video-bytes")
    artifacts = [
        {
            "id": "a1",
            "type": "clip",
            "study": "s",
            "participant": "P01",
            "file": "clip_a1.mp4",
            "start": 0,
            "end": 5,
        }
    ]

    monkeypatch.setattr(spreadsheet, "generate_list", lambda *a, **kw: [{"desc": "x"}])
    monkeypatch.setattr(utils, "resolve_output_path", lambda name: tmp_path / name)

    def fake_process(clips, **kwargs):
        # User clicks Cancel while ffmpeg is still writing clips.
        server._timeline_viewer_cancel_event.set()
        return (1, artifacts)

    monkeypatch.setattr(pipeline, "process_clips", fake_process)

    resp = client.post("/studio/api/timeline-viewer", json={})
    assert resp.status_code == 200
    assert resp.get_json()["cancelled"] is True
    assert not clip_file.exists()
    server._timeline_viewer_cancel_event.clear()


def test_api_gallery_cancel_discards_captures(client, monkeypatch, tmp_path):
    """A cancel mid-build unlinks gallery captures already written to disk."""
    import utils
    import video

    monkeypatch.setattr(server, "_sheet_context", object())
    (tmp_path / "v.mp4").write_text("v")
    cap_file = tmp_path / "gallery_0_05.png"
    cap_file.write_text("img")
    artifacts = [{"file": "gallery_0_05.png", "timestamp": 5.0, "type": "screen"}]

    monkeypatch.setattr(
        server, "_resolve_participant_sources", lambda pid: [tmp_path / "v.mp4"]
    )
    monkeypatch.setattr(utils, "resolve_output_path", lambda name: tmp_path / name)
    monkeypatch.setattr(video, "timeline_or_none", lambda paths: None)
    monkeypatch.setattr(video, "get_file_duration", lambda p: 10)

    def fake_captures(path, **kwargs):
        server._gallery_cancel_event.set()
        return artifacts

    monkeypatch.setattr(video, "generate_interval_captures", fake_captures)

    resp = client.post("/studio/api/gallery", json={"participant": "P01"})
    assert resp.status_code == 200
    assert resp.get_json()["cancelled"] is True
    assert not cap_file.exists()
    server._gallery_cancel_event.clear()


def test_api_timeline_viewer_passes_cancel_flag(client, monkeypatch):
    """process_clips must receive a callable cancel_flag so the build is
    interruptible mid-encode."""
    import pipeline
    import spreadsheet
    import viewer

    fake_ws = type("FakeWS", (), {"title": "Sheet1"})()
    monkeypatch.setattr(server, "_worksheet", fake_ws)

    captured = {}

    def fake_process_clips(clips, **kw):
        captured["cancel_flag"] = kw.get("cancel_flag")
        return (1, [{"id": "a1", "type": "clip", "study": "s", "participant": "P01"}])

    monkeypatch.setattr(spreadsheet, "generate_list", lambda *a, **kw: [{"desc": "x"}])
    monkeypatch.setattr(pipeline, "process_clips", fake_process_clips)
    monkeypatch.setattr(pipeline, "is_excel_worksheet", lambda ws: False)
    monkeypatch.setattr(viewer, "load_screenspace_events_for_viewer", list)
    monkeypatch.setattr(viewer, "finalize_timeline_data", lambda *a, **kw: {"meta": {}})
    monkeypatch.setattr(
        viewer, "generate_timeline_viewer", lambda *a, **kw: "/out/viewer.html"
    )
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)

    server._timeline_viewer_cancel_event.clear()
    resp = client.post("/studio/api/timeline-viewer", json={})
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True
    assert callable(captured["cancel_flag"])


def test_api_timeline_viewer_short_circuits_after_cancel(client, monkeypatch):
    """When cancel is signaled during process_clips, the handler returns
    {cancelled: true} and never writes the viewer HTML."""
    import pipeline
    import spreadsheet
    import viewer

    fake_ws = type("FakeWS", (), {"title": "Sheet1"})()
    monkeypatch.setattr(server, "_worksheet", fake_ws)

    def fake_process_clips(clips, **kw):
        server._timeline_viewer_cancel_event.set()
        return (1, [{"id": "a1", "type": "clip", "study": "s", "participant": "P01"}])

    generated_calls = []

    monkeypatch.setattr(spreadsheet, "generate_list", lambda *a, **kw: [{"desc": "x"}])
    monkeypatch.setattr(pipeline, "process_clips", fake_process_clips)
    monkeypatch.setattr(pipeline, "is_excel_worksheet", lambda ws: False)
    monkeypatch.setattr(
        viewer,
        "generate_timeline_viewer",
        lambda *a, **kw: generated_calls.append(True) or "/out/viewer.html",
    )

    server._timeline_viewer_cancel_event.clear()
    try:
        resp = client.post("/studio/api/timeline-viewer", json={})
    finally:
        server._timeline_viewer_cancel_event.clear()

    assert resp.status_code == 200
    data = resp.get_json()
    assert data == {"ok": False, "cancelled": True}
    assert generated_calls == []


def test_api_gallery_short_circuits_after_cancel(client, monkeypatch, tmp_path):
    """When cancel is signaled during capture extraction, the handler returns
    {cancelled: true} and never writes the gallery HTML."""
    import video
    import viewer

    vid = tmp_path / "video.mp4"
    vid.write_bytes(b"x")
    monkeypatch.setattr(server, "_sheet_context", object())
    monkeypatch.setattr(server, "_resolve_participant_sources", lambda p: [vid])

    def fake_captures(*a, **kw):
        server._gallery_cancel_event.set()
        return [{"file": "f.png", "timestamp": 0.0, "type": "screen"}]

    generated_calls = []

    monkeypatch.setattr(video, "generate_interval_captures", fake_captures)
    monkeypatch.setattr(
        viewer,
        "generate_gallery_viewer",
        lambda *a, **kw: generated_calls.append(True) or "/out/gallery.html",
    )

    server._gallery_cancel_event.clear()
    try:
        resp = client.post(
            "/studio/api/gallery",
            json={"participant": "P01", "format": "screen", "interval": 10},
        )
    finally:
        server._gallery_cancel_event.clear()

    assert resp.status_code == 200
    data = resp.get_json()
    assert data == {"ok": False, "cancelled": True}
    assert generated_calls == []


def test_api_timeline_viewer_discards_sheet_clips_on_cancel_during_intake(
    client, monkeypatch
):
    """A cancel during intake discards the already-generated sheet clips (they
    never enter _generated_artifacts) and writes no viewer HTML."""
    import pipeline
    import spreadsheet
    import viewer

    fake_ws = type("FakeWS", (), {"title": "Sheet1"})()
    monkeypatch.setattr(server, "_worksheet", fake_ws)
    monkeypatch.setattr(server, "_generated_artifacts", [])

    sheet_artifacts = [{"id": "a1", "type": "clip", "study": "s", "participant": "P01"}]
    monkeypatch.setattr(spreadsheet, "generate_list", lambda *a, **kw: [{"desc": "x"}])
    monkeypatch.setattr(
        pipeline, "process_clips", lambda *a, **kw: (1, sheet_artifacts)
    )
    monkeypatch.setattr(pipeline, "is_excel_worksheet", lambda ws: False)

    def fake_intake(items, **kw):
        # Cancel arrives while the intake clips are being generated.
        server._timeline_viewer_cancel_event.set()
        return []

    generated_calls = []
    monkeypatch.setattr(server, "_generate_intake_clips", fake_intake)
    monkeypatch.setattr(
        viewer,
        "generate_timeline_viewer",
        lambda *a, **kw: generated_calls.append(True) or "/out/viewer.html",
    )

    server._timeline_viewer_cancel_event.clear()
    try:
        resp = client.post(
            "/studio/api/timeline-viewer",
            json={
                "include_intake": True,
                "intake_items": [{"participant": "P02", "start": 1, "end": 2}],
            },
        )
    finally:
        server._timeline_viewer_cancel_event.clear()

    assert resp.status_code == 200
    assert resp.get_json() == {"ok": False, "cancelled": True}
    assert generated_calls == []
    assert server._generated_artifacts == []


def test_api_gallery_short_circuits_when_cancelled_before_finalize(
    client, monkeypatch, tmp_path
):
    """A cancel after capture extraction (during the finalize window) still
    prevents the gallery HTML from being written."""
    import video
    import viewer

    vid = tmp_path / "video.mp4"
    vid.write_bytes(b"x")
    monkeypatch.setattr(server, "_sheet_context", object())
    monkeypatch.setattr(server, "_resolve_participant_sources", lambda p: [vid])
    monkeypatch.setattr(
        video,
        "generate_interval_captures",
        lambda *a, **kw: [{"file": "f.png", "timestamp": 0.0, "type": "screen"}],
    )
    monkeypatch.setattr(video, "get_file_duration", lambda *a, **kw: 30)

    def fake_finalize(*a, **kw):
        # Cancel arrives after extraction, while finalizing.
        server._gallery_cancel_event.set()
        return {"meta": {}}

    generated_calls = []
    monkeypatch.setattr(viewer, "finalize_gallery_data", fake_finalize)
    monkeypatch.setattr(
        viewer,
        "generate_gallery_viewer",
        lambda *a, **kw: generated_calls.append(True) or "/out/gallery.html",
    )

    server._gallery_cancel_event.clear()
    try:
        resp = client.post(
            "/studio/api/gallery",
            json={"participant": "P01", "format": "screen", "interval": 10},
        )
    finally:
        server._gallery_cancel_event.clear()

    assert resp.status_code == 200
    assert resp.get_json() == {"ok": False, "cancelled": True}
    assert generated_calls == []


def test_api_gallery_multi_video_captures_all_parts_with_global_times(
    client, monkeypatch, tmp_path
):
    """A multi-video participant's gallery spans every part with global times."""
    import video
    import viewer

    a = tmp_path / "a.mp4"
    b = tmp_path / "b.mp4"
    a.write_bytes(b"x")
    b.write_bytes(b"x")
    monkeypatch.setattr(server, "_sheet_context", object())
    monkeypatch.setattr(server, "_resolve_participant_sources", lambda p: [a, b])
    monkeypatch.setattr(
        video,
        "timeline_or_none",
        lambda paths: [(str(a), 80, 0), (str(b), 120, 80)],
    )

    calls = []

    def fake_captures(path, **kw):
        calls.append(path)
        return [{"file": "f.png", "timestamp": 0.0, "type": "screen"}]

    monkeypatch.setattr(video, "generate_interval_captures", fake_captures)

    captured = {}

    def fake_finalize(artifacts, **kw):
        captured["artifacts"] = artifacts
        captured["duration"] = kw.get("video_duration")
        return {"meta": {}}

    monkeypatch.setattr(viewer, "finalize_gallery_data", fake_finalize)
    monkeypatch.setattr(
        viewer, "generate_gallery_viewer", lambda *a, **kw: "/out/gallery.html"
    )

    server._gallery_cancel_event.clear()
    resp = client.post(
        "/studio/api/gallery",
        json={"participant": "P01", "format": "screen", "interval": 10},
    )

    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True
    assert calls == [str(a), str(b)]  # captured each part
    # Part 2's captures are shifted by its cumulative start (80s).
    timestamps = sorted(art["timestamp"] for art in captured["artifacts"])
    assert timestamps == [0.0, 80.0]
    assert captured["duration"] == 200  # total timeline duration


def test_api_gallery_multi_video_intervals_globally_aligned(
    client, monkeypatch, tmp_path
):
    """A part boundary keeps the global capture grid evenly spaced.

    A first part whose duration isn't a multiple of the interval must not push
    the next part's grid off the global cadence (the pre-fix behavior restarted
    each part's timestamps at 0 then shifted by the cumulative start).
    """
    import video
    import viewer

    a = tmp_path / "a.mp4"
    b = tmp_path / "b.mp4"
    a.write_bytes(b"x")
    b.write_bytes(b"x")
    monkeypatch.setattr(server, "_sheet_context", object())
    monkeypatch.setattr(server, "_resolve_participant_sources", lambda p: [a, b])
    # Part 1 is 95s (not a multiple of the 10s interval); part 2 starts at 95.
    monkeypatch.setattr(
        video,
        "timeline_or_none",
        lambda paths: [(str(a), 95, 0), (str(b), 120, 95)],
    )

    requested: dict[str, list[int]] = {}

    def fake_captures(path, *, timestamps=None, **kw):
        requested[path] = list(timestamps or [])
        return [
            {"file": f"{path}-{ts}.png", "timestamp": float(ts), "type": "screen"}
            for ts in (timestamps or [])
        ]

    monkeypatch.setattr(video, "generate_interval_captures", fake_captures)

    captured = {}

    def fake_finalize(artifacts, **kw):
        captured["artifacts"] = artifacts
        return {"meta": {}}

    monkeypatch.setattr(viewer, "finalize_gallery_data", fake_finalize)
    monkeypatch.setattr(
        viewer, "generate_gallery_viewer", lambda *a, **kw: "/out/gallery.html"
    )

    server._gallery_cancel_event.clear()
    resp = client.post(
        "/studio/api/gallery",
        json={"participant": "P01", "format": "screen", "interval": 10},
    )

    assert resp.status_code == 200
    # Each part is asked for an interval-aligned local grid.
    assert requested[str(a)] == list(range(0, 95, 10))  # 0,10,...,90
    assert requested[str(b)] == list(range(5, 120, 10))  # 5,15,...,115
    # Global timestamps are evenly spaced by the interval across the boundary.
    globals_ = sorted(art["timestamp"] for art in captured["artifacts"])
    diffs = {round(hi - lo, 6) for lo, hi in itertools.pairwise(globals_)}
    assert diffs == {10.0}


def test_api_timeline_viewer_returns_409_when_busy(client, monkeypatch):
    """A second timeline-viewer build is rejected while one holds the slot, so
    the two builds can't clobber the shared cancel event."""
    fake_ws = type("FakeWS", (), {"title": "Sheet1"})()
    monkeypatch.setattr(server, "_worksheet", fake_ws)
    assert server._try_claim_busy("timeline_viewer")
    try:
        resp = client.post("/studio/api/timeline-viewer", json={})
    finally:
        server._release_busy("timeline_viewer")
    assert resp.status_code == 409
    assert resp.get_json()["ok"] is False


def test_api_gallery_returns_409_when_busy(client, monkeypatch):
    """A second gallery build is rejected while one holds the slot."""
    monkeypatch.setattr(server, "_sheet_context", object())
    assert server._try_claim_busy("gallery")
    try:
        resp = client.post(
            "/studio/api/gallery",
            json={"participant": "P01", "format": "screen", "interval": 10},
        )
    finally:
        server._release_busy("gallery")
    assert resp.status_code == 409
    assert resp.get_json()["ok"] is False


# ---- Stash API tests ----


def test_api_stashes_get_empty(client, monkeypatch):
    monkeypatch.setattr(server, "_load_stashes", list)
    resp = client.get("/studio/api/stashes")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["stashes"] == []


def test_api_stashes_create(client, monkeypatch):
    saved = []
    monkeypatch.setattr(server, "_load_stashes", list)
    monkeypatch.setattr(server, "_save_stashes", lambda s: saved.append(s))

    items = [
        {"participant": "P01", "row": 5, "segDuration": 30},
        {"participant": "P02", "row": 7, "segDuration": 45},
    ]
    resp = client.post(
        "/studio/api/stashes",
        json={"action": "create", "items": items, "name": "My reel"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    stash = data["stash"]
    assert stash["name"] == "My reel"
    assert stash["count"] == 2
    assert stash["totalDuration"] == 75
    assert stash["id"].startswith("stash_")
    assert "createdAt" in stash
    assert len(saved) == 1
    assert len(saved[0]) == 1


def test_api_stashes_create_default_name(client, monkeypatch):
    monkeypatch.setattr(server, "_load_stashes", lambda: [{"id": "stash_old"}])
    monkeypatch.setattr(server, "_save_stashes", lambda s: None)

    resp = client.post(
        "/studio/api/stashes",
        json={"action": "create", "items": [{"segDuration": 10}]},
    )
    data = resp.get_json()
    assert data["stash"]["name"] == "Stash 2"


def test_api_stashes_create_empty_items_400(client, monkeypatch):
    monkeypatch.setattr(server, "_load_stashes", list)
    resp = client.post(
        "/studio/api/stashes",
        json={"action": "create", "items": []},
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "No items" in data["error"]


def test_api_stashes_update_name(client, monkeypatch):
    existing = [{"id": "stash_abc", "name": "Old name", "items": [], "count": 0}]
    saved = []
    monkeypatch.setattr(server, "_load_stashes", lambda: list(existing))
    monkeypatch.setattr(server, "_save_stashes", lambda s: saved.append(s))

    resp = client.post(
        "/studio/api/stashes",
        json={"action": "update", "id": "stash_abc", "name": "New name"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["stash"]["name"] == "New name"
    assert saved[0][0]["name"] == "New name"


def test_api_stashes_update_404(client, monkeypatch):
    monkeypatch.setattr(server, "_load_stashes", list)
    resp = client.post(
        "/studio/api/stashes",
        json={"action": "update", "id": "stash_nope", "name": "x"},
    )
    assert resp.status_code == 404


def test_api_stashes_delete(client, monkeypatch):
    existing = [
        {"id": "stash_aaa", "name": "A"},
        {"id": "stash_bbb", "name": "B"},
    ]
    saved = []
    monkeypatch.setattr(server, "_load_stashes", lambda: list(existing))
    monkeypatch.setattr(server, "_save_stashes", lambda s: saved.append(s))

    resp = client.post(
        "/studio/api/stashes",
        json={"action": "delete", "id": "stash_aaa"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert len(saved[0]) == 1
    assert saved[0][0]["id"] == "stash_bbb"


def test_api_stashes_delete_not_found_404(client, monkeypatch):
    monkeypatch.setattr(server, "_load_stashes", list)
    resp = client.post(
        "/studio/api/stashes",
        json={"action": "delete", "id": "stash_nope"},
    )
    assert resp.status_code == 404


def test_api_stashes_unknown_action_400(client, monkeypatch):
    monkeypatch.setattr(server, "_load_stashes", list)
    resp = client.post(
        "/studio/api/stashes",
        json={"action": "bogus"},
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "Unknown action" in data["error"]


# ---- Artifact stash API tests ----


def test_api_artifact_stashes_get_empty(client, monkeypatch):
    monkeypatch.setattr(server, "_load_artifact_stashes", list)
    resp = client.get("/studio/api/artifact-stashes")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["stashes"] == []


def test_api_artifact_stashes_create(client, monkeypatch):
    saved = []
    monkeypatch.setattr(server, "_load_artifact_stashes", list)
    monkeypatch.setattr(server, "_save_artifact_stashes", lambda s: saved.append(s))

    items = [
        {"participant": "P01", "row": 5, "segDuration": 30},
        {"participant": "P02", "row": 7, "segDuration": 45},
    ]
    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "create", "items": items, "name": "My artifacts"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    stash = data["stash"]
    assert stash["name"] == "My artifacts"
    assert stash["count"] == 2
    assert stash["totalDuration"] == 75
    assert stash["id"].startswith("astash_")
    assert "createdAt" in stash
    assert len(saved) == 1
    assert len(saved[0]) == 1


def test_api_artifact_stashes_create_default_name(client, monkeypatch):
    monkeypatch.setattr(
        server, "_load_artifact_stashes", lambda: [{"id": "astash_old"}]
    )
    monkeypatch.setattr(server, "_save_artifact_stashes", lambda s: None)

    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "create", "items": [{"segDuration": 10}]},
    )
    data = resp.get_json()
    assert data["stash"]["name"] == "Stash 2"


def test_api_artifact_stashes_create_empty_items_400(client, monkeypatch):
    monkeypatch.setattr(server, "_load_artifact_stashes", list)
    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "create", "items": []},
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "No items" in data["error"]


def test_api_artifact_stashes_update_name(client, monkeypatch):
    existing = [{"id": "astash_abc", "name": "Old name", "items": [], "count": 0}]
    saved = []
    monkeypatch.setattr(server, "_load_artifact_stashes", lambda: list(existing))
    monkeypatch.setattr(server, "_save_artifact_stashes", lambda s: saved.append(s))

    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "update", "id": "astash_abc", "name": "New name"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["stash"]["name"] == "New name"
    assert saved[0][0]["name"] == "New name"


def test_api_artifact_stashes_update_404(client, monkeypatch):
    monkeypatch.setattr(server, "_load_artifact_stashes", list)
    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "update", "id": "astash_nope", "name": "x"},
    )
    assert resp.status_code == 404


def test_api_artifact_stashes_delete(client, monkeypatch):
    existing = [
        {"id": "astash_aaa", "name": "A"},
        {"id": "astash_bbb", "name": "B"},
    ]
    saved = []
    monkeypatch.setattr(server, "_load_artifact_stashes", lambda: list(existing))
    monkeypatch.setattr(server, "_save_artifact_stashes", lambda s: saved.append(s))

    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "delete", "id": "astash_aaa"},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert len(saved[0]) == 1
    assert saved[0][0]["id"] == "astash_bbb"


def test_api_artifact_stashes_delete_not_found_404(client, monkeypatch):
    monkeypatch.setattr(server, "_load_artifact_stashes", list)
    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "delete", "id": "astash_nope"},
    )
    assert resp.status_code == 404


def test_api_artifact_stashes_unknown_action_400(client, monkeypatch):
    monkeypatch.setattr(server, "_load_artifact_stashes", list)
    resp = client.post(
        "/studio/api/artifact-stashes",
        json={"action": "bogus"},
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert "Unknown action" in data["error"]


# ---- Titlecard settings tests ----


def test_api_sheet_returns_titlecard_defaults(client, monkeypatch):
    """api/sheet includes titlecardsEnabled and titlecardDuration from config."""
    import types

    import config

    ctx = types.SimpleNamespace(
        header_row=["ID", "P01"],
        id_cell=types.SimpleNamespace(row=1, col=1),
        num_participants=1,
        study_name="study",
        observation_cell=types.SimpleNamespace(col=3),
        category_cell=types.SimpleNamespace(col=4),
        severity_cell=None,
        baseline_row_idx=None,
        filename_row_idx=None,
        first_data_row_idx=2,
        sheet_data=[["study"], ["ID", "P01", "Observation", "Category"]],
    )
    monkeypatch.setattr(server, "_sheet_context", ctx)

    resp = client.get("/studio/api/sheet")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["titlecardsEnabled"] == config.TITLECARDS_ENABLED
    assert data["titlecardDuration"] == config.TITLECARD_DURATION_SECONDS


def test_api_sheet_exposes_keyword_annotations(client, monkeypatch):
    """api/sheet must include cell.keywords / row.keywords from !key annotations
    and config.annotations from utils.get_frontend_config()."""
    import types

    # Two data rows: row 2 has a !key annotation in P01, row 3 has none.
    sheet_data = [
        ["ID", "P01", "P02", "Observation", "Category"],
        ["1", "0:10-0:20 !key", "0:30", "obs A", "catA"],
        ["2", "0:40", "", "obs B", "catB"],
    ]
    ctx = types.SimpleNamespace(
        header_row=sheet_data[0],
        id_cell=types.SimpleNamespace(row=1, col=1),
        num_participants=2,
        study_name="study",
        observation_cell=types.SimpleNamespace(col=4),
        category_cell=types.SimpleNamespace(col=5),
        severity_cell=None,
        baseline_row_idx=None,
        filename_row_idx=None,
        first_data_row_idx=1,
        sheet_data=sheet_data,
    )
    monkeypatch.setattr(server, "_sheet_context", ctx)

    resp = client.get("/studio/api/sheet")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True

    # Annotations are surfaced through frontend config so the sidebar can label pills.
    annotations = data["config"]["annotations"]
    assert annotations, "config.annotations must be non-empty"
    assert {"id": "key", "token": "!key"} in annotations

    rows = data["rows"]
    assert len(rows) == 2

    # Row 0 (data row with !key in P01)
    assert rows[0]["keywords"] == ["key"]
    assert rows[0]["cells"]["P01"]["keywords"] == ["key"]
    assert rows[0]["cells"]["P02"]["keywords"] == []

    # Row 1 (no annotation)
    assert rows[1]["keywords"] == []
    assert rows[1]["cells"]["P01"]["keywords"] == []
    assert rows[1]["cells"]["P02"]["keywords"] == []


def test_api_sheet_marks_valid_timestamps_only(client, monkeypatch):
    """api/sheet distinguishes parseable timestamps from other cell text."""
    import types

    sheet_data = [
        ["ID", "P01", "P02", "Observation", "Category"],
        ["1", "1:23-1:45", "N/A", "obs valid", "catA"],
        ["2", "see notes", "0:30", "obs invalid", "catB"],
        ["3", "", "x", "obs empty-ish", "catC"],
    ]
    ctx = types.SimpleNamespace(
        header_row=sheet_data[0],
        id_cell=types.SimpleNamespace(row=1, col=1),
        num_participants=2,
        study_name="study",
        observation_cell=types.SimpleNamespace(col=4),
        category_cell=types.SimpleNamespace(col=5),
        severity_cell=None,
        baseline_row_idx=None,
        filename_row_idx=None,
        first_data_row_idx=1,
        sheet_data=sheet_data,
    )
    monkeypatch.setattr(server, "_sheet_context", ctx)

    resp = client.get("/studio/api/sheet")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True

    row_valid = data["rows"][0]["cells"]
    assert row_valid["P01"]["valid"] is True
    assert row_valid["P01"]["hasText"] is True
    assert row_valid["P02"]["valid"] is False
    assert row_valid["P02"]["hasText"] is True

    row_invalid = data["rows"][1]["cells"]
    assert row_invalid["P01"]["valid"] is False
    assert row_invalid["P01"]["hasText"] is True
    assert row_invalid["P02"]["valid"] is True
    assert row_invalid["P02"]["hasText"] is True

    row_empty = data["rows"][2]["cells"]
    assert row_empty["P01"]["valid"] is False
    assert row_empty["P01"]["hasText"] is False
    assert row_empty["P02"]["valid"] is False
    assert row_empty["P02"]["hasText"] is True


def test_api_sheet_reuses_derived_payload_for_same_context(client, monkeypatch):
    """Repeated /api/sheet calls should not re-parse unchanged sheet cells."""
    import types

    sheet_data = [
        ["ID", "P01", "Observation", "Category"],
        ["1", "0:10-0:20 !key", "obs A", "catA"],
    ]
    ctx = types.SimpleNamespace(
        header_row=sheet_data[0],
        id_cell=types.SimpleNamespace(row=1, col=1),
        num_participants=1,
        study_name="study",
        observation_cell=types.SimpleNamespace(col=3),
        category_cell=types.SimpleNamespace(col=4),
        severity_cell=None,
        baseline_row_idx=None,
        filename_row_idx=None,
        first_data_row_idx=1,
        sheet_data=sheet_data,
    )
    parse_calls = {"annotations": 0, "timestamps": 0}

    def fake_parse_cell_annotations(value):
        parse_calls["annotations"] += 1
        return value.replace(" !key", ""), set(), {"key"}

    def fake_parse_timestamps(value):
        parse_calls["timestamps"] += 1
        return [("0:10", "0:20")]

    monkeypatch.setattr(server, "_sheet_context", ctx)
    monkeypatch.setattr(
        server.utils, "parse_cell_annotations", fake_parse_cell_annotations
    )
    monkeypatch.setattr(server.utils, "parse_timestamps", fake_parse_timestamps)

    first = client.get("/studio/api/sheet").get_json()
    second = client.get("/studio/api/sheet").get_json()

    assert first["rows"] == second["rows"]
    assert parse_calls == {"annotations": 1, "timestamps": 1}


def test_api_sheet_refresh_invalidates_derived_payload(client, monkeypatch):
    """Refreshing the sheet swaps context and rebuilds the cached row payload."""
    import types

    def make_context(timestamp, observation):
        sheet_data = [
            ["ID", "P01", "Observation", "Category"],
            ["1", timestamp, observation, "catA"],
        ]
        return types.SimpleNamespace(
            header_row=sheet_data[0],
            id_cell=types.SimpleNamespace(row=1, col=1),
            num_participants=1,
            study_name="study",
            observation_cell=types.SimpleNamespace(col=3),
            category_cell=types.SimpleNamespace(col=4),
            severity_cell=None,
            baseline_row_idx=None,
            filename_row_idx=None,
            first_data_row_idx=1,
            sheet_data=sheet_data,
        )

    old_ctx = make_context("0:10-0:20", "old obs")
    new_ctx = make_context("0:30-0:40", "new obs")
    parse_calls = {"annotations": 0, "timestamps": 0}

    def fake_parse_cell_annotations(value):
        parse_calls["annotations"] += 1
        return value, set(), set()

    def fake_parse_timestamps(value):
        parse_calls["timestamps"] += 1
        return [tuple(value.split("-", 1))]

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr(server, "_sheet_context", old_ctx)
    monkeypatch.setattr(server.spreadsheet, "build_sheet_context", lambda ws: new_ctx)
    monkeypatch.setattr(
        server.utils, "parse_cell_annotations", fake_parse_cell_annotations
    )
    monkeypatch.setattr(server.utils, "parse_timestamps", fake_parse_timestamps)

    old_data = client.get("/studio/api/sheet").get_json()
    refresh = client.post("/studio/api/sheet/refresh").get_json()
    new_data = client.get("/studio/api/sheet").get_json()

    assert refresh["ok"] is True
    assert old_data["rows"][0]["observation"] == "old obs"
    assert new_data["rows"][0]["observation"] == "new obs"
    assert parse_calls == {"annotations": 2, "timestamps": 2}


def test_swap_worksheet_rollback_clears_sheet_payload_cache(monkeypatch):
    """A failed sheet swap must not leave the attempted sheet payload cached."""
    import types

    import spreadsheet

    prev_ctx = spreadsheet.SheetContext(
        sheet_data=[["ID"]],
        id_cell=types.SimpleNamespace(row=1, col=1),
        observation_cell=types.SimpleNamespace(row=1, col=1),
        category_cell=types.SimpleNamespace(row=1, col=1),
        num_participants=0,
        study_name="previous",
    )
    attempted_ctx = spreadsheet.SheetContext(
        sheet_data=[["ID"]],
        id_cell=types.SimpleNamespace(row=1, col=1),
        observation_cell=types.SimpleNamespace(row=1, col=1),
        category_cell=types.SimpleNamespace(row=1, col=1),
        num_participants=0,
        study_name="attempted",
    )

    def fake_init_studio_state(new_worksheet):
        server._worksheet = new_worksheet
        server._sheet_context = attempted_ctx
        server._sheet_payload_cache = (attempted_ctx, {"participants": [], "rows": []})

    def fail_screenspace_init(**kwargs):
        raise RuntimeError("screenspace failed")

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr(server, "_sheet_context", prev_ctx)
    monkeypatch.setattr(server, "_sheet_payload_cache", (prev_ctx, {}))
    monkeypatch.setattr(server, "_generated_artifacts", [])
    monkeypatch.setattr(server, "_generated_reels", [])
    monkeypatch.setattr(server, "_init_studio_state", fake_init_studio_state)

    import screenspace_server
    import transcripts_server

    monkeypatch.setattr(
        screenspace_server, "_init_screenspace_state", fail_screenspace_init
    )
    monkeypatch.setattr(
        transcripts_server, "_init_transcripts_state", lambda **kw: None
    )

    with pytest.raises(RuntimeError, match="screenspace failed"):
        server._swap_worksheet(object())

    assert server._sheet_context is prev_ctx
    assert server._sheet_payload_cache is None


def test_swap_worksheet_repins_workflows_and_composer(monkeypatch):
    """All five blueprints follow the swap, not just the three with a full init.

    Workflows and Composer used to keep whatever sheet the *process* started
    with — none, on a desktop launch — so a spreadsheet opened from the Start
    overlay never reached a run's NodeContext or Composer's participant list.
    """
    import types

    import composer_server
    import screenspace_server
    import spreadsheet
    import transcripts_server
    import workflows_server

    new_ctx = spreadsheet.SheetContext(
        sheet_data=[["ID"]],
        id_cell=types.SimpleNamespace(row=1, col=1),
        observation_cell=types.SimpleNamespace(row=1, col=1),
        category_cell=types.SimpleNamespace(row=1, col=1),
        num_participants=0,
        study_name="opened",
    )
    new_ws = object()

    def fake_init_studio_state(worksheet):
        server._worksheet = worksheet
        server._sheet_context = new_ctx if worksheet is not None else None

    monkeypatch.setattr(server, "_worksheet", None)
    monkeypatch.setattr(server, "_sheet_context", None)
    monkeypatch.setattr(server, "_generated_artifacts", [])
    monkeypatch.setattr(server, "_generated_reels", [])
    monkeypatch.setattr(server, "_init_studio_state", fake_init_studio_state)
    monkeypatch.setattr(
        screenspace_server, "_init_screenspace_state", lambda **kw: None
    )
    monkeypatch.setattr(
        transcripts_server, "_init_transcripts_state", lambda **kw: None
    )
    monkeypatch.setattr(workflows_server, "_sheet_context", None)
    monkeypatch.setattr(workflows_server, "_worksheet", None)
    monkeypatch.setattr(composer_server, "_sheet_context", None)

    server._swap_worksheet(new_ws)

    assert workflows_server._sheet_context is new_ctx
    assert workflows_server._worksheet is new_ws
    assert composer_server._sheet_context is new_ctx

    # ...and closing has to clear them again, or both keep a dead worksheet.
    server._swap_worksheet(None)

    assert workflows_server._sheet_context is None
    assert workflows_server._worksheet is None
    assert composer_server._sheet_context is None


def test_swap_worksheet_rollback_restores_workflows_and_composer(monkeypatch):
    """A failed swap must leave the sister blueprints on the prior sheet, not
    the half-applied one."""
    import types

    import composer_server
    import screenspace_server
    import spreadsheet
    import workflows_server

    prev_ctx = spreadsheet.SheetContext(
        sheet_data=[["ID"]],
        id_cell=types.SimpleNamespace(row=1, col=1),
        observation_cell=types.SimpleNamespace(row=1, col=1),
        category_cell=types.SimpleNamespace(row=1, col=1),
        num_participants=0,
        study_name="previous",
    )
    prev_ws = object()

    attempted_ctx = spreadsheet.SheetContext(
        sheet_data=[["ID"]],
        id_cell=types.SimpleNamespace(row=1, col=1),
        observation_cell=types.SimpleNamespace(row=1, col=1),
        category_cell=types.SimpleNamespace(row=1, col=1),
        num_participants=0,
        study_name="attempted",
    )

    def fake_init_studio_state(worksheet):
        server._worksheet = worksheet
        server._sheet_context = attempted_ctx

    monkeypatch.setattr(server, "_worksheet", prev_ws)
    monkeypatch.setattr(server, "_sheet_context", prev_ctx)
    monkeypatch.setattr(server, "_sheet_payload_cache", None)
    monkeypatch.setattr(server, "_generated_artifacts", [])
    monkeypatch.setattr(server, "_generated_reels", [])
    monkeypatch.setattr(server, "_init_studio_state", fake_init_studio_state)

    def _boom(**kwargs):
        raise RuntimeError("screenspace failed")

    monkeypatch.setattr(screenspace_server, "_init_screenspace_state", _boom)
    monkeypatch.setattr(workflows_server, "_sheet_context", prev_ctx)
    monkeypatch.setattr(workflows_server, "_worksheet", prev_ws)
    monkeypatch.setattr(composer_server, "_sheet_context", prev_ctx)

    with pytest.raises(RuntimeError, match="screenspace failed"):
        server._swap_worksheet(object())

    assert workflows_server._sheet_context is prev_ctx
    assert workflows_server._worksheet is prev_ws
    assert composer_server._sheet_context is prev_ctx


def test_api_generate_passes_titlecard_options_to_pipeline(client, monkeypatch):
    """Generate passes per-request titlecard options without mutating config."""
    import types

    import config

    monkeypatch.setattr(server, "_worksheet", object())
    original_enabled = config.TITLECARDS_ENABLED
    original_duration = config.TITLECARD_DURATION_SECONDS
    captured = {}

    cell = types.SimpleNamespace(row=5, col=2, value="1:00")

    def fake_generate_list(ws, mode, *, ctx=None, cell_specs, skip_prompts):
        return [{"participant": "P01", "cell": cell}]

    def fake_process_clips(
        clips,
        *,
        output_format,
        cancel_flag=None,
        titlecards_enabled=None,
        titlecard_duration_seconds=None,
        clear_titlecard_cache=True,
    ):
        captured["enabled"] = titlecards_enabled
        captured["duration"] = titlecard_duration_seconds
        captured["clear_titlecard_cache"] = clear_titlecard_cache
        return (1, [{"id": "a1", "type": "clip"}])

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)
    monkeypatch.setattr("spreadsheet.parse_cell_specifications", lambda t: [("P01", 5)])
    monkeypatch.setattr("pipeline.process_clips", fake_process_clips)
    _set_artifacts(monkeypatch, [])

    resp = client.post(
        "/studio/api/generate",
        json={
            "cells": ["P01.5"],
            "format": "clip",
            "titlecards_enabled": True,
            "titlecard_duration": 5,
        },
    )
    assert resp.status_code == 200
    resp.data  # consume streamed response so generator finally-block runs
    assert captured["enabled"] is True
    assert captured["duration"] == 5
    # Per-cell workers must not purge the shared endcard cache; api_generate
    # clears it once after the stream finishes.
    assert captured["clear_titlecard_cache"] is False
    assert config.TITLECARDS_ENABLED == original_enabled
    assert config.TITLECARD_DURATION_SECONDS == original_duration


def test_api_generate_titlecard_options_on_pipeline_error(client, monkeypatch):
    """Config stays unchanged when process_clips raises."""
    import types

    import config

    monkeypatch.setattr(server, "_worksheet", object())
    original_enabled = config.TITLECARDS_ENABLED
    original_duration = config.TITLECARD_DURATION_SECONDS

    cell = types.SimpleNamespace(row=5, col=2, value="1:00")

    def fake_generate_list(ws, mode, *, ctx=None, cell_specs, skip_prompts):
        return [{"participant": "P01", "cell": cell}]

    def fake_process_clips(clips, *, output_format, cancel_flag=None, **kwargs):
        raise RuntimeError("boom")

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)
    monkeypatch.setattr("spreadsheet.parse_cell_specifications", lambda t: [("P01", 5)])
    monkeypatch.setattr("pipeline.process_clips", fake_process_clips)
    _set_artifacts(monkeypatch, [])

    resp = client.post(
        "/studio/api/generate",
        json={
            "cells": ["P01.5"],
            "format": "clip",
            "titlecards_enabled": True,
            "titlecard_duration": 10,
        },
    )
    _ = resp.data  # consume streaming response to trigger finally
    assert resp.status_code == 200  # streaming response still 200
    assert config.TITLECARDS_ENABLED == original_enabled
    assert config.TITLECARD_DURATION_SECONDS == original_duration


def test_api_reel_passes_titlecard_options_to_pipeline(client, monkeypatch):
    """Reel build forwards per-request titlecard options to the reel stream helper."""
    import types

    import config

    monkeypatch.setattr(server, "_worksheet", object())
    original_enabled = config.TITLECARDS_ENABLED
    original_duration = config.TITLECARD_DURATION_SECONDS
    captured = {}

    cell = types.SimpleNamespace(row=5, col=2, value="1:00-1:30")

    def fake_generate_list(ws, mode, *, ctx=None, reel_input, skip_prompts):
        return [
            {
                "participant": "P01",
                "cell": cell,
                "times": [("1:00", "1:30")],
                "study": "study",
                "category": "cat",
                "desc": "desc",
            }
        ]

    def fake_prepare_clip(clip):
        return clip

    def fake_stream_process_reel(
        clips,
        cancel_flag,
        *,
        titlecards_enabled=None,
        titlecard_duration_seconds=None,
        token=None,
    ):
        captured["enabled"] = titlecards_enabled
        captured["duration"] = titlecard_duration_seconds
        yield json.dumps({"ok": False, "error": "stub"}) + "\n"

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)
    monkeypatch.setattr("files.prepare_clip", fake_prepare_clip)
    monkeypatch.setattr(server, "_stream_process_reel", fake_stream_process_reel)

    resp = client.post(
        "/studio/api/reel",
        json={
            "cells": ["P01.5"],
            "titlecards_enabled": True,
            "titlecard_duration": 4,
        },
    )
    assert resp.status_code == 200
    resp.data
    assert captured["enabled"] is True
    assert captured["duration"] == 4
    assert config.TITLECARDS_ENABLED == original_enabled
    assert config.TITLECARD_DURATION_SECONDS == original_duration


def test_generate_and_reel_busy_slots_are_independent():
    """Artifact generation and reel builds can claim separate busy slots."""
    server._release_busy("generate")
    server._release_busy("reel")
    assert server._try_claim_busy("generate")
    assert server._try_claim_busy("reel")
    assert not server._try_claim_busy("generate")
    assert not server._try_claim_busy("reel")
    server._release_busy("generate")
    server._release_busy("reel")


# ---- Settings API tests ----


def test_api_settings_get(client):
    """GET /api/settings returns all Studio-exposed settings with metadata."""
    resp = client.get("/studio/api/settings")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    settings = data["settings"]
    assert len(settings) > 0

    names = {s["name"] for s in settings}
    assert "REENCODING" in names
    assert "TITLECARDS_ENABLED" in names
    assert "HIGHLIGHTS_REEL_DURATION_SECONDS" in names

    for s in settings:
        assert "name" in s
        assert "value" in s
        assert "default" in s
        assert "description" in s
        assert "group" in s
        assert "type" in s
        assert "tab" in s


def test_api_settings_includes_transcription_settings(client):
    """GET /api/settings includes Transcription-tab settings."""
    resp = client.get("/studio/api/settings")
    data = resp.get_json()
    by_name = {s["name"]: s for s in data["settings"]}
    for name in (
        "TRANSCRIBE_ENABLED",
        "TRANSCRIBE_MODEL",
        "TRANSCRIBE_FORMAT",
        "TRANSCRIBE_PREWARM",
        "TRANSCRIBE_VAD_FILTER",
        "TRANSCRIBE_NO_SPEECH_THRESHOLD",
        "TRANSCRIBE_LOG_PROB_THRESHOLD",
        "TRANSCRIBE_COMPRESSION_RATIO_THRESHOLD",
        "TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD",
        "TRANSCRIBE_CONDITION_ON_PREVIOUS_TEXT",
    ):
        assert name in by_name
        assert by_name[name]["tab"] == "Transcription"
    quality = by_name["TRANSCRIBE_VAD_FILTER"]
    assert quality["group"] == "Transcription quality"
    assert quality["type"] == "bool"
    speakers = by_name["TRANSCRIBE_SPEAKERS"]
    assert speakers["group"] == "Speakers"
    assert speakers["type"] == "bool"
    assert speakers["value"] is False
    cap = by_name["TRANSCRIBE_SPEAKER_MAX"]
    assert cap["group"] == "Speakers"
    assert (cap["min"], cap["max"], cap["step"]) == (2, 8, 1)
    for name in ("TRANSCRIBE_SPEAKERS", "TRANSCRIBE_SPEAKER_MAX"):
        assert name in config.SETTINGS_DESCRIPTIONS


def test_api_settings_includes_cli_settings(client):
    """GET /api/settings includes CLI-tab Rich settings."""
    resp = client.get("/studio/api/settings")
    data = resp.get_json()
    by_name = {s["name"]: s for s in data["settings"]}
    for name in ("RICH_COLORS", "RICH_PANELS", "RICH_PROGRESS"):
        assert name in by_name
        assert by_name[name]["tab"] == "CLI"
        assert by_name[name]["type"] == "bool"


def test_api_settings_includes_screenspace_cv_resolution_scale(client):
    """GET /api/settings exposes the Screenspace CV resolution scale knob."""
    resp = client.get("/studio/api/settings")
    data = resp.get_json()
    by_name = {s["name"]: s for s in data["settings"]}
    assert "SCREENSPACE_CV_RESOLUTION_SCALE" in by_name
    s = by_name["SCREENSPACE_CV_RESOLUTION_SCALE"]
    assert s["tab"] == "Screenspace"
    assert s["type"] == "float"
    assert s["min"] == 0.25
    assert s["max"] == 4.0
    assert s["step"] == 0.25
    assert isinstance(s["value"], float)


def test_api_settings_includes_grouped_tool_nav(client):
    """GET /api/settings exposes the grouped tool-nav toggle, default on."""
    import config

    assert config.SCREENSPACE_GROUPED_TOOL_NAV is True
    assert "SCREENSPACE_GROUPED_TOOL_NAV" in config.SETTINGS_DESCRIPTIONS
    resp = client.get("/studio/api/settings")
    data = resp.get_json()
    by_name = {s["name"]: s for s in data["settings"]}
    assert "SCREENSPACE_GROUPED_TOOL_NAV" in by_name
    s = by_name["SCREENSPACE_GROUPED_TOOL_NAV"]
    assert s["tab"] == "Screenspace"
    assert s["type"] == "bool"
    assert s["value"] is True
    assert s["default"] is True


def test_api_settings_includes_source_filename_pattern(client):
    """GET /api/settings exposes the source-video filename pattern."""
    import config

    assert "SOURCE_FILENAME_PATTERN" in config.SETTINGS_DESCRIPTIONS
    resp = client.get("/studio/api/settings")
    by_name = {s["name"]: s for s in resp.get_json()["settings"]}
    s = by_name["SOURCE_FILENAME_PATTERN"]
    assert s["tab"] == "Video & Clips"
    assert s["group"] == "Source Videos"
    assert s["type"] == "str"
    assert s["default"] == "{study}_{participant}"


def test_settings_put_rejects_bad_source_filename_pattern(
    client, tmp_path, monkeypatch
):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    default = server.config.SOURCE_FILENAME_PATTERN
    monkeypatch.setattr(server.config, "SOURCE_FILENAME_PATTERN", default)
    for bad in (
        "{study}",  # missing {participant}
        "{foo}_{participant}",  # unknown placeholder
        "a/{participant}",  # path separator
        "{participant}_{participant}",  # duplicate
        "",  # empty
    ):
        resp = client.put(
            "/studio/api/settings",
            json={"settings": {"SOURCE_FILENAME_PATTERN": bad}},
        )
        assert resp.status_code == 400, bad
        assert resp.get_json()["ok"] is False
        assert server.config.SOURCE_FILENAME_PATTERN == default  # unchanged


def test_settings_put_applies_source_filename_pattern(client, tmp_path, monkeypatch):
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server.config, "INPUT_DIR", str(tmp_path), raising=False)
    default = server.config.SOURCE_FILENAME_PATTERN
    monkeypatch.setattr(server.config, "SOURCE_FILENAME_PATTERN", default)
    (tmp_path / "P01_demo.mp4").write_text("v")
    assert server.utils.discover_participant_videos() == []

    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"SOURCE_FILENAME_PATTERN": "{participant}_{study}"}},
    )
    assert resp.status_code == 200
    assert server.config.SOURCE_FILENAME_PATTERN == "{participant}_{study}"
    # Discovery honors the new pattern immediately (memo keys on the pattern).
    assert [p["id"] for p in server.utils.discover_participant_videos()] == ["P01"]


def test_api_settings_includes_provider_field(client):
    """model_select settings include a provider field."""
    resp = client.get("/studio/api/settings")
    data = resp.get_json()
    model_settings = [s for s in data["settings"] if s["type"] == "model_select"]
    assert len(model_settings) >= 2  # TRANSCRIBE_MODEL + LLM_SUMMARY_MODEL
    for s in model_settings:
        assert "provider" in s
        assert s["provider"] in ("whisper", "llm")


def test_api_settings_put_applies_values(client, monkeypatch):
    """PUT /api/settings applies values to config and returns applied dict."""
    import config

    monkeypatch.setattr(server, "_save_studio_settings", lambda o: None)
    # Capture+auto-restore: the endpoint mutates config.REENCODING; monkeypatch
    # teardown restores the original even if an assertion below fails.
    monkeypatch.setattr(config, "REENCODING", config.REENCODING)

    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"REENCODING": True}},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["applied"]["REENCODING"] is True
    assert config.REENCODING is True


def test_api_settings_put_ignores_unknown(client, monkeypatch):
    """Unknown setting names are silently ignored."""
    monkeypatch.setattr(server, "_save_studio_settings", lambda o: None)

    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"FAKE_SETTING": 42, "REENCODING": False}},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert "FAKE_SETTING" not in data["applied"]
    assert "REENCODING" in data["applied"]


def test_api_settings_put_type_coercion(client, monkeypatch):
    """Int and bool values are coerced correctly."""
    import config

    monkeypatch.setattr(server, "_save_studio_settings", lambda o: None)
    # Capture+auto-restore (see test_api_settings_put_applies_values).
    monkeypatch.setattr(
        config,
        "HIGHLIGHTS_REEL_DURATION_SECONDS",
        config.HIGHLIGHTS_REEL_DURATION_SECONDS,
    )

    resp = client.put(
        "/studio/api/settings",
        json={"settings": {"HIGHLIGHTS_REEL_DURATION_SECONDS": "120"}},
    )
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["applied"]["HIGHLIGHTS_REEL_DURATION_SECONDS"] == 120
    assert config.HIGHLIGHTS_REEL_DURATION_SECONDS == 120


def test_api_settings_put_invalid_payload(client):
    """PUT /api/settings with no settings dict returns 400."""
    resp = client.put(
        "/studio/api/settings",
        json={"not_settings": {}},
    )
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False


def test_api_settings_reset_tab(client, monkeypatch):
    """PUT {reset: 'tab:<Name>'} resets only that tab's settings to defaults."""
    import config

    monkeypatch.setattr(server, "_save_studio_settings", lambda o: None)
    original_enabled = config.TRANSCRIBE_ENABLED
    original_reenc = config.REENCODING

    # Move two settings off their defaults: one in the target tab, one outside.
    config.TRANSCRIBE_ENABLED = not server._settings_defaults["TRANSCRIBE_ENABLED"]
    config.REENCODING = not server._settings_defaults["REENCODING"]

    try:
        resp = client.put(
            "/studio/api/settings",
            json={"reset": "tab:Transcription"},
        )
        assert resp.status_code == 200
        data = resp.get_json()
        assert data["ok"] is True
        applied = data["applied"]

        # Target-tab settings are back to defaults.
        assert (
            applied["TRANSCRIBE_ENABLED"]
            == server._settings_defaults["TRANSCRIBE_ENABLED"]
        )
        assert (
            config.TRANSCRIBE_ENABLED == server._settings_defaults["TRANSCRIBE_ENABLED"]
        )

        # Settings outside that tab are not in applied and not reset.
        assert "REENCODING" not in applied
        assert config.REENCODING != server._settings_defaults["REENCODING"]
    finally:
        config.TRANSCRIBE_ENABLED = original_enabled
        config.REENCODING = original_reenc


def test_api_settings_reset_all(client, monkeypatch):
    """PUT {reset: 'all'} resets every setting to its default."""
    import config

    monkeypatch.setattr(server, "_save_studio_settings", lambda o: None)
    original_reenc = config.REENCODING
    original_enabled = config.TRANSCRIBE_ENABLED

    config.REENCODING = not server._settings_defaults["REENCODING"]
    config.TRANSCRIBE_ENABLED = not server._settings_defaults["TRANSCRIBE_ENABLED"]

    try:
        resp = client.put("/studio/api/settings", json={"reset": "all"})
        assert resp.status_code == 200
        data = resp.get_json()
        assert data["ok"] is True
        applied = data["applied"]

        for name, default in server._settings_defaults.items():
            assert applied[name] == default
            assert getattr(config, name) == default
    finally:
        config.REENCODING = original_reenc
        config.TRANSCRIBE_ENABLED = original_enabled


def test_api_settings_reset_unknown_tab(client, monkeypatch):
    """Unknown tab resets nothing and returns ok with empty applied."""
    monkeypatch.setattr(server, "_save_studio_settings", lambda o: None)
    resp = client.put("/studio/api/settings", json={"reset": "tab:Bogus"})
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["applied"] == {}


def test_load_studio_settings(monkeypatch, tmp_path):
    """_load_studio_settings reads file and applies to config."""
    import config

    monkeypatch.setattr(server.start_settings, "config_dir", lambda: tmp_path)
    # Capture+auto-restore (see test_api_settings_put_applies_values).
    monkeypatch.setattr(config, "REENCODING", config.REENCODING)

    settings_file = tmp_path / config.STUDIO_SETTINGS_FILENAME
    settings_file.write_text(json.dumps({"REENCODING": True}))

    applied = server._load_studio_settings()
    assert applied["REENCODING"] is True
    assert config.REENCODING is True


def test_load_studio_settings_missing_file(monkeypatch, tmp_path):
    """Missing settings file returns empty dict without error."""
    monkeypatch.setattr(server.start_settings, "config_dir", lambda: tmp_path)
    applied = server._load_studio_settings()
    assert applied == {}


def test_load_studio_settings_skips_invalid_card_image(monkeypatch, tmp_path):
    """A persisted card image that PUT would reject is not applied on load."""
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server.start_settings, "config_dir", lambda: tmp_path)
    monkeypatch.setattr(server.config, "TITLECARD_IMAGE", "")
    monkeypatch.setattr(server.config, "ENDCARD_IMAGE", "")

    images = tmp_path / server.config.TITLECARD_IMAGES_DIRNAME
    images.mkdir()
    (images / "good.png").write_bytes(b"x")

    settings_file = tmp_path / server.config.STUDIO_SETTINGS_FILENAME
    settings_file.write_text(
        json.dumps(
            {
                "TITLECARD_IMAGE": "../escape.png",  # traversal → rejected
                "ENDCARD_IMAGE": "good.png",  # real upload → applied
            }
        )
    )

    applied = server._load_studio_settings()
    # The bad value is skipped (config stays at default); the good one applies.
    assert "TITLECARD_IMAGE" not in applied
    assert server.config.TITLECARD_IMAGE == ""
    assert applied["ENDCARD_IMAGE"] == "good.png"
    assert server.config.ENDCARD_IMAGE == "good.png"


def test_load_studio_settings_skips_invalid_prompt(monkeypatch, tmp_path):
    """A persisted prompt that PUT would reject is not applied on load.

    Regression: the load path shared its coercion ladder with the PUT path, but
    used to lack the ``prompt`` branch, so a tampered prompt was applied without
    _validate_prompt. It must now be skipped (config keeps its default), matching
    the card_picker guard, while a valid setting alongside it still applies.
    """
    monkeypatch.setattr(server.start_settings, "config_dir", lambda: tmp_path)
    default_prompt = server.config.LLM_FRICTION_PROMPT
    monkeypatch.setattr(server.config, "LLM_FRICTION_PROMPT", default_prompt)
    monkeypatch.setattr(server.config, "REENCODING", server.config.REENCODING)

    settings_file = tmp_path / server.config.STUDIO_SETTINGS_FILENAME
    settings_file.write_text(
        json.dumps(
            {
                # Missing required {summary}/{segments}/{limit} placeholders → rejected.
                "LLM_FRICTION_PROMPT": "tampered prompt with no placeholders",
                "REENCODING": True,  # valid → applied
            }
        )
    )

    applied = server._load_studio_settings()
    assert "LLM_FRICTION_PROMPT" not in applied
    assert server.config.LLM_FRICTION_PROMPT == default_prompt
    assert applied["REENCODING"] is True
    assert server.config.REENCODING is True


def test_save_studio_settings_non_defaults_only(monkeypatch, tmp_path):
    """Only non-default values are written; all-defaults deletes the file."""
    import config

    monkeypatch.setattr(server.start_settings, "config_dir", lambda: tmp_path)
    settings_file = tmp_path / config.STUDIO_SETTINGS_FILENAME

    # Save a non-default value
    result = server._save_studio_settings({"REENCODING": True})
    assert result is not None
    assert settings_file.is_file()
    data = json.loads(settings_file.read_text())
    assert data["REENCODING"] is True

    # Save all defaults — file should be removed
    result = server._save_studio_settings(
        {"REENCODING": server._settings_defaults["REENCODING"]}
    )
    assert result is None
    assert not settings_file.is_file()


def test_api_reel_cancel_endpoint(client):
    """POST /api/reel/cancel should set the cancel event and return ok."""
    server._reel_cancel_event.clear()
    resp = client.post("/studio/api/reel/cancel")
    data = resp.get_json()
    assert resp.status_code == 200
    assert data["ok"] is True
    assert server._reel_cancel_event.is_set()


def test_api_generate_cancel_endpoint(client):
    """POST /api/generate/cancel should set the cancel event and return ok."""
    server._generate_cancel_event.clear()
    resp = client.post("/studio/api/generate/cancel")
    data = resp.get_json()
    assert resp.status_code == 200
    assert data["ok"] is True
    assert server._generate_cancel_event.is_set()


def test_api_generate_intake_streams_per_item(client, monkeypatch):
    """POST /api/generate-intake yields one NDJSON line per item with index/ok."""
    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("video.run_ffmpeg", lambda *a, **kw: True)
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)

    items = [
        {
            "participant": "P01",
            "start": 0.0,
            "end": 5.0,
            "event_type": "x",
            "event_ids": [],
            "source": "screenspace",
            "mark_ids": [],
        },
        {
            "participant": "P02",
            "start": 0.0,
            "end": 5.0,
            "event_type": "y",
            "event_ids": [],
            "source": "screenspace",
            "mark_ids": [],
        },
    ]
    resp = client.post(
        "/studio/api/generate-intake",
        json={"items": items, "format": "clip"},
    )
    assert resp.status_code == 200
    assert resp.mimetype == "application/x-ndjson"
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert len(lines) == 2
    assert {ln["index"] for ln in lines} == {0, 1}
    for ln in lines:
        assert ln["ok"] is True
        assert ln["artifact"]["participant"] in {"P01", "P02"}


def test_api_generate_intake_streams_failure(client, monkeypatch):
    """Items with no resolvable video stream an ok=false line with error."""
    monkeypatch.setattr(server, "_resolve_intake_video_paths", lambda p, s="": [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)

    items = [
        {
            "participant": "Pxx",
            "start": 0.0,
            "end": 5.0,
            "event_type": "",
            "event_ids": [],
            "source": "screenspace",
            "mark_ids": [],
        },
    ]
    resp = client.post(
        "/studio/api/generate-intake",
        json={"items": items, "format": "clip"},
    )
    assert resp.status_code == 200
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert len(lines) == 1
    assert lines[0]["ok"] is False
    assert lines[0]["index"] == 0
    assert "No video" in lines[0]["error"]


def test_api_generate_intake_rejects_empty(client):
    """POST /api/generate-intake without items returns a 400."""
    resp = client.post("/studio/api/generate-intake", json={"items": []})
    assert resp.status_code == 400
    data = resp.get_json()
    assert data["ok"] is False


def test_api_reel_streams_progress_events(client, monkeypatch, tmp_path):
    """/api/reel emits NDJSON phase events plus a final result line.

    The streaming contract is what the new Studio button-progress UI relies on
    to fill the Build Reel button as encoding advances. We mock process_reel to
    fire start/clip_done/concat/done events and assert each one appears in the
    response, followed by the {"ok": True, ...} payload.
    """
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    # Prevent the skip-existing-reel branch from short-circuiting the stream.
    monkeypatch.setattr(server, "_generated_reels", [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)

    cell = types.SimpleNamespace(row=5, col=2, value="1:00-1:30")

    def fake_generate_list(ws, mode, *, ctx=None, reel_input, skip_prompts):
        return [
            {
                "participant": "P01",
                "cell": cell,
                "times": [("1:00", "1:30")],
                "desc": "test",
                "category": "cat",
                "study": "study",
                "severity": "",
            }
        ]

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)
    monkeypatch.setattr("files.prepare_clip", lambda clip: clip)

    reel_record = {
        "id": "r1",
        "file": "reel.mp4",
        "study": "study",
        "components": [],
    }

    def fake_process_reel(
        clips_list,
        output_file=None,
        cancel_flag=None,
        progress_cb=None,
        **kwargs,
    ):
        if progress_cb is not None:
            progress_cb({"phase": "start", "total_clips": 2})
            progress_cb({"phase": "clip_done", "clip_index": 0, "total_clips": 2})
            progress_cb({"phase": "clip_done", "clip_index": 1, "total_clips": 2})
            progress_cb({"phase": "concat", "progress": 0.5})
            progress_cb({"phase": "concat", "progress": 0.99})
            progress_cb({"phase": "done"})
        return (1, [reel_record])

    monkeypatch.setattr("pipeline.process_reel", fake_process_reel)

    resp = client.post("/studio/api/reel", json={"cells": ["P01.5"]})
    assert resp.status_code == 200
    assert resp.mimetype == "application/x-ndjson"

    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    # Expect at minimum a "start", a "clip_done", a "concat", "done", and the
    # final result line — order matters for monotonic UI updates.
    phases = [ln.get("phase") for ln in lines if "phase" in ln]
    assert "start" in phases
    assert "clip_done" in phases
    assert "concat" in phases
    assert phases[-1] == "done"

    final = lines[-1]
    assert final["ok"] is True
    assert final["generated"] == 1
    assert final["reels"] == [reel_record]


def test_api_generate_intake_distinct_ids_for_same_span(client, monkeypatch):
    """Two intake items over the same participant span receive distinct artifact
    ids, so manifest dedup does not silently collapse them onto one record."""
    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("video.run_ffmpeg", lambda *a, **kw: True)
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr(
        server.files, "get_unique_filename", lambda name, file_format=None: name
    )

    same_span = {
        "participant": "P01",
        "start": 1.0,
        "end": 5.0,
        "event_type": "",
        "source": "screenspace",
        "mark_ids": [],
    }
    items = [
        {**same_span, "event_ids": ["e1"]},
        {**same_span, "event_ids": ["e2"]},
    ]
    resp = client.post(
        "/studio/api/generate-intake", json={"items": items, "format": "clip"}
    )
    assert resp.status_code == 200
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    ids = [ln["artifact"]["id"] for ln in lines]
    assert len(ids) == 2
    assert len(set(ids)) == 2


def test_api_generate_intake_persists_when_later_item_raises(client, monkeypatch):
    """A later intake item raising still persists the earlier successes via the
    generator's finally block."""
    saved: list[bool] = []
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: saved.append(True))

    def flaky(item, output_format, study, index=0):
        if index == 1:
            raise RuntimeError("boom")
        return {
            "_ok": True,
            "_error": "",
            "id": f"art{index}",
            "participant": item["participant"],
        }

    monkeypatch.setattr(server, "_process_intake_item", flaky)

    items = [
        {"participant": "P01", "start": 0, "end": 5},
        {"participant": "P02", "start": 0, "end": 5},
    ]
    resp = client.post(
        "/studio/api/generate-intake", json={"items": items, "format": "clip"}
    )
    try:
        resp.data  # the generator raises mid-stream; the finally still runs
    except RuntimeError:
        pass
    assert saved == [True]


def test_api_generate_intake_persists_on_early_client_close(client, monkeypatch):
    """Closing the response after one line still persists via the finally."""
    saved: list[bool] = []
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: saved.append(True))
    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("video.run_ffmpeg", lambda *a, **kw: True)
    monkeypatch.setattr(
        server.files, "get_unique_filename", lambda name, file_format=None: name
    )

    items = [
        {"participant": f"P0{i}", "start": 0, "end": 5, "source": "screenspace"}
        for i in range(1, 5)
    ]
    resp = client.post(
        "/studio/api/generate-intake", json={"items": items, "format": "clip"}
    )
    encoded = resp.iter_encoded()
    next(encoded)  # consume only the first NDJSON line
    resp.close()  # disconnect mid-stream
    assert saved == [True]


def test_media_cache_lru_is_threadsafe(monkeypatch):
    """Concurrent get_or_compute calls keep the LRU bounded without corrupting
    the OrderedDict mid-eviction."""
    import concurrent.futures

    cache = server._MediaCache(server._THUMBNAIL_CACHE_MAX)
    monkeypatch.setattr(server, "_thumbnail_cache", cache)

    def put(i: int) -> None:
        cache.get_or_compute((f"v{i}", i), lambda: b"jpeg")

    with concurrent.futures.ThreadPoolExecutor(max_workers=16) as pool:
        list(pool.map(put, range(2000)))

    assert len(cache._store) <= server._THUMBNAIL_CACHE_MAX


def test_media_cache_single_flight():
    """Concurrent identical cache misses must run the producer (ffmpeg) once —
    the stampede guard the plain lock-around-dict-access could not provide. A
    barrier forces every thread into the miss path before any producer returns,
    which the old release-lock-before-compute pattern would have failed."""
    import concurrent.futures
    import threading
    import time

    cache = server._MediaCache(server._THUMBNAIL_CACHE_MAX)
    workers = 8
    barrier = threading.Barrier(workers)
    call_count = [0]
    count_lock = threading.Lock()

    def produce():
        with count_lock:
            call_count[0] += 1
        time.sleep(0.05)  # hold the producer so overlapping misses would stack
        return b"jpeg"

    def fetch(_):
        barrier.wait()  # release all threads into get_or_compute together
        return cache.get_or_compute(("v", 5), produce)

    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        results = list(pool.map(fetch, range(workers)))

    assert all(r == b"jpeg" for r in results)
    assert call_count[0] == 1


def _poll_until(predicate, *, timeout=5.0, interval=0.02):
    """Spin until predicate() returns truthy or timeout elapses."""
    import time

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if predicate():
            return True
        time.sleep(interval)
    return False


def _fake_reel_clip():
    """Single-clip payload that satisfies /api/reel's setup phase."""
    import types

    cell = types.SimpleNamespace(row=5, col=2, value="1:00-1:30")
    return {
        "participant": "P01",
        "cell": cell,
        "times": [("1:00", "1:30")],
        "desc": "test",
        "category": "cat",
        "study": "study",
        "severity": "",
    }


def _setup_api_reel(monkeypatch, tmp_path, *, clips=None):
    """Shared mocks for /api/reel tests."""
    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server, "_generated_reels", [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)

    payload = [_fake_reel_clip()] if clips is None else clips
    monkeypatch.setattr(
        "spreadsheet.generate_list",
        lambda ws, mode, *, ctx=None, reel_input, skip_prompts: payload,
    )
    monkeypatch.setattr("files.prepare_clip", lambda clip: clip)
    server._reel_cancel_event.clear()


def test_api_reel_continues_worker_after_client_disconnect(
    client, monkeypatch, tmp_path
):
    """A client disconnect (browser navigation to a sibling frontend) must NOT
    abort the background encoder. The worker keeps running, persists the reel
    record to the manifest, and releases the busy slot when it completes.

    Reverses PR #353's behavior: GeneratorExit no longer sets the cancel event;
    only explicit cancels do.
    """
    import threading

    _setup_api_reel(monkeypatch, tmp_path)

    reel_record = {"id": "r1", "file": "reel.mp4", "study": "study", "components": []}
    worker_blocked = threading.Event()
    allow_finish = threading.Event()

    def fake_process_reel(clips_list, **kwargs):
        progress_cb = kwargs.get("progress_cb")
        if progress_cb is not None:
            progress_cb({"phase": "start", "total_clips": 1})
        worker_blocked.set()
        # Block briefly so the test can disconnect mid-encode.
        allow_finish.wait(timeout=5)
        return (1, [reel_record])

    monkeypatch.setattr("pipeline.process_reel", fake_process_reel)

    resp = client.post("/studio/api/reel", json={"cells": ["P01.5"]})
    try:
        stream = resp.iter_encoded()
        first = json.loads(next(stream).decode().strip())
        assert first["phase"] == "start"
        assert worker_blocked.wait(timeout=5)

        resp.close()  # simulate browser tab navigating away
        # No auto-cancel on disconnect anymore.
        assert server._reel_cancel_event.is_set() is False
        # Slot still held: worker is running.
        assert server._busy_slots["reel"] is True
    finally:
        allow_finish.set()

    # Once the worker finishes, manifest is updated and slot is released.
    assert _poll_until(lambda: server._busy_slots["reel"] is False)
    assert reel_record in server._generated_reels


def test_api_reel_busy_slot_held_during_worker_after_disconnect(
    client, monkeypatch, tmp_path
):
    """While the background worker is running, a second /api/reel POST must
    receive 409 even after the original client has disconnected. The slot only
    frees once the worker actually completes."""
    import threading

    _setup_api_reel(monkeypatch, tmp_path)

    worker_blocked = threading.Event()
    allow_finish = threading.Event()

    def fake_process_reel(clips_list, **kwargs):
        progress_cb = kwargs.get("progress_cb")
        if progress_cb is not None:
            progress_cb({"phase": "start", "total_clips": 1})
        worker_blocked.set()
        allow_finish.wait(timeout=5)
        return (1, [{"id": "r1", "file": "r.mp4"}])

    monkeypatch.setattr("pipeline.process_reel", fake_process_reel)

    resp = client.post("/studio/api/reel", json={"cells": ["P01.5"]})
    try:
        stream = resp.iter_encoded()
        next(stream)
        assert worker_blocked.wait(timeout=5)
        resp.close()

        # Worker still running → second request rejected.
        second = client.post("/studio/api/reel", json={"cells": ["P01.5"]})
        assert second.status_code == 409
    finally:
        allow_finish.set()

    # After worker exits, the slot frees and a fresh request can proceed.
    assert _poll_until(lambda: server._busy_slots["reel"] is False)
    third = client.post("/studio/api/reel", json={"cells": ["P01.5"]})
    assert third.status_code == 200


def test_api_reel_explicit_cancel_still_works(client, monkeypatch, tmp_path):
    """The Cancel button (POST /api/reel/cancel) must continue to abort an
    in-flight build. This is the only path that should kill ffmpeg."""
    import threading

    _setup_api_reel(monkeypatch, tmp_path)

    started = threading.Event()

    def fake_process_reel(clips_list, **kwargs):
        progress_cb = kwargs.get("progress_cb")
        cancel_flag = kwargs.get("cancel_flag")
        if progress_cb is not None:
            progress_cb({"phase": "start", "total_clips": 1})
        started.set()
        # Honor cancel_flag instead of running ffmpeg.
        for _ in range(500):
            if cancel_flag and cancel_flag():
                return (0, [])
            threading.Event().wait(0.01)
        return (1, [{"id": "r1", "file": "r.mp4"}])

    monkeypatch.setattr("pipeline.process_reel", fake_process_reel)

    resp = client.post("/studio/api/reel", json={"cells": ["P01.5"]})
    stream = resp.iter_encoded()
    next(stream)
    assert started.wait(timeout=5)

    cancel_resp = client.post("/studio/api/reel/cancel")
    assert cancel_resp.status_code == 200
    assert server._reel_cancel_event.is_set()

    # The worker observes the cancel and emits a cancelled-error event.
    final = json.loads(resp.data.decode().strip().split("\n")[-1])
    assert final.get("cancelled") is True
    assert _poll_until(lambda: server._busy_slots["reel"] is False)


def test_api_reel_releases_slot_on_no_clips_early_return(client, monkeypatch, tmp_path):
    """When generate_list returns no clips, the route reports the error and
    releases the slot itself — no worker takes over."""
    _setup_api_reel(monkeypatch, tmp_path, clips=[])

    resp = client.post("/studio/api/reel", json={"cells": ["P01.5"]})
    assert resp.status_code == 200
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert lines[-1]["ok"] is False
    assert "No clips" in lines[-1]["error"]
    assert server._busy_slots["reel"] is False


def test_api_reel_releases_slot_on_cached_match(client, monkeypatch, tmp_path):
    """A cached-reel hit short-circuits before the worker runs. The route must
    release the slot on this early-return path."""
    _setup_api_reel(monkeypatch, tmp_path)

    # Pre-populate a matching reel record and create the file on disk.
    (tmp_path / "cached.mp4").write_bytes(b"video")
    cached = {
        "id": "cached-id",
        "file": "cached.mp4",
        "study": "study",
        "components": [],
    }
    monkeypatch.setattr(server, "_generated_reels", [cached])
    monkeypatch.setattr("pipeline.compute_reel_id", lambda components: "cached-id")
    monkeypatch.setattr(
        "utils.build_reel_component", lambda clip, src, s, e: {"start": s, "end": e}
    )

    resp = client.post("/studio/api/reel", json={"cells": ["P01.5"]})
    assert resp.status_code == 200
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert lines[-1]["ok"] is True
    assert lines[-1]["skipped"] is True
    assert server._busy_slots["reel"] is False


def test_api_reel_releases_slot_on_pipeline_exception(client, monkeypatch, tmp_path):
    """If pipeline.process_reel raises, the worker's finally still releases the
    busy slot — the route does not need to know the worker failed."""
    _setup_api_reel(monkeypatch, tmp_path)

    def raises(*a, **kw):
        raise RuntimeError("ffmpeg exploded")

    monkeypatch.setattr("pipeline.process_reel", raises)

    resp = client.post("/studio/api/reel", json={"cells": ["P01.5"]})
    assert resp.status_code == 200
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]
    assert any(
        ln.get("ok") is False and "ffmpeg exploded" in ln.get("error", "")
        for ln in lines
    )
    assert _poll_until(lambda: server._busy_slots["reel"] is False)


def test_api_reel_direct_continues_worker_after_client_disconnect(
    client, monkeypatch, tmp_path
):
    """Same disconnect-survives-work contract for /api/reel-direct: the
    intake-reel worker continues, persists, and releases the slot."""
    import threading

    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server, "_generated_reels", [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr("video.run_ffmpeg", lambda *a, **kw: True)
    monkeypatch.setattr(
        "files.get_unique_filename", lambda name, file_format=None: name
    )

    concat_blocked = threading.Event()
    allow_finish = threading.Event()

    def fake_concat(clip_paths, reel_name, **kwargs):
        concat_blocked.set()
        allow_finish.wait(timeout=5)
        return True

    monkeypatch.setattr("video.concatenate_clips", fake_concat)
    server._reel_cancel_event.clear()

    resp = client.post(
        "/studio/api/reel-direct",
        json={"segments": [{"participant": "P01", "start": 0, "end": 5}]},
    )
    try:
        stream = resp.iter_encoded()
        first = json.loads(next(stream).decode().strip())
        assert first["phase"] == "start"
        assert concat_blocked.wait(timeout=5)
        resp.close()
        assert server._reel_cancel_event.is_set() is False
        assert server._busy_slots["reel"] is True
    finally:
        allow_finish.set()

    assert _poll_until(lambda: server._busy_slots["reel"] is False)
    assert len(server._generated_reels) == 1
    assert server._generated_reels[0]["source"] == "intake"


def test_api_reel_direct_cleans_temp_clips_after_disconnect(
    client, monkeypatch, tmp_path
):
    """The worker's finally must unlink per-segment temp files even when the
    client has disconnected before the worker completed."""
    import threading
    from pathlib import Path

    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server, "_generated_reels", [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr(
        "files.get_unique_filename", lambda name, file_format=None: name
    )

    # Track every tempfile that run_ffmpeg "produces".
    created_temps: list[str] = []

    def fake_run_ffmpeg(src, dst, *a, **kw):
        Path(dst).write_bytes(b"clip")
        created_temps.append(dst)
        return True

    monkeypatch.setattr("video.run_ffmpeg", fake_run_ffmpeg)

    concat_started = threading.Event()
    allow_finish = threading.Event()

    def fake_concat(clip_paths, reel_name, **kwargs):
        concat_started.set()
        allow_finish.wait(timeout=5)
        return True

    monkeypatch.setattr("video.concatenate_clips", fake_concat)
    server._reel_cancel_event.clear()

    resp = client.post(
        "/studio/api/reel-direct",
        json={
            "segments": [
                {"participant": "P01", "start": 0, "end": 5},
                {"participant": "P01", "start": 5, "end": 10},
            ]
        },
    )
    try:
        stream = resp.iter_encoded()
        next(stream)
        assert concat_started.wait(timeout=5)
        resp.close()
    finally:
        allow_finish.set()

    assert _poll_until(lambda: server._busy_slots["reel"] is False)
    assert created_temps  # sanity: ffmpeg ran
    for tmp in created_temps:
        assert not Path(tmp).exists(), f"temp clip {tmp} was not cleaned up"


def test_api_reel_direct_explicit_cancel_still_works(client, monkeypatch, tmp_path):
    """POST /api/reel/cancel during a reel-direct build must abort the
    per-segment loop and yield a cancelled-error event."""
    import threading

    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server, "_generated_reels", [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr(
        "files.get_unique_filename", lambda name, file_format=None: name
    )

    started = threading.Event()

    def slow_ffmpeg(src, dst, *a, **kw):
        started.set()
        for _ in range(500):
            if kw.get("cancel_flag") and kw["cancel_flag"]():
                return False
            threading.Event().wait(0.01)
        return True

    monkeypatch.setattr("video.run_ffmpeg", slow_ffmpeg)
    monkeypatch.setattr("video.concatenate_clips", lambda *a, **kw: True)
    server._reel_cancel_event.clear()

    resp = client.post(
        "/studio/api/reel-direct",
        json={"segments": [{"participant": "P01", "start": 0, "end": 5}]},
    )
    stream = resp.iter_encoded()
    next(stream)
    assert started.wait(timeout=5)

    cancel_resp = client.post("/studio/api/reel/cancel")
    assert cancel_resp.status_code == 200

    final = json.loads(resp.data.decode().strip().split("\n")[-1])
    assert final.get("cancelled") is True
    assert _poll_until(lambda: server._busy_slots["reel"] is False)


def test_api_generate_persists_artifacts_after_disconnect(
    client, monkeypatch, tmp_path
):
    """After a client disconnect mid-/api/generate, every clip whose worker
    completed (including ones still in the pool when the for-loop died) lands
    in the manifest. Validates that _extend_generated_artifacts moved into the
    per-clip worker."""
    import threading
    import time
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    _set_artifacts(monkeypatch, [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr("titlecards.clear_endcard_cache", lambda: None)
    # Force the parallel path: pool size > 1, more cells than pool workers so
    # some clips queue up behind in-flight ones.
    monkeypatch.setattr("pipeline._resolve_clip_workers", lambda: 2)

    cells = [types.SimpleNamespace(row=r, col=2, value="1:00") for r in (5, 6, 7, 8)]
    monkeypatch.setattr(
        "spreadsheet.parse_cell_specifications",
        lambda text: [("P01", r) for r in (5, 6, 7, 8)],
    )
    monkeypatch.setattr(
        "spreadsheet.generate_list",
        lambda ws, mode, *, ctx=None, cell_specs, skip_prompts: [
            {"participant": "P01", "cell": c} for c in cells
        ],
    )

    started_count = [0]
    finished_count = [0]
    started_lock = threading.Lock()

    # Hold each worker long enough that the disconnect below lands while the
    # second pool batch is still in flight. It has to be a plain sleep, not a
    # wait-for-signal: post() does not return until the *first* batch finishes,
    # and resp.close() then blocks draining the pool, so there is no point on the
    # test thread from which the first batch could ever be released. This was a
    # `proceed.wait(timeout=2)` paired with a `proceed.set()` after close() —
    # every one of the four workers provably timed out instead, the set() was
    # unreachable, and the test paid 2 batches x 2 s. The in-flight assertion
    # after close() is what keeps this constant honest if it is ever too short.
    HOLD_SECONDS = 0.15

    def fake_process_clips(clip_list, **kwargs):
        with started_lock:
            started_count[0] += 1
        time.sleep(HOLD_SECONDS)
        clip = clip_list[0]
        artifact = {
            "id": f"a{clip['cell'].row}",
            "type": "clip",
            "file": f"clip{clip['cell'].row}.mp4",
            "cellRow": clip["cell"].row,
            "cellCol": clip["cell"].col,
        }
        with started_lock:
            finished_count[0] += 1
        return (1, [artifact])

    monkeypatch.setattr("pipeline.process_clips", fake_process_clips)

    resp = client.post(
        "/studio/api/generate",
        json={
            "cells": ["P01.5", "P01.6", "P01.7", "P01.8"],
            "format": "clip",
        },
    )
    # post() returns once the first pool batch has drained and the next one has
    # been submitted, so some worker is always mid-flight here.
    assert _poll_until(lambda: started_count[0] >= 2)
    with started_lock:
        in_flight_at_disconnect = started_count[0] - finished_count[0]
    resp.close()

    # The whole point of the test is that a worker still running at disconnect
    # gets its artifact persisted. If HOLD_SECONDS is ever cut so fine that every
    # worker has already finished by now, the row assertion below would still pass
    # while proving nothing — so fail loudly on that instead.
    assert in_flight_at_disconnect >= 1, (
        "no worker was still in flight when the client disconnected; "
        f"raise HOLD_SECONDS ({started_count[0]} started, {finished_count[0]} done)"
    )
    assert _poll_until(lambda: server._busy_slots["generate"] is False)
    persisted_rows = {a["cellRow"] for a in server._generated_artifacts}
    assert persisted_rows == {5, 6, 7, 8}


def test_api_job_status_idle(client):
    """/api/job-status reports neither slot in progress when nothing is running."""
    resp = client.get("/studio/api/job-status")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["ok"] is True
    assert data["reel"]["in_progress"] is False
    assert data["generate"]["in_progress"] is False
    assert data["intake"]["in_progress"] is False
    assert "done" in data["intake"]
    assert "total" in data["intake"]
    # Snapshot fields are always present so the client can render without
    # null-checks.
    assert "phase" in data["reel"]
    assert "clips_done" in data["reel"]
    assert "total_clips" in data["reel"]
    assert "concat_progress" in data["reel"]
    assert "done" in data["generate"]
    assert "total" in data["generate"]
    # started_at lets a reattaching client show accurate elapsed time.
    assert "started_at" in data["reel"]
    assert "started_at" in data["generate"]
    assert "started_at" in data["intake"]


def test_api_job_status_reflects_reel_progress(client, monkeypatch, tmp_path):
    """While a reel worker is running, /api/job-status surfaces phase + counts
    so Studio can rebuild the progress bar after navigating back."""
    import threading

    _setup_api_reel(monkeypatch, tmp_path)

    proceed = threading.Event()
    started = threading.Event()

    def fake_process_reel(clips_list, **kwargs):
        progress_cb = kwargs.get("progress_cb")
        if progress_cb is not None:
            progress_cb({"phase": "start", "total_clips": 3})
            progress_cb({"phase": "clip_done", "clip_index": 0, "total_clips": 3})
            progress_cb({"phase": "clip_done", "clip_index": 1, "total_clips": 3})
            progress_cb({"phase": "concat", "progress": 0.4})
        started.set()
        proceed.wait(timeout=5)
        return (1, [{"id": "r1", "file": "r.mp4"}])

    monkeypatch.setattr("pipeline.process_reel", fake_process_reel)

    resp = client.post("/studio/api/reel", json={"cells": ["P01.5"]})
    try:
        assert started.wait(timeout=5)

        status = client.get("/studio/api/job-status").get_json()
        reel = status["reel"]
        assert reel["in_progress"] is True
        assert reel["endpoint"] == "reel"
        assert reel["total_clips"] == 3
        assert reel["clips_done"] == 2
        assert reel["concat_progress"] == 0.4
        assert reel["phase"] == "concat"
        assert reel["cancelling"] is False
        # Stamped when the build began so a reattach can show elapsed time.
        assert isinstance(reel["started_at"], (int, float))
    finally:
        proceed.set()
        resp.close()

    assert _poll_until(lambda: server._busy_slots["reel"] is False)
    final = client.get("/studio/api/job-status").get_json()
    assert final["reel"]["in_progress"] is False


def test_api_job_status_reflects_generate_progress(client, monkeypatch, tmp_path):
    """The generate side of /api/job-status reports done/total so Studio's
    Generate button can be re-filled after navigation. Uses the parallel
    path so the test can observe mid-flight state without driving the
    streaming response (the busy pool worker thread keeps the job alive)."""
    import threading
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    _set_artifacts(monkeypatch, [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr("titlecards.clear_endcard_cache", lambda: None)
    monkeypatch.setattr("pipeline._resolve_clip_workers", lambda: 2)

    cells = [types.SimpleNamespace(row=r, col=2, value="1:00") for r in (5, 6, 7)]
    monkeypatch.setattr(
        "spreadsheet.parse_cell_specifications",
        lambda text: [("P01", r) for r in (5, 6, 7)],
    )
    monkeypatch.setattr(
        "spreadsheet.generate_list",
        lambda ws, mode, *, ctx=None, cell_specs, skip_prompts: [
            {"participant": "P01", "cell": c} for c in cells
        ],
    )

    started_lock = threading.Lock()
    started_count = [0]
    proceed = threading.Event()

    def fake_process_clips(clip_list, **kwargs):
        with started_lock:
            started_count[0] += 1
        row = clip_list[0]["cell"].row
        if row != 5:
            # Row 5 finishes immediately; rows 6 and 7 block so the test can
            # observe the state when one clip is done and others are mid-flight.
            proceed.wait(timeout=5)
        return (
            1,
            [
                {
                    "id": f"a{row}",
                    "type": "clip",
                    "file": f"c{row}.mp4",
                    "cellRow": row,
                    "cellCol": 2,
                }
            ],
        )

    monkeypatch.setattr("pipeline.process_clips", fake_process_clips)
    server._generate_cancel_event.clear()

    resp = client.post(
        "/studio/api/generate",
        json={"cells": ["P01.5", "P01.6", "P01.7"], "format": "clip"},
    )
    try:
        # Wait for the snapshot to record at least one completed clip.
        assert _poll_until(lambda: server._generate_job_state["done"] >= 1)

        status = client.get("/studio/api/job-status").get_json()
        gen = status["generate"]
        assert gen["in_progress"] is True
        assert gen["total"] == 3
        assert gen["done"] >= 1  # at least row 5 fully processed
        assert gen["cancelling"] is False
    finally:
        proceed.set()
        resp.close()

    assert _poll_until(lambda: server._busy_slots["generate"] is False)


def _setup_single_cell_generate(monkeypatch, tmp_path, cell_value, *, generated=1):
    """Wire /api/generate down to one P01.5 cell holding *cell_value*."""
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    _set_artifacts(monkeypatch, [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)

    cell = types.SimpleNamespace(row=5, col=2, value=cell_value)
    monkeypatch.setattr("spreadsheet.parse_cell_specifications", lambda t: [("P01", 5)])
    monkeypatch.setattr(
        "spreadsheet.generate_list",
        lambda ws, mode, *, ctx=None, cell_specs, skip_prompts: [
            {"participant": "P01", "cell": cell, "desc": "obs", "category": "nav"}
        ],
    )
    monkeypatch.setattr("pipeline.process_clips", lambda *a, **kw: (generated, []))


def test_api_generate_progress_counts_artifacts_not_cells(
    client, monkeypatch, tmp_path
):
    """A multi-timestamp cell is one NDJSON line but several artifacts.

    Studio's Artifacts badge counts queue cards — one per timestamp segment — so
    the job-state total has to count segments too. Counting cells is what made
    the panel read "(58)" next to "51 / 52 cells" for one Generate click.
    """
    _setup_single_cell_generate(
        monkeypatch, tmp_path, "1:00-1:30 2:00-2:30", generated=2
    )

    resp = client.post(
        "/studio/api/generate", json={"cells": ["P01.5"], "format": "clip"}
    )
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]

    assert len(lines) == 1, "one cell still yields one line"
    assert lines[0]["ok"] is True
    # ...but the progress snapshot advances by both segments, in one step.
    assert server._generate_job_state["total"] == 2
    assert server._generate_job_state["done"] == 2


def test_api_generate_progress_counts_trimmed_override_segments(
    client, monkeypatch, tmp_path
):
    """Trimming in the queue replaces the cell's segments, and the count follows.

    The frontend posts the complete remaining segment list whenever a cell was
    edited or had cards removed, so the override — not the sheet value — decides
    how many artifacts that cell contributes.
    """
    _setup_single_cell_generate(monkeypatch, tmp_path, "1:00", generated=2)

    resp = client.post(
        "/studio/api/generate",
        json={
            "cells": ["P01.5"],
            "format": "clip",
            "overrides": {"P01.5": [[10, 40], [60, 90]]},
        },
    )
    assert resp.status_code == 200
    assert server._generate_job_state["total"] == 2
    assert server._generate_job_state["done"] == 2


def test_api_generate_progress_skips_unmatched_refs(client, monkeypatch, tmp_path):
    """A ref that resolves to no clip contributes 0 segments to both counters.

    It still gets its own "No clip found" line, but advancing on that line would
    push done past total and overfill the Generate button.
    """
    _setup_single_cell_generate(monkeypatch, tmp_path, "1:00", generated=1)

    resp = client.post(
        "/studio/api/generate", json={"cells": ["P01.5", "P09.99"], "format": "clip"}
    )
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]

    assert [line["cell"] for line in lines] == ["P01.5", "P09.99"]
    assert lines[1]["error"] == "No clip found"
    assert server._generate_job_state["total"] == 1
    assert server._generate_job_state["done"] == 1


def test_api_generate_ref_spelled_unlike_its_header_gets_one_line(
    client, monkeypatch, tmp_path
):
    """A resolved ref must not also draw a trailing "No clip found".

    The result line echoes the sheet header the ref resolved to, while the
    unmatched-ref sweep compares against the ref as posted — and the two only
    match case-insensitively. Exact-match comparison emitted both lines for one
    cell, which double-advanced the readout and painted succeeded cards failed.
    """
    _setup_single_cell_generate(monkeypatch, tmp_path, "1:00", generated=1)

    resp = client.post(
        "/studio/api/generate", json={"cells": ["p01.5"], "format": "clip"}
    )
    lines = [json.loads(line) for line in resp.data.decode().strip().split("\n")]

    assert [line["cell"] for line in lines] == ["P01.5"]
    assert lines[0]["ok"] is True
    assert server._generate_job_state["done"] == 1


def test_api_job_status_cancelling_flag(client, monkeypatch, tmp_path):
    """When cancel has been signaled but the worker hasn't yet exited, the
    status reflects in_progress=True + cancelling=True so the UI can show
    'Cancelling…' instead of allowing another Cancel click."""
    import threading

    _setup_api_reel(monkeypatch, tmp_path)

    started = threading.Event()
    proceed = threading.Event()

    def slow_proc(clips_list, **kwargs):
        progress_cb = kwargs.get("progress_cb")
        if progress_cb is not None:
            progress_cb({"phase": "start", "total_clips": 1})
        started.set()
        proceed.wait(timeout=5)
        return (0, [])

    monkeypatch.setattr("pipeline.process_reel", slow_proc)

    resp = client.post("/studio/api/reel", json={"cells": ["P01.5"]})
    try:
        assert started.wait(timeout=5)
        client.post("/studio/api/reel/cancel")

        status = client.get("/studio/api/job-status").get_json()
        assert status["reel"]["in_progress"] is True
        assert status["reel"]["cancelling"] is True
    finally:
        proceed.set()
        resp.close()

    assert _poll_until(lambda: server._busy_slots["reel"] is False)


def test_api_generate_explicit_cancel_still_works(client, monkeypatch, tmp_path):
    """POST /api/generate/cancel during a generate must terminate the run and
    surface a cancelled event in the stream."""
    import threading
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    _set_artifacts(monkeypatch, [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr("titlecards.clear_endcard_cache", lambda: None)
    monkeypatch.setattr("pipeline._resolve_clip_workers", lambda: 1)  # sequential

    cells = [types.SimpleNamespace(row=r, col=2, value="1:00") for r in (5, 6)]
    monkeypatch.setattr(
        "spreadsheet.parse_cell_specifications", lambda text: [("P01", 5), ("P01", 6)]
    )
    monkeypatch.setattr(
        "spreadsheet.generate_list",
        lambda ws, mode, *, ctx=None, cell_specs, skip_prompts: [
            {"participant": "P01", "cell": c} for c in cells
        ],
    )

    started = threading.Event()

    def one_clip(clip_list, **kwargs):
        started.set()
        cancel_flag = kwargs.get("cancel_flag")
        if cancel_flag and cancel_flag():
            return (0, [])
        return (1, [{"id": "a", "type": "clip", "file": "x.mp4"}])

    monkeypatch.setattr("pipeline.process_clips", one_clip)
    server._generate_cancel_event.clear()

    resp = client.post(
        "/studio/api/generate", json={"cells": ["P01.5", "P01.6"], "format": "clip"}
    )
    # `client.post` does not run the stream in the background: it returns having
    # buffered the first cell's chunk, and `resp.data` below drains the rest. So
    # the cancel is posted between the two cells by program order, and the route
    # sees it when it checks between clips — no timing window to hit. This used to
    # spin a 150 x 10 ms busy-loop here to "post a cancel mid-run"; the cancel
    # provably landed only after that loop finished, so the 2 s bought nothing.
    # If Werkzeug ever drained both cells inside post(), the `cancelled`
    # assertion below fails loudly rather than passing vacuously.
    assert started.wait(timeout=5)

    cancel_resp = client.post("/studio/api/generate/cancel")
    assert cancel_resp.status_code == 200

    # The stream must eventually emit a cancelled marker and close.
    text = resp.data.decode()
    lines = [json.loads(ln) for ln in text.strip().split("\n") if ln.strip()]
    assert any(ln.get("cancelled") is True for ln in lines)
    assert _poll_until(lambda: server._busy_slots["generate"] is False)


# ---- /api/reel-direct titlecards ----


def _drain_ndjson(resp):
    """Drain an NDJSON streaming response into a list of parsed dicts."""
    text = resp.data.decode()
    return [json.loads(ln) for ln in text.strip().split("\n") if ln.strip()]


def test_api_generate_discards_artifacts_completed_after_cancel(
    client, monkeypatch, tmp_path
):
    """A worker whose ffmpeg wraps up just as the cancel event is set must
    drop its produced artifact (no manifest append, file unlinked)."""
    import types

    monkeypatch.setattr(server, "_worksheet", object())
    monkeypatch.setattr(server, "_generated_artifacts", [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr(server, "_find_existing_artifacts", lambda *a, **kw: [])
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))

    cell = types.SimpleNamespace(row=5, col=2)

    def fake_generate_list(ws, mode, **kwargs):
        return [{"participant": "P01", "cell": cell, "times": [("0:00", "0:05")]}]

    monkeypatch.setattr("spreadsheet.generate_list", fake_generate_list)

    file_a = tmp_path / "a.mp4"
    file_a.write_bytes(b"x")

    # Worker returns its artifact and ALSO trips the cancel event mid-call —
    # simulates a worker whose ffmpeg subprocess finished successfully just
    # as the cancel signal was set. The drop-after-cancel guard inside
    # _generate_and_persist should reject the artifact and unlink the file.
    def fake_process_clips(clips, *, cancel_flag=None, **kwargs):
        server._generate_cancel_event.set()
        return (1, [{"id": "a", "type": "clip", "file": "a.mp4"}])

    monkeypatch.setattr("pipeline.process_clips", fake_process_clips)
    server._generate_cancel_event.clear()

    resp = client.post(
        "/studio/api/generate",
        json={"cells": ["P01.5"], "format": "clip"},
    )
    resp.data  # drain stream
    server._generate_cancel_event.clear()

    assert _poll_until(lambda: server._busy_slots["generate"] is False)
    ids = [a.get("id") for a in server._generated_artifacts]
    assert "a" not in ids
    assert not file_a.exists()


# ---- /api/generate-intake cancellation ----


def test_api_job_status_reflects_intake_progress(client, monkeypatch, tmp_path):
    """While /api/generate-intake is running, job-status reports intake progress."""
    import threading

    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    proceed = threading.Event()
    started = threading.Event()

    def slow_item(item, output_format, study, index=0, *, cancel_flag=None):
        started.set()
        proceed.wait(timeout=5)  # hold worker until status is polled
        return {
            "id": f"intake_{index}",
            "type": "clip",
            "file": f"clip_{index}.mp4",
            "study": study,
            "participant": item["participant"],
            "_ok": True,
            "_error": "",
        }

    monkeypatch.setattr(server, "_process_intake_item", slow_item)
    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["vid.mp4"]
    )
    monkeypatch.setattr("pipeline._resolve_clip_workers", lambda: 1)

    items = [{"participant": "P01", "start": 0, "end": 5, "source": "screenspace"}]
    stream_error = []

    def run_intake():
        try:
            with client.post(
                "/studio/api/generate-intake",
                json={"items": items, "format": "clip"},
                buffered=False,
            ) as resp:
                list(resp.iter_encoded())
        except Exception as exc:
            stream_error.append(exc)

    worker = threading.Thread(target=run_intake)
    worker.start()
    try:
        assert started.wait(timeout=5)
        status = client.get("/studio/api/job-status").get_json()
        assert status["intake"]["in_progress"] is True
        assert status["intake"]["total"] == 1
        assert status["intake"]["done"] == 0
        proceed.set()
    finally:
        worker.join(timeout=10)
    assert not stream_error

    assert _poll_until(lambda: server._intake_active == 0)
    final = client.get("/studio/api/job-status").get_json()
    assert final["intake"]["in_progress"] is False


def test_api_generate_intake_cancel_endpoint_returns_ok(client):
    """POST /api/generate-intake/cancel sets the cancel event and returns ok."""
    server._intake_cancel_event.clear()
    resp = client.post("/studio/api/generate-intake/cancel")
    assert resp.status_code == 200
    assert resp.get_json() == {"ok": True}
    assert server._intake_cancel_event.is_set() is True
    server._intake_cancel_event.clear()


def test_api_generate_intake_passes_cancel_flag_to_run_ffmpeg(
    client, monkeypatch, tmp_path
):
    """_process_intake_item must forward cancel_flag into video.run_ffmpeg
    so an in-flight ffmpeg encode can be terminated."""
    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server, "_generated_artifacts", [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr(
        "files.get_unique_filename", lambda name, file_format=None: name
    )

    captured = {}

    def fake_run_ffmpeg(*args, **kwargs):
        captured["cancel_flag"] = kwargs.get("cancel_flag")
        return True

    monkeypatch.setattr("video.run_ffmpeg", fake_run_ffmpeg)
    server._intake_cancel_event.clear()

    resp = client.post(
        "/studio/api/generate-intake",
        json={"items": [{"participant": "P01", "start": 0, "end": 5}]},
    )
    resp.data
    cb = captured.get("cancel_flag")
    assert callable(cb)
    # Verify the callable tracks the intake cancel event.
    assert cb() is False
    server._intake_cancel_event.set()
    try:
        assert cb() is True
    finally:
        server._intake_cancel_event.clear()


def test_api_generate_intake_short_circuits_after_cancel(client, monkeypatch, tmp_path):
    """When cancel is signaled before the worker starts, no ffmpeg is
    invoked and a {cancelled: true} marker is yielded."""
    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server, "_generated_artifacts", [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr(
        "files.get_unique_filename", lambda name, file_format=None: name
    )

    ffmpeg_calls = []

    def fake_run_ffmpeg(*args, **kwargs):
        ffmpeg_calls.append(args)
        return True

    monkeypatch.setattr("video.run_ffmpeg", fake_run_ffmpeg)

    # Set cancel event inside _process_intake_item to simulate cancel arriving
    # during the first worker. Use a wrapper that flips the event then calls
    # through to the real function.
    real_process = server._process_intake_item

    def wrapped(item, output_format, study, index=0, *, cancel_flag=None):
        server._intake_cancel_event.set()
        return real_process(
            item, output_format, study, index=index, cancel_flag=cancel_flag
        )

    monkeypatch.setattr(server, "_process_intake_item", wrapped)
    server._intake_cancel_event.clear()

    try:
        resp = client.post(
            "/studio/api/generate-intake",
            json={
                "items": [
                    {"participant": "P01", "start": 0, "end": 5},
                    {"participant": "P01", "start": 10, "end": 15},
                    {"participant": "P01", "start": 20, "end": 25},
                ]
            },
        )
        text = resp.data.decode()
    finally:
        server._intake_cancel_event.clear()

    lines = [json.loads(ln) for ln in text.strip().split("\n") if ln.strip()]
    assert any(ln.get("cancelled") is True for ln in lines), lines
    # Only the first item should have invoked ffmpeg; remaining items must
    # short-circuit via the pre-ffmpeg cancel check.
    assert len(ffmpeg_calls) <= 1


def test_api_reel_direct_wraps_segments_when_titlecards_enabled(
    client, monkeypatch, tmp_path
):
    """Each segment must be wrapped via titlecards.wrap_clip_with_cards with the
    per-request duration before the concat list is assembled."""
    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server, "_generated_reels", [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr("video.run_ffmpeg", lambda *a, **kw: True)
    monkeypatch.setattr("video.concatenate_clips", lambda *a, **kw: True)
    monkeypatch.setattr(
        "files.get_unique_filename", lambda name, file_format=None: name
    )

    wrap_calls: list[dict] = []

    def fake_wrap(clip, clip_path, *, resolution=None, **kwargs):
        wrap_calls.append(
            {
                "desc": clip.get("desc"),
                "resolution": resolution,
                "duration": kwargs.get("titlecard_duration_seconds"),
                "enabled": kwargs.get("titlecards_enabled"),
            }
        )
        return (True, True)

    monkeypatch.setattr("titlecards.wrap_clip_with_cards", fake_wrap)
    server._reel_cancel_event.clear()

    resp = client.post(
        "/studio/api/reel-direct",
        json={
            "segments": [
                {
                    "participant": "P01",
                    "start": 0,
                    "end": 5,
                    "event_type": "first",
                },
                {
                    "participant": "P01",
                    "start": 10,
                    "end": 20,
                    "event_type": "second",
                },
            ],
            "titlecards_enabled": True,
            "titlecard_duration": 3,
        },
    )
    _drain_ndjson(resp)
    assert _poll_until(lambda: server._busy_slots["reel"] is False)
    assert len(wrap_calls) == 2
    assert wrap_calls[0]["duration"] == 3
    assert wrap_calls[0]["enabled"] is True
    # No forced resolution: wrap_clip_with_cards probes the cut clip itself, so a
    # span cut from a later source part is wrapped at that part's resolution.
    assert wrap_calls[0]["resolution"] is None
    assert wrap_calls[0]["desc"] == "first"
    assert wrap_calls[1]["desc"] == "second"


def test_api_reel_direct_skips_wrap_when_titlecards_disabled(
    client, monkeypatch, tmp_path
):
    """No wrap_clip_with_cards calls when titlecards_enabled is False."""
    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server, "_generated_reels", [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr("video.run_ffmpeg", lambda *a, **kw: True)
    monkeypatch.setattr("video.concatenate_clips", lambda *a, **kw: True)
    monkeypatch.setattr(
        "files.get_unique_filename", lambda name, file_format=None: name
    )

    wrap_calls: list[dict] = []
    monkeypatch.setattr(
        "titlecards.wrap_clip_with_cards",
        lambda *a, **kw: wrap_calls.append(kw) or True,
    )
    server._reel_cancel_event.clear()

    resp = client.post(
        "/studio/api/reel-direct",
        json={
            "segments": [{"participant": "P01", "start": 0, "end": 5}],
            "titlecards_enabled": False,
        },
    )
    _drain_ndjson(resp)
    assert _poll_until(lambda: server._busy_slots["reel"] is False)
    assert wrap_calls == []


def test_api_reel_direct_clears_endcard_cache(client, monkeypatch, tmp_path):
    """The worker's finally block must purge the endcard cache so per-request
    titlecard temp files don't leak between reel builds."""
    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("config.OUTPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server, "_generated_reels", [])
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr("video.run_ffmpeg", lambda *a, **kw: True)
    monkeypatch.setattr("video.concatenate_clips", lambda *a, **kw: True)
    monkeypatch.setattr(
        "files.get_unique_filename", lambda name, file_format=None: name
    )

    clear_calls = []
    monkeypatch.setattr("titlecards.clear_endcard_cache", lambda: clear_calls.append(1))
    server._reel_cancel_event.clear()

    resp = client.post(
        "/studio/api/reel-direct",
        json={"segments": [{"participant": "P01", "start": 0, "end": 5}]},
    )
    _drain_ndjson(resp)
    assert _poll_until(lambda: server._busy_slots["reel"] is False)
    assert len(clear_calls) == 1


# ── Card scrubber media endpoints (opt-in hover scrub) ──────────────────────


def test_card_scrubber_flag_in_sheet_payload(client):
    """/api/sheet exposes the card-scrubber flag for the frontend toggle."""
    data = client.get("/studio/api/sheet").get_json()
    assert data["cardScrubberEnabled"] is False


def test_metadata_cluster_flag_in_sheet_payload(client):
    """/api/sheet exposes the Metadata clustering flag (default on)."""
    data = client.get("/studio/api/sheet").get_json()
    assert data["metadataClusterScreenspace"] is True


def test_api_settings_includes_metadata_cluster_toggle(client):
    """GET /api/settings exposes the Metadata clustering toggle (General tab)."""
    data = client.get("/studio/api/settings").get_json()
    by_name = {s["name"]: s for s in data["settings"]}
    s = by_name["STUDIO_METADATA_CLUSTER_SCREENSPACE"]
    assert s["tab"] == "General"
    assert s["group"] == "Metadata Overview"
    assert s["type"] == "bool"
    assert s["default"] is True


@pytest.mark.parametrize(
    "path", ["/studio/api/sprite/P01", "/studio/api/clip-audio/P01"]
)
def test_scrubber_media_404_when_no_sheet(client, path):
    resp = client.get(path + "?start=0&end=5")
    assert resp.status_code == 404
    data = resp.get_json()
    assert data["ok"] is False
    assert "No spreadsheet loaded" in data["error"]


@pytest.mark.parametrize(
    "path", ["/studio/api/sprite/P01", "/studio/api/clip-audio/P01"]
)
@pytest.mark.parametrize(
    "query", ["", "?start=5&end=5", "?start=5&end=2", "?start=abc&end=9"]
)
def test_scrubber_media_400_on_bad_window(client, monkeypatch, path, query):
    monkeypatch.setattr(server, "_sheet_context", object())
    resp = client.get(path + query)
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


def test_api_sprite_happy_path(client, monkeypatch):
    from pathlib import Path

    monkeypatch.setattr(server, "_sheet_context", object())
    monkeypatch.setattr(
        server, "_resolve_clip_media_source", lambda p, s: (Path("/nope.mp4"), 0.0)
    )
    monkeypatch.setattr(
        server.video, "extract_sprite_sheet_bytes", lambda *a, **k: b"JPEGDATA"
    )
    resp = client.get("/studio/api/sprite/P01?start=0&end=5")
    assert resp.status_code == 200
    assert resp.mimetype == "image/jpeg"
    assert resp.headers["Cache-Control"] == "public, max-age=86400"
    assert resp.data == b"JPEGDATA"


def test_api_clip_audio_happy_path(client, monkeypatch):
    from pathlib import Path

    monkeypatch.setattr(server, "_sheet_context", object())
    monkeypatch.setattr(
        server, "_resolve_clip_media_source", lambda p, s: (Path("/nope.mp4"), 0.0)
    )
    monkeypatch.setattr(
        server.video, "extract_audio_segment_bytes", lambda *a, **k: b"WAVDATA"
    )
    resp = client.get("/studio/api/clip-audio/P01?start=0&end=5")
    assert resp.status_code == 200
    assert resp.mimetype == "audio/wav"
    assert resp.headers["Cache-Control"] == "public, max-age=86400"
    assert resp.data == b"WAVDATA"


def test_static_cache_headers():
    """HTML/JS/CSS revalidate every request (no-cache + 304s); svg keeps a TTL.

    Regression guards: (1) ``send_from_directory`` defaults to a bare
    ``no-cache``, which the ``_set_cache_headers`` after_request hook must
    normalize rather than treat as a deliberate cache header (server.py);
    (2) JS/CSS must NOT get a max-age — a TTL once let browsers run hour-old
    page scripts against fresh HTML after an update. The ``client`` fixture
    omits the hook, so register it explicitly here.
    """
    app = Flask(__name__)
    app.register_blueprint(server.studio_bp, url_prefix="/studio")
    app.after_request(server._set_cache_headers)
    with app.test_client() as c:
        assert c.get("/studio/utils.js").headers["Cache-Control"] == "no-cache"
        assert c.get("/studio/tokens.css").headers["Cache-Control"] == "no-cache"
        assert (
            c.get("/studio/icons/x-mark.svg").headers["Cache-Control"]
            == "public, max-age=86400"
        )
        assert c.get("/studio/").headers["Cache-Control"] == "no-cache"
        # Conditional requests answer 304 so no-cache stays cheap on localhost.
        etag = c.get("/studio/utils.js").headers.get("ETag")
        assert etag
        resp = c.get("/studio/utils.js", headers={"If-None-Match": etag})
        assert resp.status_code == 304


def test_media_route_serves_generated_artifacts(client, tmp_path, monkeypatch):
    """/studio/media/<file> serves the *current* output dir (the Overview
    Reports tab plays generated clips from it). The getter resolves per
    request, so a mid-session OUTPUT_DIR change is picked up immediately —
    unlike screenspace's startup-snapshot media dir."""
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(tmp_path))
    (tmp_path / "Study P01 clip.mp4").write_bytes(b"not-really-mp4")
    resp = client.get("/studio/media/Study%20P01%20clip.mp4")
    assert resp.status_code == 200
    assert resp.data == b"not-really-mp4"

    other = tmp_path / "elsewhere"
    other.mkdir()
    (other / "moved.mp4").write_bytes(b"x")
    monkeypatch.setattr(server.config, "OUTPUT_DIR", str(other))
    assert client.get("/studio/media/moved.mp4").status_code == 200
    assert client.get("/studio/media/Study%20P01%20clip.mp4").status_code == 404


def test_api_generate_intake_resolves_mark_ids_to_segment_text(client, monkeypatch):
    """Marks name segments; the fallback joins those segments' text."""
    import transcripts_server

    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("video.run_ffmpeg", lambda *a, **kw: True)
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr(
        transcripts_server,
        "_manifest",
        {
            "source_transcripts": {
                "P01": {
                    "transcribed_at": "2026-01-01T00:00:00+00:00",
                    "segments": [
                        {"id": "P01:0", "text": "hello there"},
                        {"id": "P01:1", "text": "not marked"},
                    ],
                }
            },
            "marks": [{"id": "m_1", "segment_id": "P01:0"}],
        },
    )

    items = [
        {
            "participant": "P01",
            "start": 0.0,
            "end": 5.0,
            "event_type": "transcript",
            "event_ids": [],
            "source": "transcript",
            "mark_ids": ["m_1"],
        }
    ]
    resp = client.post(
        "/studio/api/generate-intake", json={"items": items, "format": "clip"}
    )
    assert resp.status_code == 200
    line = json.loads(resp.data.decode().strip().split("\n")[0])
    assert line["ok"] is True
    assert line["artifact"]["transcriptText"] == "hello there"


def test_api_reel_releases_busy_slot_when_stream_never_starts(
    studio_app, client, monkeypatch, tmp_path
):
    """A response torn down before iteration must not hold the slot forever."""
    _setup_api_reel(monkeypatch, tmp_path)
    endpoint = next(
        rule.endpoint
        for rule in studio_app.url_map.iter_rules()
        if rule.rule == "/studio/api/reel" and "POST" in (rule.methods or set())
    )
    view = studio_app.view_functions[endpoint]

    with studio_app.test_request_context(
        "/studio/api/reel", method="POST", json={"cells": ["P01.5"]}
    ):
        resp = view()
        assert server._busy_slots["reel"] is True
        resp.close()

    assert server._busy_slots["reel"] is False


def test_release_busy_ignores_a_stale_token():
    """A late second release must not drop the next request's claim."""
    server._release_busy("generate")
    first = server._try_claim_busy("generate")
    assert first
    server._release_busy("generate", first)
    second = server._try_claim_busy("generate")
    assert second and second != first

    server._release_busy("generate", first)  # stale: the generator's finally
    assert server._busy_slots["generate"] is True
    assert server._try_claim_busy("generate") is None

    server._release_busy("generate", second)
    assert server._busy_slots["generate"] is False
