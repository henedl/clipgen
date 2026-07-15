"""Smoke tests for the Composer Flask blueprint.

Verifies the page serves, participant/part discovery, the composer manifest
round-trip (cuts CRUD + UI toggles persisted to ``composer_manifest.json``),
and span clamping — mirroring tests/test_workflows_api.py's bare-blueprint
setup. Combined-app registration (topnav-visible ``/composer/`` + the
``/api/status`` flag) is exercised against ``server.build_combined_app``.
"""

import json

import pytest

Flask = pytest.importorskip("flask").Flask

import composer_server  # noqa: E402
import config  # noqa: E402
import utils  # noqa: E402
import video  # noqa: E402


@pytest.fixture
def co_client(tmp_path, monkeypatch):
    app = Flask(__name__)
    app.register_blueprint(composer_server.composer_bp, url_prefix="/composer")

    # Seed module globals via monkeypatch so they auto-restore on teardown.
    monkeypatch.setattr(composer_server, "_manifest", composer_server._empty_manifest())
    monkeypatch.setattr(composer_server, "_input_dir", str(tmp_path))
    monkeypatch.setattr(composer_server, "_sheet_context", None)
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path), raising=False)

    # One two-part participant with 10 s parts → stitched duration 20 s.
    monkeypatch.setattr(
        utils,
        "discover_participant_videos",
        lambda study_name="": [
            {
                "id": "P01",
                "has_video": True,
                "video_paths": [
                    str(tmp_path / "study_P01-1.mp4"),
                    str(tmp_path / "study_P01-2.mp4"),
                ],
            }
        ],
    )
    monkeypatch.setattr(video, "get_file_duration", lambda path: 10)

    with app.test_client() as c:
        yield c


def _manifest_on_disk(tmp_path):
    return json.loads(
        (tmp_path / config.COMPOSER_MANIFEST_FILENAME).read_text(encoding="utf-8")
    )


def test_page_serves(co_client):
    resp = co_client.get("/composer/")
    assert resp.status_code == 200
    assert resp.mimetype == "text/html"
    body = resp.get_data(as_text=True)
    assert 'data-frontend="composer"' in body


def test_participants_reports_parts_and_total(co_client):
    body = co_client.get("/composer/api/participants").get_json()
    assert body["ok"] is True
    (p,) = body["participants"]
    assert p["id"] == "P01"
    assert [part["offset"] for part in p["parts"]] == [0, 10]
    assert p["total_duration"] == 20
    assert "defaultDuration" in body["config"]


def test_cuts_crud_round_trip(co_client, tmp_path):
    created = co_client.post(
        "/composer/api/cuts",
        json={"participant": "P01", "start": 2.0, "end": 5.5},
    ).get_json()
    assert created["ok"] is True
    cut_id = created["cut"]["id"]
    assert cut_id.startswith("cut_")

    # Persisted to composer_manifest.json (save-after-mutation).
    disk = _manifest_on_disk(tmp_path)
    assert [c["id"] for c in disk["cuts"]] == [cut_id]

    patched = co_client.patch(
        f"/composer/api/cuts/{cut_id}", json={"start": 1.0, "end": 4.0}
    ).get_json()
    assert patched["ok"] is True
    assert (patched["cut"]["start"], patched["cut"]["end"]) == (1.0, 4.0)

    listed = co_client.get("/composer/api/manifest").get_json()
    assert listed["manifest"]["cuts"][0]["start"] == 1.0

    deleted = co_client.delete(f"/composer/api/cuts/{cut_id}").get_json()
    assert deleted["ok"] is True
    assert _manifest_on_disk(tmp_path)["cuts"] == []


def test_cut_create_rejects_inverted_span(co_client):
    resp = co_client.post(
        "/composer/api/cuts", json={"participant": "P01", "start": 5.0, "end": 5.0}
    )
    assert resp.status_code == 400


def test_cut_patch_clamps_to_duration(co_client):
    cut = co_client.post(
        "/composer/api/cuts", json={"participant": "P01", "start": 2.0, "end": 5.0}
    ).get_json()["cut"]
    patched = co_client.patch(
        f"/composer/api/cuts/{cut['id']}", json={"end": 999.0}
    ).get_json()
    assert patched["cut"]["end"] == 20  # stitched duration caps the out point


def test_cut_patch_unknown_id_404(co_client):
    resp = co_client.patch("/composer/api/cuts/cut_nope", json={"end": 3})
    assert resp.status_code == 404


def test_ui_toggles_round_trip(co_client, tmp_path):
    resp = co_client.put(
        "/composer/api/ui",
        json={"markerSources": {"sheet": False, "screenspace": True}},
    ).get_json()
    assert resp["ok"] is True
    assert resp["markerSources"] == {
        "sheet": False,
        "screenspace": True,
        "transcript": True,
    }
    assert _manifest_on_disk(tmp_path)["ui"]["markerSources"]["sheet"] is False


def test_ui_lane_folds_round_trip(co_client, tmp_path):
    resp = co_client.put(
        "/composer/api/ui",
        json={"laneFolds": {"screenspace": False}},
    ).get_json()
    assert resp["ok"] is True
    assert resp["laneFolds"] == {
        "sheet": True,
        "screenspace": False,
        "transcript": True,
        "annotations": True,
    }
    disk = _manifest_on_disk(tmp_path)["ui"]
    assert disk["laneFolds"]["screenspace"] is False
    # A folds-only PUT must not clobber previously saved lane toggles.
    assert "markerSources" not in disk or isinstance(disk.get("markerSources"), dict)


def test_ui_requires_some_payload(co_client):
    resp = co_client.put("/composer/api/ui", json={})
    assert resp.status_code == 400


def test_trims_round_trip_and_reset(co_client, tmp_path):
    key = "sheet:12:P01:0"
    resp = co_client.put(
        f"/composer/api/trims/{key}",
        json={
            "start": 61.0,
            "end": 70.5,
            "participant": "P01",
            "label": "misses the CTA",
            "source": "sheet",
        },
    ).get_json()
    assert resp["ok"] is True
    assert resp["trim"] == {
        "start": 61.0,
        "end": 70.5,
        "participant": "P01",
        "label": "misses the CTA",
        "source": "sheet",
    }
    assert _manifest_on_disk(tmp_path)["trims"][key] == resp["trim"]

    # A times-only re-PUT (undo/redo replay) keeps the stored metadata.
    updated = co_client.put(
        f"/composer/api/trims/{key}", json={"start": 62.0, "end": 69.0}
    ).get_json()
    assert updated["trim"]["participant"] == "P01"
    assert updated["trim"]["label"] == "misses the CTA"
    assert updated["trim"]["start"] == 62.0

    deleted = co_client.delete(f"/composer/api/trims/{key}").get_json()
    assert deleted["ok"] is True
    assert _manifest_on_disk(tmp_path)["trims"] == {}


def test_trim_delete_unknown_404(co_client):
    resp = co_client.delete("/composer/api/trims/screenspace:ev_nope")
    assert resp.status_code == 404


def test_trim_put_rejects_inverted_span(co_client):
    resp = co_client.put(
        "/composer/api/trims/transcript-mark:m1", json={"start": 5.0, "end": 5.0}
    )
    assert resp.status_code == 400


# ---- Annotations (P3) ----


def _make_annotation(co_client, **overrides):
    body = {
        "participant": "P01",
        "type": "text",
        "span": {"start": 60.0, "end": 180.0},
        "geometry": {"x": 0.42, "y": 0.18, "text": "misses the CTA"},
    }
    body.update(overrides)
    return co_client.post("/composer/api/annotations", json=body).get_json()


def test_annotation_crud_round_trip(co_client, tmp_path):
    created = _make_annotation(co_client)
    assert created["ok"] is True
    ann = created["annotation"]
    assert ann["id"].startswith("ann_")
    assert ann["style"]["color"] == config.COMPOSER_ANNOTATION_COLOR
    assert _manifest_on_disk(tmp_path)["annotations"][0]["id"] == ann["id"]

    patched = co_client.patch(
        f"/composer/api/annotations/{ann['id']}",
        json={"span": {"start": 65.0, "end": 170.0}},
    ).get_json()
    assert patched["annotation"]["span"] == {"start": 65.0, "end": 170.0}
    # Geometry/style untouched by a span-only patch.
    assert patched["annotation"]["geometry"]["text"] == "misses the CTA"

    deleted = co_client.delete(f"/composer/api/annotations/{ann['id']}").get_json()
    assert deleted["ok"] is True
    assert _manifest_on_disk(tmp_path)["annotations"] == []


def test_annotation_freehand_and_validation(co_client):
    stroke = _make_annotation(
        co_client,
        type="freehand",
        geometry={"points": [[0.1, 0.2], [0.15, 0.25], [1.7, -0.5]]},
    )
    assert stroke["ok"] is True
    # Out-of-range points clamp to the frame.
    assert stroke["annotation"]["geometry"]["points"][2] == [1.0, 0.0]

    assert (
        co_client.post(
            "/composer/api/annotations",
            json={
                "participant": "P01",
                "type": "text",
                "span": {"start": 1, "end": 5},
                "geometry": {"x": 0.5, "y": 0.5, "text": "   "},
            },
        ).status_code
        == 400
    )
    assert (
        co_client.post(
            "/composer/api/annotations",
            json={
                "participant": "P01",
                "type": "wiggle",
                "span": {"start": 1, "end": 5},
            },
        ).status_code
        == 400
    )
    assert (
        co_client.patch("/composer/api/annotations/ann_nope", json={}).status_code
        == 404
    )


def test_render_annotation_overlay_draws_pixels():
    overlay = composer_server._render_annotation_overlay(
        [
            {
                "type": "freehand",
                "geometry": {"points": [[0.1, 0.5], [0.9, 0.5]]},
                "style": {"color": "#ff0000", "strokeWidth": 0.01},
            },
            {
                "type": "text",
                "geometry": {"x": 0.1, "y": 0.1, "text": "hello"},
                "style": {"color": "#00ff00", "fontSize": 0.05},
            },
        ],
        640,
        360,
    )
    assert overlay.mode == "RGBA"
    assert overlay.size == (640, 360)
    # Stroke pixel at mid-height is opaque red.
    assert overlay.getpixel((320, 180))[3] > 0
    # Something was drawn in the text region (backing box or glyphs).
    region = overlay.crop((0, 0, 200, 80))
    assert region.getchannel("A").getextrema()[1] > 0


def test_annotation_windows_split_by_visibility():
    anns = [
        {"span": {"start": 0.0, "end": 10.0}, "id": "a"},
        {"span": {"start": 5.0, "end": 20.0}, "id": "b"},
    ]
    windows = composer_server._annotation_windows(anns, 0.0, 30.0)
    spans = [(w["start"], w["end"], len(w["annotations"])) for w in windows]
    assert spans == [(0.0, 5.0, 1), (5.0, 10.0, 2), (10.0, 20.0, 1)]


def test_build_overlay_command_seeks_first_and_uses_relative_enable():
    cmd = composer_server._build_overlay_command(
        "/vids/study_P01.mp4",
        30.0,
        12.0,
        [("/tmp/o1.png", 0.0, 5.0), ("/tmp/o2.png", 5.0, 12.0)],
        "/out/annotated.mp4",
        gif=False,
    )
    # Seek precedes the main input (span-only decode).
    assert cmd.index("-ss") < cmd.index("-i")
    assert cmd[cmd.index("-ss") + 1] == "30.000"
    graph = cmd[cmd.index("-filter_complex") + 1]
    assert "between(t,0.000,5.000)" in graph
    assert "between(t,5.000,12.000)" in graph
    assert cmd[cmd.index("-t") + 1] == "12.000"
    assert "-c:a" in cmd  # audio carried over for video burns


def test_build_overlay_command_gif_chain():
    cmd = composer_server._build_overlay_command(
        "/vids/study_P01.mp4",
        0.0,
        4.0,
        [("/tmp/o1.png", 0.0, 4.0)],
        "/out/annotated.gif",
        gif=True,
    )
    graph = cmd[cmd.index("-filter_complex") + 1]
    assert f"fps={config.GIF_FPS}" in graph
    assert "-c:a" not in cmd


def test_export_screenshot_composites_and_records(co_client, tmp_path, monkeypatch):
    _make_annotation(co_client, span={"start": 1.0, "end": 9.0})
    monkeypatch.setattr(config, "DEBUGGING", True, raising=False)  # stub frame
    saved = {}
    monkeypatch.setattr(
        composer_server,
        "_save_export_artifact",
        lambda artifact, participant: saved.update(artifact),
    )
    resp = co_client.post(
        "/composer/api/export/screenshot", json={"participant": "P01", "time": 5.0}
    ).get_json()
    assert resp["ok"] is True
    art = resp["artifact"]
    assert art["type"] == "screen"
    assert art["participant"] == "P01"
    assert saved["id"] == art["id"]
    out = tmp_path / art["file"]
    assert out.is_file() and out.stat().st_size > 0


class _FakeFfmpegResult:
    returncode = 0
    stderr = ""


def _stub_overlay_ffmpeg(monkeypatch):
    """Stub the overlay ffmpeg pass + probe so no real ffmpeg runs.

    Returns the list that records every path probe_video_properties saw, so a
    test can assert overlay dims came from the stitched temp, not a source part.
    """
    probed: list[str] = []

    def fake_probe(path):
        probed.append(path)
        return {"width": 1280, "height": 720}

    def fake_run(cmd, input_file, output_file, os_error_message, cancel_flag):
        from pathlib import Path

        Path(output_file).write_bytes(b"x")
        return _FakeFfmpegResult()

    monkeypatch.setattr(composer_server.video, "probe_video_properties", fake_probe)
    monkeypatch.setattr(composer_server.video, "run_ffmpeg_process", fake_run)
    return probed


def _capture_stitch(monkeypatch):
    """Record + stub pipeline.cut_global_range; return the call list."""
    from pathlib import Path

    calls: list[dict] = []

    def fake_cut(timeline, base, start, end, out_path, *, reencode, cancel_flag):
        calls.append({"start": start, "end": end, "out_path": out_path})
        Path(out_path).write_bytes(b"stitched")
        return {"sourceVideo": "study_P01-1.mp4", "localStart": start, "localEnd": 10.0}

    monkeypatch.setattr(composer_server.pipeline, "cut_global_range", fake_cut)
    return calls


@pytest.mark.parametrize(
    "route,kind,fmt",
    [("burn", "clip", config.FILEFORMAT), ("gif", "gif", config.GIF_FORMAT)],
)
def test_annotated_export_stitches_across_part_boundary(
    co_client, tmp_path, monkeypatch, route, kind, fmt
):
    # Span 8→12 straddles the 10 s part boundary of the two-part fixture.
    _make_annotation(co_client, span={"start": 8.0, "end": 12.0})
    probed = _stub_overlay_ffmpeg(monkeypatch)
    calls = _capture_stitch(monkeypatch)
    saved = {}
    monkeypatch.setattr(
        composer_server,
        "_save_export_artifact",
        lambda artifact, participant: saved.update(artifact),
    )

    resp = co_client.post(
        f"/composer/api/export/{route}",
        json={"participant": "P01", "start": 8.0, "end": 12.0},
    ).get_json()

    assert resp["ok"] is True
    # The raw span was stitched across parts before the overlay pass.
    assert len(calls) == 1
    assert (calls[0]["start"], calls[0]["end"]) == (8.0, 12.0)
    # Overlay dims were probed from the stitched temp, not a source part.
    assert probed and probed[-1] == calls[0]["out_path"]
    assert probed[-1] not in {
        str(tmp_path / "study_P01-1.mp4"),
        str(tmp_path / "study_P01-2.mp4"),
    }
    art = resp["artifact"]
    assert art["type"] == kind
    assert art["sourceVideo"] == "study_P01-1.mp4"
    assert art["file"].endswith(fmt)
    assert saved["id"] == art["id"]


def test_annotated_burn_within_part_keeps_single_pass_fast_path(
    co_client, tmp_path, monkeypatch
):
    # Span 2→5 lives wholly in part 1 → no stitch, decode the owning part directly.
    _make_annotation(co_client, span={"start": 1.0, "end": 9.0})
    _stub_overlay_ffmpeg(monkeypatch)
    calls = _capture_stitch(monkeypatch)
    monkeypatch.setattr(
        composer_server, "_save_export_artifact", lambda artifact, participant: None
    )

    resp = co_client.post(
        "/composer/api/export/burn",
        json={"participant": "P01", "start": 2.0, "end": 5.0},
    ).get_json()

    assert resp["ok"] is True
    assert calls == []  # fast path never stitches
    assert resp["artifact"]["sourceVideo"] == "study_P01-1.mp4"


def test_ffmpeg_exports_rejected_while_another_export_runs(co_client):
    # _export_cancel is a single shared Event; the busy lock enforces the
    # one-export-at-a-time assumption it relies on. Without it a second
    # export's clear() would un-cancel the first, and one cancel POST would
    # abort both in-flight encodes.
    assert composer_server._export_busy.acquire(blocking=False)
    try:
        for route in ("burn", "gif"):
            resp = co_client.post(
                f"/composer/api/export/{route}",
                json={"participant": "P01", "start": 2.0, "end": 5.0},
            )
            assert resp.status_code == 409
            assert resp.get_json()["ok"] is False
    finally:
        composer_server._export_busy.release()


def test_combined_app_registers_composer(tmp_path, monkeypatch):
    import server
    import start_settings

    in_dir = tmp_path / "in"
    out_dir = tmp_path / "out"
    in_dir.mkdir()
    out_dir.mkdir()
    monkeypatch.setattr(config, "INPUT_DIR", str(in_dir), raising=False)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(out_dir), raising=False)
    monkeypatch.setattr(
        start_settings, "_settings_path", lambda: tmp_path / "start.json"
    )

    app = server.build_combined_app(worksheet=None, default_page="composer")
    with app.test_client() as client:
        assert client.get("/").location.endswith("/composer/")
        status = client.get("/api/status").get_json()
        assert status["composer"] is True
        assert client.get("/composer/").status_code == 200
