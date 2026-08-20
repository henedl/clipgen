"""Smoke tests for the Composer Flask blueprint.

Verifies the page serves, participant/part discovery, the composer manifest
round-trip (cuts CRUD + UI toggles persisted to ``composer_manifest.json``),
and span clamping — mirroring tests/test_workflows_api.py's bare-blueprint
setup. Combined-app registration (topnav-visible ``/composer/`` + the
``/api/status`` flag) is exercised against ``server.build_combined_app``.
"""

import json
from pathlib import Path

import pytest

Flask = pytest.importorskip("flask").Flask

import composer_server
import config
import utils
import video


@pytest.fixture(scope="module")
def co_app():
    """The Flask app, built once for the module.

    Registering the blueprint compiles ~23 Werkzeug URL rules, which dominates
    this fixture's cost — and the app object holds no per-test state: everything
    these tests touch lives in ``composer_server`` module globals, re-pinned per
    test by the function-scoped ``co_client`` below.
    """
    app = Flask(__name__)
    app.register_blueprint(composer_server.composer_bp, url_prefix="/composer")
    return app


@pytest.fixture
def co_client(co_app, tmp_path, monkeypatch):
    # Seed module globals via monkeypatch so they auto-restore on teardown.
    monkeypatch.setattr(composer_server, "_manifest", composer_server._empty_manifest())
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

    with co_app.test_client() as c:
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
    # No sheet loaded → everything is off-sheet, and has_sheet tells the frontend
    # to skip the badge rather than marking every participant.
    assert p["in_sheet"] is False
    assert body["has_sheet"] is False


def test_participants_fallback_parts_carry_offset(co_client, monkeypatch):
    """Unprobeable videos still serve offset-complete part entries.

    The client computes global time as ``currentTime + part.offset`` on every
    timeupdate; a fallback entry without the field turned the playhead NaN.
    """
    monkeypatch.setattr(video, "get_file_duration", lambda path: None)
    body = co_client.get("/composer/api/participants").get_json()
    (p,) = body["participants"]
    assert p["total_duration"] is None
    assert [part["offset"] for part in p["parts"]] == [0, 0]


def test_participants_prewarms_probes_before_the_loop(co_client, monkeypatch):
    """Every part is probed in one batch, not one ffprobe per participant.

    The per-entry probes are cache reads; serialized cold they measured 1.06 s
    for a 24-participant study and the page cannot render until the last one
    returns. Asserting the prewarm *saw every part path* is the invariant —
    timing it here would just measure the mock.
    """
    seen: list[list[str]] = []
    monkeypatch.setattr(video, "prewarm_probes", lambda paths: seen.append(list(paths)))

    body = co_client.get("/composer/api/participants").get_json()
    (p,) = body["participants"]

    assert len(seen) == 1, "prewarm must run once for the whole list, not per entry"
    assert [Path(path).name for path in seen[0]] == [
        part["name"] for part in p["parts"]
    ]


def test_prewarm_probes_dedupes_and_never_raises(monkeypatch):
    """Prewarming is pure optimization: it dedupes, and a failure is not an error.

    A path that cannot be probed simply stays uncached — the caller re-probes it
    and handles ``None`` as it always did. Raising here would turn a slow page
    into a broken one.
    """
    calls: list[str] = []

    def boom(path):
        calls.append(path)
        raise OSError("unprobeable")

    monkeypatch.setattr(video, "probe_video_properties", boom)
    # One path is already as fast as it gets, so the pool is skipped entirely.
    video.prewarm_probes(["/only.mp4"])
    assert calls == []

    video.prewarm_probes(["/a.mp4", "/a.mp4", "/b.mp4"])
    assert sorted(calls) == ["/a.mp4", "/b.mp4"]


def test_participants_reports_audio_tracks(co_client, monkeypatch):
    monkeypatch.setattr(
        video,
        "probe_video_properties",
        lambda _p: {
            "width": 1920,
            "height": 1080,
            "fps": 30.0,
            "audio_tracks": [
                {"index": 0, "label": "Microphone", "channels": 1},
                {"index": 1, "label": "System", "channels": 2},
            ],
            "audio_track_count": 2,
        },
    )
    body = co_client.get("/composer/api/participants").get_json()
    (p,) = body["participants"]
    assert p["audio_track_count"] == 2
    assert [t["label"] for t in p["audio_tracks"]] == ["Microphone", "System"]


def test_audio_track_streams(co_client, monkeypatch, tmp_path):
    track = tmp_path / "track.m4a"
    track.write_bytes(b"audio")
    monkeypatch.setattr(
        composer_server,
        "_find_participant_parts",
        lambda pid: [{"path": "/x.mp4"}] if pid == "P01" else None,
    )
    monkeypatch.setattr(video, "extract_audio_track", lambda p, idx: track)

    resp = co_client.get("/composer/api/audio-track/P01/0")
    assert resp.status_code == 200
    assert resp.mimetype == "audio/mp4"
    assert co_client.get("/composer/api/audio-track/ZZ/0").status_code == 404


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


def test_cut_patch_start_zero_survives(co_client):
    """``start: 0`` must read as a value, not a missing field.

    The guard is ``data.get("start") is not None`` — an ``if data.get("start")``
    regression would silently keep the old start on a snap-to-zero PATCH.
    """
    cut = co_client.post(
        "/composer/api/cuts", json={"participant": "P01", "start": 2.0, "end": 5.0}
    ).get_json()["cut"]
    patched = co_client.patch(
        f"/composer/api/cuts/{cut['id']}", json={"start": 0}
    ).get_json()
    assert patched["cut"]["start"] == 0.0
    assert patched["cut"]["end"] == 5.0


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


def test_ui_update_persists_media_toggles(co_client, tmp_path):
    resp = co_client.put("/composer/api/ui", json={"markerThumbnails": True}).get_json()
    assert resp["ok"] is True
    assert resp["markerThumbnails"] is True
    disk = _manifest_on_disk(tmp_path)["ui"]
    assert disk["markerThumbnails"] is True
    assert disk["markerAudioScrub"] is False  # empty-manifest default
    assert disk["followPlayhead"] is True  # empty-manifest default
    # A thumbs-only PUT must not clobber the lane toggles.
    assert disk["markerSources"] == {
        "sheet": True,
        "screenspace": True,
        "transcript": True,
    }

    resp = co_client.put("/composer/api/ui", json={"markerAudioScrub": True}).get_json()
    assert resp["markerAudioScrub"] is True
    assert _manifest_on_disk(tmp_path)["ui"]["markerAudioScrub"] is True

    resp = co_client.put("/composer/api/ui", json={"followPlayhead": False}).get_json()
    assert resp["followPlayhead"] is False
    assert _manifest_on_disk(tmp_path)["ui"]["followPlayhead"] is False


# ---- Scrubber media (sprite sheets / audio snippets) ----


@pytest.fixture
def scrub_calls(monkeypatch):
    """Arg-capturing extractor fakes + fresh media caches for the scrub routes."""
    calls = {"sprite": [], "audio": []}

    def fake_sprite(path, start, duration, cols, rows, **kwargs):
        calls["sprite"].append((path, start, duration, cols, rows))
        return b"jpegbytes"

    def fake_audio(path, start, duration):
        calls["audio"].append((path, start, duration))
        return b"wavbytes"

    monkeypatch.setattr(video, "extract_sprite_sheet_bytes", fake_sprite)
    monkeypatch.setattr(video, "extract_audio_segment_bytes", fake_audio)
    monkeypatch.setattr(composer_server, "_sprite_cache", composer_server.MediaCache(8))
    monkeypatch.setattr(composer_server, "_audio_cache", composer_server.MediaCache(8))
    return calls


def test_sprite_serves_with_cache_header(co_client, scrub_calls, tmp_path):
    resp = co_client.get("/composer/api/sprite/P01?start=1&end=3")
    assert resp.status_code == 200
    assert resp.mimetype == "image/jpeg"
    assert resp.headers["Cache-Control"] == "public, max-age=86400"
    assert resp.data == b"jpegbytes"
    (call,) = scrub_calls["sprite"]
    assert call == (
        str(tmp_path / "study_P01-1.mp4"),
        1.0,
        2.0,
        config.STUDIO_SCRUBBER_SPRITE_COLS,
        config.STUDIO_SCRUBBER_SPRITE_ROWS,
    )


def test_sprite_maps_global_time_into_part(co_client, scrub_calls, tmp_path):
    resp = co_client.get("/composer/api/sprite/P01?start=15&end=18")
    assert resp.status_code == 200
    (call,) = scrub_calls["sprite"]
    assert call[0] == str(tmp_path / "study_P01-2.mp4")
    assert (call[1], call[2]) == (5.0, 3.0)


def test_sprite_clamps_boundary_straddle(co_client, scrub_calls, tmp_path):
    # A span crossing the 10 s part boundary samples only the owning part's tail.
    resp = co_client.get("/composer/api/sprite/P01?start=8&end=15")
    assert resp.status_code == 200
    (call,) = scrub_calls["sprite"]
    assert call[0] == str(tmp_path / "study_P01-1.mp4")
    assert (call[1], call[2]) == (8.0, 2.0)


def test_audio_serves_wav_and_caps_duration(co_client, scrub_calls, monkeypatch):
    monkeypatch.setattr(config, "COMPOSER_SCRUB_MAX_AUDIO_SECONDS", 4.0)
    resp = co_client.get("/composer/api/audio/P01?start=0&end=500")
    assert resp.status_code == 200
    assert resp.mimetype == "audio/wav"
    assert resp.data == b"wavbytes"
    (call,) = scrub_calls["audio"]
    assert (call[1], call[2]) == (0.0, 4.0)


def test_scrub_routes_reject_bad_input(co_client, scrub_calls):
    assert co_client.get("/composer/api/sprite/P01").status_code == 400
    assert co_client.get("/composer/api/sprite/P01?start=5&end=5").status_code == 400
    assert co_client.get("/composer/api/sprite/P99?start=1&end=3").status_code == 404
    assert co_client.get("/composer/api/audio/P99?start=1&end=3").status_code == 404
    assert scrub_calls["sprite"] == [] and scrub_calls["audio"] == []


def test_sprite_extraction_failure_404(co_client, scrub_calls, monkeypatch):
    monkeypatch.setattr(video, "extract_sprite_sheet_bytes", lambda *a, **k: None)
    resp = co_client.get("/composer/api/sprite/P01?start=1&end=3")
    assert resp.status_code == 404


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


# ---- Annotations ----


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


def test_annotation_shape_crud_and_validation(co_client):
    created = _make_annotation(
        co_client,
        type="shape",
        geometry={
            "shape": "rect",
            "x": 1.4,  # clamps to 1.0
            "y": 0.5,
            "w": 0.3,
            "h": 0.2,
            "rotation": 405.0,  # normalizes mod 360
        },
    )
    assert created["ok"] is True
    geometry = created["annotation"]["geometry"]
    assert geometry["shape"] == "rect"
    assert geometry["x"] == 1.0
    assert geometry["rotation"] == 45.0

    # Unknown shape kind and degenerate sizes are rejected.
    for bad in (
        {"shape": "triangle", "x": 0.5, "y": 0.5, "w": 0.3, "h": 0.2},
        {"shape": "ellipse", "x": 0.5, "y": 0.5, "w": 0, "h": 0.2},
    ):
        assert (
            co_client.post(
                "/composer/api/annotations",
                json={
                    "participant": "P01",
                    "type": "shape",
                    "span": {"start": 1, "end": 5},
                    "geometry": bad,
                },
            ).status_code
            == 400
        )


def test_render_annotation_overlay_draws_shapes():
    overlay = composer_server._render_annotation_overlay(
        [
            {
                "type": "shape",
                "geometry": {
                    "shape": "rect",
                    "x": 0.5,
                    "y": 0.5,
                    "w": 0.5,
                    "h": 0.5,
                    "rotation": 0.0,
                },
                "style": {"color": "#ff0000", "strokeWidth": 0.01},
            },
            {
                "type": "shape",
                "geometry": {
                    "shape": "ellipse",
                    "x": 0.25,
                    "y": 0.25,
                    "w": 0.3,
                    "h": 0.2,
                    "rotation": 30.0,
                },
                "style": {"color": "#00ff00", "strokeWidth": 0.01},
            },
        ],
        640,
        360,
    )
    assert overlay.mode == "RGBA"
    # Rect edge midpoint (left edge at x=0.25*640=160, mid-height) is stroked.
    assert overlay.getpixel((160, 180))[3] > 0
    # Center of the rect stays unfilled (outline-only).
    assert overlay.getpixel((320, 180))[3] == 0
    # The rotated ellipse drew something in its quadrant.
    region = overlay.crop((60, 30, 260, 150))
    assert region.getchannel("A").getextrema()[1] > 0


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


def test_render_annotation_overlay_dashed_freehand_has_gaps():
    overlay = composer_server._render_annotation_overlay(
        [
            {
                "type": "freehand",
                "geometry": {"points": [[0.1, 0.5], [0.9, 0.5]]},
                "style": {
                    "color": "#ff0000",
                    "strokeWidth": 0.01,
                    "strokeStyle": "dashed",
                },
            }
        ],
        640,
        360,
    )
    row_alpha = [overlay.getpixel((x, 180))[3] for x in range(70, 570)]
    assert max(row_alpha) > 0  # dashes were drawn
    assert min(row_alpha) == 0  # ...with gaps between them (not a solid line)


def test_render_annotation_overlay_dashed_dotted_shapes_draw():
    # Rotated dashed rect + rotated dotted ellipse must render without error and
    # put down some pixels (the dash/perimeter-polygon paths, not the solid ones).
    overlay = composer_server._render_annotation_overlay(
        [
            {
                "type": "shape",
                "geometry": {
                    "shape": "rect",
                    "x": 0.5,
                    "y": 0.5,
                    "w": 0.6,
                    "h": 0.6,
                    "rotation": 15.0,
                },
                "style": {
                    "color": "#ff0000",
                    "strokeWidth": 0.008,
                    "strokeStyle": "dashed",
                },
            },
            {
                "type": "shape",
                "geometry": {
                    "shape": "ellipse",
                    "x": 0.5,
                    "y": 0.5,
                    "w": 0.5,
                    "h": 0.3,
                    "rotation": 40.0,
                },
                "style": {
                    "color": "#00ff00",
                    "strokeWidth": 0.008,
                    "strokeStyle": "dotted",
                },
            },
        ],
        640,
        360,
    )
    assert overlay.mode == "RGBA"
    assert overlay.getchannel("A").getextrema()[1] > 0


def test_sanitize_annotation_style_stroke_style():
    assert (
        composer_server._sanitize_annotation_style({"strokeStyle": "dashed"})[
            "strokeStyle"
        ]
        == "dashed"
    )
    # Unknown value falls back to the configured default.
    assert (
        composer_server._sanitize_annotation_style({"strokeStyle": "zigzag"})[
            "strokeStyle"
        ]
        == config.COMPOSER_ANNOTATION_STROKE_STYLE
    )
    # Missing → default.
    assert composer_server._sanitize_annotation_style({})["strokeStyle"] == "solid"


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


def test_build_overlay_command_honors_hardware_encoder():
    """The burn re-encode follows FFMPEG_VIDEO_ENCODER; pix_fmt stays pinned."""
    cmd = composer_server._build_overlay_command(
        "/vids/study_P01.mp4",
        0.0,
        4.0,
        [("/tmp/o1.png", 0.0, 4.0)],
        "/out/annotated.mp4",
        gif=False,
        encoder="h264_videotoolbox",
    )
    assert cmd[cmd.index("-c:v") + 1] == "h264_videotoolbox"
    assert "-q:v" in cmd
    assert "-crf" not in cmd
    assert cmd[cmd.index("-pix_fmt") + 1] == "yuv420p"


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


def test_annotated_burn_rejects_a_span_that_only_grazes_an_annotation(
    co_client, monkeypatch
):
    """Overlap under the 0.01 s window floor leaves nothing to burn.

    The annotation is "in the span" as far as _annotations_in_span is concerned,
    so the no-annotations check passes, but _annotation_windows drops the window
    — which used to hand ffmpeg an empty filter graph and a `[0:v]` label it
    never defined.
    """
    _make_annotation(co_client, span={"start": 10.0, "end": 20.0})

    def _no_ffmpeg(*_a, **_kw):
        raise AssertionError("ffmpeg must not run with nothing to overlay")

    monkeypatch.setattr(composer_server.video, "run_ffmpeg_process", _no_ffmpeg)

    resp = co_client.post(
        "/composer/api/export/burn",
        json={"participant": "P01", "start": 19.995, "end": 25.0},
    ).get_json()

    assert resp["ok"] is False
    assert "visible for long enough" in resp["error"]


def test_annotated_export_rejects_out_of_range_span(co_client, monkeypatch):
    """A multi-part span past the recording errs cleanly, never encodes.

    ``map_global_range_to_segments`` signals an out-of-range start with None
    (never ``[]``); conflating that with the single-part fast path sent the
    span down the parts-based branch and into a doomed ffmpeg run.
    """
    _make_annotation(co_client, span={"start": 24.0, "end": 30.0})

    def _no_ffmpeg(*_a, **_kw):
        raise AssertionError("ffmpeg must not run for an out-of-range span")

    monkeypatch.setattr(composer_server.video, "run_ffmpeg_process", _no_ffmpeg)

    resp = co_client.post(
        "/composer/api/export/burn",
        json={"participant": "P01", "start": 25.0, "end": 30.0},
    )
    assert resp.status_code == 400
    assert "outside the recording" in resp.get_json()["error"]


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


def test_export_cancel_sets_the_shared_event(co_client):
    composer_server._export_cancel.clear()
    resp = co_client.post("/composer/api/export/cancel").get_json()
    assert resp["ok"] is True
    assert composer_server._export_cancel.is_set()
    composer_server._export_cancel.clear()


def test_remux_routes_registered_on_composer(co_app):
    """The shared remux routes ride on the composer blueprint itself.

    tests/test_container_seekability.py exercises the route bodies against a
    throwaway blueprint; this pins the composer registration.
    """
    rules = {r.rule for r in co_app.url_map.iter_rules()}
    assert "/composer/api/remux/status" in rules
    assert "/composer/api/remux/<pid>" in rules


def test_repin_sheet_state_repoints_context(monkeypatch):
    """Worksheet swaps go through repin_sheet_state, not a re-init.

    The remux registration hands the blueprint ``lambda: _sheet_context``, so
    this repoint is what keeps remux (and participant resolution) on the newly
    opened sheet.
    """
    monkeypatch.setattr(composer_server, "_sheet_context", None)
    sentinel = object()
    composer_server.repin_sheet_state(sentinel)
    assert composer_server._sheet_context is sentinel
    composer_server.repin_sheet_state(None)
    assert composer_server._sheet_context is None


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
