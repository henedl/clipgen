import viewer


def _make_artifact(
    artifact_id, *, start=10.0, end=20.0, study="study", participant="P01"
):
    return {
        "id": artifact_id,
        "type": "clip",
        "file": f"{artifact_id}.mp4",
        "start": start,
        "end": end,
        "thumbnail": "",
        "study": study,
        "participant": participant,
        "category": "cat",
        "severity": "",
        "description": "desc",
        "cellRow": 4,
        "cellCol": 2,
        "cellA1": "B4",
        "annotations": [],
        "sourceVideo": f"{study}_{participant}.mp4",
    }


# ---- finalize_timeline_data ----


def test_finalize_timeline_data_duration_and_structure():
    arts = [_make_artifact("a1", end=20.0), _make_artifact("a2", end=50.0)]
    data = viewer.finalize_timeline_data(
        arts,
        study="s",
        participant="P01",
        worksheet_title="Sheet1",
        is_excel=True,
        mode="batch",
    )

    assert set(data.keys()) == {"meta", "config", "artifacts", "timeline"}
    assert data["timeline"]["duration"] == 50.0 * 1.05
    assert data["timeline"]["startOffset"] == 0.0
    assert data["meta"]["study"] == "s"
    assert data["meta"]["participant"] == "P01"
    assert data["meta"]["sourceFileType"] == "excel"
    assert data["meta"]["sourceSpreadsheet"] == "Sheet1"
    assert data["meta"]["mode"] == "batch"
    assert "reels" not in data


def test_finalize_timeline_data_empty_artifacts():
    data = viewer.finalize_timeline_data([])
    assert data["timeline"]["duration"] == 0.0
    assert data["artifacts"] == []


def test_finalize_timeline_data_includes_filmstrip_meta(monkeypatch):
    import config

    monkeypatch.setattr(config, "FILMSTRIP_ENABLED", False)
    data = viewer.finalize_timeline_data([])
    assert data["meta"]["filmstripEnabled"] is False

    monkeypatch.setattr(config, "FILMSTRIP_ENABLED", True)
    data = viewer.finalize_timeline_data([])
    assert data["meta"]["filmstripEnabled"] is True


def test_finalize_timeline_data_includes_reels():
    reels = [{"id": "reel_1", "file": "r.mp4"}]
    data = viewer.finalize_timeline_data([], reels=reels)
    assert data["reels"] == reels


# ---- finalize_timeline_data ----


def test_finalize_timeline_data_keeps_artifact_with_only_start():
    art = _make_artifact("a1c1s0", end=None)
    data = viewer.finalize_timeline_data([art])
    assert len(data["artifacts"]) == 1


def test_finalize_timeline_data_keeps_valid_artifacts_unmodified():
    arts = [_make_artifact("a1c1s0"), _make_artifact("a2c1s0")]
    data = viewer.finalize_timeline_data(arts)
    assert len(data["artifacts"]) == 2
    ids = {a["id"] for a in data["artifacts"]}
    assert ids == {"a1c1s0", "a2c1s0"}


# ---- finalize_gallery_data ----


def test_finalize_gallery_data_structure():
    gallery_artifacts = [{"file": "frame_10.png", "timestamp": 10.0}]
    data = viewer.finalize_gallery_data(
        gallery_artifacts,
        source_video="vid.mp4",
        video_duration=120,
        output_format="screen",
        interval=5,
    )

    assert data["meta"]["mode"] == "gallery"
    assert data["meta"]["sourceVideo"] == "vid.mp4"
    assert data["meta"]["videoDuration"] == 120
    assert data["meta"]["interval"] == 5
    assert data["meta"]["format"] == "screen"
    assert data["artifacts"] == gallery_artifacts


# ---- finalize_gallery_data bundling ----


def test_finalize_gallery_data_bundle_embeds_png(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    # Create a tiny PNG file (1x1 pixel)
    png_bytes = (
        b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01"
        b"\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00"
        b"\x00\x00\x0cIDATx\x9cc\xf8\x0f\x00\x00\x01\x01\x00"
        b"\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82"
    )
    (tmp_path / "cap_0_00.png").write_bytes(png_bytes)

    artifacts = [{"file": "cap_0_00.png", "timestamp": 0.0, "type": "screen"}]
    data = viewer.finalize_gallery_data(artifacts, bundle=True)

    assert data["meta"]["bundled"] is True
    data_uri = str(artifacts[0]["data"])
    assert data_uri.startswith("data:image/png;base64,")


def test_finalize_gallery_data_bundle_embeds_gif(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    gif_bytes = b"GIF89a\x01\x00\x01\x00\x00\xff\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x00;"
    (tmp_path / "cap_0_00.gif").write_bytes(gif_bytes)

    artifacts = [{"file": "cap_0_00.gif", "timestamp": 0.0, "type": "gif"}]
    viewer.finalize_gallery_data(artifacts, bundle=True)

    data_uri = str(artifacts[0]["data"])
    assert data_uri.startswith("data:image/gif;base64,")


def test_finalize_gallery_data_bundle_picks_mime_by_extension(tmp_path, monkeypatch):
    """Bundle MIME is chosen by file extension, not artifact type. WebM gifs
    must be embedded as video/webm so the viewer's <video> element loads them."""
    monkeypatch.chdir(tmp_path)
    cases = [
        ("cap.webp", b"RIFF\x00\x00\x00\x00WEBPVP8 ", "data:image/webp;base64,"),
        ("cap.webm", b"\x1a\x45\xdf\xa3", "data:video/webm;base64,"),
        ("cap.jpg", b"\xff\xd8\xff\xe0", "data:image/jpeg;base64,"),
    ]
    for filename, payload, expected_prefix in cases:
        (tmp_path / filename).write_bytes(payload)
        artifacts = [{"file": filename, "timestamp": 0.0, "type": "gif"}]
        viewer.finalize_gallery_data(artifacts, bundle=True)
        assert str(artifacts[0]["data"]).startswith(expected_prefix), filename


def test_finalize_gallery_data_bundle_skips_missing_file(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    artifacts = [{"file": "missing.png", "timestamp": 0.0, "type": "screen"}]
    data = viewer.finalize_gallery_data(artifacts, bundle=True)

    assert "data" not in artifacts[0]
    assert data["meta"]["bundled"] is True


def test_finalize_gallery_data_no_bundle_by_default():
    artifacts = [{"file": "cap.png", "timestamp": 0.0, "type": "screen"}]
    data = viewer.finalize_gallery_data(artifacts)

    assert data["meta"]["bundled"] is False
    assert "data" not in artifacts[0]


def test_generate_gallery_viewer_inlines_css_and_js(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    data = {"meta": {}, "artifacts": []}
    out_path = viewer.generate_gallery_viewer(data, output_basename="gallery.html")
    assert out_path is not None
    assert out_path.is_file()

    html = out_path.read_text(encoding="utf-8")
    assert '<link rel="stylesheet" href="gallery.css">' not in html
    assert '<script src="gallery.js" defer></script>' not in html
    assert "<style>" in html
    assert "<script defer>" in html
    assert "window.CLIPGEN_DATA" in html


# ---- bundled asset cache ----


def test_read_bundled_asset_caches_first_read(tmp_path):
    asset = tmp_path / "asset.txt"
    asset.write_text("original", encoding="utf-8")
    first = viewer._read_bundled_asset(str(asset))
    # Bundled assets never change at runtime; a later mutation must not be seen.
    asset.write_text("changed", encoding="utf-8")
    second = viewer._read_bundled_asset(str(asset))
    assert first == "original"
    assert second == "original"


# ---- screenspace events for viewer ----


def test_load_screenspace_events_for_viewer_caches_by_mtime(tmp_path, monkeypatch):
    import os

    import config
    import screenspace

    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    viewer._reset_screenspace_events_cache()

    screenspace.save_screenspace_manifest(
        {},
        [],
        [
            {
                "id": "ev_keep",
                "detector": "change",
                "participant": "P01",
                "time_in": 1.0,
                "time_out": 2.0,
                "excluded": False,
            },
            {
                "id": "ev_drop",
                "detector": "color",
                "participant": "P02",
                "time_in": 3.0,
                "time_out": 4.0,
                "excluded": True,
            },
        ],
    )

    load_calls = []
    real_load = screenspace.load_screenspace_manifest

    def counting_load():
        load_calls.append(1)
        return real_load()

    monkeypatch.setattr(screenspace, "load_screenspace_manifest", counting_load)

    first = viewer.load_screenspace_events_for_viewer()
    second = viewer.load_screenspace_events_for_viewer()
    # Excluded events are filtered out; the second call is served from cache.
    assert [e["id"] for e in first] == ["ev_keep"]
    assert [e["id"] for e in second] == ["ev_keep"]
    assert len(load_calls) == 1

    # Rewriting the manifest bumps its mtime and invalidates the cache.
    # Pass pins explicitly so save does not read the manifest (this test counts loads).
    screenspace.save_screenspace_manifest(
        {},
        [],
        [
            {
                "id": "ev_new",
                "detector": "scene",
                "participant": "P03",
                "time_in": 5.0,
                "time_out": 6.0,
                "excluded": False,
            }
        ],
        pins={},
    )
    manifest_path = tmp_path / config.SCREENSPACE_MANIFEST_FILENAME
    st = manifest_path.stat()
    os.utime(manifest_path, ns=(st.st_atime_ns, st.st_mtime_ns + 1_000_000_000))

    third = viewer.load_screenspace_events_for_viewer()
    assert [e["id"] for e in third] == ["ev_new"]
    assert len(load_calls) == 2


def test_load_screenspace_events_for_viewer_sanitizes_nonfinite_times(monkeypatch):
    import math

    import screenspace

    viewer._reset_screenspace_events_cache()
    # A non-finite time (as cv2/numpy can produce) embedded via json.dumps with
    # allow_nan=True emits bare NaN/Infinity, which JSON.parse rejects — blanking
    # the whole exported viewer. The transform must neutralize it to None.
    monkeypatch.setattr(
        screenspace,
        "load_screenspace_manifest",
        lambda: {
            "events": [
                {
                    "id": "ev_bad",
                    "detector": "change",
                    "participant": "P01",
                    "time_in": float("inf"),
                    "time_out": float("nan"),
                    "excluded": False,
                }
            ]
        },
    )
    # Force a cache miss regardless of any on-disk manifest mtime state.
    viewer._reset_screenspace_events_cache()
    events = {e["id"]: e for e in viewer.load_screenspace_events_for_viewer()}
    assert events["ev_bad"]["timeIn"] is None
    assert events["ev_bad"]["timeOut"] is None
    # And the payload is now strict-JSON serializable (no NaN/Infinity tokens).
    import json

    json.dumps(events, allow_nan=False)
    assert not any(
        isinstance(v, float) and not math.isfinite(v)
        for v in (events["ev_bad"]["timeIn"], events["ev_bad"]["timeOut"])
    )


def test_load_screenspace_events_for_viewer_includes_navigational(
    tmp_path, monkeypatch
):
    import config
    import screenspace

    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    viewer._reset_screenspace_events_cache()

    screenspace.save_screenspace_manifest(
        {},
        [],
        [
            {
                "id": "ev_boundary",
                "detector": "boundary",
                "participant": "P01",
                "time_in": 12.0,
                "time_out": 12.0,
                "excluded": False,
                "navigational": True,
                "metadata": {"distance": 22},
            },
            {
                "id": "ev_change",
                "detector": "change",
                "participant": "P01",
                "time_in": 5.0,
                "time_out": 5.0,
                "excluded": False,
            },
        ],
    )

    events = {e["id"]: e for e in viewer.load_screenspace_events_for_viewer()}
    # Boundary events carry navigational so the viewer can render thin ticks.
    assert events["ev_boundary"]["navigational"] is True
    # Non-navigational detectors default to False (absent in the source dict).
    assert events["ev_change"]["navigational"] is False
