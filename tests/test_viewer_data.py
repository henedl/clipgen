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


# ---- finalize_insights_viewer_data ----


def _make_insight(insight_id, status="draft", referenced_artifact_ids=None):
    return {
        "id": insight_id,
        "title": f"Insight {insight_id}",
        "status": status,
        "causes": {"narrative": "", "artifacts": referenced_artifact_ids or []},
        "behaviors": {"narrative": "", "artifacts": []},
        "impacts": {"narrative": "", "artifacts": []},
    }


def test_finalize_insights_viewer_filters_to_final():
    ins_list = [
        _make_insight("i1", status="final", referenced_artifact_ids=["a1"]),
        _make_insight("i2", status="draft", referenced_artifact_ids=["a2"]),
        _make_insight("i3", status="final", referenced_artifact_ids=["a3"]),
    ]
    all_artifacts = [_make_artifact("a1"), _make_artifact("a2"), _make_artifact("a3")]

    data = viewer.finalize_insights_viewer_data(ins_list, all_artifacts, study="s")

    assert len(data["insights"]) == 2
    ids = {i["id"] for i in data["insights"]}
    assert ids == {"i1", "i3"}

    art_ids = {a["id"] for a in data["artifacts"]}
    assert art_ids == {"a1", "a3"}


def test_finalize_insights_viewer_shows_all_when_no_finals():
    ins_list = [
        _make_insight("i1", status="draft", referenced_artifact_ids=["a1"]),
        _make_insight("i2", status="draft", referenced_artifact_ids=["a2"]),
    ]
    all_artifacts = [_make_artifact("a1"), _make_artifact("a2")]

    data = viewer.finalize_insights_viewer_data(ins_list, all_artifacts, study="s")

    assert len(data["insights"]) == 2
    art_ids = {a["id"] for a in data["artifacts"]}
    assert art_ids == {"a1", "a2"}


# ---- HTML generation (gallery + insights viewer) ----


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


def test_generate_insights_viewer_inlines_css_and_js(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    data = {"meta": {}, "insights": [], "artifacts": []}
    out_path = viewer.generate_insights_viewer(data, output_basename="insights.html")
    assert out_path is not None
    assert out_path.is_file()

    html = out_path.read_text(encoding="utf-8")
    assert '<link rel="stylesheet" href="insights-viewer.css">' not in html
    assert '<script src="insights-viewer.js" defer></script>' not in html
    assert "<style>" in html
    assert "<script defer>" in html
    assert "window.CLIPGEN_DATA" in html
