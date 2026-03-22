import viewer


def _make_artifact(artifact_id, *, start=10.0, end=20.0, study="study", participant="P01"):
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
        arts, study="s", participant="P01", worksheet_title="Sheet1", is_excel=True, mode="batch"
    )

    assert set(data.keys()) == {"meta", "artifacts", "timeline"}
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


def test_finalize_timeline_data_includes_reels():
    reels = [{"id": "reel_1", "file": "r.mp4"}]
    data = viewer.finalize_timeline_data([], reels=reels)
    assert data["reels"] == reels


# ---- finalize_gallery_data ----


def test_finalize_gallery_data_structure():
    gallery_artifacts = [{"file": "frame_10.png", "timestamp": 10.0}]
    data = viewer.finalize_gallery_data(
        gallery_artifacts, source_video="vid.mp4", video_duration=120, output_format="screen", interval=5
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
