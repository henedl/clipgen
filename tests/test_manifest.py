import json

import config
import viewer


def _make_artifact(artifact_id, description="desc", study="study", participant="P01"):
    return {
        "id": artifact_id,
        "type": "clip",
        "file": f"{artifact_id}.mp4",
        "start": 10.0,
        "end": 20.0,
        "thumbnail": "",
        "study": study,
        "participant": participant,
        "category": "cat",
        "description": description,
        "cellRow": 4,
        "cellCol": 2,
        "cellA1": "B4",
        "annotations": [],
        "sourceVideo": "study_P01.mp4",
    }


def test_load_manifest_returns_empty_when_no_file(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    assert viewer.load_manifest_artifacts() == []


def test_load_manifest_returns_empty_on_malformed_json(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    (tmp_path / config.MANIFEST_FILENAME).write_text("not json at all")
    assert viewer.load_manifest_artifacts() == []


def test_save_and_load_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    artifacts = [_make_artifact("a4c2s0"), _make_artifact("a5c2s0")]

    path = viewer.save_manifest(artifacts, study="study", mode="batch")
    assert path is not None
    assert path.is_file()

    loaded = viewer.load_manifest_artifacts()
    assert len(loaded) == 2
    ids = {a["id"] for a in loaded}
    assert ids == {"a4c2s0", "a5c2s0"}


def test_save_manifest_merges_cumulatively(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))

    # First save: two artifacts
    viewer.save_manifest([_make_artifact("a4c2s0"), _make_artifact("a5c2s0")])

    # Second save: one new, one overlapping
    viewer.save_manifest(
        [_make_artifact("a5c2s0", description="updated"), _make_artifact("a6c2s0")]
    )

    loaded = viewer.load_manifest_artifacts()
    assert len(loaded) == 3
    by_id = {a["id"]: a for a in loaded}
    assert by_id["a5c2s0"]["description"] == "updated"
    assert "a4c2s0" in by_id
    assert "a6c2s0" in by_id


def test_save_manifest_deduplicates_by_id(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    # Same id twice in one call — last one wins
    a1 = _make_artifact("a4c2s0", description="first")
    a2 = _make_artifact("a4c2s0", description="second")

    viewer.save_manifest([a1, a2])
    loaded = viewer.load_manifest_artifacts()
    assert len(loaded) == 1
    assert loaded[0]["description"] == "second"


def test_manifest_contains_valid_timeline_data_structure(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    viewer.save_manifest(
        [_make_artifact("a4c2s0")],
        study="mystudy",
        participant="P01",
        worksheet_title="Sheet1",
        is_excel=True,
        mode="batch",
    )

    raw = json.loads((tmp_path / config.MANIFEST_FILENAME).read_text())
    assert "meta" in raw
    assert "artifacts" in raw
    assert "timeline" in raw
    assert raw["meta"]["study"] == "mystudy"
    assert raw["meta"]["sourceFileType"] == "excel"
    assert raw["meta"]["sourceSpreadsheet"] == "Sheet1"
    assert raw["timeline"]["duration"] > 0


def test_cli_manifest_flag_parsed(monkeypatch):
    import cli

    monkeypatch.setattr("sys.argv", ["clipgen.py", "-b", "--manifest"])
    args = cli.parse_arguments()
    assert args.manifest is True


def test_cli_manifest_flag_defaults_false(monkeypatch):
    import cli

    monkeypatch.setattr("sys.argv", ["clipgen.py"])
    args = cli.parse_arguments()
    assert args.manifest is False
