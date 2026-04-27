import time

import config
import insights


def _redirect_output(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))


def test_create_insight_returns_well_formed_dict():
    ins = insights.create_insight(title="Login bug", severity="High", status="draft")

    assert ins["id"].startswith("ins_")
    assert len(ins["id"]) == 12
    assert ins["title"] == "Login bug"
    assert ins["severity"] == "High"
    assert ins["status"] == "draft"
    assert ins["summary"] == ""
    assert ins["timelineContext"] == ""
    assert ins["createdAt"] == ins["updatedAt"]

    for bucket in ("causes", "behaviors", "impacts"):
        assert isinstance(ins[bucket]["narrative"], str)
        assert isinstance(ins[bucket]["artifacts"], list)
        assert ins[bucket]["artifacts"] == []


def test_create_insight_defaults():
    ins = insights.create_insight()
    assert ins["title"] == "Untitled insight"
    assert ins["severity"] == ""
    assert ins["status"] == "draft"


def test_update_insight_merges_fields():
    ins = insights.create_insight(title="Old")
    original_created = ins["createdAt"]
    original_id = ins["id"]
    time.sleep(0.01)

    result = insights.update_insight(
        [ins], original_id, {"title": "New", "severity": "Critical"}
    )

    assert result is ins
    assert ins["title"] == "New"
    assert ins["severity"] == "Critical"
    assert ins["id"] == original_id
    assert ins["createdAt"] == original_created
    assert ins["updatedAt"] >= original_created


def test_update_insight_protects_id_and_createdAt():
    ins = insights.create_insight(title="Safe")
    original_id = ins["id"]
    original_created = ins["createdAt"]

    insights.update_insight([ins], original_id, {"id": "hacked", "createdAt": "hacked"})

    assert ins["id"] == original_id
    assert ins["createdAt"] == original_created


def test_update_insight_returns_none_for_missing_id():
    ins = insights.create_insight()
    result = insights.update_insight([ins], "ins_nonexistent", {"title": "X"})
    assert result is None


def test_delete_insight_removes_and_returns_true():
    a = insights.create_insight(title="A")
    b = insights.create_insight(title="B")
    lst = [a, b]

    assert insights.delete_insight(lst, a["id"]) is True
    assert len(lst) == 1
    assert lst[0]["id"] == b["id"]


def test_delete_insight_returns_false_for_missing_id():
    ins = insights.create_insight()
    lst = [ins]
    assert insights.delete_insight(lst, "ins_nonexistent") is False
    assert len(lst) == 1


def test_save_and_load_roundtrip(tmp_path, monkeypatch):
    _redirect_output(tmp_path, monkeypatch)

    ins_list = [
        insights.create_insight(title="First", severity="High"),
        insights.create_insight(title="Second", severity="Low"),
    ]
    path = insights.save_insights_manifest({"study": "demo"}, ins_list)
    assert path is not None
    assert path.is_file()

    loaded = insights.load_insights_manifest()
    assert len(loaded["insights"]) == 2
    assert loaded["meta"]["study"] == "demo"
    assert "generatedAt" in loaded["meta"]
    assert "version" in loaded["meta"]
    titles = {i["title"] for i in loaded["insights"]}
    assert titles == {"First", "Second"}


def test_load_handles_missing_and_malformed(tmp_path, monkeypatch):
    _redirect_output(tmp_path, monkeypatch)
    manifest_path = tmp_path / config.INSIGHTS_MANIFEST_FILENAME

    # Missing file
    result = insights.load_insights_manifest()
    assert result == {"meta": {}, "insights": []}

    # Malformed JSON
    manifest_path.write_text("not json at all")
    result = insights.load_insights_manifest()
    assert result == {"meta": {}, "insights": []}
