import json
import threading
import time

import config
import utils
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

    raw = json.loads((tmp_path / config.MANIFEST_FILENAME).read_text())["clips"]
    assert "meta" in raw
    assert "artifacts" in raw
    assert "timeline" in raw
    assert raw["meta"]["study"] == "mystudy"
    assert raw["meta"]["sourceFileType"] == "excel"
    assert raw["meta"]["sourceSpreadsheet"] == "Sheet1"
    assert raw["timeline"]["duration"] > 0


def test_save_and_load_reels_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    reel = {
        "id": "reel_abcd1234",
        "file": "study_reel.mp4",
        "study": "study",
        "description": "Reel: 2 segments",
        "components": [
            {
                "cellRow": 3,
                "cellCol": 2,
                "participant": "P01",
                "sourceVideo": "study_P01.mp4",
                "start": 10.0,
                "end": 20.0,
                "category": "cat",
                "description": "desc1",
                "severity": "",
            },
            {
                "cellRow": 4,
                "cellCol": 2,
                "participant": "P01",
                "sourceVideo": "study_P01.mp4",
                "start": 30.0,
                "end": 40.0,
                "category": "cat",
                "description": "desc2",
                "severity": "",
            },
        ],
    }

    path = viewer.save_manifest([], new_reels=[reel], study="study", mode="reel")
    assert path is not None

    _, loaded_reels = viewer.load_manifest_both()
    assert len(loaded_reels) == 1
    assert loaded_reels[0]["id"] == "reel_abcd1234"
    assert len(loaded_reels[0]["components"]) == 2

    # Verify reels key is in raw JSON
    raw = json.loads((tmp_path / config.MANIFEST_FILENAME).read_text())["clips"]
    assert "reels" in raw
    assert len(raw["reels"]) == 1


def test_save_manifest_merges_reels_and_artifacts(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))

    # First save: one artifact
    viewer.save_manifest([_make_artifact("a4c2s0")])

    # Second save: one reel, no new artifacts
    reel = {
        "id": "reel_abcd1234",
        "file": "reel.mp4",
        "study": "study",
        "description": "Reel: 1 segment",
        "components": [
            {
                "cellRow": 3,
                "cellCol": 2,
                "participant": "P01",
                "sourceVideo": "study_P01.mp4",
                "start": 10.0,
                "end": 20.0,
                "category": "cat",
                "description": "desc",
                "severity": "",
            }
        ],
    }
    viewer.save_manifest([], new_reels=[reel])

    assert len(viewer.load_manifest_artifacts()) == 1
    assert len(viewer.load_manifest_both()[1]) == 1


def test_load_manifest_both_returns_empty_when_no_file(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    assert viewer.load_manifest_both() == ([], [])


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


def _make_reel(reel_id):
    return {
        "id": reel_id,
        "file": f"{reel_id}.mp4",
        "study": "study",
        "description": "Reel: 1 segment",
        "components": [
            {
                "cellRow": 3,
                "cellCol": 2,
                "participant": "P01",
                "sourceVideo": "study_P01.mp4",
                "start": 10.0,
                "end": 20.0,
                "category": "cat",
                "description": "desc",
                "severity": "",
            }
        ],
    }


def test_save_manifest_concurrent_writes_keep_every_id(tmp_path, monkeypatch):
    """Concurrent save_manifest() calls must not last-writer-wins a partial merge.

    Each thread saves a distinct artifact and reel. The real finalize step is
    wrapped with a small sleep to widen the load->write window: without the
    write lock every thread would load the same empty manifest during that
    window and only one record of each kind would survive.
    """
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))

    real_finalize = viewer.finalize_timeline_data

    def slow_finalize(*args, **kwargs):
        time.sleep(0.05)
        return real_finalize(*args, **kwargs)

    monkeypatch.setattr(viewer, "finalize_timeline_data", slow_finalize)

    n = 8
    barrier = threading.Barrier(n)
    errors = []

    def worker(i):
        try:
            barrier.wait(timeout=5)
            viewer.save_manifest(
                [_make_artifact(f"a{i}")], new_reels=[_make_reel(f"reel{i}")]
            )
        except Exception as exc:
            errors.append(exc)

    threads = [threading.Thread(target=worker, args=(i,)) for i in range(n)]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout=10)

    assert not errors, f"worker threads raised: {errors}"
    assert not any(t.is_alive() for t in threads)

    artifact_ids = {a["id"] for a in viewer.load_manifest_artifacts()}
    reel_ids = {r["id"] for r in viewer.load_manifest_both()[1]}
    assert artifact_ids == {f"a{i}" for i in range(n)}
    assert reel_ids == {f"reel{i}" for i in range(n)}


# ---- Section store (utils.load/save_manifest_section) ----


def test_store_sections_round_trip_and_sorted(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    utils.save_manifest_section("zeta", {"n": 1, "s": "multi\nline"})
    utils.save_manifest_section("alpha", [1, 2])
    raw = json.loads((tmp_path / config.MANIFEST_FILENAME).read_text())
    assert list(raw) == ["alpha", "zeta"]
    assert raw["zeta"] == {"n": 1, "s": "multi\nline"}
    assert utils.load_manifest_section("alpha") == [1, 2]
    assert utils.load_manifest_section("missing", default="d") == "d"
    assert utils.manifest_sections() == {"alpha", "zeta"}


def test_store_load_returns_fresh_objects(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    utils.save_manifest_section("one", {"items": []})
    utils.load_manifest_section("one")["items"].append("leak")
    assert utils.load_manifest_section("one") == {"items": []}


def test_store_picks_up_external_rewrite(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    utils.save_manifest_section("one", 1)
    path = tmp_path / config.MANIFEST_FILENAME
    path.write_text(json.dumps({"one": 2, "two": 3}))
    assert utils.load_manifest_section("one") == 2
    assert utils.manifest_sections() == {"one", "two"}


def test_store_identical_save_skips_the_write(tmp_path, monkeypatch):
    """An idempotent save must not rewrite the file or bump its mtime.

    Startup rewrites and debounced persists with no delta fire often; a
    phantom mtime bump would also force every mtime-gated consumer
    (workflow triggers, viewer events cache) to re-parse for nothing.
    """
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    path = tmp_path / config.MANIFEST_FILENAME
    utils.save_manifest_section("one", {"a": 1})
    utils.save_manifest_section("two", [1, 2])
    stamp = path.stat().st_mtime_ns
    assert utils.save_manifest_section("one", {"a": 1}) == path
    assert path.stat().st_mtime_ns == stamp
    # Removing a section that is not stored is equally a no-op.
    assert utils.save_manifest_section("ghost", None) == path
    assert path.stat().st_mtime_ns == stamp
    # A real change still writes.
    utils.save_manifest_section("one", {"a": 2})
    assert utils.load_manifest_section("one") == {"a": 2}
    assert json.loads(path.read_text())["one"] == {"a": 2}


def test_store_reindent_cache_yields_identical_file(tmp_path, monkeypatch):
    """Cached re-indented section texts must produce the same bytes as a cold write."""
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    path = tmp_path / config.MANIFEST_FILENAME
    big = {"rows": [{"i": i, "text": "line\nbreak"} for i in range(50)]}
    utils.save_manifest_section("big", big)
    utils.save_manifest_section("small", 1)  # re-indents "big" via the cache
    warm = path.read_text()
    utils._reset_manifest_cache()  # cold path: no cached indent texts
    utils.save_manifest_section("small", 2)
    utils.save_manifest_section("small", 1)
    assert path.read_text() == warm


def test_store_corrupt_file_reads_as_empty_and_blocks_saves(tmp_path, monkeypatch):
    """A bad read must never let the next save wipe the other sections."""
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
    path = tmp_path / config.MANIFEST_FILENAME
    path.write_text("not json")
    assert utils.load_manifest_section("one", default=0) == 0
    assert utils.manifest_sections() == set()
    assert utils.save_manifest_section("one", {"a": 1}) is None
    assert utils.save_manifest_section("one", None) is None
    assert path.read_text() == "not json"
    # Fixing the file on disk re-enables saves.
    path.write_text(json.dumps({"two": 2}))
    assert utils.save_manifest_section("one", 1) is not None
    assert utils.manifest_sections() == {"one", "two"}
