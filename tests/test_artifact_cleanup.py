"""Tests for output-dir hygiene: empty-manifest removal and stale-temp sweeping.

Covers ``utils.remove_json_manifest`` / ``utils.sweep_stale_temp_artifacts`` and
the Workflows launch-time cleanup that reclaims a stale empty manifest left by a
prior abandoned session.
"""

import config
import utils
import workflows
import workflows_server


class TestRemoveJsonManifest:
    def test_deletes_file_and_tmp_sibling(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        (tmp_path / "thing.json").write_text("{}")
        (tmp_path / "thing.json.tmp").write_text("partial")
        utils.remove_json_manifest("thing.json")
        assert not (tmp_path / "thing.json").exists()
        assert not (tmp_path / "thing.json.tmp").exists()

    def test_missing_file_is_noop(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        # Must not raise when nothing is on disk.
        utils.remove_json_manifest("absent.json")


class TestSweepStaleTempArtifacts:
    def test_sweeps_only_our_orphans(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        # Our orphans: an atomic-write tmp sibling and a reel temp-clip.
        (tmp_path / "screenspace_manifest.json.tmp").write_text("partial")
        (tmp_path / (config.TEMP_ARTIFACT_PREFIX + "ab12.mp4")).write_bytes(b"")
        # User files that must survive.
        (tmp_path / "keep.json").write_text("{}")
        (tmp_path / "keep.mp4").write_bytes(b"data")

        utils.sweep_stale_temp_artifacts()

        assert not (tmp_path / "screenspace_manifest.json.tmp").exists()
        assert not (tmp_path / (config.TEMP_ARTIFACT_PREFIX + "ab12.mp4")).exists()
        assert (tmp_path / "keep.json").exists()
        assert (tmp_path / "keep.mp4").exists()

    def test_missing_output_dir_is_noop(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path / "nope"))
        # Must not raise when the output dir does not exist.
        utils.sweep_stale_temp_artifacts()


class TestWorkflowsLaunchCleanup:
    def test_init_removes_stale_empty_manifest(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        # Keep the launch hermetic — no watch daemon / disk discovery.
        monkeypatch.setattr(workflows_server, "_start_watch_thread", lambda: None)
        monkeypatch.setattr(workflows_server, "_seed_watch_seen", lambda: None)

        manifest = tmp_path / config.WORKFLOWS_MANIFEST_FILENAME
        manifest.write_text('{"blueprints": [], "stashes": [], "runs": []}')
        (tmp_path / (config.WORKFLOWS_MANIFEST_FILENAME + ".tmp")).write_text("partial")

        workflows_server._init_workflows_state()

        assert not manifest.exists()
        assert not (tmp_path / (config.WORKFLOWS_MANIFEST_FILENAME + ".tmp")).exists()

    def test_init_keeps_nonempty_manifest(self, tmp_path, monkeypatch):
        monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))
        monkeypatch.setattr(workflows_server, "_start_watch_thread", lambda: None)
        monkeypatch.setattr(workflows_server, "_seed_watch_seen", lambda: None)

        workflows.save_workflows_manifest(
            [{"id": "bp1", "name": "Real", "nodes": [{"id": "n1"}], "edges": []}]
        )
        manifest = tmp_path / config.WORKFLOWS_MANIFEST_FILENAME
        assert manifest.is_file()

        workflows_server._init_workflows_state()

        assert manifest.is_file()
