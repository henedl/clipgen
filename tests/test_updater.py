"""Unit tests for the self-updater's pure and offline-testable halves.

Version parsing, asset selection, install-shape detection, helper-script
rendering, the download stream (fake urlopen), the GitHub check (fake urlopen
with 200 / 304 / offline), zip extraction, and the Inno relaunch gate. The
apply paths that mount a DMG or spawn a helper are hand-verified on a real Mac
and Windows box; see the module docstring in source/updater.py.
"""

from __future__ import annotations

import email.message
import hashlib
import io
import json
import urllib.error
import zipfile
from pathlib import Path

import pytest

import config
import start_settings
import updater
import utils


@pytest.fixture(autouse=True)
def _isolated_state(monkeypatch, tmp_path):
    updater.reset_for_tests()
    monkeypatch.setattr(start_settings, "config_dir", lambda: tmp_path / "cfg")
    monkeypatch.setattr(config, "UPDATE_CHECK_ON_LAUNCH", True)
    yield
    updater.reset_for_tests()


# ---- versions and assets -----------------------------------------------------


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("v1.2.3", (1, 2, 3)),
        ("1.2.3", (1, 2, 3)),
        ("v0.17.10", (0, 17, 10)),
        ("0.0.0+unknown", None),
        ("v1.2", None),
        ("v1.2.3-rc1", None),
        ("", None),
    ],
)
def test_parse_version(text, expected):
    assert updater.parse_version(text) == expected


def test_is_newer_compares_numerically():
    assert updater.is_newer("v1.2.10", "1.2.9")
    assert not updater.is_newer("v1.2.9", "1.2.9")
    assert not updater.is_newer("v0.0.1", "0.0.0+unknown")
    assert not updater.is_newer("dev", "1.0.0")


def test_pick_asset_matches_the_exact_name():
    release = {
        "tag": "v0.18.0",
        "assets": [
            {"name": "clipgen-v0.18.0-macos.dmg.sha256"},
            {"name": "clipgen-v0.18.0-macos.dmg"},
            {"name": "clipgen-v0.18.0-setup.exe"},
            {"name": "clipgen-v0.18.0-windows.zip"},
        ],
    }

    def name(shape):
        asset = updater.pick_asset(release, shape)
        return asset["name"] if asset else None

    assert name("mac-app") == "clipgen-v0.18.0-macos.dmg"
    assert name("win-inno") == "clipgen-v0.18.0-setup.exe"
    assert name("win-zip") == "clipgen-v0.18.0-windows.zip"
    assert updater.pick_asset(release, "unsupported") is None
    assert updater.pick_asset({"tag": "v0.18.0", "assets": []}, "mac-app") is None


# ---- install shape -----------------------------------------------------------


def test_source_checkout_is_unsupported(monkeypatch):
    monkeypatch.delattr("sys.frozen", raising=False)
    assert updater.install_shape() == "unsupported"
    assert updater.install_root() is None
    assert updater.is_supported() is False


def test_mac_bundle_shape(monkeypatch):
    monkeypatch.setattr("sys.frozen", True, raising=False)
    monkeypatch.setattr(
        "sys.executable", "/Applications/clipgen.app/Contents/MacOS/clipgen"
    )
    assert updater.install_shape() == "mac-app"
    assert updater.install_root() == Path("/Applications/clipgen.app")


def test_windows_shapes_split_on_the_uninstaller(monkeypatch, tmp_path):
    app = tmp_path / "clipgen"
    (app / "lib").mkdir(parents=True)
    exe = app / "clipgen.exe"
    exe.write_bytes(b"")
    monkeypatch.setattr("sys.platform", "win32")
    monkeypatch.setattr("sys.frozen", True, raising=False)
    monkeypatch.setattr("sys.executable", str(exe))
    monkeypatch.setattr("sys._MEIPASS", str(app / "lib"), raising=False)
    assert updater.install_shape() == "win-zip"
    (app / "unins000.exe").write_bytes(b"")
    assert updater.install_shape() == "win-inno"
    assert updater.install_root() == app.resolve()


def test_supported_needs_the_desktop_window(monkeypatch):
    monkeypatch.setattr(updater, "install_shape", lambda: "mac-app")
    monkeypatch.setattr(utils, "GUI_LAUNCH", False)
    assert updater.is_supported() is False
    monkeypatch.setattr(utils, "GUI_LAUNCH", True)
    assert updater.is_supported() is True


# ---- helper scripts ----------------------------------------------------------


def test_mac_helper_waits_swaps_and_relaunches():
    script = updater.render_mac_helper(
        4242,
        Path("/Applications/clip gen.app"),
        Path("/Applications/.clipgen-update/clip gen.app"),
        Path("/tmp/apply.log"),
    )
    assert script.startswith("#!/bin/sh")
    assert 'while kill -0 "$PID"' in script
    assert "'/Applications/clip gen.app'" in script
    assert script.index('mv "$LIVE" "$LIVE.old"') < script.index('mv "$STAGED" "$LIVE"')
    assert 'mv "$LIVE.old" "$LIVE"' in script  # rollback
    assert "xattr -dr com.apple.quarantine" in script
    assert script.rstrip().endswith('open -n "$LIVE"')


def test_win_helper_waits_renames_and_relaunches():
    script = updater.render_win_helper(
        4242,
        Path(r"D:\tools\my clipgen"),
        Path(r"D:\tools\my clipgen.new"),
        Path(r"C:\cfg\apply.log"),
    )
    assert "Wait-Process -Id $target" in script
    assert "'D:\\tools\\my clipgen'" in script
    assert "-NewName 'my clipgen.old'" in script
    assert script.index("'my clipgen.old'") < script.index("-NewName 'my clipgen'")
    assert "Start-Process -FilePath (Join-Path $root 'clipgen.exe')" in script
    assert "Set-Content -LiteralPath $log" in script


def test_inno_relaunch_is_gated_on_the_switch():
    """CI installs with /VERYSILENT and no /RELAUNCH; that must stay GUI-free."""
    iss = (Path(__file__).resolve().parent.parent / "build" / "clipgen.iss").read_text(
        encoding="utf-8"
    )
    assert 'Filename: "{app}\\clipgen.exe"; Flags: nowait; Check: WantRelaunch' in iss
    assert "ExpandConstant('{param:RELAUNCH|0}') = '1'" in iss


# ---- zip extraction ----------------------------------------------------------


def _portable_zip(path: Path, members: dict[str, bytes]) -> Path:
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in members.items():
            zf.writestr(name, data)
    return path


def test_extract_bundle_strips_the_top_folder(tmp_path):
    archive = _portable_zip(
        tmp_path / "u.zip",
        {
            "clipgen/clipgen.exe": b"exe",
            "clipgen/lib/a.dll": b"dll",
            "clipgen/lib/": b"",
        },
    )
    dest = tmp_path / "clipgen.new"
    assert updater.extract_bundle(archive, dest) is None
    assert (dest / "clipgen.exe").read_bytes() == b"exe"
    assert (dest / "lib" / "a.dll").read_bytes() == b"dll"


def test_extract_bundle_rejects_escapes_and_missing_exe(tmp_path):
    bad = _portable_zip(tmp_path / "bad.zip", {"clipgen/../evil.exe": b"x"})
    assert "unexpected archive member" in str(
        updater.extract_bundle(bad, tmp_path / "a")
    )
    no_exe = _portable_zip(tmp_path / "noexe.zip", {"clipgen/lib/a.dll": b"x"})
    assert (
        updater.extract_bundle(no_exe, tmp_path / "b") == "archive has no clipgen.exe"
    )


# ---- download ----------------------------------------------------------------


class _FakeResponse(io.BytesIO):
    def __init__(self, payload: bytes, headers: dict[str, str] | None = None):
        super().__init__(payload)
        self.headers = headers or {}


def _asset(payload: bytes, **overrides):
    asset = {
        "name": "clipgen-v9.9.9-macos.dmg",
        "size": len(payload),
        "sha256": hashlib.sha256(payload).hexdigest(),
        "url": "https://example.invalid/asset",
    }
    asset.update(overrides)
    return asset


def test_download_streams_verifies_and_renames(monkeypatch, tmp_path):
    payload = b"x" * 5000
    monkeypatch.setattr(
        updater.urllib.request,
        "urlopen",
        lambda req, timeout: _FakeResponse(
            payload, {"Content-Length": str(len(payload))}
        ),
    )
    monkeypatch.setattr(updater, "_DOWNLOAD_CHUNK", 1024)
    seen = []
    result = updater.download_update(
        _asset(payload), on_progress=lambda c, t: seen.append((c, t))
    )
    assert result == tmp_path / "cfg" / "updates" / "clipgen-v9.9.9-macos.dmg"
    assert result is not None
    assert result.read_bytes() == payload
    assert seen[-1] == (5000, 5000)
    assert not list(result.parent.glob("*.part"))


def test_download_drops_a_corrupt_file(monkeypatch, tmp_path):
    payload = b"y" * 100
    monkeypatch.setattr(
        updater.urllib.request, "urlopen", lambda req, timeout: _FakeResponse(payload)
    )
    assert updater.download_update(_asset(payload, sha256="0" * 64)) is None
    assert updater.download_update(_asset(payload, size=99)) is None
    assert not list((tmp_path / "cfg" / "updates").glob("*"))


def test_download_survives_a_network_error(monkeypatch, tmp_path):
    def boom(req, timeout):
        raise urllib.error.URLError("offline")

    monkeypatch.setattr(updater.urllib.request, "urlopen", boom)
    assert updater.download_update(_asset(b"z")) is None


# ---- check -------------------------------------------------------------------


def _release_payload(tag="v9.9.9"):
    return {
        "tag_name": tag,
        "html_url": f"https://github.com/henedl/clipgen/releases/tag/{tag}",
        "assets": [
            {
                "name": f"clipgen-{tag}-macos.dmg",
                "size": 10,
                "digest": "sha256:" + "a" * 64,
                "browser_download_url": "https://example.invalid/dmg",
            }
        ],
    }


def test_check_latest_persists_and_honours_the_cooldown(monkeypatch):
    calls = []

    def fake_urlopen(req, timeout):
        calls.append(dict(req.header_items()))
        body = json.dumps(_release_payload()).encode()
        return _FakeResponse(body, {"ETag": '"abc"'})

    monkeypatch.setattr(updater.urllib.request, "urlopen", fake_urlopen)
    release = updater.check_latest(force=True)
    assert release is not None
    assert release["tag"] == "v9.9.9"
    assert release["assets"][0]["sha256"] == "a" * 64
    state = start_settings.load_config_json(updater.STATE_FILENAME)
    assert state["etag"] == '"abc"' and state["latest"]["tag"] == "v9.9.9"
    # Inside the cooldown a launch check answers from the file.
    cached = updater.check_latest(force=False)
    assert cached is not None and cached["tag"] == "v9.9.9"
    assert len(calls) == 1
    # Forced: revalidates with the stored ETag.
    updater.check_latest(force=True)
    assert calls[1]["If-none-match"] == '"abc"'


def test_check_latest_skips_when_disabled_and_treats_304_as_cached(monkeypatch):
    monkeypatch.setattr(config, "UPDATE_CHECK_ON_LAUNCH", False)
    assert updater.check_latest(force=False) is None

    start_settings.save_config_json(
        updater.STATE_FILENAME,
        {
            "last_check": 0,
            "etag": '"e"',
            "latest": updater._normalize_release(_release_payload()),
        },
    )

    def not_modified(req, timeout):
        raise urllib.error.HTTPError(
            req.full_url, 304, "Not Modified", email.message.Message(), None
        )

    monkeypatch.setattr(updater.urllib.request, "urlopen", not_modified)
    revalidated = updater.check_latest(force=True)
    assert revalidated is not None and revalidated["tag"] == "v9.9.9"


def test_check_latest_returns_none_offline(monkeypatch):
    def boom(req, timeout):
        raise urllib.error.URLError("offline")

    monkeypatch.setattr(updater.urllib.request, "urlopen", boom)
    assert updater.check_latest(force=True) is None


def test_run_check_moves_to_available_or_idle(monkeypatch):
    monkeypatch.setattr(updater, "install_shape", lambda: "mac-app")
    monkeypatch.setattr(utils, "get_version", lambda: "0.1.0")
    monkeypatch.setattr(
        updater,
        "check_latest",
        lambda *, force: updater._normalize_release(_release_payload()),
    )
    updater.run_check(force=True)
    snap = updater.status()
    assert snap["phase"] == "available"
    assert snap["version"] == "v9.9.9"
    assert snap["asset"] == "clipgen-v9.9.9-macos.dmg"
    assert snap["total"] == 10

    monkeypatch.setattr(utils, "get_version", lambda: "9.9.9")
    updater.run_check(force=True)
    snap = updater.status()
    assert (
        snap["phase"] == "idle" and snap["checked"] is True and snap["version"] is None
    )


def test_offline_check_is_never_up_to_date(monkeypatch):
    monkeypatch.setattr(updater, "install_shape", lambda: "mac-app")
    monkeypatch.setattr(updater, "check_latest", lambda *, force: None)
    updater.run_check(force=False)
    assert updater.status()["checked"] is False
    updater.run_check(force=True)
    snap = updater.status()
    assert snap["phase"] == "error"
    assert snap["checked"] is True and snap["error"] == "Could not reach GitHub"


def test_run_check_reports_a_missing_asset(monkeypatch):
    monkeypatch.setattr(updater, "install_shape", lambda: "win-zip")
    monkeypatch.setattr(utils, "get_version", lambda: "0.1.0")
    monkeypatch.setattr(
        updater,
        "check_latest",
        lambda *, force: updater._normalize_release(_release_payload()),
    )
    updater.run_check(force=True)
    snap = updater.status()
    assert snap["phase"] == "error" and snap["error"] == "No download for this platform"


def test_skip_hides_the_release_until_a_manual_check(monkeypatch):
    monkeypatch.setattr(updater, "install_shape", lambda: "mac-app")
    monkeypatch.setattr(utils, "get_version", lambda: "0.1.0")
    monkeypatch.setattr(
        updater,
        "check_latest",
        lambda *, force: updater._normalize_release(_release_payload()),
    )
    updater.run_check(force=True)
    assert updater.status()["phase"] == "available"
    assert updater.skip_version() is True
    snap = updater.status()
    assert snap["phase"] == "idle" and snap["skipped"] == "v9.9.9"
    assert (
        start_settings.load_config_json(updater.STATE_FILENAME)["skipped"] == "v9.9.9"
    )
    # A launch check keeps it hidden; a manual check forgets the skip.
    updater.run_check(force=False)
    assert updater.status()["phase"] == "idle"
    updater.run_check(force=True)
    snap = updater.status()
    assert snap["phase"] == "available" and snap["skipped"] is None
    assert start_settings.load_config_json(updater.STATE_FILENAME)["skipped"] is None
    assert updater.skip_version() is True
    # Nothing to skip once idle.
    assert updater.skip_version() is False


def test_status_reports_the_auto_check_setting(monkeypatch):
    monkeypatch.setattr(config, "UPDATE_CHECK_ON_LAUNCH", False)
    assert updater.status()["auto_check"] is False


# ---- startup sweep -----------------------------------------------------------


def test_sweep_reports_the_helper_log_and_drops_stale_files(monkeypatch, tmp_path):
    monkeypatch.setattr(utils, "get_version", lambda: "1.0.0")
    monkeypatch.setattr(updater, "install_root", lambda: None)
    updates = tmp_path / "cfg" / "updates"
    updates.mkdir(parents=True)
    (updates / updater.APPLY_LOG).write_text("could not move the old app aside")
    (updates / "tmpabc.part").write_bytes(b"")
    (updates / "clipgen-v0.9.0-macos.dmg").write_bytes(b"old")
    (updates / "clipgen-v1.1.0-macos.dmg").write_bytes(b"new")
    updater.sweep_updates_dir()
    assert updater.status()["last_error"] == "could not move the old app aside"
    assert sorted(p.name for p in updates.iterdir()) == ["clipgen-v1.1.0-macos.dmg"]
