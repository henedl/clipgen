"""Guard the source layout: `source/` holds every module, and `pyproject.toml`
lists them all.

The project is a flat set of top-level modules in ``source/``, imported by bare
name (``import config``). ``source/`` is not a package — no ``__init__.py`` — it
is just put on ``sys.path`` by the ``clipgen.py`` launcher for real runs and by
``tests/conftest.py`` under pytest.

Two drifts, both silent from the source tree:

* A module missing from ``[tool.setuptools] py-modules`` imports fine locally, but
  ``uv pip install .`` ships only the listed ones, so installed and frozen
  environments die with ``ModuleNotFoundError``.
* A new module in the *repo root* also imports fine locally (the root is still
  ``sys.path[1]``) and PyInstaller still bundles it, so nothing fails until an
  install — by which time the flat root has quietly grown back.
"""

import itertools
import os
import re
import subprocess
import sys
import tomllib
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
_SOURCE = _ROOT / "source"


def _source_modules() -> set[str]:
    """Top-level importable modules: ``source/*.py`` minus private/dunder files."""
    return {p.stem for p in _SOURCE.glob("*.py") if not p.name.startswith("_")}


def _listed_modules() -> set[str]:
    data = tomllib.loads((_ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    return set(data["tool"]["setuptools"]["py-modules"])


def test_py_modules_covers_every_source_module() -> None:
    missing = _source_modules() - _listed_modules()
    assert not missing, (
        "source/ modules missing from pyproject [tool.setuptools] py-modules "
        f"(install would break `import` of these): {sorted(missing)}"
    )


def test_py_modules_has_no_phantom_entries() -> None:
    phantom = _listed_modules() - _source_modules()
    assert not phantom, (
        "pyproject py-modules lists modules with no matching source/*.py file: "
        f"{sorted(phantom)}"
    )


def test_package_dir_points_at_source() -> None:
    """``py-modules`` names resolve to ``source/<name>.py``, not the repo root."""
    data = tomllib.loads((_ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    assert data["tool"]["setuptools"]["package-dir"] == {"": "source"}


def test_repo_root_holds_only_the_launcher() -> None:
    """No product module may sit beside ``clipgen.py`` again."""
    root_modules = {p.name for p in _ROOT.glob("*.py")}
    assert root_modules == {"clipgen.py"}, (
        "the repo root must hold only the clipgen.py launcher; product modules "
        f"belong in source/. Found: {sorted(root_modules)}"
    )


def test_source_dir_is_not_a_package() -> None:
    """``source/__init__.py`` would break bare-name imports and ty's discovery.

    With an ``__init__.py`` present, ``source/`` becomes a package: ``import
    config`` no longer resolves from it, and ty stops treating it as a module
    root. Both failures are wholesale rather than local, so guard the absence.
    """
    assert not (_SOURCE / "__init__.py").exists()


def _fetch_binaries_module():
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "fetch_binaries", _ROOT / "build" / "fetch_binaries.py"
    )
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_vendor_pins_cover_both_platforms_and_are_wellformed() -> None:
    """The desktop bundle ships pinned ffmpeg/ffprobe + llama-server; a
    malformed or partial PINS table only surfaces on the (slow, per-platform)
    release build."""
    pins = _fetch_binaries_module().PINS
    required_targets = {
        "macos-arm64": {"ffmpeg", "ffprobe", "llama-server"},
        "windows-x64": {"ffmpeg.exe", "ffprobe.exe", "llama-server.exe"},
    }
    assert set(pins) == set(required_targets)
    for plat, archives in pins.items():
        targets = set()
        for archive in archives:
            assert archive["url"].startswith("https://"), archive["url"]
            assert _looks_like_sha256(archive["sha256"]), archive["url"]
            for member in archive["members"].values():
                assert _looks_like_sha256(member["sha256"]), member["target"]
                targets.add(member["target"])
        # Superset: llama-server rides with its dylib/DLL closure, whose file
        # list is the pin's concern, not this test's.
        assert required_targets[plat] <= targets, plat
        assert len(targets) == sum(len(a["members"]) for a in archives), (
            f"{plat}: duplicate member targets"
        )


def _looks_like_sha256(value: str) -> bool:
    return len(value) == 64 and all(c in "0123456789abcdef" for c in value)


def test_vendor_pin_member_paths_belong_to_their_archive() -> None:
    """A half-bumped entry (new URL, stale ``llama-b…/`` member paths) only
    fails at extraction on the slow per-platform release build."""
    pins = _fetch_binaries_module().PINS
    for archives in pins.values():
        for archive in archives:
            archive_name = archive["url"].rsplit("/", 1)[1]
            for member in archive["members"]:
                top, sep, _rest = member.partition("/")
                if sep:
                    assert top in archive_name, f"{member} is not from {archive_name}"


def test_vendor_pin_entries_round_trip_through_repin() -> None:
    """``--repin`` rewrites an entry by rendering it; the rendering must match
    the file's (ruff-formatted) layout or every bump produces a noisy diff."""
    module = _fetch_binaries_module()
    text = (_ROOT / "build" / "fetch_binaries.py").read_text(encoding="utf-8")
    entries = [a for archives in module.PINS.values() for a in archives]
    entries += module.OCR_MODEL_PINS + module.SPEAKER_MODEL_PINS
    for entry in entries:
        indent = 8 if "members" in entry else 4
        rendered = "".join(module._render_entry(entry, indent))
        assert rendered in text, f"rendering drifted for {entry['url']}"


def test_member_target_strips_dylib_minor_versions() -> None:
    """``--repin`` matches new archive members to old targets by this rule."""
    target = _fetch_binaries_module()._member_target
    assert target("llama-b10588/libggml.0.21.0.dylib") == "libggml.0.dylib"
    assert target("llama-b10588/llama-server") == "llama-server"
    assert target("ffmpeg-8.1.2-essentials_build/bin/ffmpeg.exe") == "ffmpeg.exe"
    assert target("ggml-cpu-x64.dll") == "ggml-cpu-x64.dll"


def test_ocr_model_pins_match_the_ocr_module() -> None:
    """The vendored rec models must cover every non-default family the OCR
    module maps languages to, at the PP-OCR version its dev fallback asks for."""
    import screenspace_ocr

    pins = _fetch_binaries_module().OCR_MODEL_PINS
    families = {
        model
        for model in screenspace_ocr._OCR_LANG_TO_MODEL.values()
        if model != screenspace_ocr._OCR_MODEL_DEFAULT
    }
    assert {pin["target"] for pin in pins} == {f"{m}_rec.onnx" for m in families}
    for pin in pins:
        family = pin["target"].removesuffix("_rec.onnx")
        # Mirrors _build_ocr_reader: japan has no PP-OCRv5 ONNX model upstream.
        version = "PP-OCRv4" if family == "japan" else "PP-OCRv5"
        assert f"/{version}/" in pin["url"], pin["url"]
        assert f"{family}_{version}_rec_mobile.onnx" in pin["url"], pin["url"]


def _pyproject() -> dict:
    return tomllib.loads((_ROOT / "pyproject.toml").read_text(encoding="utf-8"))


def test_pyinstaller_is_pinned_once() -> None:
    """The ``build`` extra is the only place PyInstaller's version lives."""
    build = _pyproject()["project"]["optional-dependencies"]["build"]
    assert len(build) == 1 and re.fullmatch(r"pyinstaller==[\d.]+", build[0]), build
    for rel in (
        ".github/workflows/build-binaries.yml",
        "README.md",
        "agents/skills/build/SKILL.md",
    ):
        text = (_ROOT / rel).read_text(encoding="utf-8")
        assert "pyinstaller==" not in text, f"{rel} restates the PyInstaller pin"
    workflow = (_ROOT / ".github" / "workflows" / "build-binaries.yml").read_text(
        encoding="utf-8"
    )
    assert "uv sync --locked --extra build" in workflow


def test_opencv_pin_agrees_across_pyproject_and_ci() -> None:
    """The macOS leg rebuilds opencv from its sdist: floor, sdist URL, cache
    key, and extracted dir must all name the same version."""
    dep = next(
        d
        for d in _pyproject()["project"]["dependencies"]
        if d.startswith("opencv-python-headless")
    )
    match = re.search(r">=([\d.]+)", dep)
    assert match, dep
    version = match.group(1)
    workflow = (_ROOT / ".github" / "workflows" / "build-binaries.yml").read_text(
        encoding="utf-8"
    )
    assert f"opencv_python_headless-{version}.tar.gz" in workflow
    assert f"opencv-noffmpeg-{version}-" in workflow
    assert f"/tmp/opencv_python_headless-{version}" in workflow
    other = set(re.findall(r"opencv[_-]python[_-]headless-(\d+(?:\.\d+)*)", workflow))
    assert other == {version}, other


def test_ruff_pin_agrees_across_pyproject_and_ci() -> None:
    required = _pyproject()["tool"]["ruff"]["required-version"]
    version = required.removeprefix(">=")
    workflow = (_ROOT / ".github" / "workflows" / "tests.yml").read_text(
        encoding="utf-8"
    )
    pinned = set(re.findall(r"uvx ruff@([\d.]+)", workflow))
    assert pinned == {version}, (required, pinned)


def test_pin_health_workflow_probes_every_url() -> None:
    """Pinned downloads live on other people's servers; a vanished asset only
    shows up on a release-day cache miss unless something probes weekly."""
    workflow = (_ROOT / ".github" / "workflows" / "pin-health.yml").read_text(
        encoding="utf-8"
    )
    assert "schedule:" in workflow
    assert "workflow_dispatch:" in workflow
    assert "build/fetch_binaries.py --check-urls" in workflow
    assert "uv lock --check" in workflow


def test_spec_guards_and_upx_excludes_the_vendored_tools() -> None:
    """The spec must refuse to build without the fetched binaries (a silent
    skip ships an app that dies on its startup ffmpeg check), and UPX must
    never touch them (it corrupts signed mach-O / packed ffmpeg.exe)."""
    spec_text = (_ROOT / "build" / "clipgen.spec").read_text(encoding="utf-8")
    assert "fetch_binaries.py" in spec_text, "vendor guard missing from spec"
    _upx = 'upx_exclude=["ffmpeg*", "ffprobe*", "llama*", "libllama*", "libggml*", "ggml*", "libomp*", "mtmd*"]'
    assert spec_text.count(_upx) == 2, (
        "both EXE and COLLECT must exclude the bundled tools from UPX"
    )
    assert "binaries=binaries" in spec_text, (
        "Analysis must collect the vendored binaries list"
    )


def test_spec_advertises_macos_14_minimum() -> None:
    """Wheels in the lockfile are macosx_14_0; 11.0 was a false Gatekeeper floor."""
    spec_text = (_ROOT / "build" / "clipgen.spec").read_text(encoding="utf-8")
    assert '"LSMinimumSystemVersion": "14.0"' in spec_text
    assert '"LSMinimumSystemVersion": "11.0"' not in spec_text


def test_ci_verifies_all_vendored_ocr_models() -> None:
    """Post-build smoke must require japan/korean rec, not only latin."""
    yml = (_ROOT / ".github" / "workflows" / "build-binaries.yml").read_text(
        encoding="utf-8"
    )
    module = _fetch_binaries_module()
    for pin in module.OCR_MODEL_PINS + module.SPEAKER_MODEL_PINS:
        assert pin["target"] in yml, pin["target"]


def test_speaker_model_pin_matches_speakers_module() -> None:
    """One 16 kHz sherpa-onnx export, named as speakers.py looks it up."""
    import speakers

    pins = _fetch_binaries_module().SPEAKER_MODEL_PINS
    assert len(pins) == 1
    pin = pins[0]
    assert pin["target"] == speakers.MODEL_FILENAME
    assert pin["url"].startswith(
        "https://github.com/k2-fsa/sherpa-onnx/releases/download/"
    )
    assert pin["url"].endswith("_16k.onnx") or "wespeaker" in pin["url"]
    assert re.fullmatch(r"[0-9a-f]{64}", pin["sha256"])
    spec_text = (_ROOT / "build" / "clipgen.spec").read_text(encoding="utf-8")
    assert '"speaker_models"' in spec_text


def test_spec_never_strips_or_packs_on_windows() -> None:
    """``strip``/``upx`` must stay macOS-only, or Windows ships a dead exe.

    PyInstaller shells out to whatever ``strip`` is on PATH. The GitHub windows
    runner carries a GNU binutils ``strip``, which happily rewrites MSVC-built
    PE DLLs into images the Windows loader rejects — ``python312.dll`` included.
    The build stays green (``strip`` returns 0), every file is present and
    plausibly sized, and the shipped ``clipgen.exe`` dies before any Python runs
    with ``Failed to load Python DLL ... LoadLibrary: Invalid access to memory
    location``. That shipped in every Windows build until the ``_strip`` gate.
    """
    spec_text = (_ROOT / "build" / "clipgen.spec").read_text(encoding="utf-8")
    assert '_strip = sys.platform == "darwin"' in spec_text, (
        "the spec must gate `strip` to macOS; GNU strip corrupts Windows PE DLLs"
    )
    assert spec_text.count("strip=_strip") == 2, (
        "both EXE and COLLECT must take `strip` from the platform gate"
    )
    assert "strip=True" not in spec_text, (
        "`strip=True` is unconditional and breaks the Windows bundle"
    )
    assert spec_text.count("upx=False") == 2, (
        "UPX corrupts the same PE DLLs; keep it off in both EXE and COLLECT"
    )


def test_build_workflow_smoke_launches_both_bundles() -> None:
    """A green PyInstaller build proves nothing about whether the exe launches.

    The bootloader loads ``python312.dll`` before argv is parsed, so ``--help``
    is enough to catch a bundle whose DLLs were corrupted at build time — the
    one failure mode the feature-verification steps all miss.
    """
    workflow = (_ROOT / ".github" / "workflows" / "build-binaries.yml").read_text(
        encoding="utf-8"
    )
    assert "dist/clipgen/clipgen.exe --help" in workflow, (
        "the Windows leg must smoke-launch the frozen exe"
    )
    assert "./dist/clipgen.app/Contents/MacOS/clipgen --help" in workflow, (
        "the macOS leg must smoke-launch the frozen binary"
    )


def test_windows_contents_dir_is_lib_everywhere() -> None:
    """COLLECT renames PyInstaller's ``_internal`` to ``lib``. Runtime code is
    name-agnostic (everything derives from ``sys._MEIPASS``), but the workflow's
    Windows verification steps address the folder literally — a stale path there
    fails CI at build time, and stale docs mislead users."""
    spec_text = (_ROOT / "build" / "clipgen.spec").read_text(encoding="utf-8")
    assert 'contents_directory="lib"' in spec_text
    workflow = (_ROOT / ".github" / "workflows" / "build-binaries.yml").read_text(
        encoding="utf-8"
    )
    assert "_internal" not in workflow, (
        "the workflow still addresses PyInstaller's default contents dir"
    )
    assert "dist/clipgen/lib/bin/ffmpeg.exe" in workflow


def test_installer_script_is_tracked_and_wired() -> None:
    """``build/*`` is gitignored with an allowlist, so a missing ``!`` entry
    means the .iss exists locally but never commits; and the installer must be
    versioned from CI and land in the release/artifact globs."""
    gitignore = (_ROOT / ".gitignore").read_text(encoding="utf-8")
    assert "!build/clipgen.iss" in gitignore.splitlines()

    iss = (_ROOT / "build" / "clipgen.iss").read_text(encoding="utf-8")
    assert "AppVersion={#AppVer}" in iss, (
        "the version must come from ISCC /DAppVer, never hardcoded in the .iss"
    )
    assert "PrivilegesRequired=lowest" in iss, (
        "the installer is per-user by design; no UAC prompt"
    )

    workflow = (_ROOT / ".github" / "workflows" / "build-binaries.yml").read_text(
        encoding="utf-8"
    )
    assert "clipgen.iss" in workflow, "CI must compile the installer"
    assert workflow.count("dist/*-setup.exe") >= 2, (
        "the installer must be in both the artifact path and the release glob"
    )


def test_gitignore_allowlists_the_fetch_script() -> None:
    """``build/*`` is gitignored with a ``!`` allowlist; without its entry the
    fetch script exists locally but never lands in a commit."""
    gitignore = (_ROOT / ".gitignore").read_text(encoding="utf-8")
    assert "!build/fetch_binaries.py" in gitignore.splitlines()


def test_gitignore_allowlists_the_release_notes_script() -> None:
    """Same allowlist trap as the fetch script: without the ``!`` entry the
    renderer exists locally, never lands in a commit, and CI keeps publishing
    releases with the old empty body."""
    gitignore = (_ROOT / ".gitignore").read_text(encoding="utf-8")
    assert "!build/release_notes.py" in gitignore.splitlines()


def test_release_notes_workflow_fetches_full_history() -> None:
    """``build/release_notes.py`` reads the tag range with ``git log``. Under the
    default shallow checkout that returns nothing, so the release publishes with
    an empty What's Changed and no error anywhere."""
    workflow = (_ROOT / ".github" / "workflows" / "release-notes.yml").read_text(
        encoding="utf-8"
    )
    assert "fetch-depth: 0" in workflow


def test_release_notes_workflow_updates_an_existing_release() -> None:
    """build-binaries.yml publishes assets on the same tag with an action that
    *creates* the release when absent. If this workflow skips a release that
    already exists, losing that race leaves the body permanently empty — which is
    the behaviour the create-or-update branch replaced."""
    workflow = (_ROOT / ".github" / "workflows" / "release-notes.yml").read_text(
        encoding="utf-8"
    )
    assert "gh release edit" in workflow
    assert "skipping (idempotent rerun)" not in workflow


def test_build_workflow_does_not_set_a_release_body() -> None:
    """A ``body`` input on the asset-upload step would overwrite the generated
    notes; leaving it unset is what makes the two workflows order-independent."""
    workflow = (_ROOT / ".github" / "workflows" / "build-binaries.yml").read_text(
        encoding="utf-8"
    )
    # A YAML key, not the prose warning in the comment above that step.
    assert re.search(r"(?m)^\s*body:", workflow) is None


def test_ui_suite_stays_opt_in() -> None:
    """ui-check/SKILL.md: browser tests never run in /check or CI.

    Two independent locks per file plus the pytest.ini exclusion; a CI
    workflow naming ``tests/ui`` would pull Playwright into every run.
    """
    pytest_ini = (_ROOT / "tests" / "pytest.ini").read_text(encoding="utf-8")
    assert re.search(r"(?m)^norecursedirs = .*\bui\b", pytest_ini)
    for name in ("test_ui_smoke.py", "test_ui_journeys.py"):
        text = (_ROOT / "tests" / "ui" / name).read_text(encoding="utf-8")
        assert "CLIPGEN_UI_CHECK" in text, name
        assert "pytest.mark.ui" in text, name
    for workflow in (_ROOT / ".github" / "workflows").glob("*.yml"):
        assert "tests/ui" not in workflow.read_text(encoding="utf-8"), workflow.name


def test_no_markers_beyond_ui() -> None:
    """test/SKILL.md: the tests/ui split is the only slow-path marker."""
    pytest_ini = (_ROOT / "tests" / "pytest.ini").read_text(encoding="utf-8")
    markers = re.findall(r"(?m)^    (\w+):", pytest_ini)
    assert markers == ["ui"], markers


def test_xdist_stays_out_of_the_lock() -> None:
    """test/SKILL.md: pytest-xdist is CI/local opt-in, never a dependency."""
    data = tomllib.loads((_ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    declared = list(data["project"].get("dependencies", []))
    for extra in data["project"].get("optional-dependencies", {}).values():
        declared.extend(extra)
    assert not [d for d in declared if "xdist" in d], declared


def test_context_dir_stays_gitignored() -> None:
    """ui-check/SKILL.md: the agent scratch subtree must never be committed."""
    gitignore = (_ROOT / ".gitignore").read_text(encoding="utf-8")
    assert ".context/" in gitignore.splitlines()


def test_license_notice_is_tracked_and_bundled() -> None:
    """The bundle ships GPL/LGPL/MPL software, so the notice must travel with it.

    Three ways it could silently go missing, all covered here: dropping out of
    the ``build/*`` gitignore allowlist, dropping out of the spec's ``datas``
    (which is what ``--licenses`` reads inside a frozen app), or the file simply
    not being there. None of these break a build or a test run on their own.
    """
    gitignore = (_ROOT / ".gitignore").read_text(encoding="utf-8")
    assert "!build/THIRD-PARTY-LICENSES" in gitignore.splitlines()

    spec_text = (_ROOT / "build" / "clipgen.spec").read_text(encoding="utf-8")
    assert 'datas += [("THIRD-PARTY-LICENSES", ".")]' in spec_text, (
        "the notice must be in the spec's datas or `--licenses` finds nothing "
        "in a frozen bundle"
    )

    notice = _ROOT / "build" / "THIRD-PARTY-LICENSES"
    assert notice.is_file()
    text = notice.read_text(encoding="utf-8")
    # The copyleft sections carry actual obligations; a truncated or
    # accidentally-regenerated notice would still look plausible without them.
    for heading in (
        "GPL-3.0-OR-LATER (bundled ffmpeg and ffprobe executables)",
        "LGPL-2.1 (FFmpeg DLL bundled in the Windows opencv-python-headless wheel)",
        "MOZILLA PUBLIC LICENSE 2.0",
    ):
        assert heading in text, f"THIRD-PARTY-LICENSES lost its {heading!r} section"
    # The macOS cv2 is self-built with -DWITH_FFMPEG=OFF precisely so the DMG
    # carries no in-process GPL code (plans/OPENCV-SELF-COMPILE-PLAN.md); a
    # revived GPL-FFmpeg-inside-opencv section would mean that regressed.
    assert "FFmpeg bundled inside opencv-python-headless" not in text, (
        "THIRD-PARTY-LICENSES regrew the GPL-FFmpeg-inside-cv2 section; the "
        "macOS build compiles cv2 without FFmpeg and the DMG is conveyed MIT"
    )


def test_license_notice_sections_are_well_formed() -> None:
    """Every heading sits between exactly one rule above and one below.

    The file grew two malformed rules before anyone looked (a doubled rule, and
    a heading with no opening rule). Nothing reads the file programmatically
    today, but the layout is the only thing making a 1700-line notice navigable.
    """
    rule = "=" * 79
    lines = (
        (_ROOT / "build" / "THIRD-PARTY-LICENSES")
        .read_text(encoding="utf-8")
        .splitlines()
    )
    rules = [i for i, line in enumerate(lines) if line == rule]
    assert rules, "no section rules found"

    doubled = [i for a, i in itertools.pairwise(rules) if i == a + 1]
    assert not doubled, f"consecutive section rules at lines {[i + 1 for i in doubled]}"

    for i in rules:
        # A rule either opens a section (heading + closing rule follow) or
        # closes one (heading + opening rule precede). Anything else is a stray.
        opens = i + 2 < len(lines) and lines[i + 2] == rule
        closes = i >= 2 and lines[i - 2] == rule
        assert opens or closes, f"stray section rule at line {i + 1}"


def test_launcher_bootstraps_from_a_foreign_cwd() -> None:
    """``clipgen.py`` must find ``source/`` without help from the environment.

    Runs with ``PYTHONPATH`` cleared and from a directory that is not the repo,
    which is the shape of every real invocation. ``--help`` exits inside
    ``parse_arguments()``, before ``main`` chdirs or touches any media.
    """
    env = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    result = subprocess.run(
        [sys.executable, str(_ROOT / "clipgen.py"), "--help"],
        cwd=Path(sys.prefix),
        env=env,
        check=False,
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stderr or result.stdout
    assert "clipgen" in result.stdout
