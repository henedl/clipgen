---
name: clipgen-bump-pins
description: Update the pinned ffmpeg / llama.cpp / OCR-model / PyInstaller versions the desktop bundle ships
---

# Bumping pinned binaries

Every third-party binary in the desktop bundle is pinned by URL + SHA256 in
`build/fetch_binaries.py` (`PINS` per platform, `OCR_MODEL_PINS` for the RapidOCR
recognition models). PyInstaller is pinned once, in the `build` extra of `pyproject.toml`.
Nothing updates these automatically — Dependabot cannot see `fetch_binaries.py` — so a
bump is a deliberate PR. The weekly `pin-health` workflow only tells you when a pinned
URL has vanished.

## When to bump

| Pin | Bump when | Not when |
|---|---|---|
| llama.cpp (`b…` release, both platforms) | a feature or fix needs a newer build, or ~quarterly alongside a release; re-run the router-mode gate validation | upstream cut a new build (it does, several times a day) |
| ffmpeg (macOS: Martin Riedl; Windows: gyan.dev) | an upstream point release (8.1.x → 8.2) or a feature gap | — never pin a moving alias (`release-essentials`, `latest`) |
| RapidOCR models | only together with a `rapidocr` bump in `pyproject.toml`; re-derive URLs + hashes from the new wheel's `default_models.yaml` | on their own |
| PyInstaller | with the regular dependency refresh, as its own PR | — |

One pin per `build(deps):` PR. No `build/VERSION` bump.

## Procedure (ffmpeg / llama.cpp)

1. Pick the new asset URL on the provider. Permanent URLs only: gyan.dev
   `/builds/packages/ffmpeg-<ver>-essentials_build.zip` (the `full_build` is `.7z`, which
   the stdlib-only fetch script cannot extract), Martin Riedl's per-build paths, llama.cpp
   GitHub release assets. llama.cpp needs both platforms' URLs.
2. `uv run build/fetch_binaries.py --repin <url> [<url> ...]` — finds the entry each URL
   replaces (same host; platform token in the filename), checks the archive against the
   provider's `<url>.sha256` when one is published, re-hashes every member the old entry
   pinned, and rewrites the entry in place. A member missing from the new archive aborts
   with the list — the provider's closure changed, edit by hand.
3. Update the provenance block at the top of `fetch_binaries.py` and the matching
   paragraph in `build/THIRD-PARTY-LICENSES`.
4. `rm -rf build/vendor && uv run build/fetch_binaries.py` — the real fetch, against the
   rewritten hashes.
5. `uv run --extra dev pytest -c tests/pytest.ini tests/test_packaging.py` — member paths
   must belong to their archive, OCR pins must match `screenspace_ocr`, and the rendered
   entries must round-trip.
6. Open the PR, then `gh workflow run build-binaries.yml --ref <branch>`. CI greps the
   in-bundle ffmpeg for libx264 / libwebp / libvpx-vp9 / drawtext and runs
   `llama-server --version`; Windows can only be verified this way, say so in the PR.
7. Download the artifact and launch it the way a user does (`open dist/clipgen.app`) —
   [build](../build/SKILL.md) explains why `--help` proves nothing.

## Procedure (OCR models)

Bump `rapidocr` in `pyproject.toml`, `uv lock`, then open the installed wheel's
`default_models.yaml` (onnxruntime section) and copy each rec model's URL + SHA256 into
`OCR_MODEL_PINS`. `--repin` works here too once the URL is known. Keep the PP-OCR version
rule in step with `screenspace_ocr._build_ocr_reader` (v4 for japan, v5 otherwise; the
test enforces it).

## Procedure (PyInstaller / opencv / ruff)

- PyInstaller: edit the `build` extra in `pyproject.toml`, `uv lock`. Nothing else names
  the version; the test fails if something starts to.
- opencv: the `>=` floor in `pyproject.toml` must equal the sdist URL, cache key, and
  extracted-dir name in `build-binaries.yml` — four places, one test.
- ruff: `required-version` in `pyproject.toml` must equal the `uvx ruff@…` pin in
  `tests.yml`.

## Health check

`uv run build/fetch_binaries.py --check-urls` probes every pinned URL. The `pin-health`
workflow runs it Mondays and on dispatch, plus `uv lock --check`.
