# clipgen-test — Run the test suite

## Command

```
uv run --extra dev pytest -c tests/pytest.ini
```

Run from the project root. The `--extra dev` flag is required: pytest is only in the optional `dev` extra and is not installed by `uv sync` alone.

To filter to specific tests, append the file or test name:
```
uv run --extra dev pytest -c tests/pytest.ini tests/test_utils_timestamps.py
uv run --extra dev pytest -c tests/pytest.ini -k "test_batch"
```

## Key fixtures (conftest.py)

- `make_clip(**overrides)` — builds a minimal clip record dict; pass keyword args to override any field
- `fake_sheet_meta()` — returns a mock sheet metadata object

## Choosing what kind of test to write

Pick one kind and write it in the right place; do not mix kinds in one file.

| Kind | Example | Default `/check`? |
|---|---|---|
| **Ratchet** | `test_frontend_satellite_wiring.py`, `test_packaging.py`, `test_shared_constants.py` | yes, must stay fast |
| **Domain** | `test_utils_timestamps.py` | yes |
| **API** | `test_studio_api.py` route + JSON under mocks | yes |
| **I/O** | real `process_clips` cut of a 2s `testsrc` in `test_pipeline_io.py` | yes, max four, `skipif` no ffmpeg |
| **Journey** | Studio generate → artifact appears | no, `tests/ui/` only |

Teaching copies: domain → `tests/test_utils_timestamps.py`; API → a *small* test in
`tests/test_studio_api.py` (not the whole file); journey → `tests/ui/test_ui_smoke.py`
+ `tests/ui/_ui_states.py`.

Selection rules:

- **Source scans are frozen.** No new `assert "<code>" in src` unless the bug is
  "this symbol vanished" — and then prefer a shrink-only allowlist (`KNOWN_DEAD` in
  `tests/test_js_dead_functions.py` is the pattern). A frontend test pins a **rule**
  or a **shipped bug**, never a spelling.
- The only additive-by-default rule is the existing one below: every new CLI mode,
  flag, or selector gets at least one smoke test.
- Bug was "page booted but the click did nothing" → add a journey in
  `tests/ui/test_ui_journeys.py`, and only if the journey list is still ≤ 5.
- Bug was "ffmpeg produced garbage / succeeded with partial output" → add an I/O
  test in `tests/test_pipeline_io.py`, not another mocked argv assert. Check
  `tests/test_clip_pipeline.py` first; the silent-wrong reel class is already
  locked under mocks there.
- Do not add pytest markers (`unit` / `integration`). The `tests/ui/` directory
  split is the slow-path marker.

## CLI test pattern

CLI argument tests use a local `_args(**overrides)` helper that builds a full argparse namespace with defaults, then overrides specific keys. Look for this pattern in `tests/test_cli_args.py` and `tests/test_cli_modes.py` when adding new flag tests.

## Test coverage areas

CLI args, CLI modes, clip pipeline, file/artifact handling, Google/Excel adapters, manifest, selectors, spreadsheet generation, studio API, titlecards, transcripts, timestamp utilities, video commands, viewer data, viewer inlining, Screenspace, shared constants.

Every new CLI mode, flag, or selector needs at least one smoke test.

## Speed

The suite runs on every `/check` and every push, so a slow test is a permanent tax. If a test
you added shows up in `--durations=20`, work through
[test-perf](../test-perf/SKILL.md) — it covers the four patterns that have actually cost this
repo minutes of CI (waits whose release never fires, per-item regex over a whole corpus,
uncached scan helpers, real production poll intervals) and how to fix one without quietly
turning the test into a no-op.

For a faster local inner loop (not a CI change — `pytest-xdist` stays out of `pyproject.toml`):

```
uv run --extra dev --with pytest-xdist pytest -c tests/pytest.ini -n auto
```
