# clipgen-test — Run the test suite

## Command

```
uv run --extra dev pytest -c tests/pytest.ini
```

Run from the project root. The `--extra dev` flag is required — pytest is only in the optional `dev` extra and is not installed by `uv sync` alone.

To filter to specific tests, append the file or test name:
```
uv run --extra dev pytest -c tests/pytest.ini tests/test_utils_timestamps.py
uv run --extra dev pytest -c tests/pytest.ini -k "test_batch"
```

## Key fixtures (conftest.py)

- `make_clip(**overrides)` — builds a minimal clip record dict; pass keyword args to override any field
- `fake_sheet_meta()` — returns a mock sheet metadata object

## CLI test pattern

CLI argument tests use a local `_args(**overrides)` helper that builds a full argparse namespace with defaults, then overrides specific keys. Look for this pattern in `tests/test_cli_args.py` and `tests/test_cli_modes.py` when adding new flag tests.

## Test coverage areas

CLI args, CLI modes, clip pipeline, file/artifact handling, Google/Excel adapters, insights, manifest, selectors, spreadsheet generation, studio API, titlecards, transcripts, timestamp utilities, video commands, viewer data, viewer inlining, Screenspace, shared constants.

Every new CLI mode, flag, or selector needs at least one smoke test.
