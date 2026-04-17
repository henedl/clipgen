# Cloud specific instructions

## Environment overview

clipgen is a self-contained Python CLI tool with no databases or Docker services. System dependencies (`ffmpeg`, `ffprobe`, Python 3.12) are pre-installed on the VM. The update script handles `uv` installation and Python dependency sync only.

## Dependency installation

- Dependencies are installed via `uv pip install ".[dev]" --torch-backend cpu`.
- The `--torch-backend cpu` flag avoids downloading ~2.5 GB of CUDA/nvidia packages that are not needed in this environment.
- Do **not** use plain `uv sync` for Cloud Agent sessions — it pulls GPU-capable torch and is much slower.

## Running commands

| Task | Command |
|------|---------|
| Lint | `uvx ruff check` and `uvx ruff format --check` |
| Type check | `uvx ty check` |
| Tests | `uv run --no-sync pytest -c tests/pytest.ini` |
| Run app (CLI help) | `uv run clipgen.py --help` |
| Run Screenspace UI | `uv run clipgen.py --screenspace -i DIR -o DIR` (no spreadsheet needed) |
| Run Insights UI | `uv run clipgen.py --insights -i DIR -o DIR` (reads from manifest) |
| Run Studio UI | `uv run clipgen.py --studio` (requires a spreadsheet) |

All web UIs serve on `http://127.0.0.1:8089`. The Flask server starts automatically when launching `--studio`, `--screenspace`, or `--insights`.

## Gotchas

- `uv run` auto-creates `.venv` if missing; always prefer `uv run` over activating the venv manually.
- The first `uv run clipgen.py` invocation after a CPU-only pip install may re-resolve and download GPU torch packages via `uv.lock`. To avoid this, use `uv run --no-sync` after the initial `uv pip install ".[dev]" --torch-backend cpu` setup, or run commands that don't trigger a full sync (like `uvx` for ruff/ty).
- Google Sheets integration requires OAuth credentials not available in Cloud Agent VMs. Use local Excel files or mocked tests instead.
- DBus errors in the Flask server log are harmless — they come from Chrome attempting DBus connections in the headless VM environment.
