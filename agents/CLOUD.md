# Cloud specific instructions

## Environment overview

clipgen is a self-contained Python CLI tool with no databases or Docker services. System dependencies (`ffmpeg`, `ffprobe`, Python 3.12) are pre-installed on the VM; only `uv` and Python packages need to be installed at session start.

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

`--no-sync` on the tests command is a cloud-specific optimization: it relies on the prior `uv pip install ".[dev]" --torch-backend cpu` step having already installed pytest and CPU torch, and skips the lockfile resolution that would otherwise re-pull the (much larger) GPU torch packages. Locally — where you have not done the cpu-only install — use AGENTS.md's `uv run --extra dev pytest ...` instead.

All web UIs serve on `http://127.0.0.1:8089`. The Flask server starts automatically when launching `--studio`, `--screenspace`, or `--insights`.

## Gotchas

- `uv run` auto-creates `.venv` if missing; always prefer `uv run` over activating the venv manually.
- The first `uv run clipgen.py` invocation after a CPU-only pip install may re-resolve and download GPU torch packages via `uv.lock`. To avoid this: use `uv run --no-sync` for any command after the initial `uv pip install ".[dev]" --torch-backend cpu`, or use `uvx` (which doesn't trigger a sync) for tools like ruff and ty.
- Google Sheets integration requires OAuth credentials not available in Cloud Agent VMs. Use local Excel files or mocked tests instead.
- DBus errors in the Flask server log are harmless — they come from Chrome attempting DBus connections in the headless VM environment.
