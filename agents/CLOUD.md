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
| Run Transcripts UI | `uv run clipgen.py --transcripts -i DIR -o DIR` (no spreadsheet needed) |
| Run Studio UI | `uv run clipgen.py --studio` (requires a spreadsheet) |

`--no-sync` on the tests command is a cloud-specific optimization: it relies on the prior `uv pip install ".[dev]" --torch-backend cpu` step having already installed pytest and CPU torch, and skips the lockfile resolution that would otherwise re-pull the (much larger) GPU torch packages. Locally — where you have not done the cpu-only install — use AGENTS.md's `uv run --extra dev pytest ...` instead.

All web UIs serve on `http://127.0.0.1:8089`. The Flask server starts automatically when launching `--studio`, `--screenspace`, or `--transcripts`.

## Gotchas

- `uv run` auto-creates `.venv` if missing; always prefer `uv run` over activating the venv manually.
- The first `uv run clipgen.py` invocation after a CPU-only pip install may re-resolve and download GPU torch packages via `uv.lock`. To avoid this: use `uv run --no-sync` for any command after the initial `uv pip install ".[dev]" --torch-backend cpu`, or use `uvx` (which doesn't trigger a sync) for tools like ruff and ty.
- Google Sheets integration requires OAuth credentials not available in Cloud Agent VMs. Use local Excel files or mocked tests instead.
- DBus errors in the Flask server log are harmless — they come from Chrome attempting DBus connections in the headless VM environment.

## Git commits and PR attribution (Cloud Agent)

Cloud Agent workspaces set `core.hooksPath` in `.git/config` to a **Cursor-managed hooks directory** that rewrites commits: it can force **author/committer** to `Cursor Agent <cursoragent@cursor.com>` and append **`Co-authored-by:`** trailers. That shows up as agent attribution on GitHub.

**Policy:** commits from Cloud Agents on this repo should look like normal maintainer commits (no Cursor author, no `Co-authored-by` trailers).

**How:** point git at an **empty hooks directory** so the agent hooks do not run, and set author/committer explicitly. The directory must exist and contain **no** executable hook scripts.

```bash
mkdir -p /tmp/noagenthooks

export GIT_AUTHOR_NAME="Henrik"
export GIT_AUTHOR_EMAIL="henedl@users.noreply.github.com"
export GIT_COMMITTER_NAME="Henrik"
export GIT_COMMITTER_EMAIL="henedl@users.noreply.github.com"

git -c core.hooksPath=/tmp/noagenthooks commit ...
```

**Rewrite the last commit** (e.g. before `git push`) if hooks already ran:

```bash
mkdir -p /tmp/noagenthooks
MSG=$(git log -1 --format=%B | grep -vi '^co-authored-by:')
printf '%s\n' "$MSG" > /tmp/gitmsg.txt

GIT_AUTHOR_NAME=Henrik GIT_AUTHOR_EMAIL=henedl@users.noreply.github.com \
GIT_COMMITTER_NAME=Henrik GIT_COMMITTER_EMAIL=henedl@users.noreply.github.com \
  git -c core.hooksPath=/tmp/noagenthooks commit --amend --no-gpg-sign \
  --author="Henrik <henedl@users.noreply.github.com>" -F /tmp/gitmsg.txt
```

(Adjust **name/email** if the maintainer changes.)

### Pull requests — do not use `open_git_pr`

The Cursor Automation `open_git_pr` MCP tool (and `gh` in this VM, which authenticates as the `cursor` GitHub App) **always opens PRs as `cursor[bot]`**. GitHub does not let you change PR author after the fact.

**Policy:** push the branch with maintainer-authored commits, then **stop**. Do **not** call `open_git_pr`. Leave PR creation to the maintainer.

**Maintainer opens the PR** from the compare URL (substitute the branch name):

`https://github.com/henedl/clipgen/compare/master...<branch>?expand=1`

If a `cursor[bot]` PR was opened by mistake, close it and open a fresh one from that URL while logged in as the maintainer. The commits on the branch are already correctly attributed if the hook bypass above was used.
