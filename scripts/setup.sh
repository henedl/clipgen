#!/usr/bin/env sh
# Sync the current worktree's Python venv via uv (Conductor `scripts.setup`).
#
# Nothing here downloads packages in the normal case: uv materializes .venv from
# the shared unpacked-wheel cache (~/.cache/uv/archive-v0) and the shared
# uv-managed interpreter, cloning files copy-on-write on APFS. A fresh worktree
# therefore costs cache reads, not network.
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

cd "$REPO_ROOT"

case "$(uname -s)" in
Darwin)
  # --frozen: install straight from uv.lock with no resolution pass, and fail
  #   loudly rather than silently rewriting the lock if pyproject.toml drifted.
  # --extra dev: front-load pytest into setup, where Conductor shows progress,
  #   instead of the first /check. `uv run` is inexact by default, so a later
  #   plain `uv run clipgen.py` will not strip these back out.
  uv sync --frozen --extra dev
  ;;
*)
  # Linux (Conductor cloud). uv.lock pins GPU-capable torch, so a plain
  # `uv sync` here would pull ~2.5 GB of CUDA wheels that nothing uses. Use the
  # cpu-only install documented in agents/CLOUD.md instead.
  uv venv
  uv pip install ".[dev]" --torch-backend cpu
  ;;
esac

echo "Venv ready at $REPO_ROOT/.venv"
