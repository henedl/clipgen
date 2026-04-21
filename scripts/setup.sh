#!/usr/bin/env sh
# Sync the current worktree's Python venv via uv.
# macOS only for now.
set -e

case "$(uname -s)" in
Darwin) ;;
*)
  echo "Unsupported OS: setup.sh is macOS-only for now." >&2
  exit 1
  ;;
esac

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

cd "$REPO_ROOT"
uv sync
echo "Venv ready at $REPO_ROOT/.venv"
