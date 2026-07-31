#!/usr/bin/env sh
# Run clipgen Studio against ~/Projects/clipgen with the 'clipgen-test' sheet.
# Opens the native desktop window; pass --browser to get a browser tab instead.
# macOS only for now.
set -e

case "$(uname -s)" in
Darwin) ;;
*)
  echo "Unsupported OS: run.sh is macOS-only for now." >&2
  exit 1
  ;;
esac

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

cd "$REPO_ROOT"

# Media comes from the root checkout (the multi-GB clipgen-test_P*.mp4 sources
# are shared, never copied per worktree); the workbook is the per-worktree copy
# Conductor drops in via .worktreeinclude, resolved against cwd.
MEDIA_DIR="${CONDUCTOR_ROOT_PATH:-$HOME/Projects/clipgen}"

exec uv run clipgen.py -i "$MEDIA_DIR" -s "clipgen-test" --desktop --studio "$@"
