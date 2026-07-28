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
exec uv run clipgen.py -i "$HOME/Projects/clipgen/" -s "clipgen-test" --desktop --studio "$@"
