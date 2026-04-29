#!/usr/bin/env sh
# Install git hooks that run lint checks before commit.
# Idempotent — re-run any time. Hooks live in .git/hooks/ and are not
# tracked, so this is the only way to keep them in sync across machines.
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
HOOKS_DIR="$(git -C "$REPO_ROOT" rev-parse --git-common-dir)/hooks"

mkdir -p "$HOOKS_DIR"

cat > "$HOOKS_DIR/pre-commit" <<'HOOK'
#!/usr/bin/env sh
# Installed by scripts/install-git-hooks.sh. Run lint on staged JS files.
set -e

staged_js=$(git diff --cached --name-only --diff-filter=ACM | grep -E '^assets/web/.*\.js$' || true)
if [ -z "$staged_js" ]; then
  exit 0
fi

if command -v bun >/dev/null 2>&1; then
  bun run lint:js
elif command -v npx >/dev/null 2>&1; then
  npx eslint assets/web --ext .js
else
  echo "pre-commit: skipping JS lint (neither bun nor npx found)" >&2
fi
HOOK

chmod +x "$HOOKS_DIR/pre-commit"
echo "Installed pre-commit hook at $HOOKS_DIR/pre-commit"
