# clipgen-bump — Bump the version on feat: PRs

The version lives in [build/VERSION](../../../build/VERSION). It is **not** auto-bumped by CI. When you ship a `feat:` PR, you (the agent) are responsible for bumping the patch number as part of the same PR.

## When to bump

- `feat: …` (or `feat(scope): …`, `feat!:`) → bump the patch number.
- `fix: …`, `docs: …`, `chore: …`, `refactor: …`, `test: …`, `build: …`, `ci: …`, or untyped titles → **do not** bump.

The human may also bump `build/VERSION` manually at any time (any segment — patch, minor, or major). Do not undo or "normalize" a manual bump.

## How to bump

1. Read `build/VERSION` (e.g. `0.10.123`).
2. Increment the patch segment by 1 (`0.10.123` → `0.10.124`).
3. Write the new value back to `build/VERSION` (single line, trailing newline).
4. Stage `build/VERSION` alongside the feature changes and include it in the PR.

## What not to do

- Do not bump on `fix:` / `refactor:` / `docs:` / `chore:` / `test:` / `build:` / `ci:` PRs.
- Do not edit version strings in `CLAUDE.md`, `README.md`, or other docs — they reference `build/VERSION` or `utils.get_version()` by convention.
- Do not introduce a new CI workflow to do this for you.
