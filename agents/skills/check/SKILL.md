# clipgen-check — Full pre-commit check pipeline

Run this before every `git commit` to catch formatting, lint, type, and test failures.

## Steps (run in order, stop on first failure)

1. **Format check** — `uv run ruff format --check` on all modified `.py` files.
   - If any would be reformatted: run `uv run ruff format` on them, then re-check.
2. **Lint** — `uv run ruff check --fix`
3. **Type check** — `uv run ty check`
4. **Tests** — `uv run --extra dev pytest -c tests/pytest.ini`

Report pass/fail per stage. If all pass, remind the agent:

> If changes are substantive (bug fix or new feature), title the PR with `feat:` or `fix:` so [version-bump.yml](../../../.github/workflows/version-bump.yml) auto-bumps `build/VERSION` on merge. `docs:`, `chore:`, `refactor:`, `test:`, `build:`, `ci:`, or untyped titles do not bump.

## Common failures

- **Ruff format**: auto-fixed by running `uv run ruff format`. Most common cause of CI failures.
- **ty `union-attr`**: narrow the Optional before use with `assert x is not None` rather than `# type: ignore`.
- **ty JSON dicts**: after `isinstance(item, dict)`, use `cast(Dict[str, Any], item)`.
- **Tests**: run a single file to isolate — `uv run --extra dev pytest -c tests/pytest.ini tests/test_foo.py`.
