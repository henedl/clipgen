# clipgen-bump — Increment the patch version

## When to bump

**Do bump** for: bug fixes, new features, any user-visible behavior change.
**Do NOT bump** for: docs-only, comment-only, or refactor-only changes with no behavior change.

## How to bump

1. Read `VERSIONNUM` from `config.py` (e.g. `"0.9.4"`)
2. Increment the **last segment only**: `"0.9.4"` → `"0.9.5"`
3. Write the new value back to `config.py`
4. Print: `VERSIONNUM: 0.9.4 → 0.9.5`

## Rules

- Never edit version strings in `CLAUDE.md`, `README.md`, or any docs — they reference `config.py` by convention, not by literal value.
- Only bump the patch (third) segment. Major and minor bumps are made manually by the human.
