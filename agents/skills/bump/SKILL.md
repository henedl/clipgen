# clipgen-bump — Version bumping is automated

Do not bump the version manually. The version lives in [build/VERSION](../../../build/VERSION) and is auto-bumped on merge to `master` by [.github/workflows/version-bump.yml](../../../.github/workflows/version-bump.yml) based on the PR title.

## How to trigger a bump

- Title the PR with `feat: …` or `fix: …` (with optional scope, e.g. `feat(viewer): …`) → patch bump.
- `docs: …`, `chore: …`, `refactor: …`, `test: …`, `build: …`, `ci: …`, or untyped titles → no bump.

## What not to do

- Do not edit `build/VERSION` by hand.
- Do not edit version strings in `CLAUDE.md`, `README.md`, or any docs — they reference `build/VERSION` or `utils.get_version()` by convention, not by literal value.
- Major and minor bumps are still made manually by the human (edit `build/VERSION` directly in a one-off PR).
