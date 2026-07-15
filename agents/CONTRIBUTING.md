# Pull requests

How we write PRs and commits, for agents and humans. Aim for **minimal, structured, scannable**.
The PR body lives in [.github/PULL_REQUEST_TEMPLATE.md](../.github/PULL_REQUEST_TEMPLATE.md); agents
mirror it when passing `--body` to `gh pr create`.

## Title

`type(scope): imperative description`

- **type** — one of `feat`, `fix`, `refactor`, `perf`, `docs`, `test`, `chore`, `ci`, `build`, `style`.
- **scope** (optional) — the subsystem: `screenspace`, `transcripts`, `studio`, `web`, `server`,
  `viewer`, `files`, … Lowercase and singular. Omit for cross-cutting changes.
- Imperative mood ("add", not "added"), no trailing period, ≤ ~70 chars.
- GitHub appends `(#NNN)` on squash-merge. Don't hand-add the PR number.

## Body

Two sections, each kept only if it earns its place:

- **Summary** — what changed and *why*, in 1–2 sentences. Add one bullet per part only when there
  are several distinct parts. Backtick paths, endpoints, and config keys.
- **Test plan** — `[x]` for automated checks you ran, `[ ]` for manual steps still needed. Be
  honest about what isn't verified (e.g. "⚠️ UI not browser-checked").

Cut anything that doesn't help a reviewer: no Motivation/Solution/Alternatives scaffolding, and
don't restate the diff in prose.

## Commits

Commit subjects use the same `type(scope): imperative description` rules as PR titles (above).
On squash-merge the PR title becomes the commit subject and GitHub appends `(#NNN)`. Add a body
for non-trivial changes (bullet the notable changes; flag regressions or gotchas); keep simple
fixes to a single line. The `type` also drives versioning: `feat:` bumps the patch in
[build/VERSION](../build/VERSION), others don't. See [skills/bump/SKILL.md](skills/bump/SKILL.md).

## Before opening

- `/check` is green: ruff format + lint, ty, full test suite. See [skills/check/SKILL.md](skills/check/SKILL.md).
- `feat:` PRs bump the patch in [build/VERSION](../build/VERSION). See [skills/bump/SKILL.md](skills/bump/SKILL.md).
- Self-review against [CODE-REVIEW.md](CODE-REVIEW.md).
