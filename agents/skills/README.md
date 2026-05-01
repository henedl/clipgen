# Agent skills

Skill procedures for development workflows and CLI usage. Each skill is also available as a slash command (e.g. `/check`, `/generate`) via the `.claude/skills/` symlink.

## Development

- @agents/skills/check/SKILL.md — full pre-commit pipeline (ruff → ty → tests)
- @agents/skills/test/SKILL.md — run the test suite
- @agents/skills/bump/SKILL.md — increment patch version
- @agents/skills/new-mode/SKILL.md — checklist for adding a CLI mode or flag
- @agents/skills/new-screenspace-tool/SKILL.md — checklist for adding a Screenspace tool
- @agents/skills/new-thinking-agent/SKILL.md — checklist for adding an Ollama thinking agent
- @agents/skills/sync-constants/SKILL.md — audit Python ↔ JS constant mirroring

## Using clipgen

- @agents/skills/generate/SKILL.md — translate intent into a clipgen CLI command
- @agents/skills/screenspace/SKILL.md — headless Screenspace analysis workflow
- @agents/skills/transcribe/SKILL.md — transcription and thinking agent workflow
- @agents/skills/debug/SKILL.md — diagnostic checklist for common issues
