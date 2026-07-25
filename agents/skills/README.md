# Agent skills

Skill procedures for development workflows and CLI usage. Each skill is also available as a slash command (e.g. `/check`, `/generate`) via the `.claude/skills/` symlink.

## Development

- [check](check/SKILL.md) — full pre-commit pipeline (ruff → ty → tests)
- [test](test/SKILL.md) — run the test suite
- [bump](bump/SKILL.md) — increment patch version
- [new-mode](new-mode/SKILL.md) — checklist for adding a CLI mode or flag
- [new-screenspace-tool](new-screenspace-tool/SKILL.md) — checklist for adding a Screenspace tool
- [new-thinking-agent](new-thinking-agent/SKILL.md) — checklist for adding an Ollama thinking agent
- [sync-constants](sync-constants/SKILL.md) — audit Python ↔ JS constant mirroring
- [carve-satellite](carve-satellite/SKILL.md) — carve a JS hub into hub + satellite without ReferenceErrors
- [split-module](split-module/SKILL.md) — split a Python god-file into facade + siblings

## Using clipgen

- [generate](generate/SKILL.md) — translate intent into a clipgen CLI command
- [screenspace](screenspace/SKILL.md) — headless Screenspace analysis workflow
- [transcribe](transcribe/SKILL.md) — transcription and thinking agent workflow
- [debug](debug/SKILL.md) — diagnostic checklist for common issues
