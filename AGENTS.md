# Project learnings for agents

## Learned User Preferences

- Prefer placing generic index/letter conversion utilities (e.g. index_to_letter, letter_to_index) in utils.py rather than in domain-specific modules like files.py.

## Learned Workspace Facts

- **Version bump:** When making substantive code changes (fixes or features), increment the patch (last number) of `VERSIONNUM` in [config.py](config.py); see [CLAUDE.md](CLAUDE.md) § Version.
