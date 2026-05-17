# clipgen-new-mode — Add a new CLI mode or flag

Missing mode-detection is the most common integration bug when adding new flags. Follow this checklist in order.

## Checklist

1. **Argparse definition** (`cli.py`)
   - Add the argument with help text and appropriate `action`/`type`/`nargs`
   - If mutually exclusive with other modes, add to the correct `add_mutually_exclusive_group()`

2. **Mode detection** (`cli.py`)
   - Find the mode-detection block (function or inline logic that maps args → mode string)
   - Add a branch for the new flag
   - **Critical**: if this step is skipped, the mode will never dispatch even though argparse accepts it

3. **Mode dispatch** (`cli.py` `main()`)
   - Add a dispatch case that calls the appropriate handler when the mode is detected

4. **Web UI (if applicable)** (`server.py`)
   - Register the blueprint: `app.register_blueprint(...)`
   - Add it to the mutual-exclusion check at startup (only one of `--studio`, `--screenspace`, `--transcripts` can be active)

5. **Tests**
   - Add at least one smoke test in `tests/test_cli_modes.py` (or a new test file for larger modes)
   - Test that the flag is accepted, dispatches correctly, and that incompatible flag combinations raise errors

6. **Version bump**
   - Increment patch in `build/VERSION` (see [agents/skills/bump/SKILL.md](../bump/SKILL.md))
