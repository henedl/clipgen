# clipgen-new-screenspace-tool — Add a Screenspace analysis tool

Eight files need touching. Work through this checklist in order.

## Checklist

1. **Analysis engine** (`screenspace.py`)
   - Implement the tool function following the pattern of existing tools (color, change, similarity, etc.)
   - Register it in the tool dispatch dict

2. **REST API** (`screenspace_server.py`)
   - Add endpoint(s) for queuing and retrieving results for the new tool
   - Follow the existing endpoint naming pattern

3. **Design token** (`assets/web/tokens.css`)
   - Add `--color-task-{name}` to the Screenspace task color block
   - Read it in JS via `getComputedStyle(document.documentElement).getPropertyValue("--color-task-" + type)` — never hardcode hex

4. **Frontend UI** (`assets/web/screenspace.js`)
   - Add the tool to the tool selector
   - Add result rendering logic

5. **Timeline viewer** (`viewer.py`) — _skip for timelapse_
   - Add the task type to `SS_DETECTOR_COLORS` and `SS_DETECTOR_ICON_PATHS`
   - Timelapse produces a single output file and does NOT need entries here

6. **CLI flag** (`cli.py`)
   - Add `--ss-task {name}` as a valid choice to the screenspace task argparse argument

7. **CLI tests** (`tests/test_cli_screenspace_args.py`)
   - Add a test that `--ss-task {name}` is accepted with the expected parameters

8. **Unit tests** (`tests/test_screenspace.py`, `tests/test_screenspace_api.py`)
   - Add tests for the engine function and the API endpoint

## Notes

- Icon for the tool: pick a Heroicon from `assets/icons/` (kebab-case, e.g. `eye.svg`). Use the CSS `mask-image` pattern — see `XREF_BADGES` in `utils.js` for the canonical example.
- Tool-specific CLI parameters (e.g. `--ss-target-color`, `--ss-threshold`) go in the screenspace args group in `cli.py`.
