# clipgen-new-screenspace-tool — Add a Screenspace analysis tool

Nine files need touching. Work through this checklist in order.

`screenspace.py` is a **re-export facade** — do not implement there. The scan goes in
`screenspace_scans.py`, the `AnalysisTool` subclass in `screenspace_tools.py`, and any new
public/test-touched name must be added to the facade's re-export block (see
[split-module/SKILL.md](../split-module/SKILL.md)).

## Checklist

1. **Analysis engine** (`screenspace_scans.py` + `screenspace_tools.py`)
   - Implement `scan_{name}` following the pattern of existing scans (color, change, similarity, etc.); scans never call each other
   - Add the `AnalysisTool` subclass and register it in the `TOOLS` registry; wire per-frame dispatch in `check_frame_for_tool` / `score_frame_for_tool`
   - Re-export the new public names from `screenspace.py`
   - **Reuse primitives, don't reimplement.** Static-frame skipping, template preparation, resolution scaling and the like already exist in `screenspace_primitives.py`. `scan_scene` grew its own mean-luminance static skip while every sibling used `_frame_is_static()`'s per-pixel absdiff, and silently skipped visually different frames (`d96c77a1`)

2. **REST API** (`screenspace_server.py`)
   - Add endpoint(s) for queuing and retrieving results for the new tool
   - Follow the existing endpoint naming pattern; use `server_utils`' `ok`/`err`/`json_endpoint`/`parse_number_arg` rather than a hand-rolled guard block
   - If the route is polled, add it to `_QUIET_POLL_PATHS` or it doubles the access-log volume (`7116973d`)

3. **Design token** (`assets/web/tokens.css`)
   - Add `--color-task-{name}` to the Screenspace task color block
   - Read it in JS via `getComputedStyle(document.documentElement).getPropertyValue("--color-task-" + type)`. Never hardcode hex (`tests/test_js_color_discipline.py` ratchets)

4. **Frontend UI** (`assets/web/screenspace*.js`)
   - Add the tool to the tool selector, and result rendering to `screenspace-results.js`
   - **Update every parallel registry, not just the selector.** A new tool typically needs entries in the confidence (`hasConf`) map, the category hue map, the task-Edit parameter-restore branch, and the fast-scan description whitelist. One PR's own self-review found three such omissions at once (`7d10862b`); the fast-scan marker leaked onto Boundary because the mode was stamped for every non-timelapse tool instead of consulting the whitelist (`178e7bf8`)
   - Any user-editable parameter must live in `state`, not only in the input element — see the "DOM is not the state store" rule in [CODE-REVIEW.md](../../CODE-REVIEW.md); this class was re-fixed four times in Multitool

5. **Model-view preview** (`screenspace_preview.py`), _if the tool gets one_
   - `build_preview` / `build_overlay_layer` must **call the same primitive the scan calls**, never re-derive the math. A preview that diverges from the scan reads to the user as "the tool is wrong", not "the preview is wrong", and has cost four separate fixes: frame-timestamp drift (`30192acd`, `4886888d`), Flow/Canny computed at a different resolution (`684fd235`), and a blurred-vs-binarized template mask (`111f7747`)

6. **Timeline viewer** (`viewer.py`), _skip for timelapse_
   - Add the task type to `SS_DETECTOR_COLORS` and `SS_DETECTOR_ICON_PATHS`
   - Timelapse produces a single output file and does NOT need entries here

7. **CLI flag** (`cli_args.py`)
   - Add `--ss-task {name}` as a valid choice to the screenspace task argparse argument

8. **CLI tests** (`tests/test_cli_screenspace_args.py`)
   - Add a test that `--ss-task {name}` is accepted with the expected parameters

9. **Unit tests** (`tests/test_screenspace.py`, `tests/test_screenspace_api.py`)
   - Add tests for the engine function and the API endpoint
   - `mock.patch` targets must name the **owning sibling module**, not the facade — re-exporting only rebinds (`5683a96`)

## Notes

- Icon for the tool: pick a Heroicon from `assets/icons/` (kebab-case, e.g. `eye.svg`). Use the CSS `mask-image` pattern. See `XREF_BADGES` in `utils.js` for the canonical example.
- Tool-specific CLI parameters (e.g. `--ss-target-color`, `--ss-threshold`) go in `_add_screenspace_args` in `cli_args.py`.
