# Code review checklist

Patterns distilled from recurring post-review and post-merge fixes across the project's history.

## Frontend (JS/CSS)

- **CSS toggle completeness**: Every JS class toggle (`.hidden`, `.active`, `.disabled`, etc.) must have a corresponding CSS rule. Verify the rule exists in the stylesheet, not just the JS call.
- **Falsy-safe DOM helpers**: Use `!== undefined` or `!= null` instead of `if (x)` when `0`, `""`, or `false` are valid values.
- **Event listener cleanup**: When rebuilding UI or re-initializing components (e.g. color pickers, modals), remove previous `document`-level listeners before adding new ones. Store references for cleanup.
- **UI state after DOM rebuilds**: If a function rebuilds a container's `innerHTML`, re-apply transient UI state (filmstrip mode, toggle states, scroll position) after the rebuild.
- **Async race conditions**: For video seeking, image loading, or any async chain that can be re-triggered before completion, use a generation counter to reject stale callbacks. Coalesce rapid-fire requests with `requestAnimationFrame`.
- **Canvas/rendering performance**: RAF-throttle canvas draws and mouse-tracking renders. Cache `getBoundingClientRect()` results instead of calling in loops. Pause polling when the tab is hidden (`document.hidden`).
- **Flex layout**: Elements inside flex containers need explicit `flex: 1` or `min-width: 0` to avoid zero-width collapse. Verify new elements are visible after adding them to flex parents.
- **Autocomplete off on text inputs**: Every `<input type="text">` (static or dynamic) must have `autocomplete="off"` to prevent browser autofill (e.g. contact names). For static HTML use the attribute directly; for JS-created inputs set `.autocomplete = "off"` after creation.
- **Polling + render gates**: A control gated on poll-driven state must re-render when *any* input to the gate changes. Watch both the immediately-rendered source (e.g. `state.summaryText`) **and** the poll-lagged source (e.g. `state.participants[].agents.summary`) — reading only the lagged one leaves the control stuck until reload (the friction Re-run gate, `5683a96`).

## Backend (Python / Flask)

- **Route parameter types**: Prefer string route parameters with manual `float()`/`int()` parsing over Flask's `<float:x>` converter — JS may send integers where Flask expects floats, causing silent 404s.
- **JSON serialization safety**: Filter `math.isfinite()` on any float derived from OpenCV or numpy before including in JSON responses. Non-finite floats produce invalid JSON that `JSON.parse` silently drops.
- **numpy/ndarray in JSON**: Exclude numpy arrays and other non-serializable objects from manifest saves and API responses. Convert to lists or omit.
- **Dependency manifests**: When importing a new package, immediately add it to `pyproject.toml`. Missing dependencies surface as silent task failures.
- **All call sites**: When modifying a shared function's signature or adding a new parameter, grep for every call site — not just the one you're working on. Functions like `finalize_timeline_data()` have 5+ callers across CLI, Studio, Viewer, and Screenspace.

## Type checking (ty)

`ty` is a blocking CI gate. These rules prevent the most common typecheck failures.

- **Narrow Optional before use**: When a variable can be `None` (e.g. `cap: Optional[cv2.VideoCapture]`, `proc.stdout`, a lookup return), add `assert x is not None` before the first use — with a comment explaining the invariant (e.g. `# guaranteed by stdout=PIPE`). Do not use `# type: ignore` instead.
- **JSON dicts need `cast`**: Iterating over dicts from JSON, `isinstance(item, dict)` narrows to `dict[Unknown, Unknown]`, not `Dict[str, Any]`. After the isinstance guard, use `cast(Dict[str, Any], item)`. Annotate the source list explicitly: `steps: list[dict[str, Any]] = data.get("steps", [])`.
- **Avoid None-initialized result lists**: `[None] * n` forces `List[Optional[T]]` and requires narrowing at every use site. Define a typed empty sentinel (e.g. `_EMPTY: T = (0, [])`) and pre-fill with that.
- **cv2 output parameters**: cv2 type stubs reject `None` for output-array parameters (e.g. `calcOpticalFlowFarneback`). Pass a pre-allocated `np.zeros(...)` array instead.
- **Hoist annotations above branches**: Annotating a variable inside one branch of an if/else does not carry to the other. Declare the annotation before the if (`region: Dict[str, Any]`), then assign in each branch.
- **`list[T] | None` vs `list[T | None]`**: For optional list parameters, write `details: list[str] | None = None`. The form `list[str | None] = None` declares a non-optional list of nullable elements — a different type.
- **Narrow properly, don't suppress**: Replace `# type: ignore[union-attr]` and similar with proper narrowing (`assert`, `isinstance`, `if is not None`). Suppressions hide real bugs.

## Integration

- **Data contract completeness**: When creating records consumed by the frontend (artifacts, events, tasks), include all fields the renderer expects — even optional ones. Missing fields cause empty/broken cards.
- **New flags in mode detection**: When adding a CLI flag, verify it appears in the mode-detection logic (`cli.py`), not just in argparse definition. See [agents/skills/new-mode/SKILL.md](skills/new-mode/SKILL.md) for the full checklist.
- **Bundled/frozen paths**: Use `utils.get_bundled_assets_root()` for asset resolution, never raw `Path(__file__).parent`. Test that asset paths resolve in both source and PyInstaller environments.
- **No duplicated constants between Python and JS.** Any value that lives in `config.py` (or a Python helper) and that the frontend also needs — severity labels, default clip duration, annotation keyphrases (`!key`), ignored timestamp tokens (`x`) — must flow through `utils.get_frontend_config()`, not be hardcoded in JS. Procedure for adding a new mirrored constant: [agents/skills/sync-constants/SKILL.md](skills/sync-constants/SKILL.md).

## Refactors (carve / split)

The two most error-prone refactors in this repo each have a dedicated skill and an automated guard. Apply these whenever a diff moves code between modules.

- **JS hub→satellite carve completeness**: Each page script is its own IIFE scope, so a bare cross-file function call throws `ReferenceError` at runtime (invisible to `node --check`) — shipped 3× (`e4f67b2`, `8c7f347`). Every hub-called function defined in a satellite needs a same-named hub delegator (or a late-bound `SS.fn(...)` call); every moved `var` must route through `state.`/the namespace, never a bare cross-file read. Guarded by `tests/test_frontend_satellite_wiring.py`. Procedure: [agents/skills/carve-satellite/SKILL.md](skills/carve-satellite/SKILL.md).
- **Python god-file split**: New modules must be listed in `pyproject.toml [tool.setuptools] py-modules` (guarded by `tests/test_packaging.py`); the facade must re-export every public **and test-touched private** name; and test `mock.patch` targets must point at the owning sibling, not the facade (re-export only rebinds — `5683a96`). Procedure: [agents/skills/split-module/SKILL.md](skills/split-module/SKILL.md).
