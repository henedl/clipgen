# Code review checklist

Patterns distilled from recurring post-review and post-merge fixes across the project's history.

## Frontend (JS/CSS)

- **CSS toggle completeness**: Every JS class toggle (`.hidden`, `.active`, `.disabled`, etc.) must have a corresponding CSS rule. Verify the rule exists in the stylesheet, not just the JS call.
- **Falsy-safe DOM helpers**: Use `!== undefined` or `!= null` instead of `if (x)` when `0`, `""`, or `false` are valid values.
- **Event listener cleanup**: When rebuilding UI or re-initializing components (e.g. color pickers, modals), remove previous `document`-level listeners before adding new ones. Store references for cleanup.
- **The DOM is not the state store**: Any *user-editable* value must live in `state` and be written there on input; a render function may only read state. If a value can only be recovered by reading it back off an input element, every future mutation path (add / remove / reorder / import / toggle) is a new way to lose it — this was re-fixed four times in Screenspace Multitool alone (`ceb0ed5c`, `10a9e348`, `256eda28`, `12d32e8f`). `snapshotMultitoolStepValues()` is the reference pattern.
- **UI state after DOM rebuilds**: If a function rebuilds a container's `innerHTML`, re-apply transient *presentation* state (filmstrip mode, toggle states, scroll position) after the rebuild. Apply filters before restoring `scrollTop` or the browser clamps it (`bb880948`).
- **Async race conditions**: For video seeking, image loading, or any async chain that can be re-triggered before completion, use a generation counter to reject stale callbacks. Coalesce rapid-fire requests with `requestAnimationFrame`. Key the guard on the thing that changed — a completion dedup keyed by participant while old tasks linger needs to be keyed by task id instead (`d9e9ebc8`).
- **Teardown symmetry**: Every listener / timer / `ResizeObserver` / rAF registered at page or document scope needs a `pagehide` teardown, and every `activate` needs a real `deactivate` body — an empty `mapDeactivate` left a rAF sweep animating a hidden canvas (`2a387fd6`). `visibilitychange` must pause *all* pollers, not just the ones you remember (`93088113` paused two of five).
- **Tooltips — one convention**: Use the `[data-tooltip]` singleton (`clipgenInitDataTooltips` in `utils.js`). Never put both `title` and `data-tooltip` on one element (`7e340085`). Never hand-write a key hint in the label — the hotkey registry appends it from `data-hotkey` (`ed9323c9`). Native `title` does not render on `draggable="true"` rows (`0e21c928`). Guarded by `tests/test_tooltip_conventions.py`.
- **CSS specificity and source order**: `tokens.css` loads first, so a page rule of equal specificity wins on source order — that is how `.btn-primary` lost its white text when the base moved into tokens (`f643bd1b`). Never redefine a shared token in page CSS; the viewer forked 16 tokens into an old palette and never rendered with the shared dark theme (`ecfed991`). Theme fallbacks in JS branch on `data-theme` rather than hardcoding the light values (`c895bd2d`, `cdf68a9f`). A new `z-index` must account for stacking contexts an ancestor creates (`74b6e4c7`).
- **Canvas/rendering performance**: RAF-throttle canvas draws and mouse-tracking renders. Cache `getBoundingClientRect()` results instead of calling in loops. Pause polling when the tab is hidden (`document.hidden`).
- **Flex layout**: Elements inside flex containers need explicit `flex: 1` or `min-width: 0` to avoid zero-width collapse. Verify new elements are visible after adding them to flex parents.
- **Autocomplete off on text inputs**: Every `<input type="text">` (static or dynamic) must have `autocomplete="off"` to prevent browser autofill (e.g. contact names). For static HTML use the attribute directly; for JS-created inputs set `.autocomplete = "off"` after creation.
- **Polling + render gates**: A control gated on poll-driven state must re-render when *any* input to the gate changes. Watch both the immediately-rendered source (e.g. `state.summaryText`) **and** the poll-lagged source (e.g. `state.participants[].agents.summary`). Reading only the lagged one leaves the control stuck until reload (the friction Re-run gate, `5683a96`).

## Backend (Python / Flask)

- **Route parameter types**: Prefer string route parameters with manual `float()`/`int()` parsing over Flask's `<float:x>` converter. JS may send integers where Flask expects floats, causing silent 404s.
- **JSON serialization safety**: Filter `math.isfinite()` on any float derived from OpenCV or numpy before including in JSON responses. Non-finite floats produce invalid JSON that `JSON.parse` silently drops.
- **numpy/ndarray in JSON**: Exclude numpy arrays and other non-serializable objects from manifest saves and API responses. Convert to lists or omit.
- **Dependency manifests**: When importing a new package, immediately add it to `pyproject.toml`. Missing dependencies surface as silent task failures.
- **All call sites**: When modifying a shared function's signature *or its semantics*, grep for every call site, not just the one you're working on. Functions like `finalize_timeline_data()` have 5+ callers across CLI, Studio, Viewer, and Screenspace. Narrowing `get_num_participants`'s scan range on a layout assumption broke every spreadsheet load one day later (`d6eefbd9`) and then browse mode (`8246f1ce`).

## Concurrency and shared state (Python / Flask)

The most-repeated backend class in the project's history — swept three separate times and still
recurring. Every route that touches `_manifest_lock` (or any module-level mutable) is in scope.

- **Lock the whole read → check → copy → mutate → persist sequence**, not just the dict access. Snapshot with `list(...)` *inside* the lock, then release and do the serialization / regex / validation work against the snapshot. Reading an entry under the lock and then iterating `entry["segments"]` outside it produces torn output or `list changed size during iteration` (`280b1f0d`, `186ff45a`, `55e2da5b`, `9ea2a6e3`).
- **Check-then-act is a bug.** In-flight / busy gates must be an atomic check-and-set under one lock (`1b444d0d`). Cleanup in a `finally` must be gated on the identity of the cancel event the run started with, or a stop-then-restart lets the dead run's `finally` clobber its successor's slot (`863edf8f`).
- **Never persist a module-level list by reference.** Worker threads extend `_generated_artifacts` mid-serialize; take a `list(...)` snapshot (`186ff45a`).
- **Heavy I/O goes outside the lock**: capture the inputs under the lock, release, generate, then briefly reacquire to attach the results — and fire the change notification *after* attaching, or a poll fingerprint computed before the attach never re-renders (`5683a96b`).
- **One in-flight slot per shared cancel event.** A second request gets 409 rather than clobbering the first's event (`1b444d0d`).
- **Cancellation contract**: the cancel flag must stay reachable after the task leaves its registry — popping the task first made the flag lookup return `None` and the scan ran to completion pinning a CPU (`7d10862b`). A client disconnect is a cancel signal, not a kill (`e5977422`, `b76bc16b`). Re-check the cancel event *inside* the lock before writing a result, or a Stop that races the model leaves a stale one behind (`280b1f0d`).

## Caches

- **One cache = one value shape.** Two producers writing different shapes under the same key revoke each other's blob URLs (`ea7d04c7`: a 320×180 poster and an N-frame filmstrip both keyed by artifact id).
- **State the key, the invalidation trigger, and the bound.** Path-keyed video caches use `(path, mtime_ns)` — the pattern already existed in `viewer.py` / `pipeline.py` and simply wasn't applied (`cfc54e46`). Never key on `id()`; CPython reuses addresses after GC (`280b1f0d`).
- **Bound it or explain it**: an `OrderedDict` with a sibling `_*_CACHE_MAX`, a `MediaCache(max)`, or a naturally bounded key (per-participant). Guarded by `tests/test_resource_lifecycle.py`.
- **Every `URL.createObjectURL` needs a revoke site** — on replacement *and* on `pagehide` (`7c8e751b`, `c895bd2d`). Guarded (replacement only) by `tests/test_resource_lifecycle.py`.
- **Purge per-process temp caches in every build path**, not just the one you're editing — the endcard cache was purged in two of three (`5683a96b`).
- **Single-flight expensive misses.** Releasing the lock before the ffmpeg miss invites a thundering herd of duplicate work (`8793142b`).

## Failure paths

Wrong output with no error is the most user-damaging class in this codebase's history.

- **A partial-failure path must not return the success sentinel.** Reel regeneration warned on missing components, concatenated the survivors and returned `True`, silently replacing a correct reel with a shorter one (`d2b4fa8e`). A failed card concat silently kept the unwrapped clip (`52194ebe`). Multitool full-frame steps silently fell back to the parent region (`ec8fcd60`). Prefer failing loudly over warn-and-continue whenever the output would be quietly wrong.
- **A parse/probe helper that returns `None` on partial success reads as "skip"** to its caller, and the work vanishes without an error (`6bbc5b2c`: one `strptime` format applied to both ends of a range).
- **Narrow `except Exception`** to the errors actually expected so genuine bugs surface (`1b444d0d`).
- **A partial update must not overwrite full state.** A PUT carrying a subset of keys merges onto a full snapshot (`5683a96b` rewrote the settings file from the submitted subset and dropped everything else); a pause/resume branch prepends prior results rather than replacing them (`8eb42a1f` permanently dropped pre-pause detections).
- **Sanitize at the boundary, not at the site that got reported.** Non-finite floats / ndarrays reaching JSON is a written rule above and still needed the same guard at six different boundaries (`0335b652` → `dc8d26d7` → `515e21ec` → `315dde16` → `d732a710` → `8eb42a1f`).

## Type checking (ty)

`ty` is a blocking CI gate. These rules prevent the most common typecheck failures.

- **Narrow Optional before use**: When a variable can be `None` (e.g. `cap: Optional[cv2.VideoCapture]`, `proc.stdout`, a lookup return), add `assert x is not None` before the first use, with a comment explaining the invariant (e.g. `# guaranteed by stdout=PIPE`). Do not use `# type: ignore` instead.
- **JSON dicts need `cast`**: Iterating over dicts from JSON, `isinstance(item, dict)` narrows to `dict[Unknown, Unknown]`, not `Dict[str, Any]`. After the isinstance guard, use `cast(Dict[str, Any], item)`. Annotate the source list explicitly: `steps: list[dict[str, Any]] = data.get("steps", [])`.
- **Avoid None-initialized result lists**: `[None] * n` forces `List[Optional[T]]` and requires narrowing at every use site. Define a typed empty sentinel (e.g. `_EMPTY: T = (0, [])`) and pre-fill with that.
- **cv2 output parameters**: cv2 type stubs reject `None` for output-array parameters (e.g. `calcOpticalFlowFarneback`). Pass a pre-allocated `np.zeros(...)` array instead.
- **Hoist annotations above branches**: Annotating a variable inside one branch of an if/else does not carry to the other. Declare the annotation before the if (`region: Dict[str, Any]`), then assign in each branch.
- **`list[T] | None` vs `list[T | None]`**: For optional list parameters, write `details: list[str] | None = None`. The form `list[str | None] = None` declares a non-optional list of nullable elements, a different type.
- **Narrow properly, don't suppress**: Replace `# type: ignore[union-attr]` and similar with proper narrowing (`assert`, `isinstance`, `if is not None`). Suppressions hide real bugs.

## Integration

- **Data contract completeness**: When creating records consumed by the frontend (artifacts, events, tasks), include all fields the renderer expects, even optional ones. Missing fields cause empty/broken cards.
- **New flags in mode detection**: When adding a CLI flag, verify it appears in the mode-detection logic (`cli.py`), not just in argparse definition. See [agents/skills/new-mode/SKILL.md](skills/new-mode/SKILL.md) for the full checklist.
- **Bundled/frozen paths**: Use `utils.get_bundled_assets_root()` for asset resolution, never raw `Path(__file__).parent`. Test that asset paths resolve in both source and PyInstaller environments.
- **No duplicated constants between Python and JS.** Any value that lives in `config.py` (or a Python helper) and that the frontend also needs — severity labels, default clip duration, annotation keyphrases (`!key`), ignored timestamp tokens (`x`) — must flow through `utils.get_frontend_config()`, not be hardcoded in JS. Procedure for adding a new mirrored constant: [agents/skills/sync-constants/SKILL.md](skills/sync-constants/SKILL.md).
- **Parallel registries.** That hard rule is scoped to Python↔JS; most drift has been Python↔Python or JS↔JS. Adding a detector / tool / format means updating *every* parallel list, not just the catalog: filter chips, quiet-poll paths, hue and confidence maps, edit-restore branches, dropdown labels (`e1019f8f`, `ef3dfbde`, `178e7bf8`, `7116973d`). The per-subsystem lists are enumerated in [agents/skills/new-screenspace-tool/SKILL.md](skills/new-screenspace-tool/SKILL.md) and [agents/skills/new-mode/SKILL.md](skills/new-mode/SKILL.md).
- **One computation, one implementation.** A preview or overlay must call the same primitive the real pipeline calls, never reimplement the math — a preview that diverges reads to the user as "the tool is wrong", not "the tool crashed" (`30192acd`, `684fd235`, `111f7747`, `d96c77a1`).

## Refactors (carve / split)

The two most error-prone refactors in this repo each have a dedicated skill and an automated guard. Apply these whenever a diff moves code between modules.

- **JS hub→satellite carve completeness**: Each page script is its own IIFE scope, so a bare cross-file function call throws `ReferenceError` at runtime (invisible to `node --check`), shipped 3× (`e4f67b2`, `8c7f347`). Every hub-called function defined in a satellite needs a same-named hub delegator (or a late-bound `SS.fn(...)` call); every moved `var` must route through `state.`/the namespace, never a bare cross-file read. Guarded by `tests/test_frontend_satellite_wiring.py`. Procedure: [agents/skills/carve-satellite/SKILL.md](skills/carve-satellite/SKILL.md).
- **Python god-file split**: New modules must be listed in `pyproject.toml [tool.setuptools] py-modules` (guarded by `tests/test_packaging.py`); the facade must re-export every public **and test-touched private** name; and test `mock.patch` targets must point at the owning sibling, not the facade (re-export only rebinds, `5683a96`). Procedure: [agents/skills/split-module/SKILL.md](skills/split-module/SKILL.md).
