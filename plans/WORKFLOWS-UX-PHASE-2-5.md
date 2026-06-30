# Workflows frontend UX — Phases 2 & 5 (remaining work)

## Context

Part of a five-phase UX improvement pass on the Workflows canvas (`--workflows`, `/workflows/`),
which under-invested in discoverability/feedback/polish relative to the Screenspace/Studio pages.
**Phases 1, 3, 4 are already shipped** on branch `henedl/workflows-ux-improvements`:

- **Phase 1** (`fc04194`, v0.13.43) — undo/redo buttons, save-status, shortcuts legend, focus rings.
- **Phase 4** (`430acdf`, v0.13.44) — focus rings everywhere, load spinner, card hover-lift, palette
  no-results, wide cards (>3 params), port-type legend.
- **Phase 3** (`14b2831`, v0.13.45) — run-panel status icons, timestamps, status filter, expandable
  results, reconnect pill.

This file covers the **two remaining phases**. Both are independent and independently shippable as
their own `feat:` PRs (bump the patch in `build/VERSION` once each). The full original roadmap lives
at `~/.claude/plans/system-instruction-you-are-working-scalable-emerson.md`.

**Conventions** (hard requirements):
- Hub + satellite carve: `workflows.js` hub + `workflows-{nodes,canvas,wires,runs,stashes,validate}.js`,
  shared state via `window.ClipgenWorkflows` (WF). Any hub-called function defined in a satellite needs
  a same-named guarded delegator or a late-bound `WF.fn(...)` call site; any new shared mutable state
  routes through `WF.state`, never a bare cross-file `var`. Guarded by
  `tests/test_frontend_satellite_wiring.py`.
- Vanilla ES5-style JS (`.then()`, no async/await), hand-written CSS using `tokens.css` design tokens,
  Heroicons via `mask-image`. No frameworks/build tools.
- **Do not browser-test by installing Chromium.** Ask the human to verify, or hand them a DevTools
  snippet.
- Run `/check` (ruff format + lint, ty, full suite) and `node --check` on each edited `.js` before
  committing.

---

## Phase 2 — Canvas navigation *(high impact, medium effort)* — ✅ DONE (v0.13.48)

Shipped fit-to-view (button + `F`), auto-pan on node drag near the edge (world-anchored drag so the
node tracks the cursor across the pan), **and** the minimap (2c was built, not deferred — corner
overview + viewport rectangle + click/drag-to-recenter).

Large graphs are hard to navigate: there's no way to frame the whole graph, no auto-pan when dragging
a node off-screen, and no minimap. All work is in **`assets/web/workflows-canvas.js`** plus small
toolbar (HTML/CSS) and shortcuts-legend additions.

### 2a. Fit-to-view button + `f` shortcut

Add `fitToView()` to the canvas satellite. It frames all `state.nodes` in the viewport:

- Compute the world-space bounding box of all node cards. Each card's pixel size is read the same way
  `focusNode()` already does (`card.offsetWidth/offsetHeight`, lines ~535-553); positions are
  `node.position.x/y`. Fall back to a default size (~200×120) for any node without a rendered card yet.
- Derive a zoom that fits the box into the canvas rect with padding (~40px), `clamp`ed to
  `ZOOM_MIN`/`ZOOM_MAX` (already defined, lines 16-17), then center it:
  `vp.x = rect.width/2 - boxCenterX*zoom` (mirrors `focusNode`'s centering math, lines 550-551).
- Call `applyViewport()` then `WF.scheduleSave()` (viewport is persisted per blueprint).
- No-op when `!state.ready` or `state.nodes` is empty (match `autoArrange`'s guards, lines 477-479).

Wiring:
- Publish `WF.fitToView = fitToView;` (bottom of the satellite, beside `WF.autoArrange`).
- Toolbar button in `workflows.html` next to `#wfCleanUp`: `<button id="wfFitView" class="btn btn-small
  wf-icon-btn">` with a `<span class="wf-btn-icon wf-fit-icon">` using `icons/arrows-pointing-out.svg`
  (mirror the `.wf-undo-icon` pattern already in `workflows.css`). Add `#wfFitView` to the
  `setToolbarDisabled` list in `workflows.js` and wire its click in `boot()` to `WF.fitToView()`.
- Keyboard: in `onKeyDown` (line ~406), add an `f` case **guarded by `!inField`** and no modifier:
  `if (e.key === "f" || e.key === "F") { if (WF.fitToView) WF.fitToView(); e.preventDefault(); }`.
  Place it after the clipboard/undo block, before the Delete handling.
- Add a "Fit to view — `F`" row to the shortcuts legend (`#wfShortcutsMenu` in `workflows.html`).

### 2b. Auto-pan while dragging a node near the viewport edge

In `startNodeDrag`'s `move(ev)` handler (lines ~239-252): when the cursor is within an edge band
(~40px) of the canvas rect, add a small per-frame nudge to `state.viewport.x/y` (away from that edge)
and re-`applyViewport()`. Reuse the existing RAF tick (`scheduleNodePositionFlush`) so it stays
frame-bounded — do not spawn a second loop. Cache the canvas `getBoundingClientRect()` once at drag
start (per the canvas-perf code-review rule). Keep the nudge small (a few px/frame) so it's
controllable; the dragged node's position already tracks the cursor in world space, so it keeps moving
with the pan. `scheduleSave()` already fires on mouseup — no extra save needed.

### 2c. Minimap *(optional — defer unless 2a/2b prove insufficient)*

A small fixed-corner `<canvas>` drawing node rects (scaled bounding box) + a viewport rectangle, with
click/drag-to-recenter. Higher effort (a second render surface kept in sync with pan/zoom/node-move).
**Recommend deferring**: ship 2a+2b first, only build this if real graphs are large enough to need it.
If skipped, say so explicitly in the PR so it's not mistaken for done.

### Phase 2 verification
- `node --check assets/web/workflows-canvas.js`; `/check` green.
- Browser (human): with several nodes spread out, the Fit button (and `F`) frames them all centered;
  dragging a node to the canvas edge scrolls the view; `F` does nothing while typing in a param field.

---

## Phase 5 — Multi-select participant batch *(medium impact, medium effort)* — ✅ DONE (v0.13.49)

Shipped: a checkbox-popover multi-select replaces the single `<select>` on the Video Source
participant param (`buildParticipantSelect` in `workflows-nodes.js`, reusing the hub's `bindMenuToggle`).
Value is normalized on write — a single id stays a string (server single-run path untouched), `__all__`
for everyone, an array for a subset, `[]` for none. `blueprintWantsBatch`/`startBatch` send the subset
as `participants`; an empty array raises a validation warning. Server already honored the subset; added
a `test_batch_honors_participant_subset` regression guard.

Today a Video Source's participant param is a single `<select>` offering each participant **or** the
`__all__` sentinel (`WF.ALL_PARTICIPANTS`); "All" fans out to a batch over every participant. Goal:
let the user pick an **arbitrary subset** to fan out over, keeping "All" as a shortcut. The original
authors flagged this as the intended follow-up (`plans/WORKFLOWS-PHASE2-PLAN.md` lines 128-131).

### Backend — already done, just confirm

`POST /api/batches` (`workflows_server.py:837`) already reads an optional `participants` array
(lines 865-869): `requested = data.get("participants")`; if truthy it intersects with the
with-video `available` list, else uses all. No server change needed — the batch rebinds every
`video_source` node to each participant per child run. **First step: re-read that block to confirm it
still matches before wiring the UI.**

### Frontend

**Value shape.** Let `node.params.participant` hold one of: a single id string (single run — unchanged),
the `__all__` sentinel (batch over all — unchanged), or an **array of ids** (batch over the subset).
Per the repo's no-back-compat rule, just change the shape; no migration shim. A single-element array
should behave like a single run (see `blueprintWantsBatch` below).

1. **Widget — `workflows-nodes.js`, the `participant` branch of `buildParamControl`** (lines ~179-210).
   Replace the bare `<select>` with a compact multi-select. Lightest fit that matches the
   "no `window.prompt`" direction: a small button showing the current selection summary
   ("3 participants" / "All" / "P01") that opens a popover of checkboxes (one per discovered
   participant, from `state.context.participants`) plus an "All participants" shortcut. Reuse the
   `bindMenuToggle()` open/close pattern (now in `workflows.js`; either publish it as `WF.bindMenuToggle`
   or replicate the small handler locally). On change, write the array (or `__all__`) back to
   `node.params[spec.name]` and call `WF.scheduleSave()`. Keep `autocomplete="off"` on any text input
   (none needed here). Preserve focus across the card's re-render the way the enum branch does.

2. **Batch trigger — `workflows-runs.js` `blueprintWantsBatch()`** (lines ~34-46). Currently returns
   true only when a `video_source` param equals `WF.ALL_PARTICIPANTS`. Extend: also true when the value
   is `__all__` **or** an array with length ≥ 2. An array of length 1 → single run (bind that one id).

3. **Send the subset — `workflows-runs.js` `startBatch()`** (lines ~242-266). Collect the participant
   selection and POST it: add `participants` to the `api/batches` body when the triggering
   `video_source` carries an explicit array (omit the field for `__all__` so the server's "all" branch
   runs). Resolve the list from the first `video_source` whose value is an array/`__all__` (document
   that multiple Video Sources share one batch list — the server rebinds them all per child; a
   per-source subset is out of scope).

4. **Single-run path.** When `blueprintWantsBatch()` is false and the value is a 1-element array,
   ensure the single-run launch binds that id (the server's single-run path expects a scalar
   participant on the node — coerce the 1-array to its element before/at POST, or normalize on write so
   a 1-element selection is stored as a string).

5. **Validation — `workflows-validate.js`.** Add a warning (not a blocking error) when a `video_source`
   carries an **empty** participant array (nothing to run). Follow the existing `nodeIssues` warning
   pattern.

### Phase 5 verification
- `node --check` on the three edited JS files; `/check` green (incl. `test_frontend_satellite_wiring.py`).
- If practical, a small `test_workflows_api.py` assertion that `POST /api/batches` honors a
  `participants` subset (the route already supports it — a regression guard, not new behavior).
- Browser (human): selecting 2+ participants on a Video Source and pressing Run fans out one child run
  each (only the chosen ones); "All" still runs everyone; a single selection runs once.

---

## Out of scope (intentionally descoped — do not build without a fresh ask)

Per `plans/WORKFLOWS-PHASE2-PLAN.md`: canvas **groups/comments**, **per-node cancel/retry**, cross-run
**memoization**, general **foreach/looping**, **sibling-node parallelism within a run**, and
**parallel/queued batch runs** (batches stay sequential).
