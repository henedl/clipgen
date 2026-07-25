# Workflows — Phase 3 plan (exports · canvas polish · resume · chaining)

> **Phase 3 complete — W1–W8 all shipped** (v0.14.15 → v0.14.22). Picks up the deliberately
> deferred items from `plans/WORKFLOWS-PHASE2-PLAN.md` and `plans/WORKFLOWS-UX-PHASE-2-5.md`
> across four themes: new nodes, canvas/authoring polish, run experience, and triggers/automation.
> Each workstream landed as one `feat(workflows)` commit with its own version bump.

## Status

| # | Workstream | Status | Landed |
|---|---|---|---|
| W1 | New nodes: `transcript_export` (md/srt/vtt), `data_export` (JSON/CSV over events/segments), compound AND/OR filter clauses, heatmap `output` image/gif/rolling_gif + `frames`/`window`, "Transcribe → Transcript + Data Export" recipe, viewer `export` attachment doc-cards | ✅ Done | 0.14.15 |
| W2 | Canvas navigation: two-finger-scroll pan + pinch / Ctrl-⌘-wheel zoom, Space-hold pan (incl. the shared `hotkeys.js` Space-keyup fix), snap-to-grid + neighbor alignment guides (`#wfSnapBtn`), select-all / arrow nudge / zoom-% reset, viewport-only saves off the undo stack (`scheduleViewportSave`) | ✅ Done | 0.14.16 |
| W3 | Dialogs: `openPromptDialog`/`openConfirmDialog` on `openBlockingModal`; stash name/rename prompts; confirm-before-delete for blueprints + stashes | ✅ Done | 0.14.17 |
| W4 | Sticky notes: `note` pseudo-node in `blueprint.nodes` (rides save/undo/copy/import), runner filter (`NOTE_NODE_TYPE`), toolbar button + `N` hotkey, validation/Clean-up exemptions | ✅ Done | 0.14.18 |
| W5 | Resume-from-failure: sidecars widened to all JSON-safe ports + `__type__` (inspector filters at read time), seed-before-skip runner fix, pure `compute_resume_plan` (descendant closure, clipRecords + heatmap `raw_results` rules), `POST /api/runs` `resumeFromRunId`, Resume button on failed/cancelled runs | ✅ Done | 0.14.19 |
| W6 | Run history + dry-run: "This blueprint / All" scope toggle with cross-blueprint rows + click-through (`pendingFocusRunId`), hover-Run preview (`computeWouldRun`, would-run glow / skip dim / "N of M steps" chip) | ✅ Done | 0.14.20 |
| W7 | Trigger chaining: `transcript_complete` + `scan_event` trigger types via mtime-gated manifest polling in the watcher; schema `{"type", "enabled"}` (old `watch_dir` renamed `new_video`, no shim); single-armed **per type**; `TRIGGER_TYPES` via `/api/catalog`; toolbar type-picker menu; `triggerType` in run snapshots | ✅ Done | 0.14.21 |
| W8 | Parallel batches: opt-in `config.WORKFLOWS_BATCH_WORKERS` (default 1, clamp 4) pooling `_run_batch_child`; per-child deep copy of batch seeds (latent shared-mutation fix regardless of workers) | ✅ Done | 0.14.22 |

## Still open (demand-driven)

- Gallery-viewer node; compression output param (backends exist).
- foreach/subgraph iteration; compound predicate widget beyond two clauses.
- Per-node cancel/retry (resume covers the retry half; mid-run per-node cancel would need runner
  surgery).
- A general `TriggerMonitor` UI (armed-trigger history beyond the toolbar hint + run badges).
- Parallel/queued *batches* (only children within one batch parallelize).

## Verification notes

- Full pre-commit pipeline (`ruff format` + `ruff check` + `ty` + the whole pytest suite) ran green
  per workstream; new coverage lives in `tests/test_workflows_{executors,runner,collection_ops,api,
  frontend_source}.py`.
- ⚠️ Browser-facing work (W2–W4, W6, the W7 picker) is source-asserted but not browser-verified —
  per the no-headless-browser rule, needs a human pass: pan/zoom/snap feel, dialogs, notes,
  history click-through, dry-run hover, trigger menu.
