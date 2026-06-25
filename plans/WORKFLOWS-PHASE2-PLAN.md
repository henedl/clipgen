# Workflows — Phase 2 plan (recipes · batch · catalog · triggers)

> **Living doc — converged on posture, still iterating on detail.** v1 (`plans/WORKFLOWS-PLAN.md`,
> M0–M5) proves the cross-domain chain with a curated node set and a sequential, in-process run
> engine. Phase 2's shape is now set by a round of decisions (recorded below): a **recipe/template
> layer on top of a steadily-growing catalog**, **whole-study batch**, and a **narrow trigger**, with
> several heavier ideas deliberately cut. Detail (result layout, storage contract, trigger binding)
> is still open — see *Open questions*.

## Context

The v1 vertical slice answers "can a researcher chain transcribe → find-word → scan → clip across
domains, with live progress?" Phase 2 makes it a daily tool by optimizing for **"get my clips and
insights fast, across the whole study"** — not by becoming a general automation platform. The bias
throughout follows the repo's grain: small/ephemeral user base, "just re-run," thin server / thick
client, no persistence complexity, grow by demonstrated demand.

## Decisions settled (this round)

| Question | Decision |
|---|---|
| **Posture** | Recipe/template layer **on top of** a steadily-growing catalog. Defer broad automation; build the one concrete piece (triggers). |
| **Batch model** | **Whole-graph fan-out** — a launch-time "run for all participants" that runs the entire blueprint once per participant, results grouped per participant. Not a composable `foreach`. |
| **Memoization / partial re-run** | **Skip it.** Re-runs stay honest and simple (re-transcribe, re-scan). |
| **Triggers** | **Build, narrow-first** — one watch-dir watcher that auto-runs a designated blueprint when a new video lands in `-i`. Chaining + general signals are a later seam, not v2.0. |
| **Templates vs stashes** | **Unify** — a template *is* a seeded stash. Ship headline recipes as built-in stashes; one concept, one palette, one storage path. |
| **First catalog tranche** | Artifact formats (screenshots + GIFs + titlecards/endcards) · **highlights reel** · the 6 remaining Screenspace detectors + multitool/timelapse/heatmaps. Outputs/export bucket → demand-driven. |

## Depends on v1 finishing

Phase 2 builds directly on two pending v1 milestones — **M4 (run engine: `WorkflowRunner`, SSE,
run history)** and **M5 (stashes: save/instantiate sub-graphs)**. Templates need M5's stash storage;
batch and triggers need M4's runner. Land both before starting here.

## Gaps carried out of v1 (the starting state)

1. **Frontend `canConnect` ignores `ADAPTERS`.** `workflows-wires.js:151` is exact-match; the runner
   *does* coerce (`events→clipRecords`, `segments→timeRange`, `video→scalar`), so the UI rejects
   wires that would actually run. The most misleading gap → fixed first (P1).
2. **`ss_scan` exposes 4 of 10 detectors**; no multitool/timelapse/heatmap.
3. **`make_clips` emits `clip` only** despite `process_clips` supporting `screen`/`gif`; no highlights
   reel, titlecards, compression.
4. **One participant per run** — no whole-study batch.
5. **Thin run history**; large per-node results stripped with no fetch-on-demand path.
6. **No pre-run validation surfacing** — missing params / unmet context only fail at runtime.

---

## Workstreams

Rough order; P1 is the unblock, P2–P4 are the core experience, P5 is interleaved, P6 is last.

### P1 — Adapter ↔ UI parity *(small; do first)*
Serve the `ADAPTERS` keys through `GET /api/catalog` and widen `canConnect` to accept exact-match
**or** a registered adapter; render adapter-bridged wires with a distinct cue (dashed / "↳ coerced").
Unblocks every multi-type recipe.

### P2 — Catalog tranche *(incremental, append-only; each node = one `NodeType` + executor)*
Build only the demanded buckets; skills `new-screenspace-tool` / `new-mode` cover the mechanics.

| Bucket | Nodes | Underlying function(s) |
|---|---|---|
| Artifact formats | **screenshots** · **gif** · **titlecards/endcards** toggle | `pipeline.process_clips` (clip/screen/gif), `titlecards.wrap_clip_with_cards` |
| Highlights | **highlights reel** (severity / uniqueness / annotation scoring, default 180s budget) | the `-H` reel path in `pipeline`/`spreadsheet` |
| Screenspace breadth | numbers · template · flow · scene · inactivity · boundary detectors; **multitool** chain; **timelapse**; **heatmap** | `screenspace_tools.TOOLS[*]`, `screenspace_multitool`, `screenspace_heatmap` |

*(Deferred to demand: data export, gallery viewer, transcript export md/srt/vtt — all exist in the
backend; add when a recipe needs them.)*

### P3 — Whole-study batch *(the multiplier)*
A launch-time **"run for all participants"** that fans the entire active blueprint out, one run per
participant, with results grouped per participant under one batch view.
- Runner: instantiate the blueprint per participant (rebind the `video_source`/participant param),
  reuse the existing per-run `WorkflowRunner`; a thin batch coordinator tracks N runs.
- **Continue-on-error is mandatory here** — one bad participant must not sink the batch.
- Surfacing: a batch summary (N done / failed / skipped) + drill into any participant's run.
- Concurrency: default sequential (Whisper/Ollama are single-resource); a tuning knob may allow a
  small fan-out cap for ffmpeg-bound graphs — *open question, see below.*

### P4 — Recipe / template layer *(the posture's centerpiece; templates = seeded stashes)*
- Ship the headline graphs as **built-in stashes** (the M5 stash system, pre-seeded): e.g.
  "Transcribe → Find Word → Scan → Clips → Viewer", "Scan → Highlights reel", "Transcribe → Summary
  + Citations". One palette, no parallel "template" concept.
- "New from recipe" instantiates a built-in stash onto a fresh blueprint with sensible defaults.
- Built-ins are read-only seeds; instantiating copies (non-destructive, reusing the M5 copy-on-save).

### P5 — Authoring & run-history UX *(interleaved, ongoing)*
- **Pre-run validation panel** — missing params, unmet `requires`, dangling required inputs, cycles —
  before Run, not as a runtime failure.
- **Run-history UX** — re-run, cancel/retry a node, expandable per-node results.
- **Large-result storage** — store full per-node outputs in `run_id/node_id` sidecar files, fetched
  lazily; keep the SSE snapshot to counts + status. *(Define the contract before P2 emits big
  artifacts — see open questions.)*
- Canvas niceties as they earn their place (groups/comments, copy-paste) — not a priority.

### P6 — Triggers, narrow-first *(last; the one automation piece)*
- A single **watch-dir watcher**: when a new video appears in `-i`, auto-run a **designated**
  blueprint (typically a fan-out recipe). Debounce + dedup so one file doesn't fire twice.
- Fill the reserved `blueprint.trigger` field with a minimal schema (`{type:"watch_dir", enabled}`)
  that leaves room for `transcript_complete` / `scan_event` chaining and a general monitor later —
  **build the narrow case, keep the seam.**
- Safeguards: enable/disable per blueprint, a "triggered runs" view, a dry-run preview.

---

## Explicitly out of scope (cut, with rationale)

- **Memoization / cross-run caching** — chose "just re-run"; avoids cache-invalidation risk.
- **General control-flow suite** (map / filter / accumulate / switch) — graph-as-programming-language
  scope sink; the existing **gate** plus whole-graph fan-out covers the real need.
- **Composable `foreach` node** — whole-graph fan-out (P3) covers the common case far more simply.
- **General `TriggerMonitor`** — start with the narrow watch-dir (P6); generalize only on demand.
- **A separate "templates" system** — folded into stashes (P4).
- **Sibling-node parallelism inside a single run** — stays sequential; only the *batch* fan-out
  across participants might parallelize (open question).

---

## Open questions remaining

1. **Batch result layout (P3):** per-participant output subfolders + a combined batch view? How do
   batch artifacts avoid filename collisions across participants?
2. **Large-result storage contract (P5):** sidecar files keyed by `run_id/node_id` — settle the shape
   before P2 ships big outputs (highlights reels, screenshot sets).
3. **Batch concurrency (P3):** strictly sequential, or a small fan-out cap for ffmpeg-bound graphs
   that respects Whisper/Ollama single-resource contention?
4. **Trigger binding (P6):** how is a watch-dir bound to its blueprint — a per-blueprint toggle, or a
   single "active trigger" slot? What's the UX for "this graph runs automatically"?
5. **Recipe set (P4):** which 3–5 built-in recipes ship first?

## Reference (where the capability already lives)

- Artifact pipeline: `pipeline.process_clips` (clip/screen/gif), `process_reel`, the `-H` highlights
  path; `titlecards.py`; gallery/timeline in `viewer.py`; `data_export.write_export_bundle`.
- Screenspace: `screenspace_tools.py:TOOLS` (12 tools), `screenspace_multitool`, `screenspace_heatmap`.
- Thinking agents: `thinking_agents.AGENTS` (+ `new-thinking-agent` skill).
- Engine seams: `workflows.py` (`ADAPTERS`, `_EXECUTORS`, `NODE_TYPES`); `blueprint.trigger`
  (reserved null) + stash storage in `workflows_manifest.json`.
