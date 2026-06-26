# Workflows — Phase 2 plan (recipes · batch · catalog · triggers)

> **Living doc — posture set; P1/P2/P5 detail now pinned.** v1 (`plans/WORKFLOWS-PLAN.md`,
> M0–M5) proves the cross-domain chain with a curated node set and a sequential, in-process run
> engine. An earlier round set Phase 2's shape (a **recipe/template layer on top of a steadily-growing
> catalog**, **whole-study batch**, a **narrow trigger**, several heavier ideas cut). A second round
> (recorded below) pinned the **node catalog + connection model (P1/P2)** and the **validation +
> result-storage detail (P5)**. Batch layout and trigger binding stay open — see *Open questions*.

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

**M4 (run engine: `WorkflowRunner`, SSE, run history)** has landed (#467) — batch, triggers, and P5
build on its runner. **M5 (stashes: save/instantiate sub-graphs)** is still schema-only
(`workflows_manifest.json` has the `stashes` key, no CRUD/UI yet); P4's built-in recipes need it. Land
M5 before P4.

## Gaps carried out of v1 (the starting state)

1. ~~**Frontend `canConnect` ignores `ADAPTERS`.**~~ ✅ **Resolved (P1).** `canConnect` now consults
   the served `ADAPTERS` table and coerced wires render dashed with a "↳ coerced" tooltip.
2. ~~**`ss_scan` exposes 4 of 10 detectors** *and passes `parameters={}`*.~~ ✅ **Resolved (P2).**
   Replaced by ten per-detector nodes whose real params reach the scan; multitool/timelapse/heatmap added.
3. ~~**`make_clips` emits `clip` only.**~~ ✅ **Resolved (P2).** `output_format` enum + `titlecards`/
   `titlecard_duration` params; `highlights` selector node added. (Compression still demand-driven.)
4. **One participant per run** — no whole-study batch.
5. **Thin run history**; large per-node results stripped with no fetch-on-demand path.
6. **No pre-run validation surfacing** — missing params / unmet context only fail at runtime.

---

## Workstreams

Rough order; P1 is the unblock, P2–P4 are the core experience, P5 is interleaved, P6 is last.

### P1 — Adapter ↔ UI parity *(small; do first)* — ✅ Done
Serve the `ADAPTERS` keys through `GET /api/catalog` (alongside the catalog) and widen
`canConnect(out,in)` to `out === in || ADAPTERS.has([out,in])`; render adapter-bridged wires with a
distinct cue (dashed / "↳ coerced"). Unblocks *scan → clips* and every multi-type recipe.
**Shipped** (`workflows.serialize_adapters` + `workflows_server` `/api/catalog`; `workflows-wires.js`
`canConnect` + `.wf-wire-coerced`; `workflows.js` `state.adapters`). Guarded by
`test_catalog_serves_adapter_pairs`.

### P2 — Catalog tranche *(each node = one `NodeType` + executor; pinned shapes below)* — ✅ Done
Mechanics: skills `new-screenspace-tool` / `new-mode`. **Shipped** — all node rows below landed; the
single `ss_scan` was replaced by ten per-detector nodes (`_make_ss_executor` factory reading real
params via `_build_ss_scan_params`), `make_clips` no longer hardcodes `output_format`, and
timelapse/heatmap render in a new viewer **Attachments** panel. Pinned node set:

| Node(s) | Ports (in → out) | Params / notes | Reuses |
|---|---|---|---|
| ✅ **`make_clips`** (changed) | `clips?`/`video?`/`timeRange?` → `artifacts` | + `output_format` enum {clip,screen,gif} · `titlecards` bool · `titlecard_duration`. **Done** — passes through to `process_clips`. | `pipeline.process_clips`, `titlecards.wrap_clip_with_cards` |
| ✅ **`highlights`** (new) | `clipRecords` → `clipRecords` | composable selector: scores severity/uniqueness/annotation, truncates to `budget` (default 180s); wire before `build_reel`/`make_clips`. Uniqueness scored vs. existing output-dir artifacts. **Done.** | `-H` path: `spreadsheet.score_and_truncate_clips`, `files.discover_clips` |
| ✅ **10 per-detector SS nodes** (replaced `ss_scan`) | `video` (+`region?`,`timeRange?`) → `events` | `ss_text`·`ss_color`·`ss_change`·`ss_similarity`·`ss_numbers`·`ss_template`·`ss_flow`·`ss_scene`·`ss_inactivity`·`ss_boundary`; shared `_make_ss_executor(tool)` reads node params into `scan_params` via `_build_ss_scan_params` (fixed the dead-params gap). Reference detectors (template/similarity/scene) self-extract their reference from the region at `reference_seconds`. Detectors attach `raw_results` for the heatmap node. **Done.** | `screenspace_tools.TOOLS[*]` |
| ✅ **`multitool`** (new) | `video` (+`region?`) → `events` | param = ordered `step-list` sub-editor (the one new compound widget, in `workflows-nodes.js`; reuses each `ss_<type>` catalog param). Step types limited to the 6 check_frame detectors. **Done.** | `screenspace_multitool.scan_multitool` |
| ✅ **`timelapse`** (new) | `video` (+`region?`) → `artifacts` | emits `type:"timelapse"`; params speedup/format/sample_interval. **Done.** | `screenspace_scans.generate_timelapse` |
| ✅ **`heatmap`** (new) | `events` → `artifacts` | emits `type:"heatmap"`; param `style` ∈ {template,flow,change}, reads upstream `raw_results`. **Done** (rolling-GIF `window` deferred). | `screenspace_heatmap.*` |
| ✅ **`measure`** (new) | `events?`/`clipRecords?`/`segments?` → `scalar` | `metric` ∈ {count,max_confidence,total_duration}. **Activates the `gate`.** **Done.** | — |

**Connection model:** *no new `media` type* — `timelapse`/`heatmap` emit `artifacts` carrying a new
`type` value; `viewer.finalize_timeline_data` + the viewer JS branch on `type` so clips/screens/gifs go
on the timeline and timelapse/heatmap land in a separate attachments panel. *No new adapters* — every
new node connects via existing types and the 6 existing adapters.

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
- Ship the headline graphs as **built-in stashes** (the M5 stash system, pre-seeded). **Chosen
  starters:** "Transcribe → Find Word → Make Clips → Viewer" and "Sheet Selection → Highlights →
  Build Reel → Viewer". (Demand-driven later: Scan → Clips → Reel; Scan → Measure → Gate → Reel;
  Transcribe → Summary + Citations + Friction; Template/Flow → Heatmap; Video → Timelapse.) One
  palette, no parallel "template" concept.
- "New from recipe" instantiates a built-in stash onto a fresh blueprint with sensible defaults.
- Built-ins are read-only seeds; instantiating copies (non-destructive, reusing the M5 copy-on-save).

### P5 — Authoring & run-history UX *(designed; land the storage contract with P2)*
- **Pre-run validation panel** (client-side) — aggregates today's scattered cues; **errors disable
  Run, warnings don't**. *Errors:* cycle (port `topo_order`'s Kahn loop to JS; the server 400 stays a
  backstop) · unwired required input (`requiredInputsSatisfied`) · unmet context (`nodeContextMet`) ·
  **empty required param** (new `required:true` flag on `ParamSpec`, e.g. `find_word.word`,
  `video_source.participant`, search strings). *Warnings:* heatmap needs a template/flow/change
  upstream (data-driven `requiresUpstream` declaration) · orphan node · gate with no scalar source.
  Recompute on **every edit** (edge/param/blueprint change), not just load; each row links to its node.
- **Large-result sidecar storage** — write *inspectable* node outputs (`artifacts` w/ paths, `events`,
  `segments`, `summary`/`citations`/`friction`, reel `manifest`, `viewerHtml`; **not** plumbing types —
  which also dodges serializing clipRecords' gspread `Cell`) to
  `<output_dir>/workflow_runs/<run_id>/<node_id>.json`. The **runner** writes on each node completion
  (JSON-sanitized: drop non-finite floats / numpy / `Cell`); the snapshot adds a per-node `hasResult`
  flag. New `GET /api/runs/<run_id>/nodes/<node_id>/result` serves it; the frontend lazily fetches on
  row-expand and renders by type. Prune `workflow_runs/<run_id>/` in lockstep with the 50-run history
  cap, and **evict terminal runners from `_runs`** once persisted (frees the in-memory results that leak
  today). SSE/snapshot stay counts-only.
- **Run-history UX** — re-run, cancel/retry a node, expandable per-node results (backed by the sidecars).
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
2. **Batch concurrency (P3):** strictly sequential, or a small fan-out cap for ffmpeg-bound graphs
   that respects Whisper/Ollama single-resource contention?
3. **Trigger binding (P6):** how is a watch-dir bound to its blueprint — a per-blueprint toggle, or a
   single "active trigger" slot? What's the UX for "this graph runs automatically"?

*Resolved this round:* **large-result storage** → `run_id/node_id` sidecars, inspectable outputs only,
runner-written, lazily fetched (P5). **Recipe set** → two starters chosen (P4).

## Reference (where the capability already lives)

- Artifact pipeline: `pipeline.process_clips` (clip/screen/gif), `process_reel`, the `-H` highlights
  path; `titlecards.py`; gallery/timeline in `viewer.py`; `data_export.write_export_bundle`.
- Screenspace: `screenspace_tools.py:TOOLS` (12 tools), `screenspace_multitool`, `screenspace_heatmap`.
- Thinking agents: `thinking_agents.AGENTS` (+ `new-thinking-agent` skill).
- Engine seams: `workflows.py` (`ADAPTERS`, `_EXECUTORS`, `NODE_TYPES`); `blueprint.trigger`
  (reserved null) + stash storage in `workflows_manifest.json`.
