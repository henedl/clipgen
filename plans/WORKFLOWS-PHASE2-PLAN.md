# Workflows — Phase 2 plan (recipes · batch · catalog · triggers)

> **Phase 2 complete — P1–P6 all shipped.** v1 (`plans/WORKFLOWS-PLAN.md`, M0–M5) proves the
> cross-domain chain with a curated node set and a sequential, in-process run engine. Phase 2 made it a
> daily tool: a **recipe/template layer on top of a steadily-growing catalog** (P1/P2/P4), a
> **whole-study batch** (P3), authoring + run-history UX (P5), and a **narrow watch-dir trigger** (P6,
> the final workstream). All open questions are resolved (see *Open questions*); remaining items are
> demand-driven follow-ups noted inline per workstream.

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
build on its runner. **M5 (stashes: save/instantiate sub-graphs)** has now **landed alongside P4**:
`GET/POST/PUT/DELETE /api/stashes` (combined-manifest CRUD mirroring blueprint CRUD), a
`workflows-stashes.js` satellite with a sidebar stash section (drag-to-canvas + click-to-add, client-
side id remap), and the read-only built-in recipes served from `workflows.BUILTIN_STASHES`.

## Gaps carried out of v1 (the starting state)

1. ~~**Frontend `canConnect` ignores `ADAPTERS`.**~~ ✅ **Resolved (P1).** `canConnect` now consults
   the served `ADAPTERS` table and coerced wires render dashed with a "↳ coerced" tooltip.
2. ~~**`ss_scan` exposes 4 of 10 detectors** *and passes `parameters={}`*.~~ ✅ **Resolved (P2).**
   Replaced by ten per-detector nodes whose real params reach the scan; multitool/timelapse/heatmap added.
3. ~~**`make_clips` emits `clip` only.**~~ ✅ **Resolved (P2).** `output_format` enum + `titlecards`/
   `titlecard_duration` params; `highlights` selector node added. (Compression still demand-driven.)
4. ~~**One participant per run** — no whole-study batch.~~ ✅ **Resolved (P3).** "Run all" fans the
   blueprint out one sequential run per participant, grouped under one batch card.
5. ~~**Thin run history**; large per-node results stripped with no fetch-on-demand path.~~ ✅
   **Resolved (P5).** Inspectable per-node results are written to
   `<output_dir>/workflow_runs/<run_id>/<node_id>.json` sidecars and lazily fetched on row-expand;
   pruned in lockstep with the 50-run cap. (P3 already evicted terminal runners.)
6. ~~**No pre-run validation surfacing** — missing params / unmet context only fail at runtime.~~ ✅
   **Resolved (P5).** A client-side Issues panel aggregates cycle / unwired-input / unmet-context /
   empty-required-param errors (disabling Run) plus warnings; recomputed on every edit.

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
on the timeline and timelapse/heatmap land in a separate attachments panel.

**P2 follow-ups (shipped same PR):** ✅
- **`time_range` node** (new Source) — manual in/out entry (`MM:SS`/range syntax), outputs `timeRange`
  to feed SS scan windows or `make_clips`. ✅
- **`clipRecords → timeRange` adapter** — lets `sheet_selection` (or any clipRecords) drive SS scan
  windows ("scan only the cells' timestamps"). The one new adapter (7 total now). ✅
- **Full-frame fallback for Screenspace** — an unwired/empty region now scans the whole frame instead
  of a silent zero-size no-op: shared `workflows._resolve_region_coords` (covers all SS detectors,
  multitool, timelapse, and the self-extract reference path), and CLI `--ss-task TYPE PARTICIPANT`
  with REGION omitted defaults to `full_frame`. ✅

*(Deferred to demand: data export, gallery viewer, transcript export md/srt/vtt — all exist in the
backend; add when a recipe needs them.)*

### P3 — Whole-study batch *(the multiplier)* — ✅ Done
Fans the active blueprint out one run per participant, results grouped under one batch card.
**Trigger lives in the Video Source node**, not a separate button: its participant dropdown gains an
**"All participants"** option (sentinel `WF.ALL_PARTICIPANTS`); the single **Run** button detects it
(`blueprintWantsBatch`) and launches a batch instead of one run. **Shipped:**
- `workflows.bind_participant(blueprint, participant)` deep-copies + rebinds every `video_source`
  node; `blueprint_participant_nodes` gates the endpoint. `WorkflowRunner` carries `participant`/
  `batch_id` in its snapshot. A thin coordinator (`workflows_server._run_batch`) runs N child
  `WorkflowRunner`s **sequentially** on one daemon thread; batches are *derived* by grouping runs on a
  `batchId` tag (no new manifest key), live coordination in `_batches`.
- Endpoints: `POST/GET /api/batches`, `GET/POST /api/batches/<id>[/cancel]`, `/stream` (SSE).
- **Continue-on-error** enforced (a failing participant leaves the rest running); a **cancel**
  short-circuits remaining children to `cancelled`.
- Frontend: a batch summary card (N done / failed / cancelled) expands into per-participant rows;
  clicking a participant drills into its run (canvas tint).
- **`has_video` parity:** `/api/catalog` `context.participants` is filtered to participants with a
  real video file — the dropdown (and "All participants") match exactly the set the batch fans over.
- **Runner eviction** (a P5 item folded in): terminal runners are dropped from `_runs` (batch + single
  run) once persisted — fixes the in-memory result leak that batch would have multiplied.
- **Batch-safe history cap:** `_trim_run_history` evicts whole *units* (a batch's children together,
  newest unit + live batches always kept) instead of per-record, so a batch larger than the 50-run
  cap is never split (it would 404 on drill-in). Children persist on execution, not pre-queued.
- *Resolved open questions:* concurrency = **strictly sequential**; output = **flat dir** (every
  output node already routes through `files.get_unique_filename`, so `{study}_{participant}` clip
  names + auto-suffixed reels/viewers avoid collisions — no per-participant subfolders needed).

**Follow-up (future):** replace the single "All participants" option with a **custom multi-select
dropdown** — pick an arbitrary subset of participants to fan out over, not just one or all. The batch
endpoint already accepts a `participants` list; only the param widget + a richer participant param
value (a list) are needed. Keep "All" as a convenience shortcut.

### P4 — Recipe / template layer *(the posture's centerpiece; templates = seeded stashes)* — ✅ Done
- **Shipped** (with M5) — the two headline graphs ship as **read-only built-in stashes**
  (`workflows.BUILTIN_STASHES`, served by `GET /api/stashes` ahead of the user's stashes, *not*
  persisted — code, not data): "Transcribe → Find Word → Make Clips → Viewer" and "Sheet Selection →
  Highlights → Build Reel → Viewer". (Demand-driven later: Scan → Clips → Reel; Scan → Measure → Gate
  → Reel; Transcribe → Summary + Citations + Friction; Template/Flow → Heatmap; Video → Timelapse.)
  One palette — built-ins and user stashes share the sidebar list and the same instantiate path.
- Instantiating (drag-to-canvas or click-to-add) stamps a fresh copy: the satellite remaps every node
  id, rewrites edges through the id map, and offsets positions. Non-destructive — built-ins carry a
  `builtin:true` flag and the CRUD routes reject renaming/deleting them (403).

### P5 — Authoring & run-history UX — ✅ Done
- **Pre-run validation panel** (client-side) — ✅ a new `workflows-validate.js` satellite aggregates
  the scattered cues into a `#wfValidation` Issues panel; **errors disable Run, warnings don't**, and
  it recomputes on every edit (the hub's `scheduleSave` + `openBlueprint` call `WF.refreshValidation`).
  *Errors:* cycle (`WF.graphHasCycle`, a JS port of `topo_order`'s Kahn loop; the server 400 stays a
  backstop) · unwired required input · unmet context (`nodeContextMet`) · **empty required param**
  (new `required:true` flag on `ParamSpec` — `find_word.word`, `video_source.participant`,
  `time_range.ranges`, `ss_text.search_string`). *Warnings:* heatmap style needs a matching
  template/flow/change (or multitool) upstream · orphan node · gate with no scalar source. Each row
  links to its node via `WF.focusNode` (select + pan-to-centre). `WF.nodeIssues` is the single source
  of truth, shared with the on-card `.disabled`/`.invalid` cue in `workflows-nodes.js`.
- **Large-result sidecar storage** — ✅ the runner writes *inspectable* node outputs (filtered by
  declared output-port type — `artifacts`/`events`/`segments`/`summary`/`citations`/`friction`/
  `manifest`/`viewerHtml`/`scalar`; plumbing types like `clipRecords` are dropped, dodging the gspread
  `Cell`, and an `events` value's heavy `raw_results` rider is projected out) to
  `<output_dir>/workflow_runs/<run_id>/<node_id>.json` on each node completion (JSON-sanitized via
  `utils.sanitize_floats`). The snapshot adds a per-node `hasResult` flag;
  `GET /api/runs/<run_id>/nodes/<node_id>/result` serves the file; the frontend lazily fetches on
  row-expand and renders by type (artifact/event/segment lists, reel manifest, viewer path,
  summary/citation/friction text, scalar). `workflow_runs/<run_id>/` dirs are pruned in lockstep with
  the 50-run cap (`_trim_run_history` now returns dropped ids). SSE/snapshot stay counts-only.
- **Run-history UX** — ✅ **re-run** (a button on terminal run cards relaunching the same blueprint)
  + expandable per-node results backed by the sidecars. **Deferred:** per-node cancel/retry — it
  requires partial/memoized re-run, which the plan explicitly cut ("just re-run"); whole-run cancel
  already exists (Stop). Canvas niceties (groups/comments, copy-paste) stay out — not a priority.

### P6 — Triggers, narrow-first *(last; the one automation piece)* — ✅ Done
- **Shipped** — a single **watch-dir watcher** (polling daemon thread in `workflows_server`, no new
  dependency; mirrors the screenspace worker's `daemon=True` posture). When a *new* participant video
  lands in `-i`, it auto-runs the single **armed** blueprint as one run bound to the just-arrived pid
  (`workflows.bind_participant`) — **not** a whole-study batch. Poll interval =
  `config.WORKFLOWS_WATCH_POLL_SECONDS` (5s); a pid must stat identically across two consecutive polls
  before firing (the partial-copy guard), and the seen-set is seeded at startup so the pre-existing
  backlog and re-added files never fire (no retro-fire when arming later).
- **Single active trigger:** `blueprint.trigger` holds `{type:"watch_dir", enabled}`; arming one
  blueprint disarms every other (enforced in the dedicated `PUT /api/blueprints/<id>/trigger`, which
  also rejects arming a graph with a cycle or no `video_source`). The `type` field keeps the seam for
  `transcript_complete` / `scan_event` chaining + a general monitor later.
- **Run-launch refactor:** the single-run lifecycle is now `_launch_run(blueprint, participant,
  triggered)`, shared by `POST /api/runs` and the watcher. Triggered runs carry a `triggered` flag on
  the runner snapshot.
- **UX:** a per-blueprint "Auto-run on new video" toolbar toggle (single-active, gated on a valid
  graph with a Video Source — re-gated by the validate satellite on every edit) + a "⚡ triggered"
  badge on auto-launched runs in the existing run history (no separate panel). Guarded by
  `test_workflows_frontend_source.py` + the watcher/CRUD unit tests in `test_workflows_api.py`.
- **Field-test fixes (same PR):** (1) the combined web server now forces `utils.NO_INPUT_MODE` on at
  launch — server-driven clip generation (triggered runs *and* Studio generate) was blocking a daemon
  thread on the interactive fuzzy-match `input()` prompt; it now skips-and-reports. (2)
  `utils.participant_id_from_source_name` rejects ids containing whitespace, so a Finder/Explorer
  duplicate (`study_P03 copy.mp4`) is no longer a phantom participant that auto-fires a run. (3) the
  run panel runs a low-frequency idle discover poll so a triggered run (which this client never
  started) surfaces live instead of only on a manual reload.
- **Deferred (future seam):** per-blueprint enable history view beyond the badge; a dry-run preview;
  `transcript_complete` / `scan_event` chaining; a general `TriggerMonitor`.

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

*(none — P6 was the last workstream; its open question is resolved below.)*

*Resolved (P6 build):* **trigger binding** → a **single active trigger** across all blueprints
(`blueprint.trigger = {type:"watch_dir", enabled}`); arming one disarms the rest. The "this graph runs
automatically" UX is a per-blueprint **"Auto-run on new video"** toolbar toggle (armed = accent bolt),
gated on a valid graph with a Video Source. A new video fires **one run for the just-arrived
participant** (not a batch); auto-launched runs show a ⚡ badge in the existing run history.

*Resolved (P3 build):* **batch concurrency** → strictly sequential. **batch output layout** → flat
output dir (existing `files.get_unique_filename` + `{study}_{participant}` naming dodges collisions;
no per-participant subfolders). *Resolved earlier:* **large-result storage** → `run_id/node_id`
sidecars, inspectable outputs only, runner-written, lazily fetched (P5). **Recipe set** → two
starters chosen (P4).

## Reference (where the capability already lives)

- Artifact pipeline: `pipeline.process_clips` (clip/screen/gif), `process_reel`, the `-H` highlights
  path; `titlecards.py`; gallery/timeline in `viewer.py`; `data_export.write_export_bundle`.
- Screenspace: `screenspace_tools.py:TOOLS` (12 tools), `screenspace_multitool`, `screenspace_heatmap`.
- Thinking agents: `thinking_agents.AGENTS` (+ `new-thinking-agent` skill).
- Engine seams: `workflows.py` (`ADAPTERS`, `_EXECUTORS`, `NODE_TYPES`); `blueprint.trigger`
  (reserved null) + stash storage in `workflows_manifest.json`.
