# Workflows — node-based scripting frontend

Planning doc. No code is written yet. Scope is a new top-level web frontend (4th tab, next to
Studio / Screenspace / Transcripts). Track progress by checking off the milestone boxes below.

## Context

clipgen today exposes its capabilities through three siloed UIs (Studio, Screenspace, Transcripts)
and a large CLI surface. Each silo is powerful alone, but there is no way to **chain** capabilities
across them — e.g. "transcribe a video, find every time the user says a word, then run a Screenspace
scan only in those windows and cut clips from the hits," or "only enqueue a scan if the source video
is longer than N minutes." Researchers currently do this by hand across three tabs and the terminal.

**Workflows** is a new top-level frontend (4th tab, next to Studio / Screenspace / Transcripts) — a
free-form 2D node canvas where users drag "blueprint cards" (each wrapping one backend action),
wire typed outputs into typed inputs, save the canvas, reuse sub-graphs as named **stashes**, and
**run** the resulting graph with live progress and results. The outcome: the full CLI/back-end
capability set becomes composable, automatable, and cross-domain — without a build step, framework,
or new heavy dependency.

**v1 decisions (locked with the user):**
- **Vertical slice with a working run engine** — a curated node set across all three domains that
  proves cross-domain chaining end-to-end, not the full catalog. Broaden the catalog in phase 2.
- **Dual launch context** — `--workflows` accepts a spreadsheet (`-s`) **and/or** video dirs
  (`-i`/`-o`); each node declares its required context and greys out when unmet.
- **Free-form 2D canvas** — drag cards anywhere, SVG wires between typed ports, pan/zoom, positions
  persist. Built from scratch in vanilla JS (project rule: no React/TS/build tools).
- **Sub-graph stashes in v1** — save a selected group of connected cards as a reusable named stash.
- **Triggers are deferred** to a later phase, but the blueprint schema reserves a `trigger` seam now.

**Non-goals (v1):** auto-launching triggers + condition monitoring; the *complete* node catalog
(only a curated subset ships); sibling-node parallelism in the executor (sequential ready-set is
fine — Whisper/Ollama are single-resource anyway); editing/versioning of run history.

---

## Architecture overview

### New mode wiring (standard path — see `agents/skills/new-mode/SKILL.md`)
- **`cli.py`** — add `--workflows` flag; add to mode-detection + the mode-dispatch block in `main()`
  (mirror the `--screenspace` / `--transcripts` branches that call `server.start_combined_server`).
  Unlike Studio it must *not* hard-require a spreadsheet; allow `-s` and/or `-i`/`-o`.
- **`server.py`** — in `start_combined_server`, register the new blueprint
  (`combined.register_blueprint(workflows_server.workflows_bp, url_prefix="/workflows")`) and add it
  to the mutual-exclusion check; surface `"workflows": True` in the status endpoint.
- **`assets/web/topnav.js`** — append `{ id: "workflows", label: "Workflows", href: "/workflows/" }`
  to the `SURFACES` array (line ~20). That's the entire nav change.
- **`pyproject.toml`** — add `workflows` and `workflows_server` to `[tool.setuptools] py-modules`
  (flat layout; `tests/test_packaging.py` guards this — a miss ships a broken wheel).
- **Version** — `feat:`, so bump the patch in `build/VERSION` (`agents/skills/bump/SKILL.md`).

### New backend modules
| Module | Role |
|---|---|
| `workflows.py` | `NODE_TYPES` catalog (declarative registry + executors), the typed-port `ADAPTERS` table, and `WorkflowRunner` (DAG topo-sort + execution). Plus `load/save_workflows_manifest`. |
| `workflows_server.py` | Flask blueprint at `/workflows`: static routes, `/api/catalog`, blueprint/stash CRUD, run lifecycle + SSE/polling. Holds module-level run state + SSE client registry (mirrors `screenspace_server.py`). |

### New frontend (hub + satellites, mirroring Screenspace/Transcripts/Studio)
Shared namespace `window.ClipgenWorkflows` (**WF**). Load order is a contract — satellites that
*destructure* a hub/satellite fn at load must load after it; reach a later-loading owner late-bound
as `WF.fn(...)`. Route shared mutable state through `WF.state` (the same `var`-ReferenceError gotcha
that bit the Screenspace and Transcripts carves — see AGENTS.md workspace facts).

| File | Role |
|---|---|
| `workflows.html` | Page shell: `<!-- CLIPGEN_HEAD_HERE -->`, `<topnav-mount data-frontend="workflows">`, canvas + sidebar + run panel containers, script tags in load order. |
| `workflows.css` | Imports `tokens.css` + `topnav.css`; canvas/card/wire/port styling using design tokens only (no raw px/rem for shared values). |
| `workflows.js` | Hub: `WF.state` (nodes, edges, viewport pan/zoom, selection, catalog), boot, save/load orchestration, guarded delegators to satellites. |
| `workflows-canvas.js` | Pan/zoom, drag cards from sidebar palette onto canvas, node positioning, marquee selection. Reuses drag patterns from `screenspace-tasks.js` and canvas-coord mapping from `screenspace.js`. |
| `workflows-wires.js` | SVG connector layer: port hit-testing, drag-to-connect, **type validation** against the catalog, wire rendering/rerouting on node move. |
| `workflows-nodes.js` | Render a card generically from a catalog `NodeType` (ports + param inputs); param editors by `ParamSpec` type (number/enum/color/region/etc.). Greys out nodes whose `requires` context is unmet. |
| `workflows-runs.js` | Run/results panel: "Run" button, SSE subscription + polling fallback, per-node status/progress, terminal results (viewer link, artifact thumbs, event/segment counts). |
| `workflows-stashes.js` | Sidebar palette of node types **and** premade stashes; "save selection as stash"; instantiate a stash onto the canvas. |

### Reused building blocks (do not reinvent)
- **Persistence:** `utils.load_json_manifest()` / `utils.save_json_manifest()` (atomic temp+replace).
- **SSE + polling:** copy `screenspace_server.api_tasks_stream` + `api_tasks_list` (per-client
  `queue.Queue`, coalesce-drain, 15s keepalive) and the `_notify_sse_clients` overflow logic.
- **Frontend primitives:** `tokens.css`, `primitives.css`/`primitives.js` (`createBtn`), `utils.js`
  (`showToast`, `el`, `qs`, `CLIPGEN_CONFIG`), icon **mask-image** pattern (`/workflows/icons/`).
- **Stash CRUD shape:** `server.py:_handle_stash_crud` + Screenspace region-stash endpoints.

---

## Data model

### Typed ports (the "wire")
Ports are typed from the natural data contracts already in the codebase. Edge `from`→`to` is legal
when port types match, or a registered adapter exists:

`video` · `participant` · `region` · `events` · `transcript` · `segments` · `summary` ·
`citations` · `friction` · `timeRange` · `timestamps` · `clipRecords` · `artifacts` · `scalar`
(number/bool/str) · terminal `viewerHtml` / `manifest`.

**Adapters** (`ADAPTERS: dict[(PortType, PortType), Callable]` in `workflows.py`) — applied by the
runner when an output type ≠ consuming input type:
- `events → clipRecords` / `events → timeRange` (generalize `cli._build_clusters_from_ss_events`)
- `segments → timeRange` (filter by matched text — the headline glue; mirrors `api_search`)
- `timeRange → clipRecords` (thin wrap of `files.build_clip_records`)
- `transcript → segments` (projection), `video → scalar` (duration via `video` probe)

### Node catalog — single source of truth (`NODE_TYPES` in `workflows.py`)
Declarative `TypedDict` registry, modelled on `thinking_agents.AGENTS` (data-driven, enumerable) but
carrying typed ports the DAG needs. Each entry:

```python
class NodeType(TypedDict):
    id: str            # "transcribe", "ss_scan", "find_word", "make_clips", "timeline_viewer", "gate"
    label: str
    domain: Literal["artifact", "screenspace", "transcript", "thinking", "control"]
    category: str
    inputs: list[Port]            # {name, type, optional}
    outputs: list[Port]
    params: list[ParamSpec]       # {name, type, default, min/max/choices, label}
    requires: list[Literal["sheet", "videoDir"]]
    execute: Callable[[NodeContext, dict, dict], dict]   # (ctx, inputs, params) -> {out_port: value}
```
`NodeContext` carries `on_progress(float)`, `cancel_event: threading.Event`, and
`cancel_flag()->bool` so each executor forwards whichever form its callee wants (scans want
`cancel_flag`; thinking agents want `cancel_event`). **Adding a node in phase 2 = append one
`NodeType` + executor; zero frontend edits** (same property `AGENTS` advertises).

Executors are thin adapters over existing pure functions:
- `transcribe` → `transcripts.transcribe_video` / `transcribe_timeline`
- `ss_scan` (tool selected by param) → `screenspace_tools.TOOLS[type].scan(...)` →
  `screenspace_manifest` event generation; accepts optional `timeRange` to scan only those windows
- `summarize` / `citations` / `friction` → `thinking_agents.*`
- `find_word` → search `segments` text → `{timeRange}` (the cross-domain glue)
- `gate` → reads a `scalar`, returns `{pass: bool}`; downstream of the false branch is skipped
- `make_clips` → `files.build_clip_records(...)` + `pipeline.process_clips(...)` → `{artifacts}`
- `timeline_viewer` → `viewer.generate_timeline_viewer(...)` → `{viewerHtml}` (terminal)

### Catalog → frontend
A dedicated **`GET /workflows/api/catalog`** endpoint serializes `NODE_TYPES` minus the `execute`
callables. **Not** routed through `utils.get_frontend_config()` — the catalog is large and
Workflows-specific; bolting it onto the shared config would pollute the `tests/test_shared_constants.py`
contract. (Any genuinely *mirrored* scalar constants still go through `get_frontend_config`.)

### Persistence — `workflows_manifest.json`
```json
{
  "blueprints": [ { "id", "name", "nodes": [{id,type,params,position:{x,y}}],
                    "edges": [{from,fromPort,to,toPort}], "viewport":{x,y,zoom},
                    "trigger": null } ],
  "stashes":    [ { "id", "name", "nodes":[...], "edges":[...], "createdAt" } ],
  "runs":       [ { "id", "blueprintId", "status", "nodeStates":{...}, "startedAt", "completedAt" } ]
}
```
Canvas autosaves (debounced) to a blueprint-CRUD endpoint — positions + config + viewport. Stashes
reuse the existing copy-on-save / restore-non-destructive stash pattern. `trigger` is reserved/null
in v1 (the deferred-trigger seam).

---

## Execution engine (`WorkflowRunner`)

**One `WorkflowRunner` per run, on a daemon thread** (mirrors `AgentOrchestrator.run_agent`). It
calls the underlying pure functions **directly** — it does *not* enqueue into `ScreenspaceWorker` /
`TranscriptWorker` (those are unordered/single-domain queues; routing a cross-domain DAG through
them would mean inventing fake task types and teaching foreign domains). Every domain op already
accepts the uniform `on_progress` + `cancel_flag`/`cancel_event` contract, so direct calls give the
runner clean end-to-end progress and cancellation.

- **Toposort + ready-set** (Kahn); reject cycles at submit (400). A node runs when all upstream are
  `completed`; if any upstream is `failed`/`skipped`, the node is `skipped` (reuses the dependency
  gating spirit of `transcripts_server._agent_dependencies_met`).
- **Sequential ready-set by default** — sidesteps Whisper/Ollama single-resource contention;
  intra-node parallelism (e.g. `process_clips`' own thread pool) still applies. Sibling parallelism
  is a later opt-in, not v1.
- **Per-node state** `{status ∈ queued/running/completed/failed/skipped, progress, error,
  started_at, completed_at}` + per-`(node,port)` result store. Throttle progress notifies (~0.5s,
  copy `screenspace_worker`).
- **Cancellation** — one run-wide `threading.Event`; passed as `cancel_flag=event.is_set` to
  scan/clip nodes and as `cancel_event` to thinking-agent nodes; checked between nodes.
- **Streaming** — `GET /workflows/api/runs/<id>/stream` (SSE, copy `api_tasks_stream`) +
  `GET /workflows/api/runs/<id>` polling fallback. Payload = deep-copied node-state map with private
  `_`-keys and large blobs (raw frames, full segment lists) stripped — ship counts + status; include
  terminal results (viewer URL, manifest path) on completion.

---

## The one real backend gap: `files.build_clip_records(...)`

Sheet-free clip generation **already works**: `files.prepare_clip` has a pre-parsed fast path — when
`clip["times"]` is pre-filled it never touches `clip["cell"]`. The CLI `--ss-clips` /
`--transcript-clips` paths (`cli._run_ss_clips`, `_make_synthetic_clip_record`,
`_build_clusters_*`) already build synthetic `ClipRecord`s and run them through
`pipeline.process_clips` — the exact Workflows pattern, but private to `cli.py` and single-range.

**Action:** add one public helper in `files.py` (next to `prepare_clip`, which owns the clip-record
contract — avoids a backwards `workflows → cli` import) and refactor the two `cli.py` paths onto it
(collapses existing duplication):

```python
def build_clip_records(*, participant, source_filename, time_ranges: list[tuple[float, float]],
                       description, category="workflow", study="", severity="",
                       cell_col=_WORKFLOW_CELL_COL, cell_row_base=0,
                       cluster_gap=None, pad_pre=0.0, pad_post=0.0,
                       max_duration=None) -> list[ClipRecord]:
    """Build sheet-free ClipRecords from explicit (start,end) second ranges; pre-fills `times`
    (H:MM:SS) so process_clips runs without a live spreadsheet. cluster_gap merges adjacent ranges."""
```
Minimal working record shape: synthetic `cell` (negative `row`, a reserved `col=3` for Workflows so
artifact IDs don't collide with `--ss-clips` col 1 / `--transcript-clips` col 2), pre-filled `times`,
exact `source_filename` basename (so `pipeline._check_source_video` exact-match wins, skipping fuzzy
resolution), empty annotations. Convert seconds via `utils.seconds_to_timestamp(int(s),
force_hours=True)`; cluster via existing `utils.cluster_spans`.

---

## Curated v1 node set (proves the headline path + a gate)

Headline graph wires end-to-end with these:
`Video Source → Transcribe → Find Word (→timeRange) → SS Scan (windowed →events) → Make Clips
(→artifacts) → Timeline Viewer`, with an optional `Gate` reading `Video→scalar(duration)` that skips
the scan branch when duration ≤ N.

- **Sources:** Video Source (pick participant/videos from `-i`) · Sheet Selection (cells/lines/
  category → `clipRecords`; `requires:["sheet"]`) · Region (define/pick a Screenspace region).
- **Transcript/Thinking:** Transcribe · Summarize · Citations · Friction · Find Word.
- **Screenspace:** SS Scan (tool by param — ship text / color / change / similarity in v1) → events.
- **Artifact:** Make Clips · Build Reel · Timeline Viewer (terminal).
- **Control:** Gate (condition on a scalar) · (adapters handle events→clips, segments→timeRange).

---

## Milestones (each independently committable)

- [x] **M0 — Mode scaffold.** `--workflows` flag + dispatch (`cli.py`), blueprint registration
      (`server.py`), topnav entry, `pyproject.toml` py-modules, empty `workflows.html/css/js`,
      packaging + CLI-mode smoke tests. Ships an empty, reachable Workflows tab.
- [x] **M1 — Canvas core.** Pan/zoom, sidebar palette, drag cards onto canvas, select/move,
      persist positions + viewport, blueprint save/load (`workflows_manifest.json`). No wires yet.
      **Pulled forward from M2** (scope decision): the real `NODE_TYPES` registry +
      `GET /api/catalog` endpoint, generic card rendering from `NodeType` (label, domain accent,
      static port markers), and palette grey-out by `requires`. Also ships a **multi-blueprint
      switcher** (create/name/switch/delete). `NODE_TYPES` is declarative-only here (each node's
      `execute` lands in M3).
- [x] **M2 — Typed ports + wires.** SVG connectors with drag-to-connect + **type validation**,
      interactive **param editors** by `ParamSpec`, and on-canvas node grey-out/validation. (The
      catalog endpoint + NodeType card rendering shipped in M1; M1 renders ports as non-interactive
      anchors with `data-*` hooks so M2 only adds connector behavior.) New `workflows-wires.js`
      satellite (SVG layer transformed in lockstep with `#wfWorld`, drag-to-connect, exact-type
      `canConnect()` seam for M3 adapters, wire select + Delete-key + floating × button); param
      editors (number/enum/bool/participant-dropdown/string) on cards; `.disabled`/`.invalid`
      validation cues. One backend touch: `GET /api/catalog` `context.participants`.
- [x] **M3 — Node catalog + executors.** `NODE_TYPES` + executors wrapping existing functions;
      `files.build_clip_records` + refactor `cli.py` paths onto it; `ADAPTERS` table. Adds a
      `NodeContext` (dirs + sheet + `on_progress`/`cancel_event`/`cancel_flag`) and an executor per
      node, each a thin adapter keyed by output-port name; every domain value embeds a **source
      descriptor** so the pure `ADAPTERS` (value→value) can still reach a clip's source. The one
      backend gap `files.build_clip_records` is now public (synthetic sheet-free records); the
      Screenspace multi-video scan loop was extracted to `screenspace_worker.dispatch_tool_scan`
      and shared by `ss_scan` so multi-part participants work. `sheet_selection` reuses the pure
      `spreadsheet.generate_list` headless path. Executors are validated by direct invocation
      (`tests/test_workflows_executors.py`); the `WorkflowRunner` + run endpoints stay M4.
- [ ] **M4 — Run engine.** `WorkflowRunner` (toposort, ready-set, per-node state, cancel), SSE +
      polling, run/results panel, run history in manifest.
- [ ] **M5 — Stashes.** Save selected sub-graph as named stash; stash palette; instantiate onto
      canvas.
- [ ] **Phase 2 (out of scope here):** full node catalog; triggers + condition monitoring.

---

## Critical files

**New:** `workflows.py`, `workflows_server.py`, `assets/web/workflows.{html,css,js}` +
`workflows-{canvas,wires,nodes,runs,stashes}.js`, `tests/test_workflows_{frontend_source,api}.py`.

**Modified:** `cli.py` (flag + dispatch; refactor `_run_ss_clips`/`_run_transcript_clips` onto
`build_clip_records`), `server.py` (blueprint registration + mutual exclusion + status),
`assets/web/topnav.js` (SURFACES), `files.py` (`build_clip_records`), `pyproject.toml` (py-modules),
`build/VERSION` (patch bump), `tests/test_cli_modes.py`, `tests/test_packaging.py`.

**Reference (read, mirror — do not edit):** `screenspace_worker.py` + `screenspace_server.py` (SSE/
progress/cancel patterns), `thinking_agents.py` (`AGENTS` registry shape), `pipeline.py` +
`files.py:prepare_clip` (clip-record contract), `cli.py:_make_synthetic_clip_record` (sheet-free
precedent), `utils.py` (`load/save_json_manifest`, `get_frontend_config`, `register_static_routes`).

---

## Verification

- **Automated (`/check`: ruff → ty → pytest):**
  - `tests/test_cli_modes.py` — `--workflows` accepted, dispatches, mutually exclusive with other modes.
  - `tests/test_packaging.py` — `workflows`/`workflows_server` present in py-modules.
  - `tests/test_workflows_frontend_source.py` — globs `workflows*.js`; asserts `window.ClipgenWorkflows`,
    hub publishes state + key fns, no bare cross-file `var` reads (the carve gotcha).
  - `tests/test_workflows_api.py` — `/api/catalog` returns serializable node types; blueprint + stash
    CRUD round-trips; a small run executes a 2–3 node DAG using **`config.DEBUGGING`** (already stubs
    ffmpeg in `video.py` and transcripts in `transcripts.py`) so no Whisper/ffmpeg/Ollama needed —
    assert per-node status transitions and adapter wiring.
  - `tests/test_files.py` (or new) — `build_clip_records` produces records `process_clips` accepts; the
    refactored `--ss-clips`/`--transcript-clips` still pass their existing tests.
- **Manual (browser — ask the user to test per project rule; no headless browsers):**
  1. `uv run clipgen.py --workflows -i INPUT_DIR -o OUTPUT_DIR` → `http://127.0.0.1:8089/workflows/`.
  2. Drag the headline graph (Video → Transcribe → Find Word → SS Scan → Make Clips → Timeline Viewer),
     wire ports, Run, watch live per-node progress, open the resulting viewer.
  3. Add a Gate on video duration; confirm the scan branch skips when the condition is false.
  4. Save the canvas, reload the page, confirm positions/config/wires persist; save a sub-graph as a
     stash and re-instantiate it.
  5. Launch with `-s` (spreadsheet) and confirm Sheet-Selection nodes enable while video-only nodes
     still work; launch without `-s` and confirm sheet-required nodes grey out.

## Open considerations
- **Trigger seam (deferred):** blueprint `trigger` field stays null in v1; a future `TriggerMonitor`
  watches signals (new video in `-i`, transcript completed, scan event) to auto-launch a run. Keep
  the field; build nothing.
- **Sibling-node parallelism:** intentionally sequential in v1 (resource contention). Revisit per-domain.
- **Large result payloads:** keep raw frames / full segment lists out of the SSE stream — counts +
  status only; fetch full results on demand.
