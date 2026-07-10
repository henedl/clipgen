# Screenspace → DaVinci Resolve integration

**Status: not started** — investigation done (July 2026), phases below are ready to act on.

Goal: make Screenspace's analysis usable from DaVinci Resolve — full functionality (region
drawing, calibration, tool selection, per-tool parameters), with results landing in Resolve
as **markers** and/or **cuts**.

## Investigation findings

### OpenFX (C++) is the wrong vehicle — rejected

Resolve's image processing *is* extended via C++ OpenFX plugins, but OFX plugins are per-frame
pixel **effects** inside the render graph. They have no API access to the timeline, markers, or
cutting — they can only transform pixels handed to them. (The one commercial OFX product that
syncs Resolve markers, FlowCaster, does it via its own side channel, not through OFX.) A native
OFX port would additionally require:

- rewriting `screenspace_primitives.py`'s cv2/numpy logic in C++;
- re-authoring every param table in OFX's limited param descriptors;
- no realistic path for the region-drawing/calibration UX or the EasyOCR (torch) tools.

### The Python scripting API is the right vehicle

Resolve ships a Python scripting API (`DaVinciResolveScript`, loaded from the Resolve install)
that provides exactly the two write paths we need, verified against the v19/v20 API docs:

- **Markers**: `Timeline.AddMarker(frameId, color, name, note, duration, customData)` and
  `MediaPoolItem.AddMarker(...)` (frameId in source frames for media-pool clips; markers use a
  fixed color palette).
- **Cuts**: `MediaPool.AppendToTimeline([{"mediaPoolItem": item, "startFrame": in, "endFrame": out}])`
  appends trimmed segments to the current timeline; combined with
  `MediaPool.CreateEmptyTimeline(name)` this builds a cut-down timeline programmatically.
- Clip matching / frame math: `MediaPoolItem.GetClipProperty("File Path")` and
  `GetClipProperty("FPS")`. Screenspace events carry seconds (float), never frames, so
  conversion is `frame = round(seconds * fps)`.
- `MediaPool.ImportMedia([path])` to add a source video that isn't in the media pool yet.

**Constraint:** external Python scripting requires **DaVinci Resolve Studio** and the setting
Preferences → System → General → "External scripting using: Local". The free edition only
allows the internal console (and, since 19.1, UIManager script GUIs are Studio-only too).
Workflow Integration plugins (phase 2) are also Studio-only. Fallback for the free edition:
file interchange — Resolve imports markers from an EDL via Timeline → Import → Timeline
Markers from EDL.

### Our side is already port-ready

- **The engine runs headless today.** `cli.py --ss-task` and `workflows.py` call
  `scan_*(video_path, pixel_region, **params, on_result=...)` with no Flask involved;
  `screenspace_scans.py` / `_frames.py` / `_primitives.py` have no server coupling.
  Resolve integration is therefore an **output adapter**, not a port of the CV logic.
- **Result shape is ready for markers/cuts.** Normalized `ScreenspaceEvent` records
  (`screenspace_manifest.py::create_event`) carry `time_in`/`time_out` (seconds, 2 dp),
  `detector`, `event_type`, `confidence`, `metadata`, `excluded`, `navigational`.
  Point events have `time_in == time_out`; color/inactivity produce real spans.
  `data_export.py` already loads these for export.
- **No declarative parameter schema exists.** Param names/types/defaults/ranges are split
  between Python `scan_*` signatures and hand-written JS (`screenspace.js`,
  `screenspace-multitool-params.js`). Rebuilding that UI natively (plus the region editor)
  is the single biggest cost of any "native" port — which is exactly what the approach
  below avoids by keeping the existing web UI as the control surface.
- **OCR is cleanly excludable** if a slimmer install ever matters: only `TextTool`/
  `NumbersTool` touch `screenspace_ocr.py` (the sole EasyOCR/torch entry point).

### Chosen shape

Keep Screenspace whole — its own ffmpeg decode, its own web UI, all region/param
functionality — and bridge to Resolve:

- **Phase 1**: a `resolve_bridge.py` adapter pushes finished results into a running Resolve
  (markers and/or a generated cut timeline), plus a reverse hook to preselect the clip Resolve
  is looking at. UI sits in a browser next to Resolve.
- **Phase 2**: a Workflow Integration plugin (Electron panel, Studio-only) hosts the same
  locally-served Screenspace UI **inside** Resolve's window — full region drawing in Resolve
  with near-zero UI re-authoring.

---

## Phase 1 — Resolve bridge

### 1. New module `resolve_bridge.py`

Top-level module (add to `pyproject.toml [tool.setuptools] py-modules`). Pure adapter — no
Flask, no cv2. All Resolve imports lazy so tests/CI never need the real module.

- `connect_to_resolve()` — import `DaVinciResolveScript` via the standard path dance
  (`RESOLVE_SCRIPT_API`/`RESOLVE_SCRIPT_LIB` env vars; macOS default
  `/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting/Modules`).
  Return the `resolve` object or a clear, user-facing error (Resolve not running / external
  scripting disabled / free edition).
- `find_media_pool_item(media_pool, video_path)` — walk media-pool folders matching
  `GetClipProperty("File Path")` (fallback: filename match); `ImportMedia` if absent.
- `push_clip_markers(item, events)` — per event: `frameId = round(time_in * fps)`,
  `duration = max(1, round((time_out - time_in) * fps))`, `name = event_type`,
  `note` = compact metadata summary, `customData` = event id. Detector → nearest color in
  Resolve's fixed marker palette (map lives here; same spirit as the viewer's
  `SS_DETECTOR_COLORS`).
- `build_cut_timeline(media_pool, item, events, name, pad_seconds)` —
  `CreateEmptyTimeline` + one `AppendToTimeline` per event; point events padded to
  `pad_seconds` (default `config.DEFAULT_DURATION_SECONDS`). Optionally `AddMarker` at each
  segment start — positions in a timeline we built are trivially known, which avoids the
  fragile source→existing-timeline offset math (deferred, see out-of-scope).
- `get_current_clip_path(resolve)` — reverse direction: current/selected media-pool clip's
  `File Path`, so the UI can offer "analyze the clip Resolve is looking at".
- Event sourcing: read the screenspace manifest the same way `data_export.py` does; respect
  `excluded`, skip `navigational` (boundary) events by default.

### 2. Surfacing (both thin)

- **Screenspace UI**: "Send to Resolve" action in the results panel
  (`screenspace-results.js` satellite) → new endpoint `POST /screenspace/api/resolve/push`
  in `screenspace_server.py` with `{video, task_id or event filter, mode: "markers"|"cuts"|"both"}`.
  Thin route: load events, call the bridge, return per-event success counts or a clear error
  when Resolve isn't reachable. Server and Resolve run on the same machine, so this works
  from the browser. Plus `GET /screenspace/api/resolve/current-clip` for the preselect hook.
- **CLI**: `--resolve-push` flag (markers/cuts selector), following
  [agents/skills/new-mode/SKILL.md](../agents/skills/new-mode/SKILL.md) (argparse + mode
  detection in `cli.py` + smoke test). Reads the screenspace manifest from `-o`.

### 3. Free-edition fallback (same PR or follow-up)

`export_resolve_edl(events, fps, path)` — plain-text CMX-style EDL with locator (`* LOC:`)
lines, importable in **any** Resolve edition via Timeline → Import → Timeline Markers from
EDL. No new dependencies (skip OTIO). Exposed next to the existing `/api/export/events`
outputs.

### 4. Tests

- Unit tests with a mocked `DaVinciResolveScript` (inject via `sys.modules`): seconds→frame
  conversion (incl. 23.976 fps rounding), span vs point duration, excluded/navigational
  filtering, palette mapping, EDL golden test.
- CLI smoke test per the new-mode checklist; endpoint test with the bridge patched.
- No Resolve in CI — bridge imports must stay lazy.

### 5. Docs + version

`feat:` PR → patch bump in `build/VERSION`. Add the bridge row to
[agents/ARCHITECTURE.md](../agents/ARCHITECTURE.md); note the Studio requirement in README's
Screenspace section.

### Verification

1. `uv run --extra dev pytest -c tests/pytest.ini` (new unit + smoke tests); `/check` before commit.
2. End-to-end needs a human with Resolve **Studio** open (Resolve can't run headless in CI):
   scan a sample video, push, confirm (a) markers on the media-pool clip at correct
   times/colors, (b) a generated timeline with the expected cuts, (c) the EDL imports markers
   into a free-edition timeline.

### Phase 1 checklist

- [ ] `resolve_bridge.py` (+ `py-modules` entry)
- [ ] `POST /screenspace/api/resolve/push` + results-panel button
- [ ] `GET /screenspace/api/resolve/current-clip` + preselect hook
- [ ] `--resolve-push` CLI flag (new-mode checklist)
- [ ] EDL marker export (free-edition fallback)
- [ ] Tests (mocked bridge, CLI smoke, EDL golden)
- [ ] Docs + version bump
- [ ] Manual end-to-end in Resolve Studio

---

## Phase 2 — Workflow Integration panel (sketch; own plan when phase 1 ships)

A Workflow Integration plugin is an **Electron app** installed under
`.../Blackmagic Design/DaVinci Resolve/Workflow Integration Plugins/` (Studio-only), listed
in Resolve's Workspace → Workflow Integrations menu. Blackmagic ships `SamplePlugin`
(sandboxed model, recommended) in the developer docs, and the `WorkflowIntegration.node`
Node module exposes the same scripting API to the panel's JS.

Shape:

- The panel window is a webview loading `http://127.0.0.1:8089/screenspace/` — the entire
  existing UI (region drawing, boolean region editing, calibration, params, results) appears
  inside Resolve unchanged.
- The panel launches (or health-checks and reuses) the clipgen server as a child process:
  `uv run clipgen.py --screenspace -i ... -o ...`. Needs a small settings surface for the
  input/output dirs, or it can read them from a config file.
- `WorkflowIntegration.node` is used for Resolve-side niceties the browser can't do in
  phase 1: jump Resolve's playhead to a clicked event, react to the current timeline/clip,
  and optionally push markers directly from JS instead of round-tripping through the Python
  bridge.
- Packaging: plugin manifest + Electron app dir copied into the plugins folder; document a
  manual install (small user base, no installer needed).

Open questions to resolve then: server lifecycle ownership (panel-spawned vs user-started),
whether marker-push stays in the Python bridge (one code path) or moves to the node module,
and Electron version pinning against Resolve's compatibility modes.

## Explicitly out of scope / not chosen

- **OpenFX C++ plugin** — rejected (see findings).
- **Declarative parameter schema extraction** — only needed for a *native* (non-webview)
  Resolve UI, which we're not building; skip unless that changes.
- Decoding via Resolve's media engine instead of our ffmpeg pipes.
- Source→*existing* timeline marker math (v1 puts markers on the media-pool clip and on
  generated cut timelines).

## References

- Scripting API reference: <https://resolvedevdoc.readthedocs.io/en/latest/API_basic.html>
- Studio-only external scripting: <https://forum.blackmagicdesign.com/viewtopic.php?f=21&t=113252>
- Workflow Integrations docs: <https://resolvedevdoc.readthedocs.io/en/latest/readme_workflow.html>
- Workflow Integration node-module walkthrough: <https://blog.corrlabs.com/2023/01/blackmagic-workflow-integration-node.html>
- OFX-in-Resolve example (image effects only): <https://docs.gyroflow.xyz/app/video-editor-plugins/davinci-resolve-openfx>
