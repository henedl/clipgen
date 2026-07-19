# Changelog

Notable changes per release. Headings follow `## <version> — <YYYY-MM-DD> — <tool>` where the tool is one of `Studio`, `Screenspace`, `Transcripts`, `Workflows`, or `Core`. The first bolded line is the title; everything after is the body.

## v0.14.14 — 2026-07-18 — Screenspace
**Colorized Change, SSIM, and Flow model-view previews**
The Change diff preview is JET-colorized with an on-frame overlay that tints changed pixels without darkening the live frame; Similarity adds an SSIM difference map and score; Flow arrows are color-coded by magnitude.

## v0.14.13 — 2026-07-17 — Core
**Filter and panel commands in the palette**
Filter clears, sidebar/panel toggles, and drawer actions across Studio, Screenspace, Transcripts, Composer, Workflows, and Overview are searchable in the command palette with visibility gates so only relevant commands surface.

## v0.14.13 — 2026-07-17 — Studio
**Region-aware keyboard navigation**
Stash/clear hotkeys, filter-sidebar and artifact/reel panel collapse, and a region-aware cursor: `1`–`5` jump between the filter list, queues, and stash lists with Enter activating each target.

## v0.14.13 — 2026-07-17 — Screenspace
**Panel and tool-tab hotkeys**
Collapse the bottom panel and cycle tool tabs from the keyboard; Alt-hold hint chips dim for disabled controls and combo glyphs render evenly in the cheatsheet.

## v0.14.12 — 2026-07-16 — Core
**Alt-hold shortcut hints**
Hold Alt to see combo chips on tagged controls across every page; Studio shows action hints on the keyboard-browsed cell or card; a uniform "?" cheatsheet button replaces per-page help popovers.

## v0.14.11 — 2026-07-16 — Core
**Global command palette**
Cmd+Shift+P or Cmd+K opens a Spotlight-style palette for page navigation, participant jumps, chrome actions, and recents across all six hub pages; deep links honor `#tab=` and `#P07` arrival hashes.

## v0.14.10 — 2026-07-16 — Core
**Rebindable hotkeys across all frontends**
A shared hotkey registry unifies defaults (Space, j/k, g, seek, fine-step) and auto-generates the `?` cheatsheet; Settings → Hotkeys lets you rebind, with conflicts resolved per binding and overrides persisted across sessions.

## v0.14.9 — 2026-07-15 — Studio
**Content-aware tooltips**
Queue action buttons and intake controls show context-aware tooltips that reflect card count, selected format, and state instead of stale nudges; Add-all buttons match the blue solid CTA style of Add-all to Reel.

## v0.14.8 — 2026-07-15 — Core
**Composer lane in Convergence and Metadata search**
Convergence adds a per-participant Composer swim lane from cut pairs; Metadata gains a search box that highlights matches and scrolls to the target row; Detect boundaries moves to a Screenspace topnav quick action.

## v0.14.7 — 2026-07-13 — Core
**Five new Map visualizations**
Color-by choropleth, shift-click pairwise compare arcs, direct-axes 5D scatter, session trajectories with replay comets, and auto-labeled cluster hulls on the Overview Map tab.

## v0.14.2 — 2026-07-13 — Core
**Annotated exports across recording parts**
Composer burn and GIF exports stitch spans that cross a multi-part boundary into one continuous clip before the overlay pass, so annotations render correctly across seams.

## v0.14.0 — 2026-07-13 — Core
**Overview page with 3D similarity Map**
A new Overview tab gathers cohort-level lenses: a 3D Map positions participants by PCA over sheet timestamps, transcript marks, and Screenspace events, with click-to-explain, session replay, and drill-down; Metadata and Convergence move here from Studio.

## v0.13.61 — 2026-07-12 — Core
**Composer — source-video cutting and annotations**
A Composer page cuts source video with named in/out pairs, non-destructive marker trims, and canvas annotations; cuts and trims feed Studio Artifact/Reel queues, and annotated screenshot/GIF/burn exports render via overlay.

## v0.13.61 — 2026-07-12 — Transcripts
**Friction before Summary**
Deterministic friction scores populate the heatmap, timeline band, and stat chips immediately; LLM-refined moments still require Summary, and search results drop cross-reference badges that crowded the narrow dropdown.

## v0.13.60 — 2026-07-10 — Screenspace
**Live magic-wand tolerance scrub**
Press-drag-release on the magic wand: horizontal drag scrubs flood-fill tolerance with a live contour preview on the canvas; release commits a new region or applies Shift/Alt boolean combine; Escape cancels mid-scrub.

## v0.13.59 — 2026-07-10 — Screenspace
**Boolean edits on unsaved canvas regions**
Shift/Alt/Shift+Alt add, subtract, and intersect now target the pending region drawn on the video before it is saved. Refine a rough shape in place with no server round-trip.

## v0.13.59 — 2026-07-10 — Screenspace
**Auto-generated task and event names**
New tasks get descriptive names from their params (e.g. `Text "checkout" · header`, `Color: blue · HUD`) instead of generic `type: region` labels; task cards, run pill, results switcher, and timeline tooltip show the stored name.

## v0.13.58 — 2026-07-10 — Screenspace
**Boolean region editing**
Combine shaped regions with Photoshop-style modifiers (Shift add, Alt subtract, Shift+Alt intersect) or merge shift-selected regions; multi-contour combo shapes persist when the result is not axis-aligned.

## v0.13.57 — 2026-07-09 — Transcripts
**Optional severity on flagged segments**
Marks gain a severity dropdown in the pop-over, a colored dot on segment rows, and a `mark_severities` export column; Studio intake filters by severity; Metadata adds a transcript severity distribution chart.

## v0.13.57 — 2026-07-09 — Transcripts
**Transcription progress on the timeline**
While a participant is being transcribed, a faint dot texture covers the un-transcribed portion of the video timeline and wipes away left-to-right in sync with decode progress.

## v0.13.57 — 2026-07-09 — Studio
**Screenspace clusters in Metadata**
The Metadata tab counts time-adjacent Screenspace event clusters instead of raw per-second events so dense scans do not dominate tables and charts; toggle in settings with live re-render and clustered export field names.

## v0.13.56 — 2026-07-09 — Workflows
**Clip padding and max duration on artifact nodes**
Make Clips and Build Reel nodes expose pad start, pad end, and max duration params: nudge clip boundaries inward or outward or cap segment length without new queue cards.

## v0.13.55 — 2026-07-09 — Screenspace
**Lasso and magic-wand region selectors**
Draw freehand lasso or flood-fill magic-wand regions alongside rectangles; shaped polygons rasterize to masks for color, change, flow, OCR, template, and scene tools, with dimmed model-view previews outside the polygon.

## v0.13.55 — 2026-07-09 — Workflows
**Live minimap and zoom controls**
The corner minimap mirrors the canvas view in real time with an edge-safe view frame; floating zoom-in, zoom-out, and fit-to-content buttons; restores canvas drag and wheel-zoom after deferred script loading.

## v0.13.54 — 2026-07-09 — Core
**CLI friction agent and --settings**
Run the friction thinking agent from the command line with `--friction`; open the interactive settings editor with `--settings` before a run; improved fallback to a local Excel file when Google auth fails.

## v0.13.53 — 2026-07-06 — Transcripts
**Streaming AI summaries**
Summary generation streams tokens into the Transcripts panel as the model produces them instead of showing a spinner for the whole run.

## v0.13.52 — 2026-07-06 — Transcripts
**Faster Whisper transcription**
VAD on by default skips silence without clipping quiet speech; beam size 2 and configurable CPU threads speed decode on many-core machines. All tunable in Studio → Transcription settings.

## v0.13.51 — 2026-07-04 — Core
**Unified motion for toasts and Studio overlays**
Toast and Studio overlay cards animate through the shared ClipgenMotion engine with fade/pop entrances; duplicate CSS keyframes retired, and overlays no longer snap in when revealed from hidden.

## v0.13.50 — 2026-07-03 — Core
**Animated stash and delete exits**
Screenspace region pills and Studio queue, artifact, and reel cards play staged exit animations on stash, delete, and clear before the list re-renders, with reduced-motion fallbacks.

## v0.13.49 — 2026-06-30 — Workflows
**Multi-select participant batch**
Video Source participant param is a checkbox popover. Pick any subset to fan out a batch instead of only one participant or all; empty selection raises a validation warning.

## v0.13.48 — 2026-06-30 — Workflows
**Canvas navigation**
Fit-to-view button and F shortcut frame all nodes; dragging near the canvas edge auto-pans; a corner minimap shows the graph with click/drag-to-recenter.

## v0.13.47 — 2026-06-29 — Screenspace
**Boundary flags above the timeline**
Scene boundaries render as flag glyphs in a rail above the timeline instead of in-band ticks; hover shows a tooltip and locator hairline, click seeks to the boundary time.

## v0.13.46 — 2026-06-29 — Transcripts
**Editable thinking-agent prompts**
View, edit, and reset summary, citations, and friction prompts in Settings → Summaries → Agent prompts; edits persist and apply on the next run, with placeholder validation.

## v0.13.45 — 2026-06-28 — Workflows
**Canvas discoverability and run panel polish**
Toolbar undo/redo, autosave status, shortcuts legend, and focus rings; run panel shows per-node status icons, duration, result filters, expandable result lists, and a Reconnecting pill when SSE drops; collapsible port-type legend and wider param-heavy node cards.

## v0.13.43 — 2026-06-28 — Studio
**Feedback, persistence, and accessibility polish**
Stash saves animate the new card in; generation failures show per-cell reasons; sidebar filter selections persist across reload; focus-visible rings and modal focus traps; overlay cards fade in with reduced-motion support.

## v0.13.43 — 2026-06-28 — Transcripts
**Keyboard review loop and AI-trust cues**
j/k/arrows move and seek segments, m marks, 1–6 set friction category, n/p jump between marks; auto-follow scrolls the active segment during playback; marks update optimistically; friction panel warns when segments were edited after AI analysis.

## v0.13.43 — 2026-06-28 — Screenspace
**Virtualized results and clearer task feedback**
Large result lists lazy-render in chunks; loading indicators during fetch and participant switch; failed-task errors expand on click; SSE drop shows a one-shot toast; exclude/include toggles revert and toast on failure.

## v0.13.42 — 2026-06-28 — Workflows
**Run notes, skip fixes, and batch efficiency**
Completed nodes surface non-fatal run notes when output is degraded; skip propagation respects optional merge inputs; validation warns on unwired or mismatched filter values; batch runs reuse participant-independent sheet data; watch-dir polling idles when disarmed; empty canvas offers built-in recipes and an armed auto-run indicator.

## v0.13.41 — 2026-06-27 — Workflows
**Detect node, interval captures, and editor power tools**
Unified Detect node over per-detector types; Interval Captures samples screenshots or GIFs across a range; per-node Ollama and Whisper model levers; palette search; blueprint JSON import/export; copy/paste/duplicate, mute, undo/redo; Run split-button with Run to here; middle-mouse pan and colour-coded title bars.

## v0.13.27 — 2026-06-27 — Workflows
**Collection control nodes**
Filter, partition, merge, limit, and dedup nodes thin or combine the collections flowing through a graph: gate clip selections before Make Clips, cap artifacts before the viewer, or branch matched vs. unmatched streams.

## v0.13.26 — 2026-06-27 — Workflows
**Watch-dir auto-run**
Arm one blueprint to run automatically when a new participant video lands in the input directory; triggered runs show a ⚡ badge in run history.

## v0.13.22 — 2026-06-27 — Workflows
**Pre-run validation and inspectable results**
An Issues panel blocks Run on wiring or param errors; completed nodes expose lazy-loaded result sidecars; Re-run replays a finished graph; timelapse and heatmap land in the viewer Attachments pane and Build Reel → Viewer renders playable reel cards.

## v0.13.20 — 2026-06-26 — Workflows
**Expanded catalog, batch runs, and stashes**
Per-detector Screenspace nodes, highlights selector, multitool/timelapse/heatmap/measure, and adapter-aware dashed wires; Video Source "All participants" fans out a whole study; save sub-graphs as named stashes or start from two built-in recipes.

## v0.13.16 — 2026-06-25 — Workflows
**Workflows mode: node canvas and run engine**
A fourth top-level tab chains clip, Screenspace, and transcript actions on an infinite pan/zoom canvas: drag nodes from a catalog, wire typed ports, edit params, Run with live per-node progress, and skip branches via Gate control edges.

## v0.13.11 — 2026-06-23 — Core
**Card scrubber on hover**
Sweep a queue or viewer card thumbnail to scrub frames with audio and a waveform playhead; toggle on in Studio settings or the exported timeline viewer header.

## v0.13.10 — 2026-06-23 — Screenspace
**Scene-aware boundary segmentation**
The Boundary detector gains hybrid scene metrics, hierarchical Scene A1/B2 labels, and a post-run consolidation pass; boundaries surface in results, Studio intake, Convergence, Metadata, and the timeline viewer.

## v0.13.6 — 2026-06-22 — Studio
**Clip-length intake timeline markers**
Screenspace and Transcript intake density timelines size each marker by its clip span so longer selections read wider at a glance.

## v0.13.5 — 2026-06-22 — Transcripts
**Model install consent and dynamic Ollama pickers**
Whisper and Ollama models now require explicit confirmation before downloading. Summary, citations, and friction pickers list installed Ollama models, friction can use a separate model, and pull progress is shown in-app.

## v0.13.3 — 2026-06-22 — Studio
**Titlecard and endcard background picker**
Choose a default, solid color, uploaded image, or no endcard from Settings → Video & Clips, with a live preview and reusable color picker. Selections persist and are baked into generated clips and reels.

## v0.13.2 — 2026-06-22 — Screenspace
**Rolling-window and change heatmaps**
Template tasks gain a rolling-window animation alongside static and cumulative views; Change tasks get full heatmaps. Each tool has a per-tool toggle under Settings → Screenspace → Heatmaps, and results show as a collapsible thumbnail strip.

## v0.13.1 — 2026-06-21 — Transcripts
**Cancel summary and citations**
An inline Cancel button stops summary generation or the citations pass mid-run, without leaving the tab.

## v0.13.0 — 2026-06-20 — Core
**Multiple source videos per participant**
A session can span several videos declared in the spreadsheet Filename row or auto-detected on disk; timestamps, clips, transcripts, and Screenspace events map across the full continuous timeline.

## v0.12.9 — 2026-06-20 — Studio
**Sortable Sheet Preview columns**
Cycle #, Category, Severity, and Function headers through Ascending → Descending → Off; severity sorts most-severe-first with empty values at the bottom.

## v0.12.9 — 2026-06-20 — Screenspace
**Color presence detection mode**
The Color tool can fire when a target colour appears anywhere in the region (per-pixel), with a Min area % control and presence-aware calibration, standalone, in Multitool steps, or from the CLI.

## v0.12.8 — 2026-06-20 — Studio
**Source times and severity tint on queue cards**
Every artifact and reel card shows its source start–stop time; spreadsheet-sourced cards tint their caption by row severity.

## v0.12.8 — 2026-06-20 — Core
**Elapsed time and ETA on long operations**
Screenspace tasks and transcription show elapsed plus an estimated time remaining; Studio builds and thinking agents show elapsed only. Clocks survive a page reload via server-stamped start times.

## v0.12.7 — 2026-06-20 — Screenspace
**CLI scene analysis and headless task re-run**
Run scene analysis from the command line with `--ss-task scene` and re-run any saved manifest task headlessly with `--ss-run-task`, the path for unattended multitool chains.

## v0.12.6 — 2026-06-20 — Studio
**Editable clip length**
Click a queue card's duration badge to trim symmetrically, drag in/out points, type exact times, or nudge ±30s before generating artifacts or reels.

## v0.12.5 — 2026-06-19 — Studio
**Non-blocking viewer build with cancel**
Timeline and Gallery builds run in a corner status card so the page stays interactive; Cancel stops in-flight ffmpeg work and discards partial output.

## v0.12.5 — 2026-06-19 — Screenspace
**Multitool offset windows**
Chain tools with per-step time offsets so later steps scan only a window around hits from earlier steps.

## v0.12.4 — 2026-06-19 — Screenspace
**Uploaded images in multitool template steps**
Template steps in a Multitool chain can reference uploaded reference images, not just frames captured from the video.

## v0.12.3 — 2026-06-18 — Screenspace
**Transcript tags in the sidebar**
Transcript-derived tags appear in the Screenspace sidebar for quick orientation alongside detector events.

## v0.12.2 — 2026-06-18 — Core
**Per-lane convergence offsets**
Alignment offsets can be set independently per participant lane when sessions are not coupled.

## v0.12.2 — 2026-06-18 — Screenspace
**Automated scene boundary detector**
A full-frame Boundary tool flags where visual content changes substantially (menu-to-gameplay, level transitions, loading screens ending) without drawing a region first.

## v0.12.1 — 2026-06-18 — Screenspace
**Pinned-frame calibration workflow**
Pin reference frames on the timeline, score detector sensitivity against them, get a suggested threshold, and apply it, with a calibration strip, grid controls, and integration into task creation.

## v0.11.21 — 2026-06-09 — Screenspace
**OCR accuracy controls**
Confidence gating, ROI preprocessing, integers-only Numbers mode, three-state Text normalization, and an optional confidence histogram in Results reduce false positives.

## v0.11.19 — 2026-06-07 — Transcripts
**Friction detection**
An LLM-backed Friction tab scores transcript segments for user struggle and surfaces them with severity in the analysis panel.

## v0.11.10 — 2026-05-22 — Studio
**Reels and viewers in the artifact log**
Built reels and timeline/gallery viewers appear in the artifact log alongside clips and screenshots for quick reopen.

## v0.11.9 — 2026-05-21 — Studio
**Independent artifact and reel generation**
Artifact and reel queues can generate concurrently instead of blocking each other.

## v0.11.8 — 2026-05-21 — Studio
**Button-progress generation UI**
Generate actions show inline progress on the button itself, and reel panel order follows the queue.

## v0.11.7 — 2026-05-21 — Core
**Per-participant convergence offsets**
Fine-tune audio/video alignment per participant from the Convergence tool.

## v0.11.6 — 2026-05-17 — Transcripts
**VAD and hallucination filters**
Voice-activity detection and Whisper hallucination filtering improve transcript quality on sparse or silent footage.

## v0.11.1 — 2026-05-16 — Core
**Double-click launches Studio**
Opening the bundled .app or .exe with no arguments now boots Studio directly; the Start overlay handles spreadsheet selection in the browser.

## v0.11.0 — 2026-05-16 — Core
**Refreshed Start screen**
New two-column launcher with animated brand intro, a "Recently opened" rail, a unified spreadsheet picker, an in-app changelog, and a native macOS folder picker.

## v0.10.144 — 2026-05-12 — Core
**Start overlay polish**
Auto-skip when source videos are already present, ghosted grid background, attribution footer.

## v0.10.140 — 2026-05-03 — Studio
**Drag timestamp cells to Artifact / Reel**
Cells can be dragged onto the Artifact or Reel intake with a live cascade preview of what will be generated.

## v0.10.135 — 2026-04-28 — Screenspace
**Collapsible left info panel**
Notes and top issues now live in a slim collapsible panel so the canvas keeps the bulk of the screen.
