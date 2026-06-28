# Changelog

Notable changes per release. Headings follow `## <version> — <YYYY-MM-DD> — <tool>` where the tool is one of `Studio`, `Screenspace`, `Transcripts`, `Workflows`, or `Core`. The first bolded line is the title; everything after is the body.

## v0.13.41 — 2026-06-27 — Workflows
**Detect node, interval captures, and editor power tools**
Unified Detect node over per-detector types; Interval Captures samples screenshots or GIFs across a range; per-node Ollama and Whisper model levers; palette search; blueprint JSON import/export; copy/paste/duplicate, mute, undo/redo; Run split-button with Run to here; middle-mouse pan and colour-coded title bars.

## v0.13.27 — 2026-06-27 — Workflows
**Collection control nodes**
Filter, partition, merge, limit, and dedup nodes thin or combine the collections flowing through a graph — gate clip selections before Make Clips, cap artifacts before the viewer, or branch matched vs. unmatched streams.

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
**Workflows mode — node canvas and run engine**
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
The Color tool can fire when a target colour appears anywhere in the region (per-pixel), with a Min area % control and presence-aware calibration — standalone, in Multitool steps, or from the CLI.

## v0.12.8 — 2026-06-20 — Studio
**Source times and severity tint on queue cards**
Every artifact and reel card shows its source start–stop time; spreadsheet-sourced cards tint their caption by row severity.

## v0.12.8 — 2026-06-20 — Core
**Elapsed time and ETA on long operations**
Screenspace tasks and transcription show elapsed plus an estimated time remaining; Studio builds and thinking agents show elapsed only. Clocks survive a page reload via server-stamped start times.

## v0.12.7 — 2026-06-20 — Screenspace
**CLI scene analysis and headless task re-run**
Run scene analysis from the command line with `--ss-task scene` and re-run any saved manifest task headlessly with `--ss-run-task` — the path for unattended multitool chains.

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
A full-frame Boundary tool flags where visual content changes substantially — menu-to-gameplay, level transitions, loading screens ending — without drawing a region first.

## v0.12.1 — 2026-06-18 — Screenspace
**Pinned-frame calibration workflow**
Pin reference frames on the timeline, score detector sensitivity against them, get a suggested threshold, and apply it — with a calibration strip, grid controls, and integration into task creation.

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
