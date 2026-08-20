# Changelog

Notable changes per release. Headings follow `## <version> — <YYYY-MM-DD> — <tool>` where the tool is one of `Core`, `Studio`, `Screenspace`, `Transcripts`, `Workflows`, `Composer`, or `Overview`. The first bolded line is the title; everything after is the body.

Keep bodies to one or two sentences, ideally under 35 words. Titles and bodies are rendered as **plain text** in the Start overlay's Recent updates tab — no backticks, no markdown inline syntax, or it shows up literally.

## v0.16.0 — 2026-08-20 — Core
**Choose how source video files are named**
A Settings → Video & Clips pattern with {study} and {participant} placeholders replaces the hardcoded {study}_{participant}.mp4, driving both filename construction and disk discovery. Non-default file formats work now.

## v0.15.18 — 2026-08-20 — Workflows
**Region picker, conditional params, and stricter validation**
The Region node picks from saved Screenspace regions instead of free text, param knobs hide when they do not apply, and cycle errors name the offending nodes. A refused Run now toasts instead of failing silently.

## v0.15.18 — 2026-08-20 — Transcripts
**Fix: swapping spreadsheets no longer mixes studies**
The previous study's transcription worker and thinking agents are stopped on swap, so their results cannot land in the new study's manifest. Marks resolve by segment id and invalidate on re-transcribe.

## v0.15.18 — 2026-08-20 — Core
**Fix: hardened startup**
A failed Start overlay mount no longer loops, a corrupt start.json warns instead of silently dropping filename overrides, and shared config reaches Composer annotations and the Transcripts participant list.

## v0.15.16 — 2026-08-20 — Transcripts
**Transcribe only between in/out markers**
I and O set a per-participant window on the timeline, and every transcribe path — single pill, force, Transcribe All — decodes just that range. Parts wholly outside the window are skipped, not failed.

## v0.15.16 — 2026-08-20 — Composer
**Fix: playback, exports, annotations, and timeline**
Burn and GIF exports are cancellable, in-progress text commits instead of vanishing, the playhead survives an unprobeable part, and right-click no longer starts a timeline drag.

## v0.15.15 — 2026-08-20 — Transcripts
**Normalize Audio quick action**
Batch-rewrites source videos in place with EBU R128 loudness-normalized audio, per-track and multi-part aware, parking the original as .orig until the swap validates.

## v0.15.15 — 2026-08-20 — Transcripts
**Tighter segment timestamps and a word-level highlight**
Word timing and an RMS energy snap trim Whisper's silence overshoot from segment edges, and the active transcript row now highlights word by word during playback.

## v0.15.15 — 2026-08-19 — Core
**A real macOS menu bar**
The desktop window gains File, Go, Window and Help menus plus a Settings… item, with ⌘ equivalents for the six pages, Reload Page, and Toggle Dark Mode.

## v0.15.14 — 2026-08-19 — Core
**Fix: the last clip of a session went missing**
A bare timestamp in the final minute of a recording overshot the end, and the whole clip was rejected for it. Clips now shorten to the footage that is actually there.

## v0.15.14 — 2026-08-19 — Core
**Profiling drill-down and faster color scans**
--profile-deep LABEL attaches cProfile to a labeled span, worker threads included. It found the Color tool's downsample: removing it made per-frame color scanning 7× faster and more accurate.

## v0.15.13 — 2026-08-18 — Composer
**Annotation tools become a vertical palette**
The floating pill is now a tool rail docked beside the video, with a two-color swatch pair (X swaps), a text-size chip, and Shot / GIF / Burn moved up to the subheader. Delete is Backspace.

## v0.15.13 — 2026-08-18 — Core
**Dropdowns match the rest of the app**
Every <select> now carries the app's field styling and chevron. Safari and the desktop window showed none of it before — WebKit discards border, background and radius on a native menulist.

## v0.15.12 — 2026-08-17 — Core
**Shimmer on in-flight status text**
Indeterminate "…" status labels across every page, the Start overlay, settings and the boot splash sweep a highlight once a second. Error, success and idle copy stay flat.

## v0.15.11 — 2026-08-17 — Core
**Fixes across OCR, seeking, Windows install, and Start**
Incompatible OCR language pairs are rejected at task creation, clips from containers with a nonzero start time seek correctly, Windows installer upgrades clean up dropped files, and Start overlay hotkeys and the About panel behave.

## v0.15.11 — 2026-08-17 — Core
**Fix a mis-named recording without leaving the app**
Each source-video preview row on Start gets an inline editor listing the videos in the input folder that no participant claims. Overrides persist per user and beat the spreadsheet's Filename row.

## v0.15.11 — 2026-08-17 — Core
**Third-party attribution in About**
The About tab lists every bundled component with its version and license, generated from the shipped notice so the two cannot drift.

## v0.15.10 — 2026-08-17 — Core
**A smaller, GPL-free macOS build**
PyAV, scikit-image, imagehash and their dependency trees are gone, and OpenCV is compiled without FFmpeg — roughly 230 MB less unpacked, and the DMG is no longer conveyed under GPLv3.

## v0.15.10 — 2026-08-16 — Core
**Start overlay splits into three tabs**
Open, About and Recent updates replace one long scrolling column, so the Open workspace button can no longer scroll out of view and the changelog gets full panel height. 1/2/3 switch tabs.

## v0.15.9 — 2026-08-16 — Screenspace
**RapidOCR replaces EasyOCR**
Text and Numbers detection moves to RapidOCR, dropping torch from the dependency tree — a much smaller download with the same language coverage.

## v0.15.8 — 2026-08-16 — Core
**Windows ships an installer**
The Windows download is a real installer with Start-menu and uninstall entries instead of a zip to unpack, and the app folder's _internal directory is renamed lib.

## v0.15.7 — 2026-08-16 — Core
**Opt-in profiling, and the speedups it found**
--profile reports where time goes across scans, ffmpeg, Whisper, Sheets and streaming routes, with matching spans in the browser. Six measured wins landed in Screenspace scans, heatmaps and seeks.

## v0.15.4 — 2026-08-15 — Core
**Themed scrollbars and native controls**
Scrollbars, <select> popups, checkboxes and autofill are painted for the active theme instead of the OS default — the bright grey track no longer glows through the translucent navbar.

## v0.15.3 — 2026-08-15 — Core
**--version and --licenses flags**
--version prints the bare version number for shell substitution; --licenses prints the bundled third-party notice. Both work from a frozen bundle.

## v0.15.1 — 2026-08-13 — Transcripts
**Install Ollama from inside the app on Windows**
The install dialog's Download & install now works on Windows, fetching the SHA-pinned official installer and running it silently — no UAC prompt, no console flash.

## v0.15.0 — 2026-08-13 — Core
**Fix: the desktop app could not transcribe at all**
Frozen builds shipped without faster-whisper's VAD model, so every default transcription died instantly. Source checkouts were unaffected.

## v0.14.75 — 2026-08-09 — Core
**A boot page instead of a blank wait**
The server answers immediately with a progress page naming each startup phase and reloads when ready, and transcription pills now distinguish "loading model…" from queued or running.

## v0.14.74 — 2026-08-08 — Transcripts
**One "Embed Subtitles…" modal replaces the two embed quick actions**
The two embed actions become a single dialog with a scope picker and a "Set as default track" toggle. Progress streams per file and is cancellable, and one bad id no longer sinks the batch.

## v0.14.73 — 2026-08-04 — Studio
**Read MindNode mind maps as a clip source**
Studio can open a .mindnode bundle directly — no spreadsheet — and cut clips from its timestamps, with its own Start tab and intake tab. Participants come from P##/G## headings, categories from the levels above.

## v0.14.72 — 2026-08-04 — Transcripts
**Download Ollama in-app on macOS**
When Ollama is missing, the install dialog offers a consent-gated download of the official CLI and then starts it. An Ollama already on PATH still wins, so a user-managed copy self-updates normally.

## v0.14.71 — 2026-08-04 — Core
**Desktop builds ship pinned GPL ffmpeg/ffprobe**
The macOS DMG and Windows build no longer need a separate ffmpeg install. The bundled libfreetype also fixes titlecard encoding, which fails under Homebrew's ffmpeg 8.x.

## v0.14.70 — 2026-08-03 — Screenspace
**Hover-scrub heatmap animations, paused by default**
Animated heatmap thumbs start paused and scrub through frames on hover, resting on the last one; the GIF loads only when you click play. A broken image now says so instead of showing a broken glyph.

## v0.14.69 — 2026-08-03 — Transcripts
**Transcribe All quick action**
Queues every participant that has a source video and no transcript yet, honoring each pill's model, language and audio-track overrides. Already-transcribed and queued participants are skipped.

## v0.14.68 — 2026-08-02 — Core
**Missing-dependency guidance where users can actually see it**
Windowed launches have no console, so missing ffmpeg, Google credentials and Ollama now explain themselves in the UI — with platform-aware install hints — instead of on a stdout nobody reads.

## v0.14.67 — 2026-08-01 — Core
**Overview drops the Map tab**
The WIP 3D similarity Map is gone, along with Three.js and the routes behind it. Overview now has three tabs — Metadata, Convergence and Reports — with hotkeys renumbered 1/2/3.

## v0.14.66 — 2026-08-01 — Core
**Detect fragmented OBS recordings and remux them in one click**
OBS's fragmented recordings play fine in ffmpeg but browsers cannot seek them. Screenspace, Transcripts and Composer now detect that container, warn, and offer a stream-copy remux that fixes playback in place.

## v0.14.65 — 2026-08-01 — Transcripts
**Clip marked lines, and Friction split into keyword vs AI evidence**
A new quick action clusters transcript marks and streams them to Studio's clip generator. The Friction evidence table separates keyword scores from AI-cited moments, each filterable on its own.

## v0.14.62 — 2026-07-31 — Composer
**Playback-speed button on the transport bar**
The same 0.5/1/2/3/5× cycle as Screenspace and Transcripts, next to the mute button. The rate survives a part switch, and multi-part playback no longer stalls at a seam.

## v0.14.61 — 2026-07-31 — Overview
**Reports tab with AI mini-report and clip strip**
A manual-trigger thinking agent synthesizes the transcript summary, sheet observations and marked lines into a per-participant report. [M:SS] stamps become playable chips over a clip strip cut from the sheet.

## v0.14.61 — 2026-07-31 — Transcripts
**Local-AI badge on run buttons instead of tabs**
The badge marking local-AI work moves from the Summary and Friction tabs onto the controls that actually start a run, so Citations gets labeled too.

## v0.14.60 — 2026-07-31 — Core
**Remember window size and position, and list off-sheet videos**
The desktop window restores its last size and position, with a Start setting to toggle it and Shift-at-launch to reset. Screenspace and Transcripts also list source videos on disk that no spreadsheet column claims.

## v0.14.59 — 2026-07-31 — Core
**Double-click the top bar to zoom (macOS)**
With the native title bar hidden, double-clicking the topnav triggers the system's zoom or minimize preference, and the window tracks the page theme so light mode no longer flashes white.

## v0.14.58 — 2026-07-31 — Transcripts
**Friction tab filters the transcript itself**
Off / Highlight / Isolate modes turn the tab into a control surface over the segment list, driven by a dual-bound score histogram and category chips. Moment cards collapse into a one-line jump strip.

## v0.14.57 — 2026-07-31 — Transcripts
**Audio-track picker with speech auto-detection**
Multitrack recordings no longer default to the container's first stream: pick a track from the pill dropdown, or let clipgen detect the speech-looking one by stream name.

## v0.14.56 — 2026-07-30 — Transcripts
**Local-AI badge on thinking-agent tabs**
The Summary and Friction tabs carry a glyph marking their results as generated by an on-device thinking agent.

## v0.14.55 — 2026-07-30 — Core
**Cache the Google Drive spreadsheet listing**
The Start picker, worksheet dropdown and open-by-name path each made their own rate-limited Drive round-trip. They now share a five-minute cache, with a Refresh button for staleness.

## v0.14.54 — 2026-07-30 — Core
**Faster Screenspace boot and honest loading states**
Frame-0 warming is a bounded queue instead of one ffmpeg per participant, and Transcripts, Workflows and Screenspace show skeletons during their first fetch instead of a false empty state.

## v0.14.53 — 2026-07-30 — Core
**One modal system across every page**
Every dialog opens and closes through the same animation and backdrop, and toasts fade instead of snapping. Composer's artifact log, previously a divergent copy of Studio's, rides the same path.

## v0.14.50 — 2026-07-29 — Core
**Parallel ffprobe and VideoToolbox hardware encoding**
Reel clips and multi-part sources probe concurrently (852 ms → 202 ms on 20 clips), and a new Video & Clips setting routes encoding through Apple's VideoToolbox: 48.4 s → 12.5 s on a 1080p minute.

## v0.14.48 — 2026-07-29 — Core
**One tab-aware Refresh button per page**
Overview's three Refresh buttons and Studio's per-tab refresh collapse into a single subheader control that acts on the active tab, and its spinner finally runs.

## v0.14.47 — 2026-07-28 — Core
**Preview expected source videos on Start**
Once a worksheet is selected, Start lists the source-video filenames clipgen will look for and marks which are already in the input folder — so a misnamed recording surfaces before the workspace opens.

## v0.14.46 — 2026-07-28 — Core
**The app window drops its title bar (macOS)**
The window controls move into the left end of clipgen's own top bar, and the bar becomes a drag handle. Resizing, snapping and green-button fullscreen stay native; browser launches are unchanged.

## v0.14.45 — 2026-07-28 — Core
**Name your projects, and a denser Recently-opened rail**
Start takes an optional project name that titles its entry in the left rail. Recent entries now fit name, input folder, output folder and spreadsheet on two lines, and twelve are remembered instead of eight.

## v0.14.44 — 2026-07-28 — Core
**16× faster app startup**
The bundle is now a one-dir build, so libraries load from a stable path with the OS page cache intact. Double-click to a usable window drops from ~17.6 s to ~1.1 s.

## v0.14.43 — 2026-07-28 — Core
**Fix: double-clicked app quit immediately on macOS**
A Finder-launched .app inherits no shell PATH, so Homebrew's ffmpeg was invisible and startup aborted with no window at all. The frozen launch now probes the standard bin directories and surfaces fatal errors in a dialog.

## v0.14.42 — 2026-07-27 — Core
**Desktop app: clipgen opens in its own window**
Double-clicking the bundled app opens a native window instead of spawning a Terminal and hijacking the browser. Fonts are vendored so it works offline, and exports route through a native save dialog. The CLI is unchanged.

## v0.14.36 — 2026-07-25 — Screenspace
**Spatial anchor for the magic-wand tolerance scrub**
The wand scrub paints an anchor dot at the press point, a dashed horizontal track, and a head dot that stops growing at slider min/max, with the tolerance readout at the head beside the pointer.

## v0.14.35 — 2026-07-24 — Core
**Audio volume popover with per-track mixing**
Hover the speaker icon for a 0–200% volume slider; multi-track sources get independent per-track sliders mixed in the browser. Wired into Screenspace, Transcripts and Composer.

## v0.14.35 — 2026-07-24 — Core
**Segmented capsule track control**
New .cg-segtrack primitive with a sliding thumb for mutually exclusive options; adopted at Screenspace Color Mode, Text Normalize, and the region tools (rect/lasso/wand).

## v0.14.34 — 2026-07-24 — Core
**Shift+numeral panel focus with arrow navigation**
Shift+numeral targets a panel for keyboard focus while bare numerals pick tools or actions; Screenspace's sidebar, tool, tasks and results panels gain arrow-key navigation.

## v0.14.33 — 2026-07-23 — Core
**Studio and Screenspace keyboard and slider polish**
Empty-queue focus hotkeys pulse a ghost card; Backspace/Delete removes the focused queue card; timeline step buttons seek 5s with Shift for 1s; Set In/Out bind to i/o.

## v0.14.32 — 2026-07-23 — Core
**Keyboard shortcuts across pages and Workflows toolbar polish**
New hotkeys on Studio, Screenspace, Transcripts and Workflows; blueprint rename modal, cleaner toolbar layout, custom [data-tooltip] tooltips, and Composer undo/redo on compact icon buttons.

## v0.14.31 — 2026-07-23 — Core
**Modal keyboard navigation and scoped Alt hints**
Alt-hold hint chips scope to the open modal; the Start launcher and Settings modal gain tab, list and reset hotkeys with focus trapping; the Start overlay is now a real blocking modal.

## v0.14.31 — 2026-07-23 — Core
**Sheet last-edit date and worksheet picker on Start**
Each spreadsheet dropdown entry shows an "Edited …" date, and multi-tab spreadsheets get a worksheet picker that threads through open and persists across reloads.

## v0.14.30 — 2026-07-22 — Composer
**Annotation stroke controls, multi-select, and timeline chrome**
Stroke width and Solid/Dashed/Dotted menus apply to new and selected annotations; shift-click and marquee multi-select move, restyle or delete as a group; the timeline gains lane bands and a step-track ruler.

## v0.14.29 — 2026-07-22 — Composer
**Hideable Timelines sidebar**
The right Timelines panel collapses to a thin strip with persisted state; F toggles the sidebar (matching Studio) and thumbnail strips move to S.

## v0.14.28 — 2026-07-21 — Composer
**Timeline chrome rework, shape annotations, and cut UX**
Transport and controls move above the canvas; rotatable rect/ellipse shape tools with corner and rotation handles; draggable annotation spans; double-click cuts on empty timeline space.

## v0.14.26 — 2026-07-20 — Screenspace
**Attention computational-saliency tool**
A bottom-up saliency composite predicts visual attention without eye-tracking; full-frame scans feed heatmaps and gaze-replay GIFs while timeline events fire only at confirmed attention shifts.

## v0.14.25 — 2026-07-20 — Screenspace
**Grouped tool navigation with numeral hotkeys**
Optional category dropdown chips (Difference, Detection, Classification, Attention, Utility) replace the flat 12-tab row; numeral hotkeys pick tabs or open a category then its tool.

## v0.14.24 — 2026-07-20 — Composer
**Follow the playhead when panning the zoomed timeline**
When zoomed in, seeks and playback pan the viewport minimally to keep the playhead visible; a persisted Follow toggle extends this to playback, while clicks always reveal.

## v0.14.23 — 2026-07-20 — Workflows
**Phase 3 — exports, canvas polish, resume, and trigger chaining**
New transcript and data export nodes and compound filter clauses; two-finger pan, pinch zoom, snap-to-grid and sticky notes on the canvas; resume failed runs; transcript-complete and scan-event triggers chain into new runs.

## v0.14.15 — 2026-07-20 — Composer
**Thumbnail strips and hover audio scrub on timeline markers**
Zoom-adaptive thumbnail tiles on marker bars and cut bands refetch finer frames as you zoom in; hover audio scrub with waveform, toggled with F/W and persisted.

## v0.14.14 — 2026-07-18 — Screenspace
**Colorized Change, SSIM, and Flow model-view previews**
The Change diff preview is JET-colorized with an overlay that tints changed pixels without darkening the frame; Similarity adds an SSIM difference map and score; Flow arrows are color-coded by magnitude.

## v0.14.13 — 2026-07-17 — Core
**Filter and panel commands in the palette**
Filter clears, sidebar and panel toggles, and drawer actions across all six pages are searchable in the command palette, gated so only relevant commands surface.

## v0.14.13 — 2026-07-17 — Studio
**Region-aware keyboard navigation**
Stash and clear hotkeys, filter-sidebar and artifact/reel panel collapse, and a region-aware cursor: 1–5 jump between the filter list, queues and stash lists, with Enter activating each target.

## v0.14.13 — 2026-07-17 — Screenspace
**Panel and tool-tab hotkeys**
Collapse the bottom panel and cycle tool tabs from the keyboard; Alt-hold hint chips dim for disabled controls and combo glyphs render evenly in the cheatsheet.

## v0.14.12 — 2026-07-16 — Core
**Alt-hold shortcut hints**
Hold Alt to see combo chips on tagged controls across every page; Studio shows action hints on the browsed cell or card, and a uniform "?" cheatsheet button replaces per-page help popovers.

## v0.14.11 — 2026-07-16 — Core
**Global command palette**
Cmd+Shift+P or Cmd+K opens a Spotlight-style palette for page navigation, participant jumps, chrome actions and recents across all six hub pages; deep links honor #tab= and #P07 hashes.

## v0.14.10 — 2026-07-16 — Core
**Rebindable hotkeys across all frontends**
A shared registry unifies the defaults and auto-generates the ? cheatsheet; Settings → Hotkeys lets you rebind, with conflicts resolved per binding and overrides persisted across sessions.

## v0.14.9 — 2026-07-15 — Studio
**Content-aware tooltips**
Queue action buttons and intake controls show tooltips reflecting card count, selected format and state instead of stale nudges.

## v0.14.8 — 2026-07-15 — Core
**Composer lane in Convergence and Metadata search**
Convergence adds a per-participant Composer swim lane from cut pairs, and Metadata gains a search box that highlights matches and scrolls to the target row.

## v0.14.7 — 2026-07-13 — Core
**Five new Map visualizations**
Color-by choropleth, shift-click pairwise compare arcs, direct-axes 5D scatter, session trajectories with replay comets, and auto-labeled cluster hulls on the Overview Map tab.

## v0.14.2 — 2026-07-13 — Core
**Annotated exports across recording parts**
Composer burn and GIF exports stitch spans that cross a multi-part boundary into one continuous clip before the overlay pass, so annotations render correctly across seams.

## v0.14.0 — 2026-07-13 — Core
**Overview page with 3D similarity Map**
A new Overview tab gathers cohort-level lenses: a 3D Map positions participants by PCA over sheet timestamps, transcript marks and Screenspace events; Metadata and Convergence move here from Studio.

## v0.13.61 — 2026-07-12 — Core
**Composer — source-video cutting and annotations**
A Composer page cuts source video with named in/out pairs, non-destructive marker trims, and canvas annotations; cuts and trims feed Studio's Artifact and Reel queues.

## v0.13.61 — 2026-07-12 — Transcripts
**Friction before Summary**
Deterministic friction scores populate the heatmap, timeline band and stat chips immediately; LLM-refined moments still require Summary.

## v0.13.60 — 2026-07-10 — Screenspace
**Live magic-wand tolerance scrub**
Press-drag-release on the magic wand: horizontal drag scrubs flood-fill tolerance with a live contour preview, release commits or applies a Shift/Alt boolean combine, Escape cancels.

## v0.13.59 — 2026-07-10 — Screenspace
**Boolean edits on unsaved canvas regions**
Shift/Alt/Shift+Alt add, subtract and intersect now target the pending region drawn on the video before it is saved — refine a rough shape in place with no server round-trip.

## v0.13.59 — 2026-07-10 — Screenspace
**Auto-generated task and event names**
New tasks get descriptive names from their params (e.g. Text "checkout" · header) instead of generic type: region labels, shown on task cards, the run pill, results and timeline tooltips.

## v0.13.58 — 2026-07-10 — Screenspace
**Boolean region editing**
Combine shaped regions with Photoshop-style modifiers (Shift add, Alt subtract, Shift+Alt intersect) or merge shift-selected regions; multi-contour shapes persist when the result is not axis-aligned.

## v0.13.57 — 2026-07-09 — Transcripts
**Optional severity on flagged segments**
Marks gain a severity dropdown, a colored dot on segment rows and a mark_severities export column; Studio intake filters by severity and Metadata charts the distribution.

## v0.13.57 — 2026-07-09 — Transcripts
**Transcription progress on the timeline**
While a participant is being transcribed, a faint dot texture covers the un-transcribed portion of the timeline and wipes away left-to-right in sync with decode progress.

## v0.13.57 — 2026-07-09 — Studio
**Screenspace clusters in Metadata**
Metadata counts time-adjacent Screenspace event clusters instead of raw per-second events, so dense scans do not dominate tables and charts. Toggle it in settings.

## v0.13.56 — 2026-07-09 — Workflows
**Clip padding and max duration on artifact nodes**
Make Clips and Build Reel nodes expose pad start, pad end and max duration: nudge clip boundaries or cap segment length without new queue cards.

## v0.13.55 — 2026-07-09 — Screenspace
**Lasso and magic-wand region selectors**
Draw freehand lasso or flood-fill magic-wand regions alongside rectangles; shaped polygons rasterize to masks for the color, change, flow, OCR, template and scene tools.

## v0.13.55 — 2026-07-09 — Workflows
**Live minimap and zoom controls**
The corner minimap mirrors the canvas view in real time with an edge-safe view frame, alongside floating zoom-in, zoom-out and fit-to-content buttons.

## v0.13.54 — 2026-07-09 — Core
**CLI friction agent and --settings**
Run the friction thinking agent from the command line with --friction, and open the interactive settings editor with --settings before a run.

## v0.13.53 — 2026-07-06 — Transcripts
**Streaming AI summaries**
Summary generation streams tokens into the Transcripts panel as the model produces them instead of showing a spinner for the whole run.

## v0.13.52 — 2026-07-06 — Transcripts
**Faster Whisper transcription**
VAD on by default skips silence without clipping quiet speech; beam size 2 and configurable CPU threads speed decode on many-core machines. All tunable in Studio → Transcription settings.

## v0.13.51 — 2026-07-04 — Core
**Unified motion for toasts and Studio overlays**
Toast and Studio overlay cards animate through the shared motion engine with fade/pop entrances, and overlays no longer snap in when revealed from hidden.

## v0.13.50 — 2026-07-03 — Core
**Animated stash and delete exits**
Screenspace region pills and Studio queue, artifact and reel cards play staged exit animations on stash, delete and clear before the list re-renders, with reduced-motion fallbacks.

## v0.13.49 — 2026-06-30 — Workflows
**Multi-select participant batch**
The Video Source participant param is a checkbox popover, so a batch can fan out over any subset instead of only one participant or all.

## v0.13.48 — 2026-06-30 — Workflows
**Canvas navigation**
Fit-to-view button and F shortcut frame all nodes; dragging near the canvas edge auto-pans; a corner minimap shows the graph with click/drag-to-recenter.

## v0.13.47 — 2026-06-29 — Screenspace
**Boundary flags above the timeline**
Scene boundaries render as flag glyphs in a rail above the timeline instead of in-band ticks; hover shows a tooltip and locator hairline, click seeks to the boundary.

## v0.13.46 — 2026-06-29 — Transcripts
**Editable thinking-agent prompts**
View, edit and reset the summary, citations and friction prompts in Settings → Summaries; edits persist and apply on the next run, with placeholder validation.

## v0.13.45 — 2026-06-28 — Workflows
**Canvas discoverability and run panel polish**
Toolbar undo/redo, autosave status and a shortcuts legend; the run panel gains per-node status icons, durations, result filters and a Reconnecting pill when the stream drops.

## v0.13.43 — 2026-06-28 — Studio
**Feedback, persistence, and accessibility polish**
Stash saves animate the new card in, generation failures show per-cell reasons, sidebar filter selections persist across reload, and modals gain focus traps and focus-visible rings.

## v0.13.43 — 2026-06-28 — Transcripts
**Keyboard review loop and AI-trust cues**
j/k/arrows move and seek segments, m marks, 1–6 set friction category, n/p jump between marks; the friction panel warns when segments were edited after AI analysis.

## v0.13.43 — 2026-06-28 — Screenspace
**Virtualized results and clearer task feedback**
Large result lists lazy-render in chunks, loading indicators cover fetch and participant switch, and failed-task errors expand on click.

## v0.13.42 — 2026-06-28 — Workflows
**Run notes, skip fixes, and batch efficiency**
Completed nodes surface run notes when output is degraded, validation warns on unwired or mismatched filter values, batch runs reuse participant-independent sheet data, and an empty canvas offers built-in recipes.

## v0.13.41 — 2026-06-27 — Workflows
**Detect node, interval captures, and editor power tools**
A unified Detect node replaces the per-detector types and Interval Captures samples screenshots or GIFs across a range; blueprint JSON import/export, copy/paste, mute, undo/redo, and a Run split-button with Run to here.

## v0.13.27 — 2026-06-27 — Workflows
**Collection control nodes**
Filter, partition, merge, limit and dedup nodes thin or combine the collections flowing through a graph: gate clip selections before Make Clips, or branch matched vs. unmatched streams.

## v0.13.26 — 2026-06-27 — Workflows
**Watch-dir auto-run**
Arm one blueprint to run automatically when a new participant video lands in the input directory; triggered runs show a ⚡ badge in run history.

## v0.13.22 — 2026-06-27 — Workflows
**Pre-run validation and inspectable results**
An Issues panel blocks Run on wiring or param errors, completed nodes expose lazy-loaded result sidecars, and Re-run replays a finished graph.

## v0.13.20 — 2026-06-26 — Workflows
**Expanded catalog, batch runs, and stashes**
Per-detector Screenspace nodes, highlights selector, multitool/timelapse/heatmap/measure, and adapter-aware dashed wires; Video Source "All participants" fans out a whole study, and sub-graphs save as named stashes.

## v0.13.16 — 2026-06-25 — Workflows
**Workflows mode: node canvas and run engine**
A fourth top-level tab chains clip, Screenspace and transcript actions on an infinite pan/zoom canvas: drag nodes from a catalog, wire typed ports, edit params and Run with live per-node progress.

## v0.13.11 — 2026-06-23 — Core
**Card scrubber on hover**
Sweep a queue or viewer card thumbnail to scrub frames with audio and a waveform playhead; toggle it on in Studio settings or the exported timeline viewer header.

## v0.13.10 — 2026-06-23 — Screenspace
**Scene-aware boundary segmentation**
The Boundary detector gains hybrid scene metrics, hierarchical Scene A1/B2 labels and a post-run consolidation pass; boundaries surface in results, Studio intake, Convergence, Metadata and the viewer.

## v0.13.6 — 2026-06-22 — Studio
**Clip-length intake timeline markers**
Screenspace and Transcript intake density timelines size each marker by its clip span, so longer selections read wider at a glance.

## v0.13.5 — 2026-06-22 — Transcripts
**Model install consent and dynamic Ollama pickers**
Whisper and Ollama models require explicit confirmation before downloading; the pickers list installed Ollama models, friction can use a separate model, and pull progress shows in-app.

## v0.13.3 — 2026-06-22 — Studio
**Titlecard and endcard background picker**
Choose a default, solid color, uploaded image, or no endcard from Settings → Video & Clips, with a live preview. Selections persist and are baked into generated clips and reels.

## v0.13.2 — 2026-06-22 — Screenspace
**Rolling-window and change heatmaps**
Template tasks gain a rolling-window animation alongside static and cumulative views, and Change tasks get full heatmaps, with a per-tool toggle and a collapsible thumbnail strip in results.

## v0.13.1 — 2026-06-21 — Transcripts
**Cancel summary and citations**
An inline Cancel button stops summary generation or the citations pass mid-run, without leaving the tab.

## v0.13.0 — 2026-06-20 — Core
**Multiple source videos per participant**
A session can span several videos declared in the spreadsheet Filename row or auto-detected on disk; timestamps, clips, transcripts and Screenspace events map across the full continuous timeline.

## v0.12.9 — 2026-06-20 — Studio
**Sortable Sheet Preview columns**
Cycle #, Category, Severity and Function headers through Ascending → Descending → Off; severity sorts most-severe-first with empty values at the bottom.

## v0.12.9 — 2026-06-20 — Screenspace
**Color presence detection mode**
The Color tool can fire when a target color appears anywhere in the region, with a Min area % control and presence-aware calibration — standalone, in Multitool steps, or from the CLI.

## v0.12.8 — 2026-06-20 — Studio
**Source times and severity tint on queue cards**
Every artifact and reel card shows its source start–stop time, and spreadsheet-sourced cards tint their caption by row severity.

## v0.12.8 — 2026-06-20 — Core
**Elapsed time and ETA on long operations**
Screenspace tasks and transcription show elapsed plus an estimated time remaining; Studio builds and thinking agents show elapsed only. Clocks survive a page reload.

## v0.12.7 — 2026-06-20 — Screenspace
**CLI scene analysis and headless task re-run**
Run scene analysis with --ss-task scene and re-run any saved manifest task headlessly with --ss-run-task, the path for unattended multitool chains.

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
Pin reference frames on the timeline, score detector sensitivity against them, get a suggested threshold and apply it, with a calibration strip and integration into task creation.

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
