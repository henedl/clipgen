# Changelog

Notable changes per release. One level-2 heading per version — `## <version> — <YYYY-MM-DD>`, with whitespace around the dash — then one line per change:

```
**<Tool>:** <Feat|Fix>: <what you can now do, in one plain sentence>
```

`<Tool>` is one of `Core`, `Studio`, `Screenspace`, `Transcripts`, `Workflows`, `Composer`, `Overview`. Write for someone using clipgen, not building it: name the thing by what it is called in the interface, say what changed for them, and leave the internals in the commit. Add a second sentence only when the first leaves an obvious "so what". Lines render as **plain text** in the Start overlay's Recent updates tab — no backticks or markdown inside them, or it shows up literally.

## v0.17.0 — 2026-09-03
**Screenspace:** Feat: Click a Multitool step to focus it. Model view previews that step's tool, region and reference, and its calibration track is highlighted.
**Screenspace:** Feat: Template and Shape scans run much faster when the tool has a search region.
**Screenspace:** Feat: Heatmap GIFs render several times faster, most of all on Attention and Flow scans.
**Studio:** Feat: Artifact Log rows have a Show on disk button in the desktop app.
**Studio:** Feat: The quick actions are now called Build Timeline and Build Gallery.
**Core:** Feat: Clips with title cards build faster, since the card is drawn once instead of once per frame.
**Core:** Feat: The Start window highlights the version you are running, and Settings, Hotkeys shows larger key caps in denser rows.
**Core:** Fix: The desktop window takes clicks right after launch instead of needing an app switch first.
**Core:** Fix: One failing clip no longer stops the rest of a batch.
**Core:** Fix: Regenerated GIFs keep their set length, and empty timestamp ranges are skipped instead of cut.
**Core:** Fix: A background action that fails now tells you, rather than looking like it worked.
**Core:** Fix: Temporary files from gallery and title card builds are cleaned up.
**Screenspace:** Fix: Multitool and Boundary previews render again, and a template that cannot be matched is reported the same with or without a search region.
**Transcripts:** Fix: Clipping marked lines picks up the text of every mark.

## v0.16.13 — 2026-08-28
**Screenspace:** Feat: Model view and Calibration now live in a Preview tab on the right, open by default so you can see what the tool is matching without hunting in the sidebar.
**Screenspace:** Feat: Shape has a paint-brush draw mode on the video canvas, separate Width and Height scale labels when axes are unlinked, and an eraser editor on sample thumbnails.
**Screenspace:** Feat: Hold or tap B to peek the results overlay, tap to latch it on, and N cycles which overlay layer is shown.
**Screenspace:** Fix: The timeline ruler and playhead stay sharp when the side panel collapses or you switch tabs.

## v0.16.12 — 2026-08-27
**Screenspace:** Feat: The Shape tool finds icons, logos, and buttons that changed color, theme, or size, where Template pixel matching misses them.
**Screenspace:** Feat: Tool categories in the grouped nav show their name and icon instead of bare tool names in a dropdown.
**Screenspace:** Feat: Template, Shape, Similarity, and Scene reference rows show a crop thumbnail and source region name.
**Screenspace:** Feat: Dense Template matches scan much faster instead of freezing the page.
**Transcripts:** Feat: The Friction tab stops re-scoring the whole transcript on every poll while an AI summary is running.
**Core:** Feat: Saving project state skips unchanged sections, so large manifests no longer stall every auto-save.
**Transcripts:** Fix: Dictionary CSV import handles Excel formula prefixes and a UTF-8 byte-order mark.
**Transcripts:** Fix: In and out markers stay with the study where you set them.
**Core:** Fix: Custom source video filename patterns apply in the CLI and Workflows, not only in the web UI.
**Screenspace:** Fix: Lasso masks on Shape and Template samples are kept through preview, re-run, and Workflows.

## v0.16.11 — 2026-08-24
**Transcripts:** Fix: A failed AI run now names what went wrong instead of saying it produced no result.
**Transcripts:** Fix: Summary citations point at the sentence they belong to again, and a summary that ends in a paragraph no longer folds it into the bullet list.
**Transcripts:** Fix: Friction scoring reads your corrected text, so the highlighting stops shifting once a summary runs.
**Transcripts:** Fix: Opening a different study no longer shows the previous study's transcript for a participant with the same name.
**Core:** Fix: A reel that failed to build no longer reports that it created one.

## v0.16.10 — 2026-08-23
**Transcripts:** Feat: Corrections is now the Transcript Dictionary. Known terms steer the transcription up front, rather than fixing names after you have seen them go wrong.
**Transcripts:** Feat: Take a dictionary between studies — export it as CSV, import one, or keep a saved copy any study can load.
**Core:** Feat: AI models are listed by name with a link to their model card, instead of by file id.

## v0.16.9 — 2026-08-23
**Core:** Feat: Settings, Summaries suggests small AI models and downloads them in place, with progress in the row.

## v0.16.8 — 2026-08-23
**Core:** Feat: The local AI server starts on its own when a summary needs it, instead of asking you to start it first.
**Core:** Feat: Each downloaded AI model can be shown on disk or deleted from Settings, Summaries.
**Core:** Fix: An AI failure is reported in the app rather than only in the terminal.

## v0.16.6 — 2026-08-23
**Core:** Feat: The local AI runtime ships inside clipgen, and models you already downloaded through llama.cpp, Hugging Face or Ollama are reused instead of fetched again.

## v0.16.5 — 2026-08-23
**Core:** Feat: Local AI runs on llama.cpp. Models pulled by the old runtime do not carry over, so download the one you want again.

## v0.16.4 — 2026-08-22
**Core:** Feat: All of a project's saved work now lives in one clipgen.json in the output folder. Files written by an earlier version are left alone and not read.

## v0.16.3 — 2026-08-22
**Core:** Feat: Settings are kept per user, so switching projects no longer switches your preferences. The settings file can be opened from the modal.
**Core:** Feat: The startup screen draws the clipgen mark instead of a spinner.

## v0.16.2 — 2026-08-22
**Core:** Feat: An exported viewer says which clipgen version wrote it and links back to the project.
**Core:** Feat: Cross-references moved into Settings, General, and one switch now covers every badge and hover card.

## v0.16.1 — 2026-08-21
**Core:** Feat: The desktop app shows its window immediately and connects to your spreadsheet behind the startup screen, instead of waiting with nothing on screen.

## v0.16.0 — 2026-08-20
**Core:** Feat: Choose how your source video files are named. A pattern under Settings → Video & Clips with {study} and {participant} placeholders replaces the fixed {study}_{participant}.mp4, so clipgen can find recordings you already named your own way.

## v0.15.18 — 2026-08-20
**Workflows:** Feat: The Region node picks from your saved Screenspace regions instead of asking you to type the name.
**Workflows:** Fix: A Run that is refused now says why instead of doing nothing.
**Transcripts:** Fix: Switching to a different spreadsheet no longer lets the old study's transcripts and summaries land in the new one.
**Core:** Fix: A damaged settings file now warns you instead of quietly forgetting your saved filename overrides.

## v0.15.16 — 2026-08-20
**Transcripts:** Feat: Transcribe only part of a recording. Set in and out markers with I and O.
**Transcripts:** Feat: Normalize Audio quick action to even out loudness across videos.
**Composer:** Fix: Burn and GIF exports can be cancelled, and text you were typing is kept instead of thrown away when you click elsewhere.

## v0.15.15 — 2026-08-19
**Transcripts:** Feat: Transcript timestamps now sit tight against the words instead of starting early in the silence, and the line being spoken highlights word by word.
**Core:** Feat: The desktop app has a proper Mac menu bar — File, Go, Window and Help, with ⌘1–⌘6 for the six pages.

## v0.15.14 — 2026-08-19
**Core:** Fix: The last observation of a session produced no clip. A timestamp in the final minute ran past the end of the recording and the whole clip was thrown out; it is now shortened to the footage that exists.
**Screenspace:** Feat: Colour detection got about 7× faster per frame, and slightly more accurate.

## v0.15.13 — 2026-08-18
**Composer:** Feat: The annotation tools moved into a vertical rail beside the video, with two colour slots you swap with X and a text-size control. Delete is now Backspace.
**Core:** Fix: Dropdowns look like the rest of the app instead of raw system widgets — most visibly in Safari and the desktop window, where they had no styling at all.

## v0.15.12 — 2026-08-17
**Core:** Feat: Status text that means "still working" now shimmers, so you can tell it apart from text that has stopped updating.

## v0.15.11 — 2026-08-17
**Core:** Feat: Fix a mis-named recording without leaving the app. Each source-video row on the Start screen can be pointed at a different file, and it sticks.
**Core:** Feat: The About tab lists every piece of third-party software in the build with its licence.
**Screenspace:** Fix: Picking two languages that cannot be read together is now refused when you create the task, rather than failing during the scan.
**Core:** Fix: Clips cut from some recordings started at the wrong moment; seeking is now correct for those files.
**Core:** Fix: Upgrading the Windows install no longer leaves files from the previous version behind.

## v0.15.10 — 2026-08-17
**Core:** Feat: The macOS download is about 230 MB smaller, and is no longer covered by the GPL.
**Core:** Feat: The Start screen's right side is now three tabs — Open, About and Recent updates — so the Open button cannot scroll out of reach.

## v0.15.9 — 2026-08-16
**Screenspace:** Feat: Text and number detection moved to a lighter reading engine. Same languages, far smaller download.

## v0.15.8 — 2026-08-16
**Core:** Feat: Windows gets a real installer with Start-menu and uninstall entries, instead of a zip you unpack yourself.

## v0.15.7 — 2026-08-16
**Core:** Feat: Run clipgen with --profile to see where time actually goes. Six slow spots it found have already been fixed, across scans, heatmaps and video seeking.

## v0.15.4 — 2026-08-15
**Core:** Feat: Scrollbars and other system-drawn controls follow the app's light or dark theme instead of the operating system's.

## v0.15.3 — 2026-08-15
**Core:** Feat: --version prints the version number and --licenses prints the third-party notices, both from the command line.

## v0.15.1 — 2026-08-13
**Transcripts:** Feat: Windows can install Ollama from inside the app, the same as macOS — no terminal, no admin prompt.

## v0.15.0 — 2026-08-13
**Core:** Fix: The installed desktop app could not transcribe at all. A file Whisper needs to skip silence was missing from the build, so every transcription failed immediately.

## v0.14.75 — 2026-08-09
**Core:** Feat: Starting up shows a progress page naming what it is doing, instead of a blank window, and transcription now tells you when it is loading the model rather than just waiting.

## v0.14.74 — 2026-08-08
**Transcripts:** Feat: The two subtitle actions became one Embed Subtitles dialog where you pick how many videos to do. Progress shows per file, and one failure no longer stops the rest.

## v0.14.73 — 2026-08-04
**Studio:** Feat: Cut clips straight from a MindNode mind map — no spreadsheet needed. Timestamps in notes work the same way they do in a sheet.

## v0.14.72 — 2026-08-04
**Transcripts:** Feat: If Ollama is missing, the app offers to download and start it for you, after asking.

## v0.14.71 — 2026-08-04
**Core:** Feat: The desktop downloads include ffmpeg, so you no longer have to install it separately. This also fixes title cards failing to render.

## v0.14.70 — 2026-08-03
**Screenspace:** Feat: Animated heatmaps sit still until you hover them, and scrub through as you move across. The full animation only loads when you press play.

## v0.14.69 — 2026-08-03
**Transcripts:** Feat: Transcribe All queues every participant that has a video and no transcript yet, keeping each one's model and language settings.

## v0.14.68 — 2026-08-02
**Core:** Fix: When ffmpeg, Google credentials or Ollama were missing, the explanation went to a console the desktop app does not have. It now appears in the interface, with instructions for your platform.

## v0.14.67 — 2026-08-01
**Overview:** Feat: The unfinished 3D Map tab is gone. Overview is now Metadata, Convergence and Reports.

## v0.14.66 — 2026-08-01
**Core:** Feat: Recordings made with OBS's fragmented option cannot be scrubbed in a browser — seeking lands in the wrong place until the whole file downloads. Clipgen now spots them and offers a one-click repair that does not re-encode.

## v0.14.65 — 2026-08-01
**Transcripts:** Feat: Clip Marked Lines turns your marked transcript lines into clips in one step.
**Transcripts:** Feat: The Friction tab separates what the keyword scoring found from what the AI cited, so you can weigh them apart.

## v0.14.62 — 2026-07-31
**Composer:** Feat: A playback-speed button, 0.5× to 5×, matching Screenspace and Transcripts. Speed now survives crossing between recording parts.

## v0.14.61 — 2026-07-31
**Overview:** Feat: A Reports tab writes a short per-participant report from the transcript summary, your sheet observations and your marked lines. Timestamps in it are clickable and play the moment.
**Transcripts:** Feat: The badge marking AI-generated results moved onto the buttons that start a run, so it is clear before you press rather than after.

## v0.14.60 — 2026-07-31
**Core:** Feat: The desktop window reopens at the size and position you left it. Hold Shift while launching to reset.
**Core:** Feat: Screenspace and Transcripts also list recordings on disk that no spreadsheet column claims, so a stray file is visible instead of silently ignored.

## v0.14.59 — 2026-07-31
**Core:** Feat: Double-clicking the top bar zooms or minimises the window, following your Mac's setting.

## v0.14.58 — 2026-07-31
**Transcripts:** Feat: The Friction tab now filters the transcript itself — highlight the struggling passages in place, or hide everything else — instead of listing them separately.

## v0.14.57 — 2026-07-31
**Transcripts:** Feat: Pick which audio track to transcribe when a recording has several, or let clipgen find the one with speech on it.

## v0.14.56 — 2026-07-30
**Transcripts:** Feat: The Summary and Friction tabs are marked as AI-generated.

## v0.14.55 — 2026-07-30
**Core:** Fix: Opening a Google spreadsheet made three separate rate-limited lookups and could stall on a slow connection. It now looks once and remembers for five minutes, with a Refresh button.

## v0.14.54 — 2026-07-30
**Core:** Feat: Screenspace opens faster with many participants, and Transcripts, Workflows and Screenspace show a loading placeholder instead of briefly claiming there is nothing there.

## v0.14.53 — 2026-07-30
**Core:** Feat: Every dialog in the app now opens, closes and fades the same way.

## v0.14.50 — 2026-07-29
**Core:** Feat: Optional hardware video encoding on Macs, under Settings → Video & Clips. A minute of 1080p re-encodes in 12 seconds instead of 48, at the cost of a larger file.
**Core:** Feat: Building a reel inspects its clips in parallel, cutting the wait before encoding starts.

## v0.14.48 — 2026-07-29
**Core:** Feat: One Refresh button per page that acts on whichever tab you are looking at, replacing four scattered ones.

## v0.14.47 — 2026-07-28
**Core:** Feat: After picking a worksheet, the Start screen lists the recordings clipgen expects and ticks off the ones it found — so a misnamed file shows up before you start work, not as missing clips afterwards.

## v0.14.46 — 2026-07-28
**Core:** Feat: The Mac app drops the separate title bar; the window buttons sit in clipgen's own top bar, which you can drag to move the window.

## v0.14.45 — 2026-07-28
**Core:** Feat: Name your projects. The name titles that project in the Recently opened list, which now also shows both folders and the spreadsheet, and remembers twelve instead of eight.

## v0.14.44 — 2026-07-28
**Core:** Feat: The app starts about 16× faster — roughly one second from double-click to a usable window, down from eighteen.

## v0.14.43 — 2026-07-28
**Core:** Fix: Double-clicking the Mac app quit instantly with no window and no message, on machines where ffmpeg came from Homebrew. It now finds ffmpeg, and anything fatal before the window opens is shown in a dialog.

## v0.14.42 — 2026-07-27
**Core:** Feat: Clipgen opens in its own window instead of launching a Terminal and taking over your browser. Fonts are built in, so it works offline. The command line is unchanged.

## v0.14.36 — 2026-07-25
**Screenspace:** Feat: The magic-wand tolerance drag shows where you started and how far you have gone, with the value next to the pointer.

## v0.14.35 — 2026-07-24
**Core:** Feat: A volume control on the speaker icon, up to 200%, with separate sliders per audio track on multi-track recordings.
**Core:** Feat: A sliding capsule control replaces stacked buttons where only one option can be active.

## v0.14.34 — 2026-07-24
**Core:** Feat: Shift plus a number jumps keyboard focus to a panel; a plain number picks a tool. Arrow keys then move within the panel.

## v0.14.33 — 2026-07-23
**Studio:** Feat: Backspace removes the focused card, timeline step buttons seek five seconds (one with Shift), and I and O set in and out points.

## v0.14.32 — 2026-07-23
**Core:** Feat: Keyboard shortcuts across Studio, Screenspace, Transcripts and Workflows, plus tooltips that work on every control.

## v0.14.31 — 2026-07-23
**Core:** Feat: The Start screen and Settings are fully keyboard-driven, and holding Alt shows the shortcuts for whatever is open.
**Core:** Feat: Spreadsheets show when they were last edited, and multi-tab files let you pick the worksheet before opening.

## v0.14.30 — 2026-07-22
**Composer:** Feat: Annotations get stroke width and solid, dashed or dotted styles, and can be selected as a group to move, restyle or delete together.

## v0.14.29 — 2026-07-22
**Composer:** Feat: The Timelines panel folds away with F, leaving the video the full width.

## v0.14.28 — 2026-07-21
**Composer:** Feat: Rectangle and ellipse annotation shapes that rotate, cuts from double-clicking the timeline, and numbered markers so you can see the order at a glance.

## v0.14.26 — 2026-07-20
**Screenspace:** Feat: An Attention tool that predicts where a viewer's eye goes, without eye-tracking hardware, and replays it as a heatmap.

## v0.14.25 — 2026-07-20
**Screenspace:** Feat: The twelve tools group into categories instead of one long tab row, reachable by number keys.

## v0.14.24 — 2026-07-20
**Composer:** Feat: When zoomed in, the timeline follows the playhead instead of leaving it off-screen.

## v0.14.23 — 2026-07-20
**Workflows:** Feat: Export nodes for transcripts and data, two-finger pan and pinch zoom, sticky notes, and the ability to resume a failed run instead of restarting it.
**Workflows:** Feat: A finished transcript or scan can now trigger the next workflow automatically.

## v0.14.15 — 2026-07-20
**Composer:** Feat: Timeline markers show thumbnails that sharpen as you zoom, and hovering scrubs the audio with a waveform.

## v0.14.14 — 2026-07-18
**Screenspace:** Feat: The Change, Similarity and Flow previews are colour-coded, so you can see what the detector is reacting to.

## v0.14.13 — 2026-07-17
**Core:** Feat: The command palette reaches filters, panels and drawers on every page, not just navigation.
**Studio:** Feat: Number keys 1–5 jump between the filter list, the queues and the stashes, and Enter acts on where you are.
**Screenspace:** Feat: Collapse the bottom panel and cycle tool tabs from the keyboard.

## v0.14.12 — 2026-07-16
**Core:** Feat: Hold Alt to see the keyboard shortcut for anything on screen, and press ? for the full list.

## v0.14.11 — 2026-07-16
**Core:** Feat: A command palette on Cmd+K for jumping between pages, participants and recent projects.

## v0.14.10 — 2026-07-16
**Core:** Feat: Keyboard shortcuts are consistent across all pages and can be rebound under Settings → Hotkeys.

## v0.14.9 — 2026-07-15
**Studio:** Feat: Queue buttons explain what they will actually do with the cards you have, instead of showing fixed text.

## v0.14.8 — 2026-07-15
**Overview:** Feat: Composer cuts appear as their own lane in Convergence, and Metadata gains a search box.

## v0.14.7 — 2026-07-13
**Overview:** Feat: Five more ways to read the Map — colour by value, compare two participants, plot your own axes, replay a session's path, and label clusters.

## v0.14.2 — 2026-07-13
**Composer:** Fix: Annotated exports that crossed from one recording part into the next came out wrong. They are now joined before the annotations are drawn.

## v0.14.0 — 2026-07-13
**Overview:** Feat: A new Overview page for looking across participants rather than at one. Metadata and Convergence move here from Studio, joined by a 3D map that places participants by how similar their sessions were.

## v0.13.61 — 2026-07-12
**Composer:** Feat: A new Composer page for cutting straight from the source video: name your in and out points, adjust markers without touching the originals, and draw annotations on the frame. Cuts feed into Studio's queues.
**Transcripts:** Feat: Friction scores appear immediately instead of waiting for the AI summary.

## v0.13.60 — 2026-07-10
**Screenspace:** Feat: Drag sideways on the magic wand to widen or narrow what it selects, with a live preview.

## v0.13.59 — 2026-07-10
**Screenspace:** Feat: Combine regions with Shift, Alt and Shift+Alt before saving them, so a rough shape can be refined in place.
**Screenspace:** Feat: Tasks name themselves from what they do, instead of "type: region".

## v0.13.58 — 2026-07-10
**Screenspace:** Feat: Add, subtract and intersect saved regions with the usual modifier keys.

## v0.13.57 — 2026-07-09
**Transcripts:** Feat: Marked lines can carry a severity, which colours the row, exports with the data and filters in Studio.
**Transcripts:** Feat: While a recording is being transcribed, the untranscribed part of the timeline is shaded and clears as it goes.
**Studio:** Feat: Metadata counts clusters of Screenspace events rather than every single one, so a dense scan no longer swamps the charts.

## v0.13.56 — 2026-07-09
**Workflows:** Feat: Clip nodes can pad their start and end or cap how long a clip may run.

## v0.13.55 — 2026-07-09
**Screenspace:** Feat: Draw freehand or flood-fill regions, not just rectangles. Every tool that took a rectangle takes these too.
**Workflows:** Feat: A live minimap and zoom controls for finding your way around a large graph.

## v0.13.54 — 2026-07-09
**Core:** Feat: Run the friction pass with --friction, and edit settings before a run with --settings.

## v0.13.53 — 2026-07-06
**Transcripts:** Feat: Summaries appear word by word as the model writes them, instead of a spinner until it finishes.

## v0.13.52 — 2026-07-06
**Transcripts:** Feat: Transcription is considerably faster: silence is skipped and more of your processor is used. Tunable under Settings → Transcription.

## v0.13.51 — 2026-07-04
**Core:** Feat: Toasts and overlay cards share one animation, and no longer snap into place.

## v0.13.50 — 2026-07-03
**Core:** Feat: Cards and pills animate out when stashed or deleted, so you can see what left.

## v0.13.49 — 2026-06-30
**Workflows:** Feat: Pick any set of participants for a batch run, not just one or all of them.

## v0.13.48 — 2026-06-30
**Workflows:** Feat: Press F to frame the whole graph, drag near an edge to pan, and use the corner minimap to jump.

## v0.13.47 — 2026-06-29
**Screenspace:** Feat: Scene boundaries show as flags above the timeline instead of ticks inside it; click one to jump there.

## v0.13.46 — 2026-06-29
**Transcripts:** Feat: Read and edit the prompts the AI uses, under Settings → Summaries.

## v0.13.45 — 2026-06-28
**Workflows:** Feat: Undo and redo in the toolbar, a visible autosave state, and per-node progress with durations while a run is going.

## v0.13.43 — 2026-06-28
**Studio:** Feat: Failed clips say why, per cell, and your sidebar filters survive a reload.
**Transcripts:** Feat: Review by keyboard — move and seek with j/k, mark with m, set a category with 1–6, and jump between marks with n/p.
**Screenspace:** Feat: Long result lists load in chunks instead of freezing the page, and failed tasks show their error when clicked.

## v0.13.42 — 2026-06-28
**Workflows:** Feat: Nodes report when their output came out degraded rather than just succeeding, and an empty canvas offers ready-made recipes.

## v0.13.41 — 2026-06-27
**Workflows:** Feat: One Detect node covers every detector, Interval Captures samples screenshots across a range, and graphs can be exported, imported, copied and undone.

## v0.13.27 — 2026-06-27
**Workflows:** Feat: Filter, merge, limit and dedup nodes to thin down what flows through a graph before the expensive steps.

## v0.13.26 — 2026-06-27
**Workflows:** Feat: A workflow can run itself when a new participant video appears in the input folder.

## v0.13.22 — 2026-06-27
**Workflows:** Feat: Problems are caught before a run starts rather than partway through, and finished nodes can be opened to see what they produced.

## v0.13.20 — 2026-06-26
**Workflows:** Feat: Many more node types, and a Video Source that fans a graph out across a whole study.
**Workflows:** Feat: Save part of a graph as a stash to reuse, or start from a built-in recipe.

## v0.13.16 — 2026-06-25
**Workflows:** Feat: A new Workflows page that chains clip, Screenspace and transcript steps into a graph you draw, then runs it with live progress.

## v0.13.11 — 2026-06-23
**Core:** Feat: Sweep across a card's thumbnail to scrub through the clip, with audio and a waveform.

## v0.13.10 — 2026-06-23
**Screenspace:** Feat: Scene detection labels boundaries hierarchically, and they carry through to Studio, Convergence, Metadata and the viewer.

## v0.13.6 — 2026-06-22
**Studio:** Feat: Intake timeline markers are as wide as the clip they represent, so long selections read as long.

## v0.13.5 — 2026-06-22
**Transcripts:** Feat: Downloading an AI model now asks first and shows progress, and the model pickers list what you actually have installed.

## v0.13.3 — 2026-06-22
**Studio:** Feat: Choose the title and end card background — a colour, your own image, or none — under Settings → Video & Clips.

## v0.13.2 — 2026-06-22
**Screenspace:** Feat: Heatmaps for the Change tool, and a rolling-window animation for Template, shown as a thumbnail strip in results.

## v0.13.1 — 2026-06-21
**Transcripts:** Feat: Cancel a summary partway through without leaving the tab.

## v0.13.0 — 2026-06-20
**Core:** Feat: A session can span several video files. Timestamps, clips, transcripts and detections all map onto one continuous timeline across them.

## v0.12.9 — 2026-06-20
**Studio:** Feat: Sort the Sheet Preview by number, category, severity or function.
**Screenspace:** Feat: The Colour tool can fire when a colour appears anywhere in the region, with a minimum-area threshold.

## v0.12.8 — 2026-06-20
**Studio:** Feat: Queue cards show their source times, and spreadsheet rows tint by severity.
**Core:** Feat: Long operations show elapsed time and an estimate of what is left, surviving a page reload.

## v0.12.7 — 2026-06-20
**Screenspace:** Feat: Run scene analysis and re-run saved tasks from the command line, for unattended work.

## v0.12.6 — 2026-06-20
**Studio:** Feat: Click a card's duration badge to change the clip's length — drag the ends, type exact times, or nudge by 30 seconds.

## v0.12.5 — 2026-06-19
**Studio:** Feat: Building a viewer no longer blocks the page, and can be cancelled.
**Screenspace:** Feat: Chain tools so a later step only looks in a window around what an earlier one found.

## v0.12.4 — 2026-06-19
**Screenspace:** Feat: Template steps in a chain can use an uploaded image, not only a frame from the video.

## v0.12.3 — 2026-06-18
**Screenspace:** Feat: Transcript tags show in the sidebar next to detector events.

## v0.12.2 — 2026-06-18
**Overview:** Feat: Alignment offsets can be set per participant lane, for sessions that were not recorded together.
**Screenspace:** Feat: A Boundary tool that finds where the screen changes substantially — menu to gameplay, a loading screen ending — without drawing a region first.

## v0.12.1 — 2026-06-18
**Screenspace:** Feat: Pin reference frames, test a detector against them, and apply the threshold it suggests.

## v0.11.21 — 2026-06-09
**Screenspace:** Feat: Controls to cut false positives from text and number reading — a confidence floor, an integers-only mode, and normalisation options.

## v0.11.19 — 2026-06-07
**Transcripts:** Feat: A Friction tab that scores transcript passages for signs of user struggle.

## v0.11.10 — 2026-05-22
**Studio:** Feat: Finished reels and viewers appear in the artifact log alongside clips, ready to reopen.

## v0.11.9 — 2026-05-21
**Studio:** Feat: Artifacts and reels can generate at the same time instead of waiting for each other.

## v0.11.8 — 2026-05-21
**Studio:** Feat: Generate buttons fill with progress as they work.

## v0.11.7 — 2026-05-21
**Overview:** Feat: Fine-tune audio and video alignment per participant in Convergence.

## v0.11.6 — 2026-05-17
**Transcripts:** Feat: Better transcripts on quiet footage — silence is skipped, and Whisper's invented text is filtered out.

## v0.11.1 — 2026-05-16
**Core:** Feat: Double-clicking the app opens Studio directly; you pick the spreadsheet in the browser.

## v0.11.0 — 2026-05-16
**Core:** Feat: A new Start screen with a Recently opened list, one spreadsheet picker for both Google and Excel, an in-app changelog and a native folder picker.

## v0.10.144 — 2026-05-12
**Core:** Feat: The Start screen skips itself when your recordings are already in place.

## v0.10.140 — 2026-05-03
**Studio:** Feat: Drag timestamp cells onto the Artifact or Reel intake, with a preview of what you will get.

## v0.10.135 — 2026-04-28
**Screenspace:** Feat: Notes and top issues fold into a slim panel, leaving the canvas most of the screen.
