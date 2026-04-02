# Transcript Integration Plan

## Overview

Transcripts in clipgen are currently a "write to disk and forget" feature. This plan transforms them into a tightly coupled property of artifacts, surfaced across all web interfaces, with a corrections system that improves transcription quality over time. The primary curation surface is a **dedicated Transcript workspace** served at `/transcripts/` (modeled after Screenspace), with a lightweight **Transcript Intake tab** in Studio for cross-referencing transcript moments during artifact work.

### Use cases driving this work

1. **Read back** — surface transcript text alongside video so researchers can review what was said without scrubbing
2. **Search / count** — keyword search across transcripts to find moments or measure frequency across participants
3. **Pull quotes** — select and export text passages as attributed evidence for stakeholder presentations

### Key architectural decisions

- **Source-video transcripts are the primary model.** `--pre-transcribe [ID...]` transcribes full source videos upfront and stores results in the manifest. No IDs = all participants enqueued. This is the primary intended workflow: transcribe everything before analysis begins, then walk away for the duration. The Transcript workspace also supports queuing transcription jobs directly from the UI.
- Per-clip transcripts are *derived* from source transcripts when available (filtering segments by time range). On-the-fly Whisper transcription is the fallback when no source transcript exists and transcription is requested.
- Source transcripts are stored in a separate `transcripts_manifest.json` (following the Screenspace manifest separation pattern), keyed by participant ID. Per-clip transcript data (the derived `transcript` field on artifact records) stays on the artifact in `clipgen_manifest.json`.
- Transcript segments are embedded on clip/reel artifact records as a `transcript` field (derived from source or generated on-the-fly).
- Standalone transcript file output stays for researchers who just want text files (secondary, opt-in via `--transcribe`).
- Study-local corrections dictionary, stored inside `transcripts_manifest.json`. Global corrections deferred until multi-study need is proven.
- Corrections feed back into Whisper as context keywords AND apply as post-processing.
- Transcript segment schema is `{id, start, end, text}` — `id` is index-based (e.g., `"P01:42"`) for provenance tracking on intake-generated artifacts. Designed to accommodate a `speaker` field later (for diarization).
- Reel transcripts are merged from constituent clip transcripts, maintaining order.
- **Dedicated workspace at `/transcripts/`** — a full-page curation environment (own Flask blueprint, own HTML/JS/CSS assets) following the Screenspace pattern. Works standalone with source media files; no spreadsheet required.
- **Transcript segments are NOT events.** Segments are dense continuous coverage (thousands per video); Screenspace events are sparse point detections (dozens per analysis). Clustering is nonsensical for continuous data. The useful interaction between transcripts and events is cross-referencing by timestamp, not unification into a single model.
- **Studio Transcript Intake tab** — a third tab in Studio's preview area (alongside Sheet Preview and Screenspace Intake) for browsing and searching transcript content without leaving Studio. Appears when source transcripts exist in the manifest. Remains a separate tab from Screenspace Intake — different interaction model (search-based vs. detection-based).

---

## Phase 1: Infrastructure

Embed transcript data on artifacts and build the corrections system.

### Pre-transcription CLI (`--pre-transcribe`)

- [ ] Add `--pre-transcribe [ID...]` CLI flag — accepts zero or more participant IDs; no IDs = enqueue all participants found in the spreadsheet
- [ ] For each target participant, transcribe their source video (`{study}_{participant}.mp4`) using `transcribe_video()`
- [ ] Store full-video transcript results in a separate `transcripts_manifest.json` under a `source_transcripts` key, keyed by participant ID
- [ ] Add `TRANSCRIPTS_MANIFEST_FILENAME` to `config.py` alongside `SCREENSPACE_MANIFEST_FILENAME`
- [ ] Add `load_transcripts_manifest()` / `save_transcripts_manifest()` to `transcripts.py` following the `load_screenspace_manifest()` / `save_screenspace_manifest()` pattern
- [ ] Schema: `{"source_transcripts": {"P01": {"segments": [...], "language": str, "model": str, "source_file": str, "transcribed_at": iso8601}}, "corrections": [...]}`
- [ ] Operation is idempotent: re-running `--pre-transcribe` skips participants already present in `source_transcripts` within the transcripts manifest; force-retranscribe with an explicit flag if needed
- [ ] Requires a spreadsheet (to resolve participant IDs and source video paths); can combine with `-s`, `-i/-o`, `-v`
- [ ] Cannot combine with mode flags or `--studio`/`--insights`

### Transcript as artifact property

- [ ] Add `transcript` field to clip artifact records in `process_clips()` — after `filter_segments()`, embed the segments list directly on the clip's artifact dict
- [ ] Source priority: if a source transcript exists in the transcripts manifest for the clip's participant, derive via `filter_segments()`; otherwise fall back to on-the-fly Whisper transcription (if `--transcribe` or `TRANSCRIBE_ENABLED`)
- [ ] Segment schema: `{"id": str, "start": float, "end": float, "text": str}` — `id` is index-based (e.g., `"P01:42"` for participant + segment index) for provenance tracking
- [ ] Add merged `transcript` field to reel artifact records — assembled from constituent clips' transcript segments during reel building, ordered to match the reel
- [ ] Keep standalone transcript file output and transcript-type manifest entries as-is (opt-in via `--transcribe`, for researchers who just want text files)

### Corrections dictionary

- [ ] Corrections stored inside `transcripts_manifest.json` under the `corrections` key (study-local only; global dictionary deferred until multi-study need is proven)
- [ ] Schema: `{"corrections": [{"id": str, "from": str, "to": str, "created": iso8601}]}`
- [ ] Load corrections from the transcripts manifest at transcription time

### Corrections integration with Whisper

- [ ] Feed correction targets (the `"to"` values) as `context_keywords` to `transcribe_video()` — these get appended to Whisper's `initial_prompt`
- [ ] Post-processing pass after transcription: apply known `"from" → "to"` corrections to segment text
- [ ] Auto-applied corrections should be flagged or logged so the researcher knows what was changed

---

## Phase 2: Transcript Workspace

Dedicated page at `/transcripts/` for transcript curation, editing, and transcription management. Follows the Screenspace architectural pattern: own Flask blueprint, own assets, registered on the combined server.

### Server and routing

- [ ] New `transcripts_server.py` — Flask blueprint registered at `/transcripts/`
- [ ] `_init_transcripts_state()` — reuse the shared participant video discovery utility (factored out of Screenspace's `_discover_participant_videos`), loads source transcripts from transcripts manifest
- [ ] Register blueprint in `start_combined_server()` alongside Studio, Insights, and Screenspace
- [ ] `--transcripts` CLI flag to launch the workspace; works standalone (no spreadsheet required) as long as source media files exist in the input directory
- [ ] Can combine with `-i/-o`, `-v`; cannot combine with mode flags, format flags, or `--viewer`

### Assets

- [ ] `transcripts.html`, `transcripts.js`, `transcripts.css` in `assets/web/`
- [ ] Served directly by Flask (no inlining), consistent with Studio and Screenspace

### Workspace layout

The workspace is a full-page environment for deep transcript work:

- [ ] **Header**: participant selector (dropdown of discovered source videos), transcription status indicator, navigation back to Studio/other workspaces
- [ ] **Main area — transcript view**: scrollable transcript for the selected participant, displayed as a segment list with timestamps on the left and text on the right
- [ ] **Video player**: inline video playback for the selected participant's source video; clicking a transcript segment seeks to that position
- [ ] **Playback-synced highlighting**: active segment highlighted as video plays, auto-scroll to keep current segment visible
- [ ] **Corrections log**: accessible from header menu, shows study-local corrections with delete action; inline editing is the primary correction flow, the log is for review and cleanup

### Transcription queue

The workspace can trigger and monitor transcription jobs, not just view results. Architecture follows the `ScreenspaceWorker` pattern (simplified):

- [ ] **`TranscriptWorker`**: thread-based background worker modeled after `ScreenspaceWorker`
  - `PriorityQueue` for task ordering
  - Task lifecycle: QUEUED → RUNNING → COMPLETED / FAILED / CANCELLED (no PAUSE/RESUME — Whisper doesn't produce meaningful partial results)
  - `on_task_complete` callback for manifest persistence (wired same as Screenspace)
  - `restore_tasks()` for loading historical task state on server restart
  - Task dict: `{id, participant, video_path, status, progress, result, created_at, completed_at}`
  - Task ID format: `tr_<8hex>` (following Screenspace's `ss_<8hex>`)
  - No task reordering (not needed for a simple "transcribe all participants" queue)
- [ ] **Queue UI**: list of participants with transcription status (not started / queued / in progress / complete)
- [ ] **Enqueue**: select one or more participants to transcribe; starts a background transcription job on the server
- [ ] **Progress**: show progress for in-flight transcription (estimated from audio duration vs. last segment's timestamp); frontend polls via same fingerprint pattern as Screenspace
- [ ] **Idempotent**: participants already transcribed show as complete; re-transcribe option available with confirmation
- [ ] **API endpoints**:
  - `POST /transcripts/api/transcribe` — enqueue participant(s) for transcription
  - `GET /transcripts/api/transcribe/status` — poll transcription job status
- [ ] Transcription results are stored to `source_transcripts` in the transcripts manifest, same as `--pre-transcribe`

### API endpoints

- [ ] `GET /transcripts/api/participants` — list discovered source videos with transcription status
- [ ] `GET /transcripts/api/transcript/<participant>` — return full source transcript segments for a participant
- [ ] `PUT /transcripts/api/transcript/<participant>/segment` — edit a segment's text, identified by `{start, end}` timestamps (not array index — indices change on re-transcription); creates a correction entry
- [ ] `GET /transcripts/api/vtt/<participant>` — serve transcript data as WebVTT for native `<track>` subtitle support (`transcripts.py` already has `_format_vtt()`)
- [ ] `GET /transcripts/api/corrections` — list all study-local corrections
- [ ] `POST /transcripts/api/corrections` — add a correction manually
- [ ] `DELETE /transcripts/api/corrections/<id>` — remove a correction
- [ ] `GET /transcripts/api/search?q=<query>` — keyword search across all transcribed participants; returns matching segments with participant ID and timestamps (Studio calls this directly via `../transcripts/api/search`, no duplicate endpoint needed)
- [ ] Update `/api/status` in `server.py` to report `transcripts: true/false` alongside `studio`, `insights`, `screenspace`

### Editing

- [ ] Click-to-edit on any transcript segment text
- [ ] On save, original and edited text are compared; if different, a correction entry is created automatically (`{"from": original, "to": edited}`)
- [ ] Edited segments are visually marked (subtle indicator that the text was corrected)
- [ ] Corrections apply immediately to all identical occurrences across all participants' transcripts (post-processing pass)

### Search

- [ ] Search bar in the workspace header — keyword search across all transcribed participants
- [ ] Results shown as a filterable list: participant, timestamp, matching segment text with query highlighted
- [ ] Click a result to jump to that participant's transcript at the matching segment
- [ ] Occurrence count displayed per participant and total

---

## Phase 3: Studio Transcript Intake

Surface transcript data inside Studio as a lightweight intake tab for cross-referencing during artifact work.

### Intake tab

- [ ] Third tab in Studio's `#sheetPreview` area: **"Transcript Intake"** alongside "Sheet Preview" and "Screenspace Intake"
- [ ] Tab only appears when source transcripts exist in the transcripts manifest (at least one participant transcribed); hidden otherwise
- [ ] Loads transcript data via the Transcript workspace API (`../transcripts/api/...`), following the same cross-blueprint pattern Studio uses for Screenspace events (`../screenspace/api/events`)

### Intake tab contents

- [ ] **Participant filter**: dropdown or chip bar to filter by participant
- [ ] **Search**: keyword search across transcript content, scoped to selected participant(s) or all; calls `../transcripts/api/search`
- [ ] **Segment list**: matching transcript segments displayed as compact cards — timestamp, participant badge, text snippet
- [ ] **"New only" toggle**: filter to segments not yet associated with any artifact (helps find uncaptured moments)
- [ ] **Drag to artifact area**: segments dragged using the standardized drag data format:
  ```javascript
  {
    participant: str,
    desc: str,             // matched text snippet (truncated)
    segStart: float,       // segment start in seconds
    segDuration: float,    // segment end - start
    source: "transcript",
    search_query: str,     // provenance: the search term that found this
    segment_ids: [str],    // provenance: segment IDs (e.g., ["P01:42", "P01:43"])
  }
  ```
- [ ] Artifacts generated from transcript intake carry `source: "transcript"`, `segment_ids: [...]`, `intake_label: "search: <query>"` — matching the Screenspace pattern in `server.py`

### Cross-referencing with Screenspace events

- [ ] **On Screenspace intake cards**: when transcript data is available, show a "transcript context" line with the text at the card's time range (client-side join by participant + timestamp)
- [ ] **On transcript search results**: when Screenspace events exist at the same timestamp, show a detector badge on the segment card (e.g., "change event detected here")
- [ ] Cross-referencing is pure client-side — both data sources available in memory, joined by participant + timestamp at display time

### Studio API additions

- [ ] `GET /studio/api/transcripts` — return transcript summary from transcripts manifest (participant list with segment counts); full segment data fetched via `../transcripts/api/transcript/<participant>`
- [ ] No duplicate search endpoint — Studio calls `../transcripts/api/search` directly

---

## Phase 4: Presentation

Surface transcripts in output-facing viewers and the Insights Builder.

### Insights Builder

- [ ] Show transcript text when browsing artifacts to attach as evidence
- [ ] Quote selection: highlight a passage from transcript segments → creates an attributed quote block
- [ ] Quote attribution includes: participant, timestamp, category, source clip
- [ ] Quotes attached to insight causes/behaviors/impacts as a first-class evidence type

### Insights Viewer (exported)

- [ ] Render quotes as evidence blocks alongside video clips in the standalone viewer
- [ ] Quote formatting: quoted text with participant attribution and timestamp

### Timeline Viewer

- [ ] **Transcript sidebar** (not a timeline track row): scrollable transcript panel alongside video playback. Transcript data is continuous — every second has text — so track markers would produce a solid unbroken bar that communicates nothing. Screenspace events work as track markers because they're sparse.
- [ ] Playback-synced transcript highlighting — active segment highlighted as video plays, auto-scroll to keep current segment visible
- [ ] Data contract: add optional `transcripts` key to `window.CLIPGEN_DATA`:
  ```javascript
  {
    "transcripts": {
      "P01": { "segments": [...], "language": "en" },
      ...
    }
  }
  ```
- [ ] Add `transcripts` parameter to `finalize_timeline_data()` and a `load_transcripts_for_viewer()` helper following the `load_screenspace_events_for_viewer()` pattern in `viewer.py`
- [ ] Type filter updated to include transcript-bearing artifacts
- [ ] Show Screenspace event badges inline with transcript text at corresponding timestamps (cross-referencing in the viewer)

---

## Future: Speaker Diarization

Not in scope for Phases 1–4, but the segment schema is designed to accommodate it.

- Extend segment schema: `{id, start, end, text, speaker}`
- Separate moderator speech from participant speech
- Enable filtering/tagging by speaker role
- Would require a diarization engine alongside or replacing Whisper (e.g., pyannote.audio)

---

## Assumptions to monitor

- **Whisper quality with context keywords** — test early with real session audio to confirm corrections-as-keywords measurably improves output
- **Manifest size** — mitigated by using a separate `transcripts_manifest.json` so the clipgen manifest isn't bloated; still worth checking that the transcripts manifest loads quickly at scale (e.g. 6+ participants with hour-long videos = 10k+ segments)
- **Auto-apply safety** — corrections applied without approval could introduce errors in edge cases (e.g., a legitimate use of the "from" text in a different context)
- **Pre-transcribe runtime** — full-session transcription is slow; the idempotency guarantee (skip already-stored participants) means re-running is safe and partial runs can be completed incrementally
- **Background transcription in workspace** — running Whisper in a background thread while serving the Flask app; need to ensure thread safety for manifest writes and that the GIL doesn't starve the server. Same single-threaded-by-nature constraint as Screenspace's task queue worker.
- **Standalone discovery** — the workspace discovers source media from the input directory without a spreadsheet. Reuses the shared video discovery utility (factored from Screenspace's `_discover_participant_videos`). Participant ID extraction from filenames (`{study}_{participant}.mp4`) must be robust to naming variations.
