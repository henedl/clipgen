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
- **Segment IDs are a manifest-layer concept.** Raw `TranscriptSegment` remains `{start, end, text}`. The `id` field (e.g. `"P01:42"`) is assigned when segments are stored in `transcripts_manifest.json`, not at transcription time. Filter results and file output use the base schema without IDs.
- **Corrections are read-time post-processing.** Raw Whisper output is stored verbatim in the manifest. Corrections are applied as a transform layer when serving segments via API or embedding on artifacts. This means re-transcription cleanly replaces raw segments and existing corrections auto-apply to new text.
- **Sidecar transcript cache is deprecated.** The `.transcript.json` files next to source videos are replaced by `transcripts_manifest.json` as the single persistence layer. Remove `save_transcript_cache()` and `load_transcript_cache()` from `transcripts.py` and their call sites in `clipgen.py:_transcribe_segments()`.
- **Re-transcription uses badge-only invalidation.** When a participant is re-transcribed, derived clip `transcript` fields become stale. The workspace surfaces a "transcript outdated" indicator. Full cascade re-derivation is deferred — researchers re-generate clips to pick up updates.

---

## Phase 1: Infrastructure

Embed transcript data on artifacts and build the corrections system.

### Pre-transcription CLI (`--pre-transcribe`)

- [x] Add `--pre-transcribe [ID...]` CLI flag — accepts zero or more participant IDs; no IDs = enqueue all participants found in the spreadsheet
- [x] For each target participant, transcribe their source video (`{study}_{participant}.mp4`) using `transcribe_video()`
- [x] Store full-video transcript results in a separate `transcripts_manifest.json` under a `source_transcripts` key, keyed by participant ID
- [x] Add `TRANSCRIPTS_MANIFEST_FILENAME` to `config.py` alongside `SCREENSPACE_MANIFEST_FILENAME`
- [x] Add `load_transcripts_manifest()` / `save_transcripts_manifest()` to `transcripts.py` following the `load_screenspace_manifest()` / `save_screenspace_manifest()` pattern
- [x] Schema: `{"source_transcripts": {"P01": {"segments": [{start, end, text}, ...], "language": str, "model": str, "source_file": str, "transcribed_at": iso8601}}, "corrections": [...]}` — segments store raw Whisper output; `id` field is added at manifest storage time (e.g. `"P01:42"`); corrections are applied as a read-time layer, not stored on segments
- [x] Operation is idempotent: re-running `--pre-transcribe` skips participants already present in `source_transcripts` within the transcripts manifest; force-retranscribe with an explicit flag if needed
- [x] Requires a spreadsheet (to resolve participant IDs and source video paths); can combine with `-s`, `-i/-o`, `-v`
- [x] Cannot combine with mode flags or `--studio`/`--insights`

### Transcript as artifact property

- [x] Add `transcript` field to clip artifact records in `process_clips()` — after `filter_segments()`, embed the segments list directly on the clip's artifact dict
- [x] Source priority: if a source transcript exists in the transcripts manifest for the clip's participant, derive via `filter_segments()`; otherwise fall back to on-the-fly Whisper transcription (if `--transcribe` or `TRANSCRIBE_ENABLED`). The sidecar `.transcript.json` cache is removed.
- [x] Segment schema in manifest: `{"id": str, "start": float, "end": float, "text": str}` — `id` is index-based (e.g., `"P01:42"` for participant + segment index) for provenance tracking. Raw `TranscriptSegment` TypedDict stays `{start, end, text}` — `id` is assigned at manifest storage, not transcription.
- [x] Add merged `transcript` field to reel artifact records — assembled from constituent clips' transcript segments during reel building, ordered to match the reel. Segment timestamps in the merged reel transcript must be offset by cumulative titlecard/transition durations. Use the reel's `components` metadata (which stores per-component start/end in reel-relative time) to compute offsets.
- [x] Keep standalone transcript file output and transcript-type manifest entries as-is (opt-in via `--transcribe`, for researchers who just want text files)

### Corrections dictionary

- [x] Corrections stored inside `transcripts_manifest.json` under the `corrections` key (study-local only; global dictionary deferred until multi-study need is proven)
- [x] Schema: `{"corrections": [{"id": str, "from": str, "to": str, "created": iso8601}]}`
- [x] Load corrections from the transcripts manifest at transcription time. Corrections are applied at read-time as post-processing on raw segments. Raw Whisper output in `source_transcripts` is never mutated. This enables clean re-transcription (replace raw segments, corrections auto-apply).

### Corrections integration with Whisper

- [x] Feed correction targets (the `"to"` values) as `context_keywords` to `transcribe_video()` — these get appended to Whisper's `initial_prompt`
- [x] Post-processing pass after transcription: apply known `"from" → "to"` corrections to segment text
- [x] Auto-applied corrections should be flagged or logged so the researcher knows what was changed

### Sidecar cache removal

- [x] Remove sidecar cache: delete `save_transcript_cache()` and `load_transcript_cache()` from `transcripts.py`; remove cache call sites in `clipgen.py:_transcribe_segments()`
- [x] Update `_transcribe_segments()` in `clipgen.py` to check `transcripts_manifest.json` for pre-existing source transcripts. New read priority: in-memory cache -> transcripts manifest -> live Whisper.

### Manifest-layer enrichment

- [x] Add `apply_corrections(segments, corrections)` function to `transcripts.py` — applies `from -> to` substitutions to segment text, returns corrected copy without mutating input
- [x] Assign segment IDs at manifest storage time: `"{participant}:{index}"` format, computed in `save_transcripts_manifest()`

---

## Phase 2: Transcript Workspace

Dedicated page at `/transcripts/` for transcript curation, editing, and transcription management. Follows the Screenspace architectural pattern: own Flask blueprint, own assets, registered on the combined server.

### Server and routing

- [x] New `transcripts_server.py` — Flask blueprint registered at `/transcripts/`
- [x] `_init_transcripts_state()` — reuse the shared participant video discovery utility (factored out of Screenspace's `_discover_participant_videos`), loads source transcripts from transcripts manifest
- [x] Factor `_discover_participant_videos()` from `screenspace_server.py` into a shared utility in `utils.py` or `files.py`; both Screenspace and Transcript workspaces call it
- [x] Always register transcript blueprint in `start_combined_server()` (like Screenspace), not gated by `--transcripts`. The `--transcripts` flag controls whether `TranscriptWorker` starts and the workspace is the active landing page.
- [x] `--transcripts` CLI flag to launch the workspace; works standalone (no spreadsheet required) as long as source media files exist in the input directory
- [x] Can combine with `-i/-o`, `-v`; cannot combine with mode flags, format flags, or `--viewer`
- [x] Video serving: add `/transcripts/media/<filename>` route serving source videos from the input directory (following Screenspace's `/screenspace/media/` pattern)
- [x] Manifest write synchronization: use `threading.Lock` for all `transcripts_manifest.json` writes (following Screenspace's `_manifest_lock` pattern in `screenspace_server.py`)

### Assets

- [x] `transcripts.html`, `transcripts.js`, `transcripts.css` in `assets/web/`
- [x] Served directly by Flask (no inlining), consistent with Studio and Screenspace

### Workspace layout

The workspace is a full-page environment for deep transcript work:

- [x] **Header**: participant selector (dropdown of discovered source videos), transcription status indicator, navigation back to Studio/other workspaces
- [x] **Main area — transcript view**: scrollable transcript for the selected participant, displayed as a segment list with timestamps on the left and text on the right
- [x] **Video player**: inline video playback for the selected participant's source video; clicking a transcript segment seeks to that position
- [x] **Playback-synced highlighting**: active segment highlighted as video plays, auto-scroll to keep current segment visible
- [x] **Corrections log**: accessible from header menu, shows study-local corrections with delete action; inline editing is the primary correction flow, the log is for review and cleanup

### Transcription queue

The workspace can trigger and monitor transcription jobs, not just view results. Architecture follows the `ScreenspaceWorker` pattern (simplified):

- [x] **`TranscriptWorker`**: thread-based background worker modeled after `ScreenspaceWorker`
  - `PriorityQueue` for task ordering
  - Task lifecycle: QUEUED → RUNNING → COMPLETED / FAILED / CANCELLED (no PAUSE/RESUME — Whisper doesn't produce meaningful partial results)
  - `on_task_complete` callback for manifest persistence (wired same as Screenspace)
  - `restore_tasks()` for loading historical task state on server restart
  - Task dict: `{id, participant, video_path, status, progress, result, created_at, completed_at}`
  - Task ID format: `tr_<8hex>` (following Screenspace's `ss_<8hex>`)
  - No task reordering (not needed for a simple "transcribe all participants" queue)
- [x] **Queue UI**: list of participants with transcription status (not started / queued / in progress / complete)
- [x] **Enqueue**: select one or more participants to transcribe; starts a background transcription job on the server
- [x] **Progress**: show progress for in-flight transcription (estimated from audio duration vs. last segment's timestamp); frontend polls via same fingerprint pattern as Screenspace
- [x] **Idempotent**: participants already transcribed show as complete; re-transcribe option available with confirmation
- [x] **API endpoints**:
  - `POST /transcripts/api/transcribe` — enqueue participant(s) for transcription
  - `GET /transcripts/api/transcribe/status` — poll transcription job status
- [x] Transcription results are stored to `source_transcripts` in the transcripts manifest, same as `--pre-transcribe`
- [x] Re-transcription staleness: when a participant is re-transcribed, flag any clip artifacts with embedded `transcript` fields for that participant as stale. Surface a "transcript outdated" badge in the workspace UI.

### API endpoints

- [x] `GET /transcripts/api/participants` — list discovered source videos with transcription status
- [x] `GET /transcripts/api/transcript/<participant>` — return full source transcript segments for a participant
- [x] `PUT /transcripts/api/transcript/<participant>/segment` — edit a segment's text, identified by segment `id` (not array index or timestamps — indices change on re-transcription, timestamps may not be unique); creates a correction entry
- [x] `GET /transcripts/api/vtt/<participant>` — serve transcript data as WebVTT for native `<track>` subtitle support (`transcripts.py` already has `_format_vtt()`)
- [x] `GET /transcripts/api/corrections` — list all study-local corrections
- [x] `POST /transcripts/api/corrections` — add a correction manually
- [x] `DELETE /transcripts/api/corrections/<id>` — remove a correction
- [x] `GET /transcripts/api/search?q=<query>` — keyword search across all transcribed participants; returns matching segments with participant ID and timestamps (Studio calls this directly via `../transcripts/api/search`, no duplicate endpoint needed)
- [x] Update `/api/status` in `server.py` to report `transcripts: true/false` alongside `studio`, `insights`, `screenspace`

### Editing

- [x] Click-to-edit on any transcript segment text
- [x] On save, original and edited text are compared; if different, a correction entry is created automatically (`{"from": original, "to": edited}`)
- [x] Edited segments are visually marked (subtle indicator that the text was corrected)
- [x] Corrections apply immediately to all identical occurrences across all participants' transcripts (post-processing pass)

### Search

- [x] Search bar in the workspace header — keyword search across all transcribed participants
- [x] Results shown as a filterable list: participant, timestamp, matching segment text with query highlighted
- [x] Click a result to jump to that participant's transcript at the matching segment
- [x] Occurrence count displayed per participant and total

---

## Phase 3: Segment Marking + Studio Transcript Intake

Pre-filter transcript content via a marking/curation system in the Transcript workspace, then surface curated marks in Studio as an intake tab for artifact generation. Follows the Screenspace Intake precedent: raw data → human curation → clustered cards → artifacts.

### Key design decisions

- **Marks are individual segment flags** stored as a separate manifest structure (like corrections). Raw segments stay immutable. Merging into groups happens at display time in Studio's clustering algorithm.
- **Optional color categories** (pain point, delight, quote, insight, task, bookmark) + optional free-text labels on each mark.
- **Studio Intake defaults to marks only**, with a toggle to show all segments.
- **Search → bulk mark** leverages the existing search API (which returns `segment_id`).
- **Full transcript text remains available in Viewers** (Phase 4) — marks only control the Transcript→Studio intake pipeline.

### Mark data model

- [x] Marks stored in `transcripts_manifest.json` under a top-level `"marks"` key (parallel to `"source_transcripts"` and `"corrections"`)
- [x] Mark schema: `{"id": "m_{hex8}", "segment_id": "P01:42", "category": str|null, "label": str|null, "created": iso8601}`
- [x] One mark per segment (POST deduplicates on `segment_id` — if mark exists, update it)
- [x] `MARK_CATEGORIES` constant in `transcripts.py`: `pain_point` (#dc2626), `delight` (#16a34a), `quote` (#2563eb), `insight` (#f97316), `task` (#8b5cf6), `bookmark` (#0891b2)
- [x] Re-transcription resilience: marks reference segment IDs that are index-based. When segment IDs change, marks API returns `valid: false` for unresolvable marks. UI shows orphaned marks dimmed with option to dismiss.

### Manifest changes (`transcripts.py`)

- [x] `_empty_transcripts_manifest()` → add `"marks": []`
- [x] `load_transcripts_manifest()` → add `"marks": data.get("marks") or []` to returned dict
- [x] `save_transcripts_manifest()` → add `marks` parameter (default `None` = preserve on-disk marks); include in serialized data

### Mark API endpoints (`transcripts_server.py`)

- [x] `GET /transcripts/api/marks` — list all marks, enriched with resolved segment data (participant, start, end, text, valid flag) + categories definition
- [x] `POST /transcripts/api/marks` — create marks; body: `{segment_ids: [...], category?, label?}`; each segment gets its own `m_{hex8}` ID
- [x] `PUT /transcripts/api/marks/<id>` — update category or label
- [x] `DELETE /transcripts/api/marks/<id>` — remove a mark; also support bulk: `DELETE /transcripts/api/marks` with `{ids: [...]}`
- [x] Update `_do_persist()` to pass `marks=_manifest.get("marks")` to `save_transcripts_manifest()`
- [x] Enrich `api_transcript()` response with per-segment marks (build `marks_by_segment_id` lookup)

### Transcript workspace marking UI

- [x] **Gutter mark column**: add a mark dot before the timestamp in each `.segment-row` — `[mark-dot] [timestamp] [text]`. Unmarked: empty circle (border only). Marked: filled circle with category color.
- [x] **`toggleMark(segmentId)`**: click gutter dot — if unmarked → POST create mark (uses `state.lastMarkCategory`); if marked → show popover
- [x] **`showMarkPopover(el, segmentId, markObj)`**: floating popover with 6 category color pills, label text input, "Remove" action. Single shared DOM element repositioned each time.
- [x] **`markAllSearchResults()`**: collect all `segment_id` from `state.searchResults.results`, POST bulk create. "Mark All" button shown in search results header next to count.
- [x] Add `markPopover` HTML element to `transcripts.html`
- [x] CSS for `.segment-mark`, `.segment-mark.marked`, `.mark-popover`, category pills

### Studio Transcript Intake tab

- [x] Third tab in Studio's `#sheetPreview` area: **"Transcript Intake"** alongside "Sheet Preview" and "Screenspace Intake"
- [x] Tab only appears when `/api/status` reports `transcripts: true`; hidden otherwise
- [x] `#transcriptIntakePanel` with: merge gap threshold input (1–60s, default 5), "Show all segments" toggle, "Add All to Artifacts"/"Add All to Reel" buttons, category filter pills, participant filter pills, canvas timeline, cards grid
- [x] `pollTranscriptIntakeMarks()` — fetch `../transcripts/api/marks`, cluster, render; polled every 10s when tab active (same pattern as `pollIntakeEvents()`)
- [x] `clusterTranscriptMarks(marks, thresholdSec)` — group by participant, merge marks whose segments are within `thresholdSec` gap (same algorithm shape as `clusterIntakeEvents()`)
- [x] Cards show: category color dot, participant + time range, truncated transcript text (~80 chars), duration badge
- [x] **Drag data format**:

  ```javascript
  {
    participant: str,
    desc: str,              // category label or "transcript"
    segStart: float,        // cluster start in seconds
    segDuration: float,     // cluster end - start
    source: "transcript",
    mark_ids: [str],        // provenance: mark IDs (e.g., ["m_abc123", "m_def456"])
  }
  ```
  
- [x] Artifacts generated from transcript intake carry `source: "transcript"`, `mark_ids: [...]` — matching the Screenspace pattern in `server.py`
- [x] Update `syncPreviewTab()` for three-tab switching with independent poll timers
- [x] `filteredTranscriptIntakeClusters()` — filter by category pills, participant pills, text search
- [x] `renderTranscriptIntakeTimeline()` — canvas timeline with category-colored markers

### Studio generation integration

- [x] Update drop handlers in `initDropTargets()` to recognize `source: "transcript"` for both artifact and reel zones
- [x] Update `onGenerateArtifacts()` to partition transcript items and route to `api/generate-intake` with `source: "transcript"`
- [x] Update `_generate_intake_clips()` in `server.py` to look up video path in `transcripts_server._participants` when `source === "transcript"`

### Transcript tooltips

Hover-to-reveal full transcript text on intake cards and other transcript-bearing UI elements.

- [x] **Studio Transcript Intake cards**: hovering a card shows a fixed-position tooltip with the full (untruncated) transcript text. Tooltip appears instantly on hover, disappears instantly on mouseout. Shared DOM element (`#trIntakeTooltip`) positioned near the card with viewport clamping.
- [x] **Tooltip toggle**: button in Studio header (`#tooltipToggle`, chat-bubble icon) next to the dark mode toggle. On by default (`state.trIntakeTooltipsEnabled`). Dims to 40% opacity when off.
- [ ] **Timeline Viewer**: extend tooltip behavior to transcript sidebar segments (Phase 4)
- [ ] **Insights Builder**: extend tooltip behavior to artifact transcript text when browsing evidence (Phase 4)
- [x] **Screenspace Intake cards**: show transcript context tooltip on Screenspace intake cards when transcript data is available at the card's time range (deferred, alongside cross-referencing)
- [x] **Toggle in all frontends**: add tooltip toggle to all viewers and workspaces where transcript data is surfaced. Scoping to Studio intake first to validate the interaction pattern.

### Cross-referencing with Screenspace and Spreadsheet events

- [x] **On Screenspace intake cards**: when transcript data is available, show a "transcript context" line with the text at the card's time range (client-side join by participant + timestamp). Also shows sheet observation text when available.
- [x] **On Transcript intake cards**: when Screenspace events exist at the same timestamp, show detector color dots. Also shows sheet observation text when available.
- [x] **On transcript search results**: when Screenspace events exist at the same timestamp, show a detector badge on the segment card (e.g., "change event detected here"). Also shows sheet observation text when available.
- [x] Cross-referencing is pure client-side — all data sources available in memory, joined by participant + timestamp at display time. Uses `findOverlappingData()` (studio.js) and `findOverlapsForSearch()` (transcripts.js).

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
- **Sidecar cache removal** — researchers who relied on `.transcript.json` files for non-manifest workflows will lose that caching. The manifest becomes the only persistence path. If this causes friction, consider a lightweight in-memory LRU cache in `_transcribe_segments()` as a session-only speedup.
- **Always-on blueprint overhead** — registering the transcript blueprint unconditionally means its routes exist even when no transcripts have been generated. Endpoints return empty results. The `/api/status` response should clearly indicate whether transcript data is available vs. whether the endpoint is reachable.
- **Mark orphaning on re-transcription** — segment IDs are index-based and reassigned on every save. Re-transcription changes indices, orphaning existing marks. The marks API returns a `valid` flag so the UI can show orphaned marks dimmed. Low cost of a few orphaned marks vs. high cost of silently deleting user curation.
- **Mark categories are hardcoded** — not user-configurable for V1. If needed later, categories can move to a `"mark_categories"` key in the manifest with the hardcoded set as defaults.
- **Clustering threshold in Studio** — the time-gap merge threshold for transcript marks (default 5s) may need different tuning than Screenspace events. Transcript segments are typically 2–5s each, so a 5s gap means adjacent marked segments always merge. Monitor whether users want finer control.
