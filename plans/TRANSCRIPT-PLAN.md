# Transcript Integration Plan

## Overview

Transcripts in clipgen are currently a "write to disk and forget" feature. This plan transforms them into a tightly coupled property of artifacts, surfaced across all web interfaces, with a corrections system that improves transcription quality over time.

### Use cases driving this work

1. **Read back** — surface transcript text alongside video so researchers can review what was said without scrubbing
2. **Search / count** — keyword search across transcripts to find moments or measure frequency across participants
3. **Pull quotes** — select and export text passages as attributed evidence for stakeholder presentations

### Key architectural decisions

- **Source-video transcripts are the primary model.** `--pre-transcribe [ID...]` transcribes full source videos upfront and stores results in the manifest. No IDs = all participants enqueued. This is the primary intended workflow: transcribe everything before analysis begins, then walk away for the duration.
- Per-clip transcripts are *derived* from source transcripts when available (filtering segments by time range). On-the-fly Whisper transcription is the fallback when no source transcript exists and transcription is requested.
- Source transcripts are stored as a new top-level section in `clipgen_manifest.json`, keyed by participant ID.
- Transcript segments are embedded on clip/reel artifact records as a `transcript` field (derived from source or generated on-the-fly).
- Standalone transcript file output stays for researchers who just want text files (secondary, opt-in via `--transcribe`).
- Two-tier corrections dictionary: study-local (default) + global (promoted).
- Corrections feed back into Whisper as context keywords AND apply as post-processing.
- Transcript segment schema is `{start, end, text}` — designed to accommodate a `speaker` field later (for diarization).
- Reel transcripts are merged from constituent clip transcripts, maintaining order.
- Initial Studio implementation is read-only; video playback integration is deferred.

---

## Phase 1: Infrastructure

Embed transcript data on artifacts and build the corrections system.

### Pre-transcription CLI (`--pre-transcribe`)

- [ ] Add `--pre-transcribe [ID...]` CLI flag — accepts zero or more participant IDs; no IDs = enqueue all participants found in the spreadsheet
- [ ] For each target participant, transcribe their source video (`{study}_{participant}.mp4`) using `transcribe_video()`
- [ ] Store full-video transcript results in `clipgen_manifest.json` under a new top-level `source_transcripts` key, keyed by participant ID
- [ ] Schema: `{"source_transcripts": {"P01": {"segments": [...], "language": str, "model": str, "source_file": str}}}`
- [ ] Operation is idempotent: re-running `--pre-transcribe` skips participants already present in `source_transcripts`; force-retranscribe with an explicit flag if needed
- [ ] Requires a spreadsheet (to resolve participant IDs and source video paths); can combine with `-s`, `-i/-o`, `-v`
- [ ] Cannot combine with mode flags or `--studio`/`--insights`

### Transcript as artifact property

- [ ] Add `transcript` field to clip artifact records in `process_clips()` — after `filter_segments()`, embed the segments list directly on the clip's artifact dict
- [ ] Source priority: if a source transcript exists in the manifest for the clip's participant, derive via `filter_segments()`; otherwise fall back to on-the-fly Whisper transcription (if `--transcribe` or `TRANSCRIBE_ENABLED`)
- [ ] Segment schema: `{"start": float, "end": float, "text": str}`
- [ ] Add merged `transcript` field to reel artifact records — assembled from constituent clips' transcript segments during reel building, ordered to match the reel
- [ ] Keep standalone transcript file output and transcript-type manifest entries as-is (opt-in via `--transcribe`, for researchers who just want text files)

### Corrections dictionary

- [ ] Define corrections file format: `clipgen_corrections.json` (study-local, in output directory) and `~/.clipgen/corrections.json` (global)
- [ ] Schema: `{"corrections": [{"from": str, "to": str, "created": iso8601}]}`
- [ ] Study-local corrections take precedence over global when both match
- [ ] Promotion mechanism: move a study-local correction to the global dictionary
- [ ] Load global + study-local corrections at transcription time

### Corrections integration with Whisper

- [ ] Feed correction targets (the `"to"` values) as `context_keywords` to `transcribe_video()` — these get appended to Whisper's `initial_prompt`
- [ ] Post-processing pass after transcription: apply known `"from" → "to"` corrections to segment text
- [ ] Auto-applied corrections should be flagged or logged so the researcher knows what was changed

---

## Phase 2: Curation

Surface transcripts in Studio and Viewer for the analysis workflow. **Initial implementation targets Studio.**

### Display

**Studio transcript panel:**

- [ ] Add a togglable transcript panel on the right side of the Studio interface; slides out when toggled, slides away when toggled off
- [ ] Panel shows only participants that have source transcripts in the manifest; participants without transcripts are not mentioned
- [ ] Empty state (no source transcripts in manifest at all): display an instruction to run `--pre-transcribe`
- [ ] Each participant with a transcript is shown as a collapsible section (folded by default); clicking the header expands/collapses it
- [ ] Transcript display uses a two-column layout: **Time** (left) | **Transcript text** (right)
- [ ] Panel is read-only in the initial implementation; editing and playback interaction deferred to later phases

**Timeline Viewer:**

- [ ] Timeline Viewer detail panel: scrollable transcript alongside video playback
- [ ] Playback-synced transcript highlighting — active segment highlighted as video plays

### Editing

- [ ] Editable transcript text in the curation UI (Studio or Viewer)
- [ ] Edits create correction entries: save `{"from": original, "to": edited}` to study-local corrections dictionary
- [ ] Option to promote a correction to the global dictionary

### Search

- [ ] Keyword search across transcript content in Studio or Viewer
- [ ] Cross-participant search: "show me every artifact where someone said X"
- [ ] Occurrence counting for keyword frequency analysis

---

## Phase 3: Presentation

Surface transcripts in output-facing viewers and the Insights Builder.

### Insights Builder

- [ ] Show transcript text when browsing artifacts to attach as evidence
- [ ] Quote selection: highlight a passage from transcript segments → creates an attributed quote block
- [ ] Quote attribution includes: participant, timestamp, category, source clip
- [ ] Quotes attached to insight causes/behaviors/impacts as a first-class evidence type

### Insights Viewer (exported)

- [ ] Render quotes as evidence blocks alongside video clips in the standalone viewer
- [ ] Quote formatting: quoted text with participant attribution and timestamp

### Timeline Viewer (reader-facing)

- [ ] Transcript visible to readers in the detail panel (not just during curation)
- [ ] Type filter updated to include transcript-bearing artifacts

---

## Future: Video Playback in Studio

Not in scope for the initial Studio implementation, but designed to be added later.

- Preview area to the right of the spreadsheet for video playback
- Clicking a transcript time or segment triggers playback at that position
- Allow creating timestamps / generating artifact cards directly from transcript selections
- Interaction model (where the player lives, how it coexists with the spreadsheet and transcript panel) to be designed when this is tackled

---

## Future: Speaker Diarization

Not in scope for Phases 1-3, but the segment schema is designed to accommodate it.

- Extend segment schema: `{start, end, text, speaker}`
- Separate moderator speech from participant speech
- Enable filtering/tagging by speaker role
- Would require a diarization engine alongside or replacing Whisper (e.g., pyannote.audio)

---

## Assumptions to monitor

- **Whisper quality with context keywords** — test early with real session audio to confirm corrections-as-keywords measurably improves output
- **Manifest size** — source transcripts stored in manifest may grow large for long sessions; worth checking at scale (e.g. 6+ participants with hour-long videos)
- **Auto-apply safety** — corrections applied without approval could introduce errors in edge cases (e.g., a legitimate use of the "from" text in a different context)
- **Pre-transcribe runtime** — full-session transcription is slow; the idempotency guarantee (skip already-stored participants) means re-running is safe and partial runs can be completed incrementally
