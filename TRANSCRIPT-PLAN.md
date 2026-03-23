# Transcript Integration Plan

## Overview

Transcripts in clipgen are currently a "write to disk and forget" feature. This plan transforms them into a tightly coupled property of artifacts, surfaced across all web interfaces, with a corrections system that improves transcription quality over time.

### Use cases driving this work

1. **Read back** — surface transcript text alongside video so researchers can review what was said without scrubbing
2. **Search / count** — keyword search across transcripts to find moments or measure frequency across participants
3. **Pull quotes** — select and export text passages as attributed evidence for stakeholder presentations

### Key architectural decisions

- Transcript segments are embedded on clip/reel artifact records as a `transcript` field (primary)
- Standalone transcript file output stays for researchers who just want text files (secondary)
- Two-tier corrections dictionary: study-local (default) + global (promoted)
- Corrections feed back into Whisper as context keywords AND apply as post-processing
- Transcript segment schema is `{start, end, text}` — designed to accommodate a `speaker` field later (for diarization)
- Reel transcripts are merged from constituent clip transcripts, maintaining order

---

## Phase 1: Infrastructure

Embed transcript data on artifacts and build the corrections system.

### Transcript as artifact property

- [ ] Add `transcript` field to clip artifact records in `process_clips()` — after `filter_segments()`, embed the segments list directly on the clip's artifact dict
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

Surface transcripts in Studio and Viewer for the analysis workflow.

### Display

- [ ] Studio: show transcript text when inspecting a generated artifact (cell detail / artifact card)
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

## Future: Speaker Diarization

Not in scope for Phases 1-3, but the segment schema is designed to accommodate it.

- Extend segment schema: `{start, end, text, speaker}`
- Separate moderator speech from participant speech
- Enable filtering/tagging by speaker role
- Would require a diarization engine alongside or replacing Whisper (e.g., pyannote.audio)

---

## Assumptions to monitor

- **Whisper quality with context keywords** — test early with real session audio to confirm corrections-as-keywords measurably improves output
- **Manifest size** — not a concern currently, but worth checking with a large study (hundreds of clips with embedded transcript segments)
- **Auto-apply safety** — corrections applied without approval could introduce errors in edge cases (e.g., a legitimate use of the "from" text in a different context)
