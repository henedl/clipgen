# AI switches

## Summary

Two switches in Settings → General, for users and organizations that do not want models running.

| Setting | Default | Off blocks |
| --- | --- | --- |
| `AI_ENABLED` | on | Thinking agents (summary, citations, friction moments, report), speaker labels, PII detection, Screenspace Text / Numbers OCR, and the LLM and Redact model downloads |
| `AI_TRANSCRIPTION_ENABLED` | on | Whisper transcription and Whisper model downloads |

Off never hides or deletes existing work. Transcripts, summaries, reports and speaker labels stay readable and editable. Stored redaction spans and manual redaction terms keep applying, so a switch can never un-redact a transcript. The deterministic friction scorer, corrections, marks, search and Embed Subtitles are unaffected. Deleting downloaded models still works.

## Decisions (maintainer, 2026-11-23)

- Transcription gets its own switch; everything else sits under `AI_ENABLED`.
- OCR is included under `AI_ENABLED`. No build-process change.
- Existing AI output stays visible while a switch is off.
- Toggling a switch reloads the page, with Settings reopened on General.
- Each switch has a "Remove downloaded models" action.
- The Start overlay asks once, on first run.

## Status

| Phase | Status |
| --- | --- |
| 1. Settings, gate helper, backend chokepoints | Done |
| 2. Route gates, orchestration, off-hook, CLI | Done |
| 3. Frontend shared gate + settings modal | Done |
| 4. Page surfaces (Transcripts, Overview, Screenspace, Workflows, palette) | Done |
| 5. Model removal action | Done |
| 6. First-run choice in the Start overlay | Done |
| 7. Tests, docs, version bump | Done |

### Changed while building

- Screenspace OCR tools stay selectable through indirect paths (restoring the last tool, editing an old task, Z/X cycling skips them). The picker items and the Run button are greyed. Blocking selection itself raised a notice on page load.
- The Model View preview for Text / Numbers stays open: it shows the preprocessed frame and runs no model.
- `--transcribe` is not refused: clip transcripts are cut from existing source transcripts, and only a fresh transcription warns. `--pre-transcribe`, `--summarize`, `--citations`, `--friction` and OCR `--ss-task` runs exit.
- A settings reset skips the AI group, so "Reset all" cannot turn a switch back on.
- A running Whisper download cannot be interrupted (it runs inside faster-whisper); LLM and Redact downloads stop within one chunk.
- A Workflows Multitool node with an OCR step is not flagged ahead of the run; the OCR chokepoint fails it at run time.

### Not verified

- Safari and the desktop app (WKWebView). The headless suite drives Chromium only.

## Design

### Backend

**Settings.** `AI_ENABLED` and `AI_TRANSCRIPTION_ENABLED` lead `STUDIO_SETTINGS` (General, group "AI", type `ai_switch`). `AI_CHOICE_MADE` (hidden) records that the first-run question was answered. Every dependent setting carries `"ai": "ai"` or `"ai": "transcription"`; `_settings_records` forwards it.

**Gate.** `utils.ai_blocked(kind)` answers whether a switch is off. Chokepoints sit at the lowest layer so no caller routes around them:

| Chokepoint | Switch | Covers |
| --- | --- | --- |
| `llm_client.generate` / `start_server` / `download_model` | ai | agents, workflow thinking nodes, LLM downloads |
| `redact.redact_entry`, `redact.download` | ai | detection and download; read-time `entry_spans` untouched |
| `speakers.diarize_entry` | ai | speaker pass everywhere |
| `screenspace_ocr._build_ocr_reader` | ai | the only OCR engine construction site |
| `transcripts._load_model`, `transcribe_video` | transcription | web queue, warmup, CLI, clip pipeline, workflows |

**Routes.** `server_utils.require_ai(kind)` answers 403 with `ai_disabled: <kind>` on routes that start model work. Read, stop, edit and delete routes stay open.

**Off-hook.** `_apply_settings_payload` cancels the matching queue, agents and downloads, stops `llama-server` and drops the cached Whisper model.

**CLI.** `cli.main` reads the two switches from `studio_settings.json`. AI flags then exit with the switch's message.

### Frontend

- `aiEnabled` / `aiTranscriptionEnabled` flow through `get_frontend_config()`. `render_index_html` also writes `html[data-ai-off]` / `html[data-transcription-off]` for first paint.
- Controls carry `data-ai="ai"` or `data-ai="transcription"`. Shared CSS greys them; one capture-phase click listener in `utils.js` stops the action and offers **Open Settings**, which opens General on the switch.
- Any API answer carrying `ai_disabled` shows the same notice.
- Settings modal: a switch card heads General; tagged rows render disabled under a note linking back to General.

## Verification

- `tests/test_ai_switch.py`: chokepoints, routes, off-hook, redaction still renders, CLI exits, source guards.
- `/ui-check`, then `shot.py` per page with each switch off.
- Safari and the desktop app need a manual pass.
