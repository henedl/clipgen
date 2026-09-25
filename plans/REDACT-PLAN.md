# Redact: PII detection for transcripts (Desert Ant Labs)

## Summary

Add an opt-in PII pass to the Transcripts tool. It runs Desert Ant Labs' **Redact** model (23M-parameter multilingual token classifier, 27 languages including Swedish) over a participant's transcript, stores the detected spans beside the text, and renders `[GIVEN_NAME_1]`-style placeholders in the page, exports, subtitles and thinking-agent input while a toggle is on. The text itself is never rewritten.

The model is an opt-in download (~25 MB) into the config dir. It first ran on `ai-edge-litert` (Google's LiteRT Python package); `source/tflite_numpy.py` now evaluates the graph in numpy, so there is no runtime dependency.

## Status

| Phase | Status |
| --- | --- |
| 0. Spike: run the TFLite model from Python | Done (2026-09-24) |
| 1. `source/redact.py`, dependency, download | Done (2026-09-24) |
| 2. Worker task, routes, manifest | Done (2026-09-24) |
| 3. Frontend, exports, agents | Done (2026-09-24) |
| 4. Tests, docs, version bump | Done (2026-09-24) |
| 5. Replace `ai-edge-litert` with the numpy evaluator `tflite_numpy.py` | Done (2026-09-25): 644/644 argmax tags and identical spans vs LiteRT over 21 multilingual texts; ~50 ms per window vs ~6 ms |

Update each row as it lands, with any descoped item or changed decision.

## Facts (verified 2026-09-24)

**Model files** at `https://huggingface.co/desert-ant-labs/redact`, tag `v0.4.0`, resolve URL `https://huggingface.co/desert-ant-labs/redact/resolve/v0.4.0/<file>`:

| File | Size | sha256 |
| --- | --- | --- |
| `redact.tflite` | 24,529,472 B | `ee36727f07e3237569e71427bfe661463a82e526f7e97e92b0fa583cad16ed27` |
| `redact_tokenizer.bin` | 391,416 B | `81bc354dc99285e05b8f70edca9df3afbd83dfa4b0cea18d0448ab17ce58aae2` |
| `labels.json` | 3,842 B | `228835734235f85ec4bf31391a296aa6e8a36684e1171e841ac94f7b3f9b430a` |

- Architecture: `BertForTokenClassification`, 6 layers, hidden 384, vocab 31,475, int8 TFLite. 89 BIOES labels over 22 families: the 20 public labels plus `ORG` (off by default) and `IMEI` (regex-only in the vendor SDK; skipped here).
- Tensor contract: inputs `serving_default_input_ids` int32 `[1,256]`, `serving_default_attention_mask` int32 `[1,256]`; output `serving_default_logits_output` float32 `[1,256,89]`.
- Tokenizer: custom `RDTK` binary (magic, 1 version byte, `<4i` unk/bos/eos/count, `count` float32 scores, `count` uint16 piece lengths, UTF-8 pieces). Unigram Viterbi over NFKC text with spaces squeezed and replaced by `▁`, a leading `▁`, unknown penalty `min(score) - 10`. bos 0, pad 1, eos 2, unk 3.
- Windowing (vendor recommendation): 254 content tokens per window, stride 64 (step 190), `min_score` 0.6, tags under 0.3 forced to `O`.
- License: Desert Ant Labs Source-Available License 1.0 (`https://license.desertant.com/1.0`). Free below 100k monthly active devices per platform; app embedding allowed; a visible "Powered by Desert Ant Labs" credit is required. The model's own third-party notices credit the MIT Multilingual-MiniLM encoder.
- Runtime: `ai-edge-litert` 2.2.0 (PyPI), wheels `macosx_12_0_arm64`, `win_amd64`, `manylinux_2_27_{x86_64,aarch64}` for cp310–cp314; 11 MB (macOS) / 18 MB (Windows). Deps: `backports.strenum`, `flatbuffers`, `numpy`, `tqdm`, `typing-extensions`, `protobuf`, `ml_dtypes`.
- Reference implementation: `github.com/rm-hull/piitag` (MIT), a Python port of the vendor SDK. Its tokenizer, window loop, BIOES decode, dedupe, hysteresis and snap logic are ported here; its deterministic checksum layer, US-address regexes and title stripping are not.

## Spike results (Phase 0)

Run with `uv run --with ai-edge-litert==2.2.0` on macOS arm64, Python 3.12:

- Tokenizer parse 8 ms; interpreter load 34 ms; one 256-token window 6–10 ms (4 threads).
- 100 kB of text: tokenize 0.13 s (24,200 tokens), 128 windows in 0.86 s. Peak RSS 104 MB. No memoisation needed.
- A second `Interpreter` on the same file works (test smoke can build its own).
- Detection on "Hi, I'm Anna Lindqvist, you can reach me at anna.lindqvist@example.se or 070-123 45 67.": GIVEN_NAME, SURNAME, EMAIL, PHONE all at 1.00. Swedish sentence: name, surname, CITY (Umeå), EMAIL. All-lowercase, unpunctuated ASR-style text still finds name, surname and city (0.95–0.99).
- Native libs in the wheel: `libLiteRt.dylib`, `libpywrap_litert_common.dylib`, `libLiteRtMetalAccelerator.dylib`, and the `_pywrap_*.so` extension modules. PyInstaller needs `collect_dynamic_libs("ai_edge_litert")`.

## Design

- **Execution** mirrors the speakers pass: a `kind: "redact"` worker task (text-only), auto-enqueued after a transcription merges when wanted, or on demand from the pill. `config.TRANSCRIBE_REDACT` is the global default; a participant's explicit `redaction.enabled` beats it.
- **Storage** in `source_transcripts[pid]`: `redaction: {enabled, detected_at, min_score, org, count, error}` and per segment `pii: [{label, start, end, score, text}]` + `pii_crc` (crc32 of the corrected text the spans index). Disabling strips `pii`/`pii_crc`.
- **Corrections**: detection runs on corrected text. At read time a crc mismatch re-anchors spans by their surface and drops the ones that vanished; a text-changing correction route requeues a redact task.
- **Numbering** at read time, per participant, server-side: same `(label, normalised surface)` → same `[LABEL_N]`. The API ships `placeholder` per span with UTF-16 offsets.
- **Model delivery**: `redact.ASSETS` pins the three files; `redact.models_dir()` is `config_dir()/redact_models`; `redact.download()` streams each file through `llm_client.stream_download`. Routes `POST /api/models/redact/download`, `GET /api/models/redact/download-status`, `DELETE /api/models/redact`. Settings → Transcription → Redaction has the install block; the pill row disables with a hint while the model is missing.
- **Applies when on**: Transcripts page (chips keep word timing), `GET /api/transcript`, search, marks, VTT + embed-subtitles, `write_transcript` outputs, `data_export`, and the thinking-agent snapshot.

## Phases

### 1. `source/redact.py`, dependency, download

- `source/redact.py`: `ASSETS`, `models_dir`, `is_redact_model_available`, `download`, `remove`, tokenizer, windowed inference, BIOES decode, span algebra, `detect_spans` (DEBUGGING stub tags `@` tokens as EMAIL), read-time helpers (`entry_spans`, `number_spans`, `render_text`, `utf16_spans`), `redact_entry`.
- `source/llm_client.py`: lift the stream loop into `stream_download(url, target, *, sha256, size, on_progress)`.
- `pyproject.toml`: `ai-edge-litert`; `redact` in `py-modules`.
- `build/clipgen.spec`: collect `ai_edge_litert` submodules and dynamic libs, guarded.
- `build/THIRD-PARTY-LICENSES`: runtime rows and a Redact model section with the credit line.

### 2. Worker task, routes, manifest

- `source/transcripts.py`: segment keys carried everywhere `speaker` is; `redact_entry`; `create_redact_task`; worker `_execute_redact_task`; formatters honour `redact`.
- `source/transcripts_server.py`: helper block, `PUT/POST /api/redact/<pid>[/regenerate|/stop]`, download routes, read routes, merge, auto-enqueue on completion, requeue on corrections, agent snapshot.
- `source/config.py`, `source/utils.py`, `assets/web/utils.js`: `TRANSCRIBE_REDACT`, `TRANSCRIBE_REDACT_MIN_SCORE`, `TRANSCRIBE_REDACT_ORG`.
- `source/data_export.py`, `source/pipeline.py`, `source/server.py` (`api_models`).

### 3. Frontend, exports, agents

- `assets/web/transcripts-redact.js` satellite: the analysis panel's **Redact** tab (switch, run/stop, detected-items list, model download) plus the transcript chips; hub, pills badge, settings-modal download block, CSS, Start overlay credit, UI fixture. Decision 2026-09-25: the feature lives in a tab next to Summary and Friction, not in the pill dropdown; the model download is offered in both the tab and Settings. Polish pass the same day: Run CTA in the empty state like Summary/Friction, no pill badge while running, per-item restore (`redaction.excluded`, keyed label + normalised surface), muted chips (a solid dim tint, inline on the baseline; a dashed border and animated static were tried and dropped as distracting), and chip margins.

### 4. Tests, docs, version bump

- `tests/test_redact.py`, worker/API tests, packaging/constants/layering/licenses ratchets, docs rows, `build/VERSION` patch bump.

## Verification

Done 2026-09-24: full suite green, `/ui-check` green (chips render on the fixture's tagged line; Settings → Transcription shows the Redaction group and the model block), and an in-process run against the real model through the routes: enable → task → `[GIVEN_NAME_1] [SURNAME_1] from [CITY_1]`, `[EMAIL_1]`, `[PHONE_1]`, the second "Anna" reuses `_1`, VTT and search carry placeholders, a correction re-anchors the chips and requeues a pass, off restores the text. `redact.download()` fetched the three files from Hugging Face in under five seconds with sha256 verified. Still unexercised: the Settings Download button in a real browser, and a Windows machine.


- Unit: the transcripts, packaging, shared-constants, import-layering, satellite-wiring and licenses suites, then the full run with `--durations=20`.
- `/ui-check` with the seeded redacted segment; probe that the page text has no raw name while redaction is on.
- Manual with the real model: seed a `clipgen.json` transcript with a name, city, email and phone, download the model from Settings, flip the pill, check chips, hover, karaoke, corrections requeue, summary prompt, VTT, export, remove.

## Out of scope (this plan)

- Workflows Transcribe node `redact` param, CLI `--redact` flag, clip descriptions.

## Future features

- **Bleep or mute PII in the audio.** Spans map to word timings (`seg["words"]`; the chip carries `data-ws`/`data-we`). A `redact.pii_windows(entry)` helper yields padded, merged `(start, end)` seconds; clip and reel cuts in `source/video.py` take a mute list and apply `volume=0:enable='between(t,a,b)'` (or a sine bleep) in the same encode. Corrected segments lose word timing, so those spans fall back to the segment window.
- **Destructive source redaction.** An explicit "Scrub source video" action: re-encode each source part with the mute windows, replace the file, rewrite segment text with placeholders, drop `pii`, mark the entry `scrubbed`. Template: the remux keep/discard-original flow in `source/remux_server.py`. Typed confirmation; refused while any task, transcription or export runs on that participant; irreversible and the UI says so.

## Deferred: Align

Desert Ant's **Align** refines word-level timestamps. Files at `huggingface.co/desert-ant-labs/align`: `align-coarse.tflite` (512 KB), `align-fine.tflite` (511 KB), `calibrator.bin` (`ALGN` v1, gradient-boosted trees, 27 features), `mel_filters.bin` (40×257 float32), `refiner_config.json` (16 kHz, n_fft 512, win 400 periodic Hann, hop 160, 40 mels, log_eps 1e-6, coarse 241 frames, fine 81 frames, byte_context 16, pad_byte 256, languages de/en/es/fr/it/ja/ko/pt/zh). Stage contract: `mel[16,1,40,width]` f32, `text_bytes[16,32]` i32, `language_id[16]` i32, `boundary_kind[16]` i32 → `logits[16,width]`; batch 16 with tail rows repeating; z-score per utterance then per crop. The Swift `refine(words, audio, sampleRate, languageCode)` entry accepts any text/start/end word list, so a Python port (~500 lines) is possible; it would need `tflite_numpy` to grow the conv/mel ops or bring LiteRT back. No Swedish. Revisit only if Whisper's word timing becomes a felt problem.
