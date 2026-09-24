# NVIDIA speech models: Nemotron 3.5 ASR, Parakeet v3, Nemotron 3 Diarization

## Summary

Add three NVIDIA speech models as optional downloads. Whisper stays the default.

| Model | Hugging Face repo | GGUF file | Size | Role |
| --- | --- | --- | --- | --- |
| Nemotron 3.5 ASR | `nvidia/nemotron-3.5-asr-streaming-0.6b` | `nemotron-3.5-asr-streaming-0.6b.q8_0.gguf` | 707 MB | Transcription, 40 language-locales, punctuation + capitalization |
| Parakeet TDT 0.6B v3 | `nvidia/parakeet-tdt-0.6b-v3` | `parakeet-tdt-0.6b-v3.q8_0.gguf` | 680 MB | Offline transcription, English + 25 European languages |
| Nemotron 3 Diarization | `nvidia/Nemotron-3-Diarization` | `Nemotron-3-Diarization.q8_0.gguf` | 102 MB | Speaker labels, up to 8 speakers |

Model licence is OpenMDW-1.1 (commercial use allowed). Sizes and filenames read from the Hugging Face tree API on 2026-09-24.

The model cards document only NeMo/PyTorch on Linux GPUs. The runtime this plan uses instead is **NeMo-Speech.cpp** (`github.com/NVIDIA/NeMo-Speech.cpp`, Apache-2.0), a llama.cpp-style C++ runtime whose `nemo-speech` CLI reads these GGUFs. No torch, NeMo, or new Python dependency.

"Drop-in" means: the app downloads a model on demand with a progress bar, and a GGUF the user copies into the ASR models dir counts as installed.

## Status

| Phase | Status |
| --- | --- |
| 0. Spike: verify the CLI contract | Not started |
| 1. Runtime pin + model downloads | Not started |
| 2. Transcription engine dispatch | Not started |
| 3. Nemotron diarization speaker engine | Not started |
| 4. Tests, docs, version bump | Not started |

Update each row as it lands, with any descoped item or changed decision.

## Facts about the runtime (verified 2026-09-24)

**Release v0.1.0** (2026-08-19), the first versioned release. Assets relevant to the bundle, each with a published `<asset>.sha256`:

| Platform | Asset | Size |
| --- | --- | --- |
| macOS arm64 | `nemo-speech-0.1.0-macos-aarch64-metal.tar.gz` | 3 MB |
| macOS arm64 (fallback) | `nemo-speech-0.1.0-macos-aarch64-cpu.tar.gz` | 3 MB |
| Windows x64 | `nemo-speech-0.1.0-windows-x86_64-vulkan.zip` | 20 MB |
| Windows x64 (alt) | `nemo-speech-0.1.0-windows-x86_64-cpu.zip` | 4 MB |
| Windows x64 (alt) | `nemo-speech-0.1.0-windows-x86_64-cuda.zip` | 101 MB |

Release URL pattern: `https://github.com/NVIDIA/NeMo-Speech.cpp/releases/download/v0.1.0/<asset>`.

**CLI, from `docs/cli.md`:**

- `nemo-speech transcribe <file> --model <name|path> --format json|srt|vtt|txt`. Word timestamps are on automatically for json/srt/vtt.
- `--diarize` adds speaker labels to transcription (1-based ids); `--diar-model <path>` picks the diarizer.
- `nemo-speech diarize <file> --format rttm --output <file>` runs diarization alone.
- `--device cuda:0|cpu|metal|vulkan:0` (alias `--backend`); auto-detected by default. `nemo-speech doctor` reports what it found.
- Input: mono or stereo PCM16 or float32 WAV, 8–96 kHz; downmixed and resampled internally. Other containers are refused.
- Progress and diagnostics go to stderr, so stdout is safe to parse.
- Local GGUF paths take precedence over its own downloads. Its own cache (`~/Library/Caches/NeMoSpeech/models`, `%LOCALAPPDATA%\NeMoSpeech\models`) is overridden by `NEMO_SPEECH_MODEL_DIR`.
- Aliases: `nemotron-3.5`, `nemotron-en`, `parakeet-v3` (offline only, refuses `--stream`/`--live`).
- Diarization presets: `--preset offline` (Sortformer v2), `--preset v3-offline` (Nemotron 3 Diarization). Threshold flags: `--onset`, `--offset`, `--pad-onset`, `--pad-offset`, `--min-duration-on`, `--min-duration-off`.

**Not documented** (Phase 0 answers these): the JSON field names, a language flag, a thread-count flag, long-audio memory behaviour, and the progress line format.

## Where clipgen assumes Whisper today

Verified against the tree on 2026-09-24.

- `source/transcripts.py`
  - `WHISPER_MODELS` is the only model catalog; `server.py`, `transcripts_server.py`, and the CLI prompt all read it.
  - `is_whisper_model_cached()` assumes the `Systran/faster-whisper-` repo prefix and reads the HF hub cache.
  - `_load_model()` builds a `WhisperModel` and caches it keyed on `_model_load_key()`.
  - `_build_transcribe_kwargs()` maps Whisper-only knobs (beam, VAD, thresholds, prompt, hotwords).
  - `transcribe_video()` decodes audio with `video.decode_audio_pcm(path, idx, start_seconds, duration_seconds)` (16 kHz mono float32), calls `model.transcribe`, derives segment bounds from word bounds, shifts by the window start, then calls `on_segment(end, seg)` and checks `cancel_flag` between segments.
  - `_flush_transcribe_profile()` tags `kind="whisper"`.
  - `label_speakers()` delegates to `speakers.diarize_entry()` with `config.TRANSCRIBE_SPEAKER_MAX`.
  - The `TranscriptWorker` `loading_model` phase calls `_load_model(task["model"])`.
- `source/transcripts_server.py`
  - `POST /api/transcribe/warmup` refuses with `reason: "model_not_cached"` via `is_whisper_model_cached`.
  - `POST /api/transcribe` refuses uncached models unless `allow_download`; it collects `override or TRANSCRIBE_MODEL` per participant.
  - `_speaker_model_ready()` gates the speaker pass on the bundled CAM++ model.
  - The only download-with-progress routes are `POST /api/models/llm/download` and `GET /api/models/llm/download-status`, backed by a `JobRegistry`.
- `source/server.py` `api_models()` returns `whisper.models` as `[{name, size_mb, description, selected, cached}]`.
- `source/cli_args.py` `--whisper-model` has hardcoded `choices`.
- `source/config.py` `STUDIO_SETTINGS` declares `TRANSCRIBE_MODEL` as `{"type": "model_select", "provider": "whisper"}`.
- Frontend:
  - `assets/web/settings-modal.js` `_loadModelsForSelect()` branches on `provider === "whisper"`.
  - `assets/web/transcripts-pills.js` builds the per-participant Model select from `data.whisper.models` and handles `model_not_cached`.
  - `assets/web/transcripts.js` `confirmModelInstall()` shows a progress bar only for LLM downloads (`downloadLlmModel()` polls `api/models/llm/download-status` each second). For Whisper it only confirms; the download happens inside the worker's `loading_model` phase.
  - `assets/web/overview-reports.js` also handles `model_not_cached`.

**Reusable pieces:**

- `llm_client.download_model(ref, on_progress)`: `urllib` download into a `.part` file, SHA256 checked against the Hugging Face LFS oid while streaming, `{status, completed, total}` progress, stale-part sweep. It writes into `llm_client.models_dir()` (`start_settings.config_dir() / "models"`), which the llama-server router scans.
- `llm_client._resolve_hf_file(ref)` resolves the repo tree and LFS sha256, but matches by quant, not by exact filename.
- `llm_client.resolve_server_bin()` is `shutil.which("llama-server")`. Frozen builds put the bundled `bin/` first on PATH, so the same lookup finds a bundled `nemo-speech`.
- `build/fetch_binaries.py` `PINS` holds per-platform archives with member hashes; `--repin` and `--check-urls` work off it. `build/clipgen.spec` derives its bundled-tool guard from `PINS`.
- `speakers.remap_speaker_ids()` and `speakers.speakers_block()` produce the manifest `speakers` block.

## 0. Spike: verify the CLI contract

No product code. Download the macOS metal asset by hand into a scratch dir under `.context/`, plus the three GGUFs. Record answers in this section.

Test inputs: `clipgen-test_P03.mp4` (the only fixture with real multi-voice speech) and one hour-long real session.

1. **JSON schema.** Capture `transcribe --format json` output for P03 with each ASR model. Save it as `tests/fixtures/nemo_speech/transcribe_p03.json`; the parser tests replay it.
2. **Segments.** Does the JSON carry segment boundaries, or only words? If only words, the parser groups words into segments (split on sentence-final punctuation or a gap over `TRANSCRIBE_SEGMENT_GAP`-style threshold; reuse an existing config knob if one fits).
3. **Long audio.** Peak RSS and wall time for 1 hour on Metal and on CPU, per model. Parakeet v3 is offline-only; decide if it needs chunking.
4. **Language.** Is there a flag for Nemotron 3.5? Is auto-detect the default? Is the detected language in the JSON?
5. **Progress.** Is stderr progress parseable (percent or timestamps)?
6. **Diarization.** Capture `diarize --format rttm --preset v3-offline` for P03 as `tests/fixtures/nemo_speech/diarize_p03.rttm`. Time it on 1 hour.
7. **Isolation.** With `--model /abs/path.gguf`, confirm no network access and no write to its own cache (run with networking off, watch `~/Library/Caches/NeMoSpeech`).
8. **Bundle.** List the tarball contents: binary, dylibs, `.metallib`. Confirm `--device metal` works when the binary runs from a directory other than the one it was extracted to (as it will inside the `.app`).
9. **Windows.** Does the vulkan build fall back to CPU without a Vulkan driver, as llama.cpp's does? If not, pin the CPU zip instead.

**Decision gate.** If (3) shows memory growth with audio length, or (5) shows no usable progress, chunk in Python: decode windows of about five minutes with `video.decode_audio_pcm`, transcribe each, offset the times. Chunking also makes cancel responsive. Record the choice here.

## 1. Runtime pin and model downloads

### Pin the binary

`build/fetch_binaries.py`:

- Add one archive entry to `PINS["macos-arm64"]` (metal asset) and one to `PINS["windows-x64"]` (vulkan or CPU, per spike item 9). List the binary and every library the spike found as `members`, each with a sha256. The archive hash comes from the published `.sha256`; `--repin` already checks that.
- Add a provenance paragraph to the module docstring, next to the llama.cpp one.
- If the `ggml-*` library names collide with llama.cpp's in the flat `bin/` dir, stop: either version-match the two or give nemo-speech its own subdir on PATH. The spike's file list decides this.

`build/clipgen.spec`: add the new binary and library globs to both `upx_exclude` lists. The tool guard derives from `PINS`, so it needs no edit.

`build/THIRD-PARTY-LICENSES`: add a SUMMARY row for NeMo-Speech.cpp (Apache-2.0) and its full text. Keep the fixed-width layout `source/licenses.py` parses; `tests/test_licenses.py` guards it. Note the OpenMDW-1.1 model licence only if models are ever bundled (this plan downloads them, so they are not redistributed).

`agents/skills/bump-pins/SKILL.md`: list the new pin.

### New module `source/nemo_speech.py`

Add it to `pyproject.toml` `[tool.setuptools] py-modules`; `tests/test_packaging.py` checks this. Imports only `config`, `utils`, `start_settings`, `llm_client` (for the downloader) at module level.

```text
NEMO_MODELS = {
    "nemotron-3.5": {repo, file, size_mb, description, kind: "asr"},
    "parakeet-v3":  {repo, file, size_mb, description, kind: "asr"},
    "nemotron-diarization": {repo, file, size_mb, description, kind: "diarizer"},
}
models_dir()        -> start_settings.config_dir() / "asr_models"
model_path(name)    -> models_dir() / NEMO_MODELS[name]["file"]
is_downloaded(name) -> model_path(name).is_file()
resolve_bin()       -> shutil.which("nemo-speech")
transcribe(wav, model, *, language, cancel_flag) -> list[TranscriptSegment] | None
diarize(wav, *, max_speakers, cancel_flag)       -> list[tuple[float, float, str]] | None
```

- The ASR models dir is separate from `llm_client.models_dir()` so the llama-server router never lists an ASR GGUF as a chat model.
- A user-copied file with the expected name counts as downloaded. That is the drop-in path.
- `transcribe()` and `diarize()` run the binary with `subprocess.Popen`, an absolute `--model` path, `--device` left on auto, and `NEMO_SPEECH_MODEL_DIR` set to `models_dir()` so a mistake never reaches its default cache. Poll `cancel_flag` and `terminate()` the process when set. Parse stdout JSON (or the RTTM file) in one function each, so a CLI format change touches one place.
- Under `config.DEBUGGING`, return stub results without running the binary, matching `transcripts.py`'s existing stubs.

### Downloads

Generalise `llm_client.download_model` rather than writing a second downloader:

- Add keyword arguments `dest_dir: Path | None = None` and `filename: str | None = None`. With `filename`, match the repo tree entry by exact path instead of by quant. Default behaviour is unchanged for every existing caller.
- `nemo_speech.download(name, on_progress)` calls it with `dest_dir=models_dir()` and the catalog filename.

Routes in `source/transcripts_server.py`, mirroring the LLM pair and listed in the module docstring's route table:

- `POST /api/models/asr/download` with `{name}`; starts a background download in its own `JobRegistry`.
- `GET /api/models/asr/download-status?name=`; returns the same `{status, completed, total, done, succeeded}` shape.

## 2. Transcription engine dispatch

One model string, no new engine setting: a `TRANSCRIBE_MODEL` (or per-participant override) found in `nemo_speech.NEMO_MODELS` with `kind == "asr"` runs through nemo-speech; anything else is Whisper.

`source/transcripts.py`:

- Add `is_model_downloaded(name)`: nemo names check `nemo_speech.is_downloaded`, others call `is_whisper_model_cached`. Switch the three existing callers (`transcripts.py`, `server.py` `api_models`, `transcripts_server.py` warmup and transcribe gates) to it.
- `_confirm_model_download()` looks up size from either catalog.
- `transcribe_video()`: keep the audio-stream probe, `_resolve_audio_index`, and window logic shared. Branch before `_load_model`:
  - Decode with `video.decode_audio_pcm`, write a temp 16 kHz mono float32 WAV (stdlib `wave` only does integer PCM, so write PCM16 after clipping, or write the float header by hand; choose in Phase 0).
  - Call `nemo_speech.transcribe`, shift times by the window start, run the existing energy edge-snap if enabled, and feed `on_segment` per segment.
  - Delete the temp WAV in a `finally`.
  - Return `TranscriptResult(segments, language, source_file, model=name)`.
- `_load_model()`, `warmup_transcription_model()`, and the worker's `loading_model` phase return early for nemo names: the subprocess loads the model per call.
- `_flush_transcribe_profile()` takes the kind as a parameter: `"whisper"` or `"nemo"`.
- `_build_transcribe_kwargs()` stays Whisper-only. Map only `TRANSCRIBE_LANGUAGE` to nemo-speech, if Phase 0 finds a flag.
- When the worker gets `allow_download` for an uncached nemo model, it downloads first under a `downloading` phase that reports progress, then transcribes.

`source/transcripts_server.py`: the warmup and transcribe gates already build `model_not_cached` payloads with `size_mb`; they only need `is_model_downloaded` and the shared size lookup. `GET /api/transcribe/model-status` reports `loaded: true` for nemo models so the prewarm UI stays quiet.

`source/server.py` `api_models()`: append the two ASR entries to `whisper.models` with an extra `engine: "nemo"` field and `cached` from `is_model_downloaded`. The existing pickers then list them without a new provider key.

`source/cli_args.py`: add `nemotron-3.5` and `parakeet-v3` to `--whisper-model` `choices` and its help string.

`source/config.py`: extend `TRANSCRIBE_MODEL` help text to name the NVIDIA options and say that Whisper-only knobs (beam, VAD, prompt, hotwords) do not apply to them.

Frontend (ES5, `.then()` chains):

- `assets/web/transcripts.js` `confirmModelInstall({kind: "whisper"})`: when the chosen model has `engine: "nemo"`, start `api/models/asr/download` and show the existing progress bar by reusing `downloadLlmModel()`'s polling with the ASR status URL. Pass the URL in rather than copying the function.
- `assets/web/settings-modal.js` and `assets/web/transcripts-pills.js`: show the catalog description and size for nemo entries, with an "NVIDIA" group label or suffix. Keep "(not downloaded)".
- User-facing strings stay within 10 words.

## 3. Nemotron 3 Diarization as a speaker engine

`source/config.py`: add `TRANSCRIBE_SPEAKER_ENGINE`, `"campp"` (default, bundled) or `"nemotron"` (downloaded on demand). Add it to `STUDIO_SETTINGS` beside `TRANSCRIBE_SPEAKERS` and to the help text. It is a plain config value; if the frontend needs the choices, send them through `utils.get_frontend_config()`.

`source/transcripts.py` `label_speakers()`:

- For `"nemotron"`, per video part: decode the whole part with `video.decode_audio_pcm`, write a temp WAV, call `nemo_speech.diarize`, offset spans by the part's start on the timeline (the same part offsets `speakers.diarize_entry` uses).
- Assign each segment the speaker with the largest total overlap. A segment with no overlapping span takes the nearest span's speaker.
- Cap distinct speakers at `min(TRANSCRIBE_SPEAKER_MAX, 8)`; merge extra labels into the nearest by overlap count.
- Build the result with `speakers.remap_speaker_ids()` and `speakers.speakers_block()`, so labels, renames, and the manifest shape stay the same.
- Report progress per part through `on_progress`.

`source/transcripts_server.py` `_speaker_model_ready()`: engine-aware. For `"nemotron"`, not downloaded gives the same 409 path, and the frontend offers the ASR download route with name `nemotron-diarization`.

The pure overlap-assignment function belongs in `speakers.py`, which stays free of subprocess code.

## 4. Tests, docs, version

Tests (no `nemo-speech` binary in CI; fake the subprocess):

- `tests/test_nemo_speech.py`:
  - JSON to `TranscriptSegment` parsing, replaying `tests/fixtures/nemo_speech/transcribe_p03.json`.
  - RTTM parsing, replaying `diarize_p03.rttm`.
  - `is_downloaded` true for a file copied into a tmp `models_dir`.
  - Cancel: a fake long-running process is terminated when `cancel_flag` flips.
  - `NEMO_SPEECH_MODEL_DIR` is set on the child env.
- `tests/test_transcripts.py`:
  - A nemo model name never calls `_load_model`'s `WhisperModel`.
  - Window offsets apply to returned segments.
  - `is_model_downloaded` routes by name.
- `tests/test_transcripts_api.py`, next to the existing LLM download tests:
  - An uncached nemo model returns `model_not_cached` with `size_mb`.
  - ASR download start and status routes.
  - `api_models` lists nemo entries with `engine` and `cached`.
- `tests/test_speakers.py` (or the existing speaker test file): overlap assignment, the 8-speaker cap, and a segment with no overlap.
- `tests/test_cli_args.py`: the new `--whisper-model` choices parse.
- `llm_client.download_model`: a `filename` + `dest_dir` case, and a check that the default path is unchanged.

Docs:

- `agents/ARCHITECTURE.md`: add a `nemo_speech.py` row and a sentence in the Transcription section.
- `agents/skills/transcribe/SKILL.md`: how to pick the NVIDIA models.
- `README.md`: one line on the optional NVIDIA models.

Version: this ships as `feat(transcripts): ...`, so bump the patch in `build/VERSION`. Leave `CHANGELOG.md` alone.

## Verification

1. `uv run build/fetch_binaries.py`, then run `nemo-speech doctor` from `build/vendor/macos-arm64/bin/`.
2. `uv run clipgen.py --transcripts -i <fixtures> -o <out>`:
   - Pick Nemotron 3.5 in Settings, confirm the download, watch the progress bar finish.
   - Transcribe P03. Check segments, word timing in the editor, marks, and search.
   - Repeat with Parakeet v3.
   - Turn on speakers with engine `nemotron`; confirm the download prompt, then distinct labels on P03.
   - Copy a GGUF into `asr_models/` by hand and confirm the picker drops "(not downloaded)".
3. CLI: `uv run clipgen.py ... --transcribe --whisper-model parakeet-v3`.
4. `/check` (ruff format, ruff check, ty, pytest). `/ui-check` on Transcripts; read the model-picker screenshots.
5. `/build` the `.app`, launch it like a user, transcribe one file with each engine, and confirm Metal loads inside the bundle (`-v` output names the device).
6. Windows: an untested platform until someone runs a CI-built bundle; record it in the PR test plan as `[ ]`.

## Risks and open questions

- **Runtime maturity.** NeMo-Speech.cpp is at v0.1.0; flags and JSON shape may change. Pin the exact release and keep parsing in one function per format.
- **Model choice guidance.** Nemotron 3.5 is a streaming model; on recorded sessions Parakeet v3 may be more accurate where its languages apply. Phase 0 should compare both on P03 and one real session, and the picker descriptions should reflect the result.
- **Shared ggml libraries.** llama.cpp and nemo-speech both ship ggml libraries into the same `bin/`. A name clash with different versions breaks one of them; see Phase 1.
- **Windows GPU.** The vulkan build's CPU fallback is unverified.
- **Bundle size.** The binaries add about 3 MB (macOS) and 4–20 MB (Windows). Models are never bundled.
