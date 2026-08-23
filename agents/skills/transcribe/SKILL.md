# clipgen-transcribe — Transcription and thinking agent workflow

## Step 1: Pre-transcribe source videos

```
uv run clipgen.py --pre-transcribe [P01 P03 ...] -i INPUT_DIR -o OUTPUT_DIR
```

Omit participant IDs to transcribe all participants. This creates `.md` transcript files (or `.srt`/`.vtt` with `--transcript-format`).

Options:
- `--whisper-model tiny|base|small|medium|large-v3` (default: `base`)
- `--transcript-format md|srt|vtt` (default: `md`)
- `--no-whisper-vad` — disable Silero VAD (`TRANSCRIBE_VAD_FILTER` is **on by default**; VAD skips long silence on recordings that are mostly quiet, which is the common UX-research case)
- `--whisper-hallucination-silence SEC` — enable hallucination silence skip when SEC > 0 (enables word timestamps; slower)

Transcription quality knobs (`TRANSCRIBE_BEAM_SIZE`, `TRANSCRIBE_VAD_FILTER` + its recall-safe tuning `TRANSCRIBE_VAD_THRESHOLD`/`TRANSCRIBE_VAD_SPEECH_PAD_MS`/`TRANSCRIBE_VAD_MIN_SILENCE_MS`, no-speech / log-probability / compression-ratio thresholds, hallucination silence threshold, condition-on-previous-text) live in `config.py` and are exposed in Studio under **Transcription → Transcription quality**. If VAD ever drops real words, lower `TRANSCRIBE_VAD_THRESHOLD` (e.g. `0.2`) or raise `TRANSCRIBE_VAD_SPEECH_PAD_MS` rather than turning VAD off. Set `TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD` to `0` to disable silence-based hallucination skip. `TRANSCRIBE_CPU_THREADS` (Studio: **Transcription**) sets CTranslate2 CPU threads; `0` = auto (all cores).

## Step 2: Generate clips with transcripts

Add `--transcribe` to any generation command:

```
uv run clipgen.py -b --transcribe -s "Study" -i INPUT_DIR -o OUTPUT_DIR
```

## Step 3: Launch transcript workspace

```
uv run clipgen.py --transcripts -i INPUT_DIR -o OUTPUT_DIR
```

UI at `http://127.0.0.1:8089/transcripts/`

## Step 4: Run thinking agents (requires the local AI server)

Summaries (Pass 1: paragraph + bullets):
```
uv run clipgen.py --summarize [P01 P03 ...] -i INPUT_DIR -o OUTPUT_DIR
```

Citations (Pass 2: requires summaries to exist first):
```
uv run clipgen.py --citations [P01 P03 ...] -i INPUT_DIR -o OUTPUT_DIR
```

To use a specific AI model: `--llm-model MODEL`

Friction detection (depends on `summary`; surfaces moments of likely interest):
- A deterministic scorer (`friction.py`) flags hesitation / confusion / frustration /
  surprise / self-correction / help-seeking per segment; a local LLM then refines the
  top candidates into ~5 "moments".
- Off by default. Enable auto-run after summaries with `LLM_FRICTION_ENABLED` (also
  in the Studio settings **Summaries → AI Summary** tab). Friction uses the same
  model as summaries/citations (`LLM_SUMMARY_MODEL`); set `LLM_FRICTION_MODEL` only
  to pin friction to a different model (blank = follow the summary model).
- Drive it from the **Transcripts UI**: the **Friction tab** is a control surface over the
  transcript below it — an `Off / Highlight / Isolate` mode switch, a score histogram whose
  marker is the threshold, category chips, "Mark all matching", and a jump strip of the LLM
  moments (whose rationales render as inline callouts under the segments they quote).
  Highlight tints matching segments and draws the timeline density band; Isolate hides
  everything else. Run/re-run/stop also live in the participant pill dropdown.
- Results persist to the `transcripts` section of `clipgen.json` under each participant's `friction` field
  and flow into `--export` (`clipgen_export_friction_moments.*` / `_friction_segments.*`).

## Step 5: Cut clips from transcript segments or marks

Turn transcript segments into video clips, no UI required:

```
uv run clipgen.py --transcript-clips [filters] [--cluster-gap N] [--clip-pre N --clip-post N] -i INPUT_DIR -o OUTPUT_DIR
```

Filters (all optional, comma-separated where listed):
- `--transcript-clips-participant P01,P02` — restrict to specific participants
- `--transcript-clips-mark insight,action` — only segments tagged with these mark categories (set via the Transcripts UI)
- `--transcript-clips-text "checkout"` — case-insensitive substring on segment text

Clustering & padding (defaults shown): same as `--ss-clips`:
`--cluster-gap 5.0`, `--clip-pre 5.0`, `--clip-post 5.0`, `--max-clip-duration 0`.

Examples:

```
# Clips for every segment marked as an "insight"
uv run clipgen.py --transcript-clips --transcript-clips-mark insight -i INPUT -o OUTPUT

# Text-search clips
uv run clipgen.py --transcript-clips --transcript-clips-text "checkout flow" -i INPUT -o OUTPUT

# One clip per segment, no clustering
uv run clipgen.py --transcript-clips --transcript-clips-participant P01 --cluster-gap 0 -i INPUT -o OUTPUT
```

When a mark filter is set, the clip's category is `mark-{category}`; otherwise `transcript`. Clips are appended to the `clips` section of `clipgen.json`.

## Step 6: Batch-mark segments by text term

Mark every transcript segment whose text contains a term, same effect as the Transcripts UI's search box + "Mark all results" button, but headless:

```
uv run clipgen.py --transcript-mark TERM --transcript-mark-category CAT [filters] -i INPUT_DIR -o OUTPUT_DIR
```

Required:
- `--transcript-mark "checkout flow"` — word or phrase to find (quote multi-word terms). Case-insensitive substring match against corrected segment text.
- `--transcript-mark-category insight` — one of `pain_point`, `delight`, `quote`, `insight`, `task`, `bookmark`.

Optional:
- `--transcript-mark-participant P01,P02` — restrict to specific participants. Omit to mark across all transcripts.
- `--transcript-mark-label "follow up"` — label written onto every created/updated mark.

Existing marks on matching segments are updated in place (category and, if given, label). Non-matching segments' marks are untouched. The created marks live in the `transcripts` section of `clipgen.json` and are immediately consumable by `--transcript-clips --transcript-clips-mark CAT`.

Example:

```
# Mark every "checkout" mention as an insight, then cut clips for them
uv run clipgen.py --transcript-mark checkout --transcript-mark-category insight -i IN -o OUT
uv run clipgen.py --transcript-clips --transcript-clips-mark insight -i IN -o OUT
```

## Notes

- `config.DEBUGGING = True` returns stub transcripts without loading the Whisper model, useful for development
- Thinking agent results are stored in the `transcripts` section of `clipgen.json` under each participant's entry
- Agents run in dependency order: both `citations` and `friction` depend on `summary` being populated first (friction runs after summary even when citations is disabled)
- `llama-server` must be installed (brew or bundled); clipgen auto-starts it on demand
