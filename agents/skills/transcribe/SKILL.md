# clipgen-transcribe — Transcription and thinking agent workflow

## Step 1: Pre-transcribe source videos

```
uv run clipgen.py --pre-transcribe [P01 P03 ...] -i INPUT_DIR -o OUTPUT_DIR
```

Omit participant IDs to transcribe all participants. This creates `.md` transcript files (or `.srt`/`.vtt` with `--transcript-format`).

Options:
- `--whisper-model tiny|base|small|medium|large-v3` (default: `base`)
- `--transcript-format md|srt|vtt` (default: `md`)
- `--no-whisper-vad` — disable Silero VAD when it is enabled (`TRANSCRIBE_VAD_FILTER` is off by default; turn VAD on in Studio or `config.py` when you want speech-only decoding on long silent recordings)
- `--whisper-hallucination-silence SEC` — enable hallucination silence skip when SEC > 0 (enables word timestamps; slower)

Transcription quality knobs (`TRANSCRIBE_VAD_FILTER`, no-speech / log-probability / compression-ratio thresholds, hallucination silence threshold, condition-on-previous-text) live in `config.py` and are exposed in Studio under **Transcription → Transcription quality**. Set `TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD` to `0` to disable silence-based hallucination skip.

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

## Step 4: Run thinking agents (requires Ollama)

Summaries (Pass 1 — paragraph + bullets):
```
uv run clipgen.py --summarize [P01 P03 ...] -i INPUT_DIR -o OUTPUT_DIR
```

Citations (Pass 2 — requires summaries to exist first):
```
uv run clipgen.py --citations [P01 P03 ...] -i INPUT_DIR -o OUTPUT_DIR
```

To use a specific Ollama model: `--ollama-model MODEL`

Friction detection (depends on `summary`; surfaces moments of likely interest):
- A deterministic scorer (`friction.py`) flags hesitation / confusion / frustration /
  surprise / self-correction / help-seeking per segment; a local LLM then refines the
  top candidates into ~5 "moments".
- Off by default. Enable auto-run after summaries with `OLLAMA_FRICTION_ENABLED = True`
  in `config.py` (model: `OLLAMA_FRICTION_MODEL`, defaults to `OLLAMA_SUMMARY_MODEL`).
- Drive it from the **Transcripts UI**: the **Friction tab** in the analysis panel (stats,
  score/category filter, top moments, "Mark all matching") and the **friction heatmap**
  toggle on the timeline; run/re-run/stop also live in the participant pill dropdown.
- Results persist to `transcripts_manifest.json` under each participant's `friction` field
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

Clustering & padding (defaults shown): same as `--ss-clips` —
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

When a mark filter is set, the clip's category is `mark-{category}`; otherwise `transcript`. Clips are appended to `clipgen_manifest.json`.

## Step 6: Batch-mark segments by text term

Mark every transcript segment whose text contains a term — same effect as the Transcripts UI's search box + "Mark all results" button, but headless:

```
uv run clipgen.py --transcript-mark TERM --transcript-mark-category CAT [filters] -i INPUT_DIR -o OUTPUT_DIR
```

Required:
- `--transcript-mark "checkout flow"` — word or phrase to find (quote multi-word terms). Case-insensitive substring match against corrected segment text.
- `--transcript-mark-category insight` — one of `pain_point`, `delight`, `quote`, `insight`, `task`, `bookmark`.

Optional:
- `--transcript-mark-participant P01,P02` — restrict to specific participants. Omit to mark across all transcripts.
- `--transcript-mark-label "follow up"` — label written onto every created/updated mark.

Existing marks on matching segments are updated in place (category and, if given, label). Non-matching segments' marks are untouched. The created marks live in `transcripts_manifest.json` and are immediately consumable by `--transcript-clips --transcript-clips-mark CAT`.

Example:

```
# Mark every "checkout" mention as an insight, then cut clips for them
uv run clipgen.py --transcript-mark checkout --transcript-mark-category insight -i IN -o OUT
uv run clipgen.py --transcript-clips --transcript-clips-mark insight -i IN -o OUT
```

## Notes

- `config.DEBUGGING = True` returns stub transcripts without loading the Whisper model — useful for development
- Thinking agent results are stored in `transcripts_manifest.json` under each participant's entry
- Agents run in dependency order: both `citations` and `friction` depend on `summary` being populated first (friction runs after summary even when citations is disabled)
- Ollama must be running (`ollama serve`) or clipgen will attempt to auto-start it
