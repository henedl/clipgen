# clipgen-transcribe — Transcription and thinking agent workflow

## Step 1: Pre-transcribe source videos

```
uv run clipgen.py --pre-transcribe [P01 P03 ...] -i INPUT_DIR -o OUTPUT_DIR
```

Omit participant IDs to transcribe all participants. This creates `.md` transcript files (or `.srt`/`.vtt` with `--transcript-format`).

Options:
- `--whisper-model tiny|base|small|medium|large-v3` (default: `base`)
- `--transcript-format md|srt|vtt` (default: `md`)

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

## Notes

- `config.DEBUGGING = True` returns stub transcripts without loading the Whisper model — useful for development
- Thinking agent results are stored in `transcripts_manifest.json` under each participant's entry
- Agents run in dependency order: `citations` depends on `summary` being populated first
- Ollama must be running (`ollama serve`) or clipgen will attempt to auto-start it
