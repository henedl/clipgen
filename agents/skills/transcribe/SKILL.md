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

## Notes

- `config.DEBUGGING = True` returns stub transcripts without loading the Whisper model — useful for development
- Thinking agent results are stored in `transcripts_manifest.json` under each participant's entry
- Agents run in dependency order: `citations` depends on `summary` being populated first
- Ollama must be running (`ollama serve`) or clipgen will attempt to auto-start it
