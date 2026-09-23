# clipgen-generate — Translate intent into a clipgen CLI command

## Required arguments (always needed)

- `-s SPREADSHEET` — Google Sheet title, index, URL, or path to local `.xlsx` file
- `-i INPUT_DIR` — directory containing source video files (default: current dir)
- `-o OUTPUT_DIR` — directory for generated artifacts (default: current dir)

## Selection modes (mutually exclusive, pick exactly one)

| Intent | Flag |
|--------|------|
| All clips | `-b` / `--batch` |
| Specific row numbers | `-l 5` or `-l 1+4+5` or `-l 1,4,5` |
| Row range | `-r 3-10` |
| By category | `-C "Usability"` or `-C "Cat1,Cat2"` |
| By participant | `-p P01` or `-p P01+P03` |
| Specific cell | `-c P01.11` or `-c P01.11+P03.11` |
| Annotated only | `-k` or `-k !key` |
| By severity | `-S Critical` or `-S Critical,High` |
| Highlights reel | `-H` or `-H 120` (seconds budget) |
| Multiple individual outputs | `-M "5, P01.11, 13-16"` |
| Combined reel from selectors | `-R "5, P01.11, 13-16"` |
| Chronological reel | `-T P01` |

## Output format (mutually exclusive, default is .mp4)

| Output | Flag |
|--------|------|
| Video clips (.mp4) | _(default, no flag)_ |
| Screenshots (.png) | `--screen` |
| Animated GIFs | `--gif` |

## Other useful flags

- `--no-input` — non-interactive mode: skip confirmation prompts and fail fast on prompts that would block on stdin (good for agent automation)
- `--manifest` — write the `clips` section of `clipgen.json` alongside artifacts
- `--viewer` — generate `clips_viewer.html` timeline viewer
- `--transcribe` — generate transcript files alongside artifacts
- `--titlecards` / `--no-titlecards` — prepend title card to each clip

## Video file naming convention

Source videos must be named `{study}_{participant}.mp4`, e.g. `mystudy_P01.mp4`. The study name comes from the spreadsheet, lowercased and filesystem-safe. Before running, confirm the files exist in `INPUT_DIR`.

## Example translations

| Natural language | Command |
|-----------------|---------|
| "clips for P01 and P03" | `uv run clipgen.py -p P01+P03 -s "Study" -i . -o ./clips` |
| "screenshots of rows 3 through 7" | `uv run clipgen.py --screen -r 3-7 -s "Study" -i . -o ./clips` |
| "everything, no prompts" | `uv run clipgen.py -b --no-input -s "Study" -i . -o ./clips` |
| "highlight reel, 2 minutes" | `uv run clipgen.py -H 120 -s "Study" -i . -o ./clips` |
| "annotated clips only" | `uv run clipgen.py -k -s "Study" -i . -o ./clips` |
| "critical and high severity" | `uv run clipgen.py -S Critical,High -s "Study" -i . -o ./clips` |
| "GIFs for the usability category" | `uv run clipgen.py --gif -C "Usability" -s "Study" -i . -o ./clips` |

## Event-driven clips (no spreadsheet)

When the user wants to cut clips from existing Screenspace events or transcript segments instead of a spreadsheet, use the manifest-driven modes, no `-s` needed:

| Natural language | Command |
|-----------------|---------|
| "clips of every change event in the dialog region" | `uv run clipgen.py --ss-clips --ss-clips-detector change --ss-clips-region dialog -i . -o ./clips` |
| "clips for every segment marked as an insight" | `uv run clipgen.py --transcript-clips --transcript-clips-mark insight -i . -o ./clips` |
| "clips wherever P01 said 'checkout flow'" | `uv run clipgen.py --transcript-clips --transcript-clips-participant P01 --transcript-clips-text "checkout flow" -i . -o ./clips` |

See `agents/skills/screenspace/SKILL.md` and `agents/skills/transcribe/SKILL.md` for the full filter and clustering reference.
