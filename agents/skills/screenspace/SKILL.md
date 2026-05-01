# clipgen-screenspace — Screenspace analysis workflow

## Launch the UI

```
uv run clipgen.py --screenspace -i INPUT_DIR -o OUTPUT_DIR
```

UI available at `http://127.0.0.1:8089/screenspace/`

## Headless CLI analysis

Run analysis tasks without the UI:

```
uv run clipgen.py --ss-task TYPE PARTICIPANT REGION [options] -i INPUT_DIR -o OUTPUT_DIR
```

### Available task types

| Type | What it measures |
|------|-----------------|
| `color` | Presence/prevalence of a target color |
| `change` | Frame-to-frame pixel change (motion) |
| `similarity` | Visual similarity to a reference frame |
| `text` | OCR — detect or match text in region |
| `numbers` | Extract numeric values from region |
| `timelapse` | Generate a timelapse of the region |
| `template` | Match a template image within region |
| `flow` | Optical flow / movement direction |
| `inactivity` | Detect periods of no change |

### Tool-specific parameters

- **color**: `--ss-target-color "#RRGGBB" --ss-tolerance H,S,V`
- **change**: `--ss-threshold 0.0-1.0`
- **text**: `--ss-text "needle"` and/or `--ss-operator contains|equals|regex`
- **numbers**: `--ss-ranges "0-100,200-300"` (expected value ranges)
- **template**: requires a saved reference frame (set via UI first)
- **timelapse**: `--ss-interval SECONDS`

### List operations

```
uv run clipgen.py --ss-list-regions -i INPUT_DIR          # show defined regions
uv run clipgen.py --ss-list-tasks -i INPUT_DIR             # all tasks
uv run clipgen.py --ss-list-tasks completed -i INPUT_DIR   # filter by status
uv run clipgen.py --ss-list-stashes -i INPUT_DIR           # saved region stashes
```

## Export results

```
uv run clipgen.py --export -i INPUT_DIR -o OUTPUT_DIR
```

Produces `screenspace_export.json` and `screenspace_export.csv` from `screenspace_manifest.json`.

## Notes

- Regions must be defined before running headless tasks (define via UI or import from a stash)
- Task statuses: `queued`, `running`, `completed`, `failed`, `cancelled`, `paused`
- Results are persisted in `screenspace_manifest.json` in the output directory
