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
| `text` | OCR: detect or match text in region |
| `numbers` | Extract numeric values from region |
| `timelapse` | Generate a timelapse of the region |
| `template` | Match a template image within region |
| `flow` | Optical flow / movement direction |
| `inactivity` | Detect periods of no change |
| `scene` | Classify frames by similarity to one or more reference scenes |
| `multitool` | Chain several tools (temporal AND/NOT with offset windows), re-run only, see below |

### Tool-specific parameters

- **color**: `--ss-target-color "#RRGGBB" --ss-tolerance H,S,V`
- **change**: `--ss-threshold 0.0-1.0`
- **text**: `--ss-text "needle"` and/or `--ss-operator contains|equals|regex`
- **numbers**: `--ss-ranges "0-100,200-300"` (expected value ranges)
- **template**: requires a saved reference frame (set via UI first)
- **timelapse**: `--ss-interval SECONDS`
- **scene**: `--ss-scene-ref NAME:TIMESTAMP[:THRESHOLD]` (repeatable; TIMESTAMP in seconds, NAME has no `:`). Each reference frame is cropped from the region at TIMESTAMP. Optional per-scene THRESHOLD overrides `--ss-threshold`.

```
uv run clipgen.py --ss-task scene P01 myregion \
    --ss-scene-ref menu:12.5 --ss-scene-ref game:30:0.8 -i INPUT -o OUTPUT
```

### Re-run a saved task (multitool, scene, any type)

Multitool chains and complex scene setups are easiest to build in the `--screenspace`
UI. Once saved to the manifest, re-run any task headlessly by id. Reference frames are
re-extracted from the source video:

```
uv run clipgen.py --ss-list-tasks -i INPUT -o OUTPUT     # find the task id
uv run clipgen.py --ss-run-task ss_abcd1234 -i INPUT -o OUTPUT
```

Re-run creates a fresh task run (new id; the original is preserved). The one case that
can't be re-run is a step built from an **uploaded** template image (no timestamp was
saved to re-extract from). The CLI reports this and exits.

### List operations

```
uv run clipgen.py --ss-list-regions -i INPUT_DIR          # show defined regions
uv run clipgen.py --ss-list-tasks -i INPUT_DIR             # all tasks
uv run clipgen.py --ss-list-tasks completed -i INPUT_DIR   # filter by status
uv run clipgen.py --ss-list-stashes -i INPUT_DIR           # saved region stashes
```

## Cut clips from events

Turn the events you generated above into video clips, no UI required:

```
uv run clipgen.py --ss-clips [filters] [--cluster-gap N] [--clip-pre N --clip-post N] -i INPUT_DIR -o OUTPUT_DIR
```

Filters (all optional, comma-separated where listed):
- `--ss-clips-detector change,color` — restrict to specific tool types
- `--ss-clips-region dialog,header` — restrict to specific regions
- `--ss-clips-participant P01,P02` — restrict to specific participants
- `--ss-clips-min-confidence 0.8` — drop low-confidence events
- `--ss-clips-event-type "login"` — case-insensitive substring on `event_type`

Clustering & padding (defaults shown):
- `--cluster-gap 5.0` — merge events within N seconds into one clip; `0` = one clip per event
- `--clip-pre 5.0 --clip-post 5.0` — pad each cluster's start/end with N seconds
- `--max-clip-duration 0` — when `>0`, split clusters longer than N seconds

Examples:

```
# Cluster all change events into clips with default 5s gap and 5s pad
uv run clipgen.py --ss-clips --ss-clips-detector change -i INPUT -o OUTPUT

# One clip per event, no clustering
uv run clipgen.py --ss-clips --ss-clips-detector change --cluster-gap 0 -i INPUT -o OUTPUT

# High-confidence dialog events for one participant
uv run clipgen.py --ss-clips --ss-clips-region dialog --ss-clips-participant P01 \
    --ss-clips-min-confidence 0.8 -i INPUT -o OUTPUT
```

Generated clips are appended to the `clips` section of `clipgen.json` (synthetic artifact ids using negative cell rows so they never collide with spreadsheet-derived clips). Reruns with the same filters update existing entries idempotently.

## Export results

```
uv run clipgen.py --export -i INPUT_DIR -o OUTPUT_DIR
```

Produces `screenspace_export.json` and `screenspace_export.csv` from the `screenspace` section of `clipgen.json`.

## Notes

- Regions must be defined before running headless tasks (define via UI or import from a stash)
- Task statuses: `queued`, `running`, `completed`, `failed`, `cancelled`, `paused`
- Results are persisted in the `screenspace` section of `clipgen.json` in the output directory
