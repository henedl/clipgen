# clipgen

clipgen is a program for quickly generating video clips, screenshots, and GIFs based on your research notes and recordings. It includes web-based interfaces for interactive clip generation (Studio), extracting and modifying transcripts (Transcripts), and video frame analysis (Screenspace).

The target audience for the program is user experience researchers and UX professionals who prefer to manage videos and analysis locally. The author intends specifically to support games user researchers conducting playtests.

clipgen is written in Python and interacts with local video files through [ffmpeg](https://www.ffmpeg.org) and expects structured data in a Google Sheets document or local Excel file. clipgen can be run from source or a compiled binary.

## How to use

### Pre-requisites

1. To run from source: install [uv](https://docs.astral.sh/uv/) and run `uv sync` to install Python dependencies.
2. Install ffmpeg and ensure it is available in your `PATH`.
3. For Google Sheets: configure Google authentication per [gspread's setup guide](https://docs.gspread.org/en/master/oauth2.html). Place `credentials.json` in `./config/gspread/` or the same folder as `clipgen.py`.

### Starting clipgen

Place your video files in the same directory as the program and name them following this syntax: `{study}_{participant}.mp4` (e.g. `mystudy_P01.mp4`).

If a participant's session spans **several video files** (a recording that broke off, or a diary study), name them with a numbered suffix (`mystudy_P01-1.mp4`, `mystudy_P01-2.mp4`, ...), and clipgen treats them as one continuous recording: a timestamp is mapped into the right file by cumulative duration (if file 1 is 1m20s long, a timestamp at `2:04` becomes `0:44` into file 2). The plain `mystudy_P01.mp4` takes precedence when it exists; numbered parts are used only when it is absent.

Then run the program interactively by launching the binary or:

```shell
uv run clipgen.py
```

clipgen can also be launched noninteractively, meaning you can script it as part of your workflows. For example:

```shell
uv run clipgen.py -H --no-input               # Highlight reel of the most severe issues
uv run clipgen.py -R "11, 13-16, P01" --no-input  # Custom reel from mixed selectors
uv run clipgen.py -b --no-input               # Batch: all clips in study
uv run clipgen.py -l 5+7+12 --no-input        # Lines: rows 5, 7, 12
uv run clipgen.py -r 5-12 --no-input          # Range: rows 5-12
uv run clipgen.py -C "Onboarding" --no-input  # Category
uv run clipgen.py --gif -b --no-input         # GIFs instead of clips
```

Run `uv run clipgen.py --help` for the full flag reference.

### Terminal modes

clipgen supports a range of generation modes, selectable interactively or via CLI flags:

- **Selection modes**: `batch` (-b), `line` (-l), `range` (-r), `category` (-C), `cell` (-c), `participant` (-p), `keyword` (-k), `severity` (-S)
- **Output formats**: clips (default), `--screen` (screenshots), `--gif` (GIFs)
- **Reels**: `reel` (-R), `chronologic` (-T), `highlights` (-H), `reellate` (interactive)
- **Gallery**: `--gallery VIDEO` — interval screenshots/GIFs with a browser-viewable gallery
- **Browse**: interactive terminal spreadsheet viewer (no output)

### About the spreadsheet

clipgen assumes that you are using a spreadsheet with a particular layout. A reference spreadsheet is [available here](https://docs.google.com/spreadsheets/d/1O51wnzRrYyz63tT6qy1HlJyVzdh9RT3t6QL5NohrcPc/edit?usp=sharing). Feel free to make a copy and use it in your studies.

Timestamps must be separated by characters ```+ , ;```
Ranges must be separated by character ```-```

An optional `Baseline time` row supports clock/absolute timestamps: add a baseline timestamp per participant column (e.g. `09:12:00`), and clipgen automatically converts those timestamps to relative offsets before cutting clips. Columns without a baseline value use relative timestamps.

An optional `Filename` row overrides the source video filename per participant column. To declare **multiple source videos** for a participant (one continuous timeline), list them plus-separated in this row — order matters: `morning.mp4 + afternoon.mp4`. This takes precedence over the on-disk `-N` auto-detection. A clip whose range straddles the boundary between two files is cut from both and stitched into a single clip.

## Features

### Studio - interactive artifact composing

clipgen can launch a web-based **Studio** interface for interactive artifact generation and reel building from your spreadsheet data. Studio opens in your browser and provides, among other things:

- An interactive spreadsheet grid with color-coded timestamp cells.
- Click cells to queue clips, screenshots, or GIFs for generation; shift-click or right-click cells for reel queue.
- Drag-to-reorder reel building from queued cells.
- Regenerate all artifacts from a saved manifest.

### Screenspace - run visual analysis to extract findings

clipgen includes **Screenspace** for analyzing video frames. Draw regions of interest and run automated analysis tasks to find patterns across the video.

Available analysis tools:

- **Color**: match a region's color
- **Change**: detect content changes
- **Similarity**: find frames matching a reference
- **Text**: OCR fuzzy search
- **Numbers**: OCR numeric comparison
- **Timelapse**: sped-up region video
- **Template**: pattern matching frame contents
- **Inactivity**: detect periods with no change

### Transcripts - generate transcripts using local models

clipgen features local **Transcripts**, generated via [faster-whisper](https://github.com/SYSTRAN/faster-whisper).

### Workflows - chain analyses on a node canvas

clipgen includes **Workflows**, a node-based canvas for scripting clipgen's capabilities without writing code. Drag blueprint cards, each wrapping one backend action (clip generation, Screenspace scans, transcription, thinking agents), onto a canvas and wire typed outputs into typed inputs to build a pipeline. clipgen executes the resulting graph in dependency order, with per-node status and inspectable results stored under `workflow_runs/` in the output directory. Built-in recipes ship as read-only stashes to start from.

A blueprint can run once, fan out across every participant in a study, or be armed to auto-run whenever a new session video lands in the input directory.

```shell
uv run clipgen.py --workflows -i ./videos -o ./out
```

### Overview - the birds-eye view of a study

The **Overview** frontend (reachable from the top navigation on any served page, at `/overview/`) gathers the cohort-level lenses in one place, as three tabs:

- **Map** renders every participant as a dot in 3D space, positioned so that spatial distance reflects behavioral similarity: computed from the timestamps in the spreadsheet (no clip generation needed), researcher marks and friction signals from Transcripts, Screenspace event rates, and session pacing. Clusters and outliers are visible at a glance; clicking a dot explains which signals set that participant apart (with deep links into Transcripts/Screenspace) and unfolds their actual moments as an in-scene burst plus a session timeline. Toggleable layers add similarity links between peers, shared category/detector anchors, and an all-moments point cloud; a session-replay control sweeps a playhead over the study and lets activity glow, mindwalk-style. Individual features can be muted from the lens, and the layout stays deterministic: the same data always produces the same map.
- **Convergence** aligns all participants' events on a shared timeline and highlights the moments where many participants do the same thing, with per-participant alignment offsets for misaligned recordings.
- **Metadata** shows aggregate statistics across every loaded session and stream.

Overview works in any launch mode; panels that need a spreadsheet show what still works without one.

```shell
uv run clipgen.py --overview -i ./videos -o ./out
```

## Third-party code

The web UIs are hand-written vanilla JavaScript with one exception: the Overview map vendors a single-file [Three.js](https://threejs.org) build (MIT) for WebGL rendering. See [assets/web/vendor/README.md](assets/web/vendor/README.md) for version and provenance. SVG icons are [Heroicons](https://heroicons.com) (MIT).

## Building from source

Cross-platform executables (macOS and Windows) are built automatically via GitHub Actions (`.github/workflows/build-binaries.yml`) on version tag pushes. To build locally with PyInstaller:

```shell
pip install pyinstaller
pyinstaller --clean --noconfirm build/clipgen.spec
```

Output: `dist/clipgen` and `dist/clipgen.app` (macOS), or `dist/clipgen.exe` (Windows). Double-clicking `clipgen.app` opens Terminal at the interactive CLI prompt; the raw binary inside is at `clipgen.app/Contents/MacOS/clipgen-bin`.

The macOS build is **unsigned** (no paid Apple Developer account). When downloaded from GitHub, macOS quarantines the app and Gatekeeper blocks the first launch. To clear the quarantine attribute, run:

```shell
xattr -dr com.apple.quarantine clipgen.app
```

Alternatively, right-click the `.app` in Finder and choose **Open** once to bypass Gatekeeper for that build.

## AI disclosure

The author has used an LLM coding agent for assistance in writing parts of this program. If you want to avoid software connected to LLMs; I get it. All code in this repository prior to 2026 was written by a human, if you would like to fork the project.
