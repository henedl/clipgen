# clipgen

clipgen is a program for quickly generating video clips, screenshots, and GIFs based on your research notes and recordings. It includes web-based interfaces for interactive clip generation (Studio), extracting and modifying transcripts (Transcripts), and video frame analysis (Screenspace).

The target audience for the program is user experience researchers and UX professionals who prefer to manage videos and analysis locally. The author intends specifically to support games user researchers conducting playtests.

clipgen is written in Python and interacts with a local video files through [ffmpeg](https://www.ffmpeg.org) and expects structured data in a Google Sheets document or local Excel file. clipgen can be run from source or a compiled binary.

## How to use

### Pre-requisites

1. To run from source: install [uv](https://docs.astral.sh/uv/) and run `uv sync` to install Python dependencies.
2. Install ffmpeg and ensure it is available in your `PATH`.
3. For Google Sheets: configure Google authentication per [gspread's setup guide](https://docs.gspread.org/en/master/oauth2.html). Place `credentials.json` in `./config/gspread/` or the same folder as `clipgen.py`.

### Starting clipgen

Place your video files in the same directory as the program and name them following this syntax: `{study}_{participant}.mp4` (e.g. `mystudy_P01.mp4`).

Then run the program interactively by launching the binary or:

```shell
uv run clipgen.py
```

clipgen can also be launched noninteractively, meaning you can script it as part of your workflows. For example:

```shell
uv run clipgen.py -R -H                # Highlight reel of the most severe issues
uv run clipgen.py -b -y                # Batch: all clips in study
uv run clipgen.py -l 5+7+12 -y         # Lines: rows 5, 7, 12
uv run clipgen.py -r 5-12 -y           # Range: rows 5–12
uv run clipgen.py -C "Onboarding" -y   # Category
uv run clipgen.py --gif -b -y          # GIFs instead of clips
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

clipgen assumes that you are using a spreadsheet with a particular layout. A reference spreadsheet is [available here](https://docs.google.com/spreadsheets/d/1O51wnzRrYyz63tT6qy1HlJyVzdh9RT3t6QL5NohrcPc/edit?usp=sharing) - feel free to make a copy and use it in your studies.

Timestamps must be separated by characters ```+ , ;```
Ranges must be separated by character ```-```

An optional `Baseline time` row supports clock/absolute timestamps: add a baseline timestamp per participant column (e.g. `09:12:00`), and clipgen automatically converts those timestamps to relative offsets before cutting clips. Columns without a baseline value use relative timestamps.

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

## Building from source

Cross-platform executables (macOS and Windows) are built automatically via GitHub Actions (`.github/workflows/build-binaries.yml`) on version tag pushes. To build locally with PyInstaller:

```shell
pip install pyinstaller
pyinstaller --clean --noconfirm build/clipgen.spec
```

Output: `dist/clipgen` (macOS) or `dist/clipgen.exe` (Windows).

## AI disclosure

The author has used an LLM coding agent for assistance in writing parts of this program. If you want to avoid software connected to LLMs; I get it. All code in this reposity prior to 2026 was written by a human, if you would like to fork the project.
