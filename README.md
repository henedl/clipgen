# clipgen

clipgen is a Python program that uses [ffmpeg](https://www.ffmpeg.org) and a Google Sheet or local Excel file to quickly generate video clips, screenshots, and GIFs based on timestamps in your research notes. It also includes web-based interfaces for interactive clip generation (Studio), structured UX findings (Insights Builder), and video frame analysis (Screenspace).

The target audience is UX Researchers and professionals who prefer to manage playtest videos locally.

## How to use

### Pre-requisites

1. Install [uv](https://docs.astral.sh/uv/) and run `uv sync` to install Python dependencies.
2. Install ffmpeg and ensure it is available in your `PATH`.
3. For Google Sheets: configure Google authentication per [gspread's setup guide](https://docs.gspread.org/en/master/oauth2.html). Place `credentials.json` in `./config/gspread/` or the same folder as `clipgen.py`.

### Starting clipgen

Place your video files in the same directory as `clipgen.py`, named `{study}_{participant}.mp4` (e.g. `mystudy_P01.mp4`). Then run:

```shell
uv run clipgen.py
```

Or launch directly with a mode flag:

```shell
uv run clipgen.py -b                    # Batch: all clips
uv run clipgen.py -l 5+7+12            # Lines: rows 5, 7, 12
uv run clipgen.py -r 5-12              # Range: rows 5–12
uv run clipgen.py -C "Onboarding"      # Category
uv run clipgen.py --screen -b          # Screenshots instead of clips
uv run clipgen.py --gif -b             # GIFs instead of clips
```

Run `uv run clipgen.py --help` for the full flag reference.

### Modes

clipgen supports a range of generation modes, selectable interactively or via CLI flags:

- **Selection modes**: `batch` (-b), `line` (-l), `range` (-r), `category` (-C), `cell` (-c), `participant` (-p), `keyword` (-k), `severity` (-S)
- **Output formats**: clips (default), `--screen` (screenshots), `--gif` (GIFs)
- **Reels**: `reel` (-R), `chronologic` (-T), `highlights` (-H), `reellate` (interactive)
- **Gallery**: `--gallery VIDEO` — interval screenshots/GIFs with a browser-viewable gallery
- **Browse**: interactive spreadsheet viewer (no output)

### Timeline viewer

clipgen can generate an interactive HTML timeline viewer that visualizes all artifacts (clips, screenshots, GIFs) from a run.

- **CLI**: Pass `--viewer` alongside any mode flag to generate the viewer after clip processing:

``` shell
uv run clipgen.py -b --viewer
uv run clipgen.py -l 5+7 --screen --viewer
```

- **Interactive mode**: During an interactive session, clipgen keeps track of all generated artifacts. You can choose the `viewer` mode from the mode selection prompt to generate a timeline viewer for everything created so far in that session.

The viewer is a standalone `clips_viewer.html` file written to the same directory as the generated artifacts. Open it in any browser (works with `file://` — no server needed). It provides:

- A timeline showing all artifacts positioned by their timestamps.
- A filterable sidebar list sorted by time.
- Filters by category, participant, and artifact type.
- A detail panel with inline video/image preview.

### Studio

clipgen can launch a web-based Studio interface for interactive artifact generation and reel building from your spreadsheet data.

- **CLI**: Pass `--studio` to launch the Studio. A spreadsheet is required (Google Sheets or Excel):

``` shell
uv run clipgen.py --studio
uv run clipgen.py --studio -s "My Study"
```

The Studio opens in your browser at `http://127.0.0.1:8089/studio/` and provides:

- An interactive spreadsheet grid with color-coded timestamp cells.
- Click cells to queue clips, screenshots, or GIFs for generation; shift-click or right-click cells for reel queue.
- Format selection: video clip (.mp4), screenshot (.png), or GIF (.gif).
- Drag-to-reorder reel building from queued cells.
- Build a timeline viewer or export a manifest from generated artifacts.
- Regenerate all artifacts from a saved manifest.
- Dark/light theme toggle.

### Insights Builder

clipgen includes an Insights Builder for authoring structured UX research findings from generated artifacts. No spreadsheet is required — it reads artifacts from a previously saved manifest.

- **CLI**: Pass `--insights` to launch the Insights Builder:

``` shell
uv run clipgen.py --insights
uv run clipgen.py --insights -i ./output -o ./output
```

The Insights Builder opens in your browser at `http://127.0.0.1:8089/insights/` and provides:

- A media library sidebar showing all artifacts from `clipgen_manifest.json`, with filters by participant, category, severity, and type.
- Hover-to-scrub previews via auto-generated sprite sheets.
- Create, edit, and delete insights with structured sections: causes, behaviors, and impacts — each with narrative text and artifact references.
- Severity and status (draft/final) per insight.
- Export to a standalone `insights_viewer.html` file that shows finalized insights with embedded artifact references.

When `--studio` is used, the Insights Builder is also available via the `/insights/` path on the same server, so both interfaces share a single port.

### Screenspace

clipgen includes a Screenspace interface for analyzing video frames. Draw regions of interest and run automated analysis tasks to find patterns across the video.

- **CLI**: Pass `--screenspace` to launch. No spreadsheet required — clipgen discovers participant videos automatically:

``` shell
uv run clipgen.py --screenspace
uv run clipgen.py --screenspace -s "My Study"
```

Screenspace opens at `http://127.0.0.1:8089/screenspace/`. Available analysis tools: **Color** (match a region's color), **Change** (detect content changes), **Similarity** (find frames matching a reference), **Text** (OCR fuzzy search), **Numbers** (OCR numeric comparison), **Timelapse** (sped-up region video). Tasks run in a pausable background queue with drag-to-reorder.

### Manifest

clipgen can write a cumulative artifact manifest (`clipgen_manifest.json`) alongside generated clips. The manifest tracks all artifacts and reels across runs, and is required by the Insights Builder and the `--regenerate` flag.

- **CLI**: Pass `--manifest` alongside any mode flag to enable manifest writing:

``` shell
uv run clipgen.py -b --manifest
uv run clipgen.py -b --viewer --manifest
```

To regenerate all media artifacts from a saved manifest (no spreadsheet needed):

``` shell
uv run clipgen.py --regenerate
```

### About the spreadsheet

clipgen assumes that you are using a spreadsheet with a particular layout. A reference spreadsheet is [available here](https://docs.google.com/spreadsheets/d/1O51wnzRrYyz63tT6qy1HlJyVzdh9RT3t6QL5NohrcPc/edit?usp=sharing) - feel free to make a copy and use it in your studies.

Timestamps must be separated by characters ```+ , ;```
Ranges must be separated by character ```-```

An optional `Baseline time` row supports clock/absolute timestamps: add a baseline timestamp per participant column (e.g. `09:12:00`), and clipgen automatically converts those timestamps to relative offsets before cutting clips. Columns without a baseline value use relative timestamps.

## Building from source

Cross-platform executables (macOS and Windows) are built automatically via GitHub Actions (`.github/workflows/build-binaries.yml`) on version tag pushes. To build locally with PyInstaller:

```shell
pip install pyinstaller
pyinstaller --clean --noconfirm build/clipgen.spec
```

Output: `dist/clipgen` (macOS) or `dist/clipgen.exe` (Windows).

## Testing

- Run the smoke test suite before releases and when adding features:
  - `uv run pytest -c tests/pytest.ini`
- Contributor rule:
  - Every new CLI mode, flag, or selector should include at least one smoke test in the same PR.

## AI Disclosure

The author has used an LLM coding agent for assistance in writing parts of this program; if you want to avoid software connected to LLMs, I get it. All code in this reposity prior to 2026 was written by a human, if you would like to fork the project.

## Credits

- Icons: [Heroicons](https://heroicons.com/) (Micro set) by Tailwind Labs, Inc. — MIT License.
