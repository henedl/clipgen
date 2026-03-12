# clipgen

clipgen is a Python program that uses the [gspread library](https://docs.gspread.org) and [ffmpeg](https://www.ffmpeg.org) to quickly generate video snippets based on timestamps in a Google Sheet or a local Excel file.

**Data flow:** Timestamps entered by you → spreadsheet → clipgen reads timestamped records (descriptions, study name, participant IDs, categories) → ffmpeg → clips or highlight reel.

The program was created to speed up data processing during playtests and is provided as-is, without promise of support. The target audience of this program are User Experience Researchers and UX professionals who prefer to manage their videos locally rather than in the cloud.

## How to use

### Pre-requisites

1. Install the required Python dependencies: ```pip install -r requirements.txt```
2. Install ffmpeg and ensure it is available via your `PATH`.
3. Configure your Google Authentication per [gspread's setup guide](https://docs.gspread.org/en/master/oauth2.html); clipgen requires you to have a Google Cloud project with API access, with a OAuth credentials file on your system.

### Starting clipgen

- Place video files in the same folder as the executable or `clipgen.py`.
- `ffmpeg` and `ffprobe` must be installed and available in `PATH`.
- Your Google `credentials.json` must be in `./config/gspread/`or same folder as executable.
- The executable uses its own folder as working directory (so local files resolve consistently).
- Launch clipgen either interactively or through command-line arguments.
- Select the Google Sheet or local Excel document you want to work on, and enjoy quick video clip generation based on your timestamped notes.

### Usage instructions

- Can be used interactively or non-interactively, via command line argument calls.
- Several modes of generating timestamps are supported:
  - Batch
  - Single or multiple lines
  - Ranges
  - Categories
- Sheets can be browsed interactively through the program; no need to have a web browser always open.
- Clipgen can also generate highlight reels based on your input, combining multiple clips into a single video file.

### Timeline viewer

clipgen can generate an interactive HTML timeline viewer that visualizes all artifacts (clips, screenshots, GIFs) from a run.

- **CLI**: Pass `--viewer` alongside any mode flag to generate the viewer after clip processing:

``` shell
python clipgen.py -b --viewer
python clipgen.py -l 5+7 --screen --viewer
```

- **Interactive mode**: During an interactive session, clipgen keeps track of all generated artifacts. You can choose the `viewer` mode from the mode selection prompt to generate a timeline viewer for everything created so far in that session.

The viewer is a standalone `clips_viewer.html` file written to the same directory as the generated artifacts. Open it in any browser (works with `file://` — no server needed). It provides:

- A timeline showing all artifacts positioned by their timestamps.
- A filterable sidebar list sorted by time.
- Filters by category, participant, and artifact type.
- A detail panel with inline video/image preview.

### About the spreadsheet

clipgen assumes that you are using a spreadsheet with a particular layout. A reference spreadsheet is [available here](https://docs.google.com/spreadsheets/d/1O51wnzRrYyz63tT6qy1HlJyVzdh9RT3t6QL5NohrcPc/edit?usp=sharing) - feel free to make a copy and use it in your studies.

Timestamps must be separated by characters ```+ , ;```
Ranges must be separated by character ```-```

You can optionally add a **baseline row for clock timestamps** using a single sheet-wide marker row:

- One row in the sheet should contain the label `Baseline time` in any cell; this row is treated as the baseline marker row.
- That same row can contain a clock timestamp for each participant column, e.g. `09:12:00` in the `P01` column.
- When a baseline cell is non-empty in the marker row, all timestamps in that participant column are treated as **clock/absolute times** and are converted to **relative offsets** by subtracting the per-column baseline time before cutting clips.
- When a baseline cell is empty in the marker row, timestamps in that participant column are treated as **relative**.
- If no `Baseline time` marker row exists at all, all participant columns are interpreted as using relative timestamps.

## Build single-file executable

clipgen can be packaged as a single-file executable with PyInstaller.

- Build each platform on that platform:
  - Build macOS binary on macOS
  - Build Windows `.exe` on Windows
- Install build dependency:
  - `pip install pyinstaller`
- Build using the included spec:
  - `pyinstaller --clean --noconfirm build/clipgen.spec`
- Output binaries:
  - macOS: `dist/clipgen`
  - Windows: `dist/clipgen.exe`

### CI build artifacts

- Cross-platform binaries are built by GitHub Actions workflow:
  - `.github/workflows/build-binaries.yml`
- Triggered on tag pushes matching `v*` and manual workflow dispatch.
- Artifacts are uploaded as:
  - `clipgen-macos`
  - `clipgen-windows`

## Testing

- Run the smoke test suite before releases and when adding features:
  - `pytest -c tests/pytest.ini`
- Contributor rule:
  - Every new CLI mode, flag, or selector should include at least one smoke test in the same PR.

## AI Disclosure

The author has used an LLM coding agent for assistance in writing parts of this program; if you want to avoid software connected to LLMs, I get it. All code in this reposity prior to 2026 was written by a human, if you would like to fork the project.
