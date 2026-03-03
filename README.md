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

- Put `clipgen.py` or the `clipgen` binary in a folder with video recordings.
- Your Google `credentials.json` file should be in `~/.config/.gspread` or the working directory.
- Launch clipgen either interactively or through command-line arguments: ```python clipgen.py``` or ```python clipgen.py --help```
- Point clipgen to your Google Sheet and enjoy quick video clip generation based on your timestamped notes.
- Alternatively, point clipgen to a local Excel file in the current working directory and enjoy all the same features.

### How to use clipgen

- Can be used interactively or non-interactively, via command line argument calls.
- Several modes of generating timestamps are supported:
  - Batch
  - Single or multiple lines
  - Ranges
  - Categories
- Sheets can be browsed interactively through the program; no need to have a web browser always open.
- Clipgen can also generate highlight reels based on your input, combining multiple clips into a single video file.

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

### Runtime expectations for distributed binaries

- Place video files and `credentials.json` in the same folder as the executable.
- `ffmpeg` and `ffprobe` must be installed and available in `PATH`.
- The executable uses its own folder as working directory (so local files resolve consistently).

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

### About the spreadsheet

clipgen assumes that you are using a spreadsheet with a particular layout. A reference spreadsheet is [available here](https://docs.google.com/spreadsheets/d/1O51wnzRrYyz63tT6qy1HlJyVzdh9RT3t6QL5NohrcPc/edit?usp=sharing) - feel free to make a copy and use it in your studies.

Timestamps must be separated by characters ```+ , ;```
Ranges must be separated by character ```-```

You can optionally add a **baseline row for clock timestamps** directly above the participant header row:

- The participant header row still contains `P01`, `P02`, etc.
- The row **above** that can contain a clock timestamp for each participant column, e.g. `09:12:00` above `P01`.
- When a baseline cell is non-empty, all timestamps in that participant column are treated as **clock/absolute times** and are converted to **relative offsets** by subtracting the baseline time before cutting clips.
- When a baseline cell is empty, timestamps in that participant column are treated as **relative** (the current behavior).

## Timeline HTML Viewer

clipgen can generate an interactive HTML timeline viewer that visualizes all artifacts (clips, screenshots, GIFs) from a run.

- **CLI**: Pass `--viewer` alongside any mode flag to generate the viewer after clip processing:

``` shell
python clipgen.py -b --viewer
python clipgen.py -l 5+7 --screen --viewer
```

- **Interactive mode**: During an interactive session, clipgen keeps track of all generated artifacts. You can choose the `viewer` mode from the mode selection prompt to generate a timeline viewer for everything created so far in that session.

The viewer is a standalone `clips_viewer.html` file written to the same directory as the generated artifacts. Open it in any browser (works with `file://` — no server needed). It provides:

- A horizontal timeline showing all artifacts positioned by their timestamps.
- A filterable sidebar list sorted by time.
- Filters by category, participant, and artifact type.
- A detail panel with inline video/image preview.
- Spreadsheet cell references for each artifact.

The viewer assets (`viewer.js`, `viewer.css`) are copied alongside the HTML automatically.

## Possible future features

- GUI
- ~~Airtable support~~
- Title/ending cards
- Watermarking
- Subtitling
- Cropping and time-lapsing! For example generate a time-lapse of part of the screen, such as the minimap in a strategy game.

## AI Disclosure

The author has used an LLM coding agent for assistance in writing parts of this program; if you want to avoid software connected to LLMs, I get it. All code in this reposity prior to 2026 was written by a human, if you would like to fork the project.
