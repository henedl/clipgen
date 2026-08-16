# clipgen

clipgen turns user research recordings into video clips, screenshots, GIFs, and highlight reels, driven by the timestamps in your research notes (a Google Sheet, a local Excel file, or MindNode mind map). Around that core it bundles interactive tools — served locally in your browser or the desktop app — for clip building, visual analysis, transcription, freeform cutting, automation, and study-level overviews.

It is built for UX researchers — especially games user researchers running playtests — who prefer to keep videos and analysis on their own machine. Everything runs locally: [ffmpeg](https://www.ffmpeg.org) for media, [faster-whisper](https://github.com/SYSTRAN/faster-whisper) for transcription, and optionally [Ollama](https://ollama.com) for AI features.

## Quick start

### Desktop app

Download the macOS `.dmg`, or on Windows the `-setup.exe` installer (recommended — per-user install with a Start Menu shortcut and uninstaller, no admin rights needed) or the portable `.zip`, from the [Releases page](https://github.com/henedl/clipgen/releases) and follow the bundled `INSTALL.txt`. ffmpeg is included (GPL builds — see `THIRD-PARTY-LICENSES` in the download); there is nothing else to install.

The macOS build is unsigned, so Gatekeeper blocks the first launch: right-click the app and choose **Open** once, or run `xattr -dr com.apple.quarantine clipgen.app`. Double-clicking the app opens clipgen in its own desktop window; the same binary is the full CLI when given arguments.

### From source

1. Install [uv](https://docs.astral.sh/uv/) and run `uv sync`.
2. Install ffmpeg: `scripts/install-ffmpeg-ollama.sh` (macOS/Linux) or `scripts/install-ffmpeg-ollama.bat` (Windows) does it for you, or use `brew install ffmpeg`, `sudo apt install ffmpeg`, or `winget install Gyan.FFmpeg`.
3. Run `uv run clipgen.py`.

### Your videos

Name recordings `{study}_{participant}.mp4` (e.g. `mystudy_P01.mp4`) and place them next to the program. If a session spans several files, add a numbered suffix (`mystudy_P01-1.mp4`, `mystudy_P01-2.mp4`, ...) and clipgen treats them as one continuous recording.

### Google Sheets (optional)

Only reading a Google Sheet needs Google auth: follow [gspread's setup guide](https://docs.gspread.org/en/latest/oauth2.html) and save the OAuth client file as `credentials.json` next to the app, in `~/.config/gspread/`, or in clipgen's config dir (`~/.config/clipgen/`; `%LOCALAPPDATA%\clipgen\` on Windows). Local Excel files need no credentials, and Screenspace, Transcripts, Composer, and Workflows need no spreadsheet at all.

### Local AI (optional)

The Transcripts thinking agents and Overview reports use a local [Ollama](https://ollama.com/download) server; everything else works without it. clipgen offers to install Ollama and pull the models the first time you use an AI feature, and starts `ollama serve` for you. Transcription itself does not need Ollama.

## The spreadsheet

clipgen expects a particular layout — make a copy of the [reference spreadsheet](https://docs.google.com/spreadsheets/d/1O51wnzRrYyz63tT6qy1HlJyVzdh9RT3t6QL5NohrcPc/edit?usp=sharing) to start. Timestamps in a cell are separated by `+`, `,`, or `;`; ranges use `-` (e.g. `1:23-1:45`).

Two optional rows: a `Baseline time` row converts clock timestamps (e.g. `09:12:00`) to relative offsets per participant, and a `Filename` row overrides the source video per participant (plus split videos).

## Command line

Launch with no flags for the interactive prompt, or script it:

```shell
uv run clipgen.py -H --no-input                      # Highlight reel of the most severe issues
uv run clipgen.py -b --no-input                      # Batch: every clip in the study
uv run clipgen.py --gif -C "Onboarding" --no-input   # GIFs for one category
```

Selection modes: batch (`-b`), line (`-l`), range (`-r`), category (`-C`), cell (`-c`), participant (`-p`), keyword (`-k`), severity (`-S`). Outputs: clips (default), `--screen`, `--gif`. Reels: custom (`-R`), chronological (`-T`), highlights (`-H`). `--gallery VIDEO` builds a browsable interval-screenshot gallery. Run `uv run clipgen.py --help` for the full reference.

## Tools

- **Studio** (`--studio`) — click timestamp cells in your spreadsheet to generate clips, screenshots, and GIFs, and build reels interactively.
- **Screenspace** (`--screenspace`) — automated visual analysis of recordings: draw a region and detect color, change, text, motion, and more across the video.
- **Transcripts** (`--transcripts`) — local transcription, with optional AI summaries.
- **Composer** (`--composer`) — cut and annotate moments that aren't in the spreadsheet, on a timeline over the full recording.
- **Workflows** (`--workflows`) — chain the tools above into automated pipelines on a node canvas.
- **Overview** (`--overview`) — study-level views: a shared participant timeline, aggregate stats, and per-participant reports.

Tools that read videos directly take input/output directories:

```shell
uv run clipgen.py --composer -i ./videos -o ./out
```

## Third-party code

Icons are [Heroicons](https://heroicons.com) and GitHub's [Octicons](https://primer.style/octicons/), both MIT-licensed.

clipgen's own source is MIT. The desktop builds bundle third-party software under MIT, BSD, Apache-2.0, HPND, MPL-2.0, LGPL, and GPL terms — notably the pinned ffmpeg/ffprobe executables, which are GPL-3.0-or-later. Full notices are in [build/THIRD-PARTY-LICENSES](build/THIRD-PARTY-LICENSES), which ships inside every release and prints via:

```shell
uv run clipgen.py --licenses
```

## Building from source

Executables are built by GitHub Actions (`.github/workflows/build-binaries.yml`) on version tag pushes. To build locally with PyInstaller:

```shell
pip install pyinstaller
uv run build/fetch_binaries.py      # pinned ffmpeg/ffprobe for the bundle (SHA256-verified)
pyinstaller --clean --noconfirm build/clipgen.spec
```

Output: `dist/clipgen.app` (macOS) or `dist/clipgen/` (Windows — keep `clipgen.exe` and `lib/` together, with `credentials.json` beside the folder).

## AI disclosure

The author used an LLM coding agent for assistance in writing parts of this program. If you want to avoid software connected to LLMs; I get it. All code in this repository prior to 2026 was written by a human, if you would like to fork the project.
