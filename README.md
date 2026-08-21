# clipgen

clipgen turns user research recordings into video clips, screenshots, GIFs, and highlight reels, driven by the timestamps in your research notes (a Google Sheet, a local Excel file, or MindNode mind map). Around that core it bundles tools for clip building, visual analysis, transcription, freeform cutting, automation, and study-level overviews.

clipgen is built for UX researchers — especially games user researchers running playtests — who want to conduct quick analysis and artifact generation. Everything runs locally.

## Quick start

### Features

- **Studio** — generate clips, screenshots, and GIFs, and build reels from structured data.
- **Screenspace** — visual analysis of recordings: detect change, text, motion, and more.
- **Transcripts** — local transcription, with optional AI summaries.
- **Composer** — quickly cut and annotate moments that aren't in your structured data.
- **Workflows** — chain the tools above into automated pipelines on a node canvas.
- **Overview** — study-level views: a shared participant timeline, aggregate stats, and per-participant reports.

### Desktop app

Download builds from the [Releases page](https://github.com/henedl/clipgen/releases) and follow the bundled `INSTALL.txt`.

The macOS build is unsigned, so Gatekeeper blocks the first launch: right-click the app and choose **Open** once, or run `xattr -dr com.apple.quarantine clipgen.app`. Double-clicking the app opens clipgen in its own desktop window; the same binary is also the full CLI when given arguments.

### From source

1. Download or clone the repository.
2. Install [uv](https://docs.astral.sh/uv/) and run `uv sync`.
3. Install ffmpeg: either manually or by `scripts/install-ffmpeg-ollama.sh` (macOS/Linux) or `scripts/install-ffmpeg-ollama.bat` (Windows).
4. Run `uv run clipgen.py`.

### Your videos

Name recordings `{study}_{participant}.mp4` (e.g. `mystudy_P01.mp4`) and direct clipgen towards them. If a session spans several files, add a numbered suffix (`mystudy_P01-1.mp4`, `mystudy_P01-2.mp4`, ...) to have clipgen treat them as a continuous recording.

### Google Sheets (optional)

Reading a Google Sheet needs Google auth: follow [gspread's setup guide](https://docs.gspread.org/en/latest/oauth2.html) and save the client file as `credentials.json` next to the app (alternately in `~/.config/gspread/`, or in clipgen's config dir `~/.config/clipgen/`; `%LOCALAPPDATA%\clipgen\` on Windows). Local Excel files need no credentials, and Screenspace, Transcripts, Composer, and Workflows need no spreadsheet at all.

#### The spreadsheet

clipgen expects a particular layout — make a copy of the [reference spreadsheet](https://docs.google.com/spreadsheets/d/1O51wnzRrYyz63tT6qy1HlJyVzdh9RT3t6QL5NohrcPc/edit?usp=sharing) to start. Timestamps in a cell are separated by `+`, `,`, or `;`; ranges use `-` (e.g. `1:23-1:45`).

Two optional rows: a `Baseline time` row converts clock timestamps (e.g. `09:12:00`) to relative offsets per participant, and a `Filename` row overrides the source video per participant (plus split videos).

### Local AI (optional)

Transcript summaries and Overview reports require [Ollama](https://ollama.com/download). clipgen offers to install Ollama and pull the models the first time you use an AI feature, and starts `ollama serve` for you. Transcription itself does not need Ollama.

## Command line

Launch with no flags for the interactive prompt, or script it:

```shell
uv run clipgen.py -H                       # Highlight reel of the most severe issues
uv run clipgen.py -b                       # Batch: every clip in the study
uv run clipgen.py --gif -C "Onboarding"    # GIFs for one category
```

Selection modes: batch (`-b`), line (`-l`), range (`-r`), category (`-C`), cell (`-c`), participant (`-p`), keyword (`-k`), severity (`-S`). Outputs: clips (default), `--screen`, `--gif`. Reels: custom (`-R`), chronological (`-T`), highlights (`-H`). `--gallery VIDEO` builds a browsable interval-screenshot gallery. Run `uv run clipgen.py --help` for the full reference.

## Building from source

To build locally with PyInstaller:

```shell
pip install pyinstaller
uv run build/fetch_binaries.py      # pinned ffmpeg/ffprobe for the bundle (SHA256-verified)
pyinstaller --clean --noconfirm build/clipgen.spec
```

## Third-party attribution

clipgen's source is MIT licensed. The desktop builds bundle third-party software under MIT, BSD, Apache-2.0, HPND, MPL-2.0, LGPL, SIL OFL, and GPL terms. Full notices are in [build/THIRD-PARTY-LICENSES](build/THIRD-PARTY-LICENSES).

Icons are [Heroicons](https://heroicons.com) and GitHub's [Octicons](https://primer.style/octicons/), both MIT-licensed.

## AI disclosure

The author used an LLM coding agent for assistance in writing parts of this program. If you want to avoid software connected to LLMs; I get it. All code in this repository prior to 2026 was written by a human, if you would like to fork the project.
