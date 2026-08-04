# License Distribution Plan

## Context

clipgen bundles 20+ third-party libraries into a PyInstaller binary. All licenses (MIT, BSD, Apache 2.0, HPND, LGPL, GPL) require attribution. The `build/THIRD-PARTY-LICENSES` file contains the full notices. This plan covers how to get that file to end users.

## Architectural decisions

### OpenCV and FFmpeg (decided 2026-04-04)

`opencv-python-headless` bundles FFmpeg shared libraries (libavcodec, libavformat, etc.) inside `cv2/.dylibs/`. The `cv2.abi3.so` binary hard-links against these at load time — they **cannot** be excluded from a PyInstaller build without breaking `import cv2`.

The opencv-python project builds FFmpeg as **LGPL 2.1** (without `--enable-gpl`). The `libx264`/`libx265` dylibs are present as transitive dynamic dependencies but are not compiled into FFmpeg, so they do not taint the FFmpeg license.

**Decision:** Keep FFmpeg libraries bundled. Comply with LGPL 2.1 terms by:
1. Including the LGPL license text (via `cv2/LICENSE-3RD-PARTY.txt` in the bundle)
2. Attributing FFmpeg and listing the bundled libraries in `build/THIRD-PARTY-LICENSES`
3. Providing source code pointers (ffmpeg.org, github.com/opencv/opencv-python)
4. Not calling `cv2.VideoCapture` or any FFmpeg-backed cv2 API — video decoding uses the bundled ffmpeg executable via subprocess

### Bundled ffmpeg/ffprobe executables (decided 2026-08-03)

The desktop builds ship pinned full-GPL static ffmpeg/ffprobe executables under `<bundle>/bin/` (fetched by `build/fetch_binaries.py`, provenance in its `PINS` block). These builds use `--enable-gpl --enable-version3`, making the executables **GPL-3.0-or-later**. This supersedes the deferred-bundling note in `plans/archive/DESKTOP-PACKAGING-PLAN.md`.

**Decision:** clipgen invokes them strictly as separate subprocesses — mere aggregation — so clipgen itself remains MIT. Comply with GPLv3 by:
1. Including the full GPLv3 text in `build/THIRD-PARTY-LICENSES` (its GPL-3.0-OR-LATER section)
2. Recording the exact build IDs, providers, and source pointers there and in `build/fetch_binaries.py`
3. Shipping `THIRD-PARTY-LICENSES` in both the DMG and, since a GPL binary now ships in it too, the Windows zip

## Steps

- [x] **1. PyInstaller `--onedir` builds**: Done — the build is one-dir and CI copies `build/THIRD-PARTY-LICENSES` into the DMG staging (macOS) and into `dist/clipgen/` before zipping (Windows).

- [x] ~~**2. PyInstaller `--onefile` builds**~~: Obsolete — the one-file shape was abandoned (see `plans/archive/DESKTOP-PACKAGING-PLAN.md`).

- [x] **3. GitHub Releases**: Done — the DMG and the Windows zip both carry `THIRD-PARTY-LICENSES` inside the published archives.

- [ ] **4. `--licenses` flag** (optional future enhancement): Add a `--licenses` CLI flag that prints the contents of the bundled THIRD-PARTY-LICENSES file to stdout. This requires adding the file to the spec's `datas` list and reading it at runtime via `utils.get_bundled_assets_root()`.
