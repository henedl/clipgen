# License Distribution Plan

## Context

clipgen bundles 20+ third-party libraries into a PyInstaller binary. All licenses (MIT, BSD, Apache 2.0, HPND, LGPL) require attribution. The `build/THIRD-PARTY-LICENSES` file contains the full notices. This plan covers how to get that file to end users.

## Architectural decisions

### OpenCV and FFmpeg (decided 2026-04-04)

`opencv-python-headless` bundles FFmpeg shared libraries (libavcodec, libavformat, etc.) inside `cv2/.dylibs/`. The `cv2.abi3.so` binary hard-links against these at load time — they **cannot** be excluded from a PyInstaller build without breaking `import cv2`.

The opencv-python project builds FFmpeg as **LGPL 2.1** (without `--enable-gpl`). The `libx264`/`libx265` dylibs are present as transitive dynamic dependencies but are not compiled into FFmpeg, so they do not taint the FFmpeg license.

**Decision:** Keep FFmpeg libraries bundled. Comply with LGPL 2.1 terms by:
1. Including the LGPL license text (via `cv2/LICENSE-3RD-PARTY.txt` in the bundle)
2. Attributing FFmpeg and listing the bundled libraries in `build/THIRD-PARTY-LICENSES`
3. Providing source code pointers (ffmpeg.org, github.com/opencv/opencv-python)
4. Not calling `cv2.VideoCapture` or any FFmpeg-backed cv2 API — video decoding uses system ffmpeg via subprocess

## Steps

- [ ] **1. PyInstaller `--onedir` builds**: The file will be in the output directory alongside the binary. No spec changes needed — just copy it to the dist folder.

- [ ] **2. PyInstaller `--onefile` builds** (current spec): The file cannot be embedded inside the binary itself. It should be distributed alongside the binary (e.g. in a zip/dmg/tar.gz archive). The spec's `datas` list could also include it so it extracts at runtime, but the standard approach is to ship it next to the executable.

- [ ] **3. GitHub Releases**: Include `THIRD-PARTY-LICENSES` as a release asset alongside the binary, or bundle both in an archive.

- [ ] **4. `--licenses` flag** (optional future enhancement): Add a `--licenses` CLI flag that prints the contents of the bundled THIRD-PARTY-LICENSES file to stdout. This requires adding the file to the spec's `datas` list and reading it at runtime via `utils.get_bundled_assets_root()`.
