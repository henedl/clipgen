# OpenCV Self-Compile Plan (WITH_FFMPEG=OFF)

**Status: researched, de-risked, and implemented 2026-08-16** (all steps in "Implementation
steps" below are done). All findings were verified empirically on macOS arm64 (M4 Pro,
macOS 15). This resolves the "needs research before anyone commits to it" open item in
[LICENSE-PLAN.md](LICENSE-PLAN.md). The heavier alternative (removing cv2 entirely) is
[DROP-OPENCV-PLAN.md](DROP-OPENCV-PLAN.md).

## Context

The official `opencv-python-headless` macOS arm64 wheels bundle GPL-3-configured FFmpeg dylibs
(upstream bug opencv/opencv-python#1260), forcing the DMG to be conveyed GPL-3.0-or-later. The
dylibs cannot be stripped post-hoc. Building the published sdist ourselves with FFmpeg disabled
is the only route that removes them while keeping cv2.

## Research results

Every LICENSE-PLAN open question, answered:

1. **Does the sdist build clean on macOS arm64? Yes — in 122 seconds.**

   ```bash
   env ENABLE_HEADLESS=1 \
       CMAKE_ARGS="-DWITH_FFMPEG=OFF -DPYTHON3_LIMITED_API=ON" \
       CMAKE_BUILD_PARALLEL_LEVEL=$(sysctl -n hw.ncpu) \
       uv build --wheel --python 3.12 --out-dir dist .   # from the unpacked sdist
   ```

   No brew packages, no preinstalled cmake needed — the sdist's custom PEP 517 backend
   auto-installs the cmake PyPI wheel; scikit-build falls back to Unix Makefiles when ninja is
   absent. 122 s wall-clock at `-j12` on an M4 Pro (the old "20 min–2 h" estimate was off by an
   order of magnitude: the build compiles ~750 translation units, not all of OpenCV's history).

   Two caveats found:
   - **The pinned 4.13.0.92 has no sdist.** 4.13.x published wheels only; the nearest 4.x with an
     sdist is **4.14.0.94** (2026-07-29). Self-compiling therefore means a pin bump. The full
     clipgen test suite (3302 tests, incl. every exact-equality cv2 oracle) passes on 4.14.0.94,
     so the bump is behaviorally free for clipgen's usage.
   - **A broken CommandLineTools install fails with `fatal error: 'string' file not found`.**
     CLT 26.3 on the dev machine ships no `CommandLineTools/usr/include/c++/v1`; clang lists that
     nonexistent dir and never falls back to the SDK copy. Workaround (harmless elsewhere):
     `CXXFLAGS="-isystem $(xcrun --show-sdk-path)/usr/include/c++/v1"`. GitHub-hosted runners run
     full Xcode and should not hit this, but keep the flag in the build step — it costs nothing.

2. **CI wall-clock: ~8–15 min uncached, seconds cached.** 122 s locally on 12 cores; standard
   `macos-latest` arm64 runners are M1 with 3 vCPUs, so expect roughly 4–6× that. Acceptable even
   on every build; with `actions/cache` keyed on (opencv version, CMAKE_ARGS, deployment target,
   python tag) a rebuild only happens on a pin bump or cache eviction.

3. **Nothing clipgen uses regresses with FFmpeg off.** Verified from the CMake summary and the
   built wheel:
   - Media I/O all **static in-tree**: build-libjpeg-turbo, libpng, libwebp, libtiff, openjp2,
     OpenEXR, zlib. PNG round-trip verified. (`AVIF: NO` — the official wheel's avif came from
     brew dylibs; clipgen encodes only PNG/JPEG via cv2, GIFs via PIL.)
   - Video I/O = AVFoundation only (clipgen never uses cv2 video I/O anyway).
   - `cv2.data.haarcascades` present (attention face channel unaffected).
   - `uv run pytest`: **3302 passed** against the self-built wheel.

4. **Is it moot (upstream fix)? Not yet.** #1260 is open, acknowledged 2026-08-14, no fix
   commits on the repo (last commit 2026-07-28), no milestone. The maintainer intends to rebuild
   macOS FFmpeg without brew and to re-audit Linux; Windows is already a separately built LGPL
   DLL. If a fixed wheel ships, this plan's CI step can be deleted again — the inverse license
   guard (below) stays valid either way.

**Measured wins** (self-built vs official 4.14.0.94 arm64 wheel):

| | official | self-built |
|---|---|---|
| wheel size | 46.5 MB | **14 MB** |
| unpacked `cv2/` | 120 MB | **44 MB** (−76 MB in the DMG payload) |
| bundled dylibs | 93 (75 MB, incl. GPL ffmpeg/x264/x265, gnutls, glib, …) | **0** — `otool -L` shows Apple system frameworks only |
| license of the artifact | forces GPL-3.0-or-later conveyance | **no GPL code at all** |
| `import cv2` (warm) | multi-second dylib chain | 0.36 s |

The wheel tags as `cp312-cp312` via `uv build` (the PEP 517 path doesn't pass
`--py-limited-api`), but the extension itself is `cv2.abi3.so` (Limited API: YES). Irrelevant
while the bundle ships Python 3.12; retag with `wheel tags` if that ever matters.

## Implementation steps (one small PR) — all landed 2026-08-16

- [x] **Bump `opencv-python-headless` to `>=4.14.0.94,<5`** in `pyproject.toml`/`uv.lock` (all
   platforms — Windows keeps the official wheel with its separately built LGPL
   `opencv_videoio_ffmpeg` DLL; the `opencv-python ; sys_platform == 'never'` override is
   untouched).
- [x] **Wheel-build steps in the macOS job of `.github/workflows/build-binaries.yml`**, before
   PyInstaller:
   - `actions/cache` keyed `opencv-noffmpeg-4.14.0.94-cp312-macos13-arm64-v1`
     (bump `-v1` to force a rebuild without a version change).
   - On miss: download the sdist from PyPI (sha256-pinned), build with
     `ENABLE_HEADLESS=1 CMAKE_ARGS="-DWITH_FFMPEG=OFF -DPYTHON3_LIMITED_API=ON"
     MACOSX_DEPLOYMENT_TARGET=13.0 CMAKE_BUILD_PARALLEL_LEVEL=$(sysctl -n hw.ncpu)
     CXXFLAGS="-isystem $(xcrun --show-sdk-path)/usr/include/c++/v1"` via
     `uv build --wheel` (13.0 matches the official wheel's deployment floor; the CXXFLAGS is
     the CLT-quirk belt-and-braces).
   - `uv pip install --force-reinstall --no-deps <wheel>`, then assert `FFMPEG` is absent from
     `cv2.getBuildInformation()` and no `cv2/.dylibs` exists in site-packages — before
     PyInstaller runs, so a bad cached wheel fails fast.
   - **Gotcha (caught by that assert on the first CI run):** a plain `uv run` re-syncs the
     project env against `uv.lock` before running, which reinstalls the *official* GPL wheel
     from PyPI over the force-reinstalled one. The workflow therefore installs with
     `uv sync --locked` once and uses `uv run --no-sync` for every later invocation
     (fetch_binaries, pyinstaller); the assert runs via `.venv/bin/python` directly.
- [x] **Inverted license guard** ("Verify no in-process FFmpeg libraries"): fails on any
   `cv2/.dylibs`/`__dot__dylibs` in the .app, or any `libav*`/`libsw*`/`libx264*`/`libx265*`
   dylib outside `bin/` (the PyAV-absence guard rides along unchanged).
- [x] **License text**: the `GPL-3.0-OR-LATER (FFmpeg bundled inside opencv-python-headless)`
   section of `build/THIRD-PARTY-LICENSES` was replaced by
   `LGPL-2.1 (FFmpeg DLL bundled in the Windows opencv-python-headless wheel)` with the full
   LGPL-2.1 text (previously a gap — the Windows DLL was mentioned but its license text never
   included); the executables section and summary table updated; the DMG INSTALL.txt LICENSING
   text conveys the app as MIT with GPL only on the aggregated executables. Heading assertions
   in `tests/test_packaging.py` updated, plus a negative assertion that the GPL-inside-cv2
   section stays gone.
- [x] **Docs**: OpenCV decision in [LICENSE-PLAN.md](LICENSE-PLAN.md) marked superseded.
   Local-dev caveat: dev venvs still install the official (GPL-tainted on macOS) wheel from
   PyPI — fine, GPL obligations attach to distribution, and clipgen distributes only the
   CI-built DMG.

## Non-goals / notes

- Local dev builds of the wheel are possible with the recipe above but not required; the license
  concern is distribution-only.
- Staying on 4.x (not 5.0.0.93, which also has an sdist) keeps `CascadeClassifier` and avoids
  the OpenCV 5 API churn the codebase currently tolerates but doesn't require.
- If upstream fixes #1260, delete step 2 and keep steps 3–4 — the inverse guard then simply
  verifies the fixed official wheel, and the pin follows the normal bump cadence.
