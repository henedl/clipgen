# License Distribution Plan

## Context

clipgen bundles 20+ third-party libraries into a PyInstaller binary. All licenses (MIT, BSD, Apache 2.0, HPND, LGPL, GPL) require attribution. The `build/THIRD-PARTY-LICENSES` file contains the full notices. This plan covers how to get that file to end users.

## Architectural decisions

### OpenCV and FFmpeg (decided 2026-04-04, **corrected 2026-08-15**)

> **The original decision rested on a false premise.** It claimed opencv-python builds FFmpeg as
> LGPL-2.1 "without `--enable-gpl`", with `libx264`/`libx265` present but "not compiled into
> FFmpeg, so they do not taint the FFmpeg license". That is wrong on macOS. Superseded by the
> text below; kept visible because the wrong version shipped in `build/THIRD-PARTY-LICENSES`
> for four months.

`opencv-python-headless` bundles FFmpeg shared libraries inside `cv2/.dylibs/`. Verified against
the exact `uv.lock` pin (4.13.0.92):

```
cv2/.dylibs/libavcodec.61.19.101.dylib
  "libavcodec license: GPL version 3 or later"
  --enable-gpl --enable-version3 --enable-libx264 --enable-libx265
```

So the libraries are **GPL-3.0-or-later**, with x264/x265 genuinely compiled in.

**Root cause is an upstream bug**, acknowledged at
[opencv/opencv-python#1260](https://github.com/opencv/opencv-python/issues/1260) (maintainer,
2026-08-14: *"it's side effect of the brew packages"*, *"We will try to fix it on our side"*).
opencv-python's macOS **arm64** workflow never sets `CONFIG_PATH`, so the GPL-stripping step in
`travis_config.sh` never runs. The wheel's own `LICENSE-3RD-PARTY.txt` and README still claim
LGPL-2.1 — upstream confirms that document is inaccurate, so it cannot be relied on.

**Scope: the macOS DMG only.** Linux wheels configure FFmpeg without `--enable-gpl`; Windows uses
a separately built LGPL `opencv_videoio_ffmpeg` DLL. `macos-latest` is arm64, so the DMG is the
affected artifact.

**They cannot be stripped**, for three independent reasons:

1. `otool -L cv2/cv2.abi3.so` shows `LC_LOAD_DYLIB` entries for libavcodec/libavformat/libavutil/
   libswscale/libavdevice — resolved by dyld **at import**, not `dlopen`ed on demand. Delete any
   and `import cv2` fails outright.
2. `libx264`/`libx265` are load-bearing one level down: `libavcodec` links them directly.
3. Even if they could be removed it would not change the license. `libavcodec` was *configured*
   `--enable-gpl`, so `avcodec_license()` reports GPLv3 whatever files are on disk. The taint is
   the build config, not the file list.

Note also that clipgen never calls `cv2.VideoCapture` or any FFmpeg-backed cv2 API (verified by
grep across `source/`) — but non-use does not change the license: the GPL obligation follows from
linking and distribution, not from invocation.

**Decision (corrected):** keep the libraries bundled — there is no supported way to drop them
short of rebuilding OpenCV — and state the position accurately:
1. Record the verified GPL-3.0-or-later status, the upstream issue, and the macOS-only scope in
   `build/THIRD-PARTY-LICENSES`
2. State plainly that the **macOS binary distribution as a whole is conveyed under
   GPL-3.0-or-later**. clipgen's own source stays MIT (MIT is GPL-compatible; nothing about
   clipgen's files is relicensed), published at github.com/henedl/clipgen
3. Provide source pointers for FFmpeg, x264, x265, and the opencv-python build
4. Guard against silent drift: `build-binaries.yml` asserts the `libavcodec license:` string of
   the cv2-bundled FFmpeg on every macOS build, so a wheel bump that changes it fails CI instead
   of leaving a false notice in the artifact (it also asserts PyAV stays *out* of the bundle,
   since its notice sections were removed — see below)

### PyAV and FFmpeg (decided 2026-08-15; superseded 2026-08-16)

PyAV (`av`, a hard dependency of faster-whisper) bundled a *third* FFmpeg in `av/.dylibs/`. Its
`avcodec_license()` reported **LGPL version 3 or later** — here the "x264/x265 present but not
linked" reading was actually correct, since enabling them would have required `--enable-gpl`.
The 2026-08-15 decision was to attribute it under a new LGPL-3.0-or-later section.

**Superseded:** PyAV is no longer shipped at all. faster-whisper only calls `av` to decode
path inputs; `transcripts.py` now feeds `transcribe()` ndarrays decoded by the pinned ffmpeg
(`video.decode_audio_pcm`) and stubs the module-scope `import av`
(`transcripts._ensure_av_stub`), so the dependency is overridden out in `pyproject.toml`
(~45 MB mac / ~69 MB Windows). The PyAV BSD entry and the whole LGPL-3.0-or-later section were
removed from `build/THIRD-PARTY-LICENSES`, and `build-binaries.yml` now fails if `av` dylibs
reappear in the bundle unattributed.

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

- [x] **4. `--licenses` flag**: Done — `utils.get_licenses_text()` reads the notice through
  `get_bundled_assets_root()` (same `build/`-vs-bundle-root two-candidate lookup as
  `get_version()`), `clipgen.spec` ships it in `datas`, and `cli._LicensesAction` prints it and
  exits. It is a custom `argparse.Action` rather than `action="version"` because that action
  takes a static string and would read ~78 KB on *every* run just to build the parser.

- [x] **5. Attribution completeness (direct dependencies)**: Done — added PyAV, pywebview,
  pyobjc-core/pyobjc-framework-Cocoa, python-bidi, certifi and tqdm, none of which were listed at
  all. This added an LGPL-3.0-or-later section and an MPL-2.0 section, and fixed two malformed
  section rules that would break any future section-aware parsing of the file.

## Open items / future candidates

None of these are scheduled. They are recorded so the next session inherits the evidence instead
of re-deriving it.

- **Track upstream opencv-python#1260.** The cheapest resolution: when a release lands whose
  macOS arm64 wheel reports LGPL, pin to it and revert the GPL sections of
  `build/THIRD-PARTY-LICENSES`. The CI license guard will fail on that build, which is the
  intended signal. No ETA — the issue was acknowledged 2026-08-14 and is unfixed.

- **Self-compiled OpenCV (needs research before anyone commits to it).**
  `CMAKE_ARGS="-DWITH_FFMPEG=OFF" ENABLE_HEADLESS=1` against the published sdist is the *only*
  route that actually removes the GPL dylibs, since they cannot be deleted post-hoc (see the
  three reasons above). It would also drop ~29 MB of the 120 MB `cv2/` tree, all of it dead
  weight — clipgen never calls cv2's video I/O. Open questions:
  - Does the sdist build clean on macOS arm64? The route is documented by upstream but far less
    exercised than the git-clone path.
  - What does a 20 min–2 h OpenCV compile do to the macOS build job's wall-clock and caching?
  - Does anything else in `cv2/` regress with FFmpeg off (`imgcodecs` keeps its own `libavif`)?
  - Is it moot by the time it is picked up, if #1260 lands first?

- [x] **Take PyAV off the decode path** — Done 2026-08-16, and it went further than this note
  anticipated: the "cannot be excluded" caveat was wrong. `faster_whisper/audio.py` imports `av`
  at module scope but only *calls* it inside `decode_audio()`, so an empty stub module
  (`transcripts._ensure_av_stub`) satisfies the import and `av` is overridden out of the
  dependency tree entirely (a ~45 MB mac / ~69 MB Windows bundle-size win on top of the
  consistency one). `video.decode_audio_pcm` decodes with the pinned ffmpeg
  (`-map 0:a:N -ac 1 -ar 16000 -f f32le`), which also deleted the non-default-audio-track
  demux workaround in `transcribe_video()`.

- **Full transitive-dependency attribution sweep.** Step 5 covered direct dependencies only. The
  bundle contains a long tail of unlisted transitives (onnxruntime, tokenizers, huggingface-hub,
  cryptography, click/itsdangerous/blinker, sympy, networkx, shapely, pyclipper, …). Most are
  permissive, so this is a completeness task rather than a risk one.
