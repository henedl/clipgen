# License Distribution Plan

## Context

clipgen bundles 20+ third-party libraries into a PyInstaller binary. All licenses (MIT, BSD, Apache 2.0, HPND, LGPL, GPL) require attribution. The `build/THIRD-PARTY-LICENSES` file contains the full notices. This plan covers how to get that file to end users.

## Architectural decisions

### OpenCV and FFmpeg (decided 2026-04-04, corrected 2026-08-15; **superseded 2026-08-16**)

> **Superseded:** the macOS build now compiles `opencv-python-headless` 4.14.0.94 from its
> published sdist with `-DWITH_FFMPEG=OFF` (research: [OPENCV-SELF-COMPILE-PLAN.md](OPENCV-SELF-COMPILE-PLAN.md)),
> so the DMG carries **no** FFmpeg libraries inside cv2 and is conveyed MIT again, with GPL
> applying only to the aggregated ffmpeg/ffprobe executables. The GPL-FFmpeg-inside-opencv
> section of `build/THIRD-PARTY-LICENSES` was replaced by an LGPL-2.1 section covering the
> Windows wheel's `opencv_videoio_ffmpeg` DLL, and the CI guard inverted: it now fails if any
> `cv2/.dylibs` or FFmpeg/x264/x265 dylib appears in the bundle outside `bin/`. The analysis
> below is kept for the record of *why* the official arm64 wheel cannot be shipped.

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

- [x] **6. In-app attribution (Start overlay → About tab)**: Done. The About tab now ends with a
  "Third-party software" list generated from the notice file's own SUMMARY table, so the UI cannot
  drift from what ships. `source/licenses.py` parses the table (`load_components()`), `/api/licenses`
  serves it, and `renderAttribution()` in `start-overlay.js` groups consecutive rows by license
  family. `--licenses` is untouched and still prints the full ~100 KB text.

  **The table is now a machine-read interface, not just prose.** Three shapes are load-bearing and
  must survive any future hand-edit: the fixed-width columns (parsed by splitting on runs of 2+
  spaces), the 2-space-indented sub-rows (`FFmpeg DLL (in cv2)`, `PP-OCR models`), and the wrapped
  license cell that continues on the next line. `tests/test_licenses.py` asserts all three against
  the real file.

  Closing the gap also required attributing three components that shipped in every bundle but
  appeared nowhere in the notice: **Heroicons** (2.2.0, MIT, © Tailwind Labs — the 316 SVGs in
  `assets/icons/`, previously credited only in `README.md`, which is not in the artifact),
  **Octicons** (19.32.0, MIT, © GitHub Inc.) and **Silero VAD** (v6, MIT, © Silero Team — bundled
  inside the faster-whisper wheel and collected by `clipgen.spec`). The two web fonts already had a
  full SIL OFL section but no SUMMARY row, so they were invisible to the new list; they have rows
  now. `assets/icons/README.md` was added to record the Heroicons provenance, including the caveat
  that the vendored files are re-exports whose geometry — not markup — matches upstream 2.2.0.

## Open items / future candidates

None of these are scheduled. They are recorded so the next session inherits the evidence instead
of re-deriving it.

- **Track upstream opencv-python#1260.** No longer load-bearing: the macOS build compiles its
  own FFmpeg-free cv2 regardless. If upstream fixes the wheel, the only optional simplification
  is deleting the wheel-build CI steps and going back to the official wheel — the inverted
  license guard stays valid either way (it would then verify the fixed wheel). Weigh that
  against the self-built wheel's other wins (−76 MB unpacked, no 93-dylib brew chain).

- [x] **Self-compiled OpenCV — researched, de-risked, and implemented 2026-08-16**, see
  [OPENCV-SELF-COMPILE-PLAN.md](OPENCV-SELF-COMPILE-PLAN.md) for the full evidence. Every open
  question answered empirically on macOS arm64: the sdist builds clean in **122 s** locally
  (~8–15 min on a standard CI runner, cacheable); requires a pin bump to 4.14.0.94 (4.13.x
  published no sdists); the result has **zero** bundled dylibs (Apple frameworks only, codecs
  static in-tree — the win is −76 MB unpacked, not the ~29 MB guessed here) and passes the full
  clipgen test suite (3302 tests). The macOS CI leg now builds and caches the wheel;
  see the superseded-decision note above for what changed where.

- **Remove cv2 entirely (reimplement its used surface)** — investigated 2026-08-16, feasible
  with no new dependencies but ~8× the effort of the self-compile; recorded as an optional
  future work package in [DROP-OPENCV-PLAN.md](DROP-OPENCV-PLAN.md).

- [x] **Take PyAV off the decode path** — Done 2026-08-16, and it went further than this note
  anticipated: the "cannot be excluded" caveat was wrong. `faster_whisper/audio.py` imports `av`
  at module scope but only *calls* it inside `decode_audio()`, so an empty stub module
  (`transcripts._ensure_av_stub`) satisfies the import and `av` is overridden out of the
  dependency tree entirely (a ~45 MB mac / ~69 MB Windows bundle-size win on top of the
  consistency one). `video.decode_audio_pcm` decodes with the pinned ffmpeg
  (`-map 0:a:N -ac 1 -ar 16000 -f f32le`), which also deleted the non-default-audio-track
  demux workaround in `transcribe_video()`.

- **Full transitive-dependency attribution sweep.** Step 5 covered direct dependencies only, and
  step 6 added the bundled *assets* (icons, fonts, VAD model) rather than any new transitives. The
  bundle still contains a long tail of unlisted ones (tokenizers, huggingface-hub, cryptography,
  click/itsdangerous/blinker, sympy, shapely, pyclipper, …). Most are permissive, so this is a
  completeness task rather than a risk one. Note the list above is stale in two places: onnxruntime
  *is* attributed now, and networkx left with scikit-image.

  This is more visible than it was: anything missing from the SUMMARY table is now also missing
  from the About tab, so the sweep would show up directly in the UI.
