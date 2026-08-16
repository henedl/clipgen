---
name: clipgen-build
description: Build, verify, and ship the desktop bundle (PyInstaller, DMG, CI artifacts)
---

# Build and packaging

Everything about turning clipgen into a shippable desktop app. Read this before touching
`build/clipgen.spec`, `.github/workflows/build-binaries.yml`, or anything path-related under a
frozen build — most of the cost here has historically gone into rediscovering the traps below.

## Build it

```bash
uv pip install pyinstaller==6.19.0          # match the CI pin
uv run --no-sync build/fetch_binaries.py    # pinned ffmpeg/ffprobe → build/vendor/ (idempotent)
uv run --no-sync pyinstaller --clean --noconfirm build/clipgen.spec
```

Takes ~90 s. Output: `dist/clipgen.app` (macOS) or `dist/clipgen/` (Windows: `clipgen.exe` +
`lib/`). The build is **one-dir** — see "Why one-dir" below before changing that.

The fetch step downloads the pinned static GPL ffmpeg/ffprobe builds (SHA256-verified, see the
`PINS` block in `build/fetch_binaries.py` for provenance) into the gitignored `build/vendor/`.
The spec hard-fails without them — a build that skipped the step must break on the build machine,
not ship an app that dies on its startup ffmpeg check. They land in `<bundle>/bin/`, which
`utils.prepend_bundled_bin_to_path()` puts at the head of PATH on frozen launches, so the app
always runs the ffmpeg it was feature-verified with. UPX is excluded for them in both `EXE` and
`COLLECT` (`upx_exclude`) — UPX corrupts signed mach-O and packed ffmpeg.exe builds.

**`strip` and `upx` are macOS-only in the spec — never re-enable them on Windows.** PyInstaller
shells out to whatever `strip`/`upx` is on PATH, and the GitHub windows runner carries a GNU
binutils `strip` that mangles MSVC-built PE DLLs (PyInstaller's docs say as much: `--strip` is
"not recommended for Windows").
The build stays green, the files all exist, and the shipped exe dies at launch with
`Failed to load Python DLL ... LoadLibrary: Invalid access to memory location` (error 998) —
that exact bug shipped in every Windows build from #639 until the spec's `_strip` gate. CI now
smoke-launches `clipgen.exe --help` on both legs, which fails on any bundle whose DLLs got
corrupted at build time (the bootloader loads `python312.dll` before argv parsing).

Stripping bought nothing to begin with: turning it off grew the Windows artifact from
440.4 MB to 442.3 MB — **+1.9 MB, 0.43%**. These are MSVC binaries whose debug info lives in
separate PDBs, so there is barely anything for `strip` to remove. Do not trade a launchable
app for that.

## The one rule: verify the way a user launches it

**A frozen app behaves differently depending on how it is started, and the convenient way to test
is the one that hides bugs.**

```bash
open dist/clipgen.app                        # ✅ what a user does
./dist/clipgen.app/Contents/MacOS/clipgen    # ❌ inherits your shell environment
```

This is not theoretical. It shipped a broken app once (#638): a Finder-launched `.app` gets
`PATH=/usr/bin:/bin:/usr/sbin:/sbin` and **does not inherit your shell PATH**, so Homebrew's
ffmpeg at `/opt/homebrew/bin` was invisible, `check_ffmpeg_tools_available()` called
`sys.exit(1)`, and the app quit with no window and nothing on screen. Every shell-based test
passed. `utils.augment_path_for_gui_launch()` exists because of this.

Corollary: **any `sys.exit()` reachable before the window opens must go through
`utils.fatal_startup_error()`**, which raises a native dialog when `utils.GUI_LAUNCH` is set.
Printing to stdout in a windowed launch is printing into the void.

To see what a Finder launch actually printed:

```bash
open -W --stdout /tmp/out.log --stderr /tmp/err.log dist/clipgen.app
```

To reproduce the Finder environment without Finder:

```bash
env -i HOME="$HOME" PATH=/usr/bin:/bin:/usr/sbin:/sbin ./dist/clipgen.app/Contents/MacOS/clipgen
```

## Measuring startup — do not use `--help`

`--help` exits before importing the heavy stack, so it is **not** a proxy for launch time. Using
it as one produced a wrong conclusion that nearly sent a session down a pointless lazy-import
refactor. Measure a real startup to first HTTP response, and abort if anything already holds the
port or you will measure a stale process:

```bash
pkill -f "clipgen.app/Contents/MacOS/clipgen"; sleep 3
lsof -nP -iTCP:8089 -sTCP:LISTEN >/dev/null && { echo "ABORT: port held"; exit 1; }
start=$(python3 -c 'import time;print(time.time())')
./dist/clipgen.app/Contents/MacOS/clipgen --studio --browser >/dev/null 2>&1 &
for i in $(seq 1 80); do curl -s -o /dev/null --max-time 1 http://127.0.0.1:8089/studio/ && break; sleep 0.5; done
python3 -c "print(f'{$(python3 -c 'import time;print(time.time())') - $start:.1f}s')"
```

Reference numbers on an M-series Mac (studio startup → first response):

| | cold | warm |
|---|---|---|
| one-dir (current) | ~3.1 s | ~1.1 s |
| one-file (old) | ~21.8 s | ~17.6 s |
| source checkout | — | 0.41 s |

The source number matters: it proves there is **no Python-side import cost worth chasing**. If a
frozen build is slow, the cause is the packaging shape, not the code.

## Why one-dir

One-file re-extracts the whole archive to a *new* temp directory on every launch, so every large
dylib (cv2, av, onnxruntime) loads cold — no OS page cache, and macOS re-validates every code signature
from scratch. That is the entire 16× difference. One-file is also deprecated in combination with
windowed mode on macOS and **becomes a hard error in PyInstaller v7.0**.

Cost of one-dir: download 261 MB → 309 MB, on disk 261 MB → 799 MB. Worth it.
With the bundled ffmpeg/ffprobe (2026-08): download ~386 MB, on disk ~925 MB — the two static
binaries add ~130 MB uncompressed. Startup is unaffected (measured 2.1 s post-bundling).

## Frozen path resolution — three different roots

Confusing these is the most common source of packaging bugs.

| Need | Use | Notes |
|---|---|---|
| Bundled assets (`assets/`, `VERSION`) | `utils.get_bundled_assets_root()` | Wraps `sys._MEIPASS`. **Never** `Path(__file__).parent` — in source runs this is `source/`'s **parent**, since the modules sit one level below the assets |
| "Next to the application" (user files, `credentials.json`) | `cli.get_runtime_working_dir()` | Not the same directory as above |

### Relative paths inside the spec resolve against two different bases

This one produced a **green build that shipped no application code**. In `build/clipgen.spec`:

- `datas` / `binaries` resolve against **SPECPATH** (`build/`), so `("../assets", "assets")`
  correctly means the repo-root `assets/`.
- `pathex` resolves against the **process CWD**. The documented build command runs from the
  repo root, so a relative entry there points somewhere else entirely.

Nothing warns you. PyInstaller prints `Build complete!`, the `.app` is the normal size (the
third-party wheels still get collected), and it dies on its first import with an empty screen.
The tell is build time — a bundle missing all first-party code finishes in ~15 s instead of
~90 s, because modulegraph never walks the real dependency tree.

`pathex` is therefore derived from `SPECPATH`, and the spec **asserts** that every
`source/*.py` appears in `a.pure` before it builds. If you see
`clipgen.spec: Analysis resolved none/some of source/`, that guard just saved a release.

`sys._MEIPASS` means different things per build shape, which is exactly why
`get_runtime_working_dir()` keys off it:

- **one-file**: an unrelated temp dir (`/var/folders/.../\_MEI123456`)
- **one-dir (Windows/Linux)**: `dist/clipgen/lib` — a *child* of the executable's directory
- **one-dir (macOS .app)**: `Contents/Frameworks` — sibling of `Contents/MacOS`, with symlinks
  into `Contents/Resources`

So on Windows one-dir the exe lives *inside* `clipgen/` next to `lib/`, and "next to the
app" is one level **up**. On macOS it is the folder containing the `.app` — never inside
`Contents/MacOS`, which is invisible in Finder and part of the code signature.

## Third-party package data must be collected explicitly

PyInstaller only bundles what a hook or the spec tells it to; **there is no hook for every
package**. faster-whisper has none, so its Silero VAD model (`assets/silero_vad_v6.onnx`,
loaded by `__file__` arithmetic with VAD on by default) was silently absent and every frozen
transcription died with `NO_SUCHFILE` — while every source run worked. The spec now has
`collect_data_files("faster_whisper")` plus a fail-loud `.onnx` guard, and CI checks the
bundle on both platforms. When a new dependency reads bundled data files at runtime, add a
`collect_data_files(...)` line *and* a build-time guard in the same commit; a missing data
file is invisible to the build and to every source-tree test.

## The bundle ships no CUDA runtime — mind libraries with their own GPU detection

Nothing in the dependency tree pulls a CUDA runtime (torch left with the RapidOCR switch),
so no `nvidia-*` wheel is present and none is collected. OCR runs on CPU through the
bundled onnxruntime.

**CTranslate2 does its own device detection.** Its wheels carry their own CUDA support, so
faster-whisper's default `device="auto"` selected CUDA on any machine with an NVIDIA GPU
and then died at the first inference with `Library cublas64_12.dll is not found or cannot
be loaded` — minutes in, after the model had downloaded and loaded. Hence
`transcripts._resolve_transcribe_device()`, which resolves "auto" to CPU when frozen.

When adding a dependency that can use a GPU, check *what it asks*: a library with its own
CUDA detection will happily pick a GPU this bundle cannot feed.

## `multiprocessing.freeze_support()` is not optional

In a frozen app `sys.executable` is the clipgen binary. Anything that touches
`multiprocessing` — tqdm's `RLock` inside faster-whisper's transcribe loop was enough — makes
CPython's resource tracker re-exec that binary with interpreter-style argv
(`-B -S -E -s -c '...'`), which lands in clipgen's argparse as
`clipgen: error: argument -S/--severity: expected one argument`. PyInstaller's runtime hook
intercepts exactly that argv shape, but only from inside `freeze_support()`. The call must
stay the **first statement** of the launcher's `__main__` block, before any clipgen import.

## CI and distribution traps

- **`actions/upload-artifact` does not preserve Unix permissions.** It forces everything to 644,
  which strips the exec bit and invalidates the code signature. A directly-uploaded `.app` will
  not launch after download. macOS therefore ships a **`.dmg`** (verified to preserve exec bit,
  symlinks, and signature through the round trip); Windows ships a **`.zip`**.
- **Actions artifacts need a login and expire after 90 days.** Tagged builds publish to **GitHub
  Releases**; that step needs `permissions: contents: write`.
- **`upload-artifact` always zips, with no opt-out** (Actions has stored artifacts as zips since
  v4), so an artifact download is always one wrapper deeper than the Release asset: a zip *around*
  the `.dmg` / `.zip`. It costs nothing in size — the outer zip is a wash or a small win — but on a
  tag it would be a redundant ~800 MB copy of what the Release already carries, so the upload steps
  are gated on `github.ref_type != 'tag'`. Dev (`workflow_dispatch`) builds keep their artifacts.
- **`fail_on_unmatched_files: true` needs a per-leg glob.** Each matrix leg lists exactly the
  files it produces in `matrix.release_glob` (`dist/*.dmg` on macOS; `dist/*.zip` plus
  `dist/*-setup.exe` on Windows), so every pattern is held to account; a shared cross-platform
  list would fail whichever leg didn't produce the other platform's file. This is the only guard
  on a tag that packaging produced anything, since the `if-no-files-found: error` uploads are
  skipped there.
- **The Windows installer is a wrapper, not a rebuild.** `build/clipgen.iss` (Inno Setup) packages
  `dist/clipgen/*` after the zip step into `dist/clipgen-<ver>-setup.exe` — per-user
  (`PrivilegesRequired=lowest`, `{localappdata}\Programs\clipgen`, no UAC), Start Menu shortcut,
  uninstaller. The version is injected via `ISCC /DAppVer=...` (never hardcoded in the .iss); the
  `AppId` GUID must never change or upgrades stop replacing the previous install. ISCC is
  preinstalled on the windows-latest image; CI compiles, then silent-installs
  (`/VERYSILENT /SUPPRESSMSGBOXES /NORESTART`) and runs the installed exe's `--help` as the smoke.
  The portable zip keeps shipping alongside it — two artifacts, one build.
- **Never publish `dist/clipgen` as a file.** Under one-dir it is a *directory*. The old raw-binary
  step was removed for exactly this reason — it would have silently published the wrong thing
  rather than failing.
- **Codesign after any post-build mutation**, then `codesign --verify --deep --strict`.

## Release verification checklist

```bash
uv run --no-sync build/fetch_binaries.py
uv run --no-sync pyinstaller --clean --noconfirm build/clipgen.spec
codesign --force --deep --sign - --timestamp=none dist/clipgen.app
codesign --verify --deep --strict dist/clipgen.app
open dist/clipgen.app                                        # window opens, no Terminal
./dist/clipgen.app/Contents/MacOS/clipgen --help             # CLI still works
# Bundled ffmpeg resolves without Homebrew on PATH (the #638 failure, inverted):
env -i HOME="$HOME" PATH=/usr/bin:/bin ./dist/clipgen.app/Contents/MacOS/clipgen --no-input -l 1 2>&1 | head -5
dist/clipgen.app/Contents/Frameworks/bin/ffmpeg -hide_banner -buildconf | grep -c enable-libx264
```

Plus: measure startup (above); DMG round-trip preserves exec bit + symlinks + signature; and
`credentials.json` beside the app is still found. Windows can only be verified via a
`workflow_dispatch` run — say so plainly in the PR rather than implying it was checked. CI also
greps the in-bundle ffmpeg's encoders/filters on both legs, so provider build-config drift fails
the build instead of silently degrading webp/vp9/titlecards.

## Related

- [agents/skills/check/SKILL.md](../check/SKILL.md) — the pre-commit gate
- [plans/archive/DESKTOP-PACKAGING-PLAN.md](../../../plans/archive/DESKTOP-PACKAGING-PLAN.md) — how
  the current shape was arrived at. Its "deferred ffmpeg bundling" note is resolved: the bundle now
  ships pinned GPL ffmpeg/ffprobe (licensing recorded in `build/THIRD-PARTY-LICENSES` and
  `plans/LICENSE-PLAN.md`)
- [desktop.py](../../../source/desktop.py) — the window host and its two JS bridges
