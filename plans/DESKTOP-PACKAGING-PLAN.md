# Desktop packaging: PATH hotfix, then onedir migration

Two follow-ups to #637 (`feat: open clipgen in a native desktop window`). They are
deliberately separate PRs: **Phase 1 is a hotfix for a shipping bug and should land on its own**;
Phase 2 is a build-system migration that needs a Windows CI run to verify.

---

# Phase 1 — Finder-launched app exits immediately (BLOCKING)

## The bug

After #637, double-clicking `clipgen.app` starts the process, never shows a window, and exits
after ~6 seconds. Reproduced and root-caused from a Console.app log
(`.context/attachments/HLUONh/`):

```
runningboardd  [...:30199] Set darwin role to: NonUserInteractive
runningboardd  [...:30199] visiblity is no
clipgen        Entering exit handler.
runningboardd  [...:30199] termination reported by launchd (0, 0, 256)
```

Status `256` is wait-status for **exit code 1** — a deliberate `sys.exit(1)`, not a crash.

**Cause: a Finder-launched `.app` does not inherit the user's shell PATH.** macOS gives GUI
processes a minimal `PATH=/usr/bin:/bin:/usr/sbin:/sbin`. Homebrew installs ffmpeg to
`/opt/homebrew/bin` (Apple Silicon) or `/usr/local/bin` (Intel) — neither is on that PATH. So
`video.check_ffmpeg_tools_available()` finds nothing and `cli.py:3798` exits 1.

**Why #637 introduced it.** The old spec shim ran `open -a Terminal "$DIR/clipgen-bin"`. Terminal
starts a login shell, which sources the user's profile and therefore *has* `/opt/homebrew/bin`.
Deleting the shim deleted that PATH inheritance. The app still works perfectly from a terminal,
which is exactly how it was tested — the failure only appears on the path real users take.

Verified both directions against a local build:

```
env -i PATH=/usr/bin:/bin:/usr/sbin:/sbin  .../MacOS/clipgen   → "Missing command(s): ffmpeg, ffprobe", exit 1
env -i PATH=/opt/homebrew/bin:/usr/bin:... .../MacOS/clipgen   → window opens
```

## Fix — two parts, both needed

### 1a. Augment PATH for frozen GUI launches

In `cli.main()`, before `check_ffmpeg_tools_available()` (and before anything else shells out),
prepend the standard package-manager bin directories when frozen on macOS:

- `/opt/homebrew/bin` (Homebrew, Apple Silicon)
- `/usr/local/bin` (Homebrew on Intel, and most manual installs)
- `/opt/local/bin` (MacPorts)

Only append entries that exist and are not already present, and only when frozen — a source run
already has the developer's real PATH and must not be second-guessed. Put the helper in `utils.py`
(it is process-environment plumbing, not CLI argument logic) and note *why* in a comment, because
"prepend some hardcoded paths" looks arbitrary without the Finder explanation.

Deliberately **not** doing the `$SHELL -l -c 'echo $PATH'` trick: it spawns a login shell on every
launch, can hang on a misbehaving profile, and adds startup latency to a path that is already slow.

### 1b. Make startup failure visible

This is the more important half. Even with 1a, *any* `sys.exit(1)` before the window opens is
currently invisible — the error is written to a stdout nobody sees. The user gets a bouncing dock
icon and silence. That is the actual defect; the missing ffmpeg is just what triggered it first.

Add a fatal-error surface for windowed launches: when a frozen macOS GUI launch is about to exit
non-zero during startup, show the message in a native dialog first (`osascript -e 'display alert
...'` is enough and needs no extra dependency). Route the two existing hard exits through it:

- `utils.validate_runtime_directories()` (`cli.py:3785`)
- `video.check_ffmpeg_tools_available()` (`cli.py:3798`) — this one should also say
  `brew install ffmpeg` in the dialog, since the console guidance is unreachable.

Windows is far less exposed (GUI processes inherit the machine/user PATH from the registry) but
the dialog path should be platform-neutral so a future failure is not silent there either.

## Phase 1 verification

1. `/check` and `/ui-check` green.
2. Build the bundle, then **launch it the way a user does** — `open dist/clipgen.app` — not from a
   shell. This is the step #637 skipped and is the whole point.
3. Reproduce the original failure deliberately: `env -i PATH=/usr/bin:/bin ... MacOS/clipgen`
   with the PATH fix reverted, confirm the dialog appears instead of a silent exit.
4. Temporarily rename `/opt/homebrew/bin/ffmpeg` and confirm the dialog names it and says how to
   install it.
5. Confirm the CLI is unaffected: `MacOS/clipgen --help` and a real `--gif` run.
6. New test: PATH augmentation only fires when frozen, only adds existing dirs, and never
   duplicates an entry already present.

---

# Phase 2 — Migrate to `--onedir`

## Why

`build/clipgen.spec` is one-file. PyInstaller now warns that one-file **plus windowed** on macOS
"clashes with macOS's security" and **becomes an error in v7.0**; `console=False` (from #637) is
what started triggering it. Beyond the deprecation, measured on a real onedir build of this repo:

| | onefile (today) | onedir |
|---|---|---|
| Launch time | 2.99 / 3.05 / 3.04 s | 0.10 / 0.09 / 0.09 s |
| DMG (download) | 261 MB | 310 MB |
| `.app` on disk | 261 MB | 799 MB |
| Files in bundle | 1 | 3,471 |
| Deprecation warning | yes | none |

**~33× faster startup** on that benchmark. One-file re-extracts the entire 261 MB archive to a
temp directory on every single launch. The cost is +49 MB of download and +538 MB on disk.

Be careful reading that table, though: `--help` is not a real launch. Measured double-click →
first HTTP response on the current one-file build:

```
warm launch 1: 18.4s
warm launch 2: 22.6s
```

So a user waits ~20 seconds staring at a bouncing dock icon. Onedir removes the ~3 s extraction
component, not the other ~15 s, which is import cost (flask, and `preload_av_libs_quietly()`
pulling in both `av` and `cv2` at `server.py`). **Onedir alone will not make this feel fast.** The
follow-up worth scoping separately is deferring that preload and the heavy imports until a
subsystem is actually used — the Studio landing page needs neither `av` nor `cv2`.

Verified the onedir build actually works: `/screenspace/` → 200, `utils.js` → 200, vendored font
→ 200, and the banner reports the bundled version. PyInstaller 6 lays it out the macOS-correct
way — code in `Contents/Frameworks/` (183 entries), data in `Contents/Resources/`, with symlinks
bridging them — which is precisely why it stops complaining.

## Work items

### 2a. Spec — `build/clipgen.spec`

`EXE(pyz, a.scripts, [], exclude_binaries=True, ...)` → `COLLECT(exe, a.binaries, a.datas, ...)`
→ `BUNDLE(coll, ...)`. (A working variant of this was already built and measured; it is a ~10-line
change.) Keep `upx_exclude` in mind if ffmpeg is ever bundled — UPX corrupts signed mach-O.

### 2b. Windows `credentials.json` regression — `cli.get_runtime_working_dir()`

**The one user-visible regression, and it is easy to miss.** Onedir on Windows produces
`dist/clipgen/clipgen.exe` next to `dist/clipgen/_internal/`. `get_runtime_working_dir()` returns
the executable's own directory — which is now *inside* the app folder rather than beside it. A
user told to "put `credentials.json` next to the application" would put it one level up, where it
would no longer be found.

macOS is already correct (#637 added the `.app` walk-up). Windows needs the equivalent: when
frozen and the executable's parent directory is the onedir payload root, resolve to its parent.
Extend the existing tests in `tests/test_cli_args.py` (`test_frozen_macos_working_dir_is_beside_the_bundle`
has the shape to copy).

Also re-check `utils.get_bundled_assets_root()` (`utils.py:652`): its docstring says "one-file
build … extracted to `sys._MEIPASS`". Under onedir `_MEIPASS` points at the payload directory
instead, so the code still works but the comment becomes wrong — update it.

### 2c. CI — `.github/workflows/build-binaries.yml`

- **Drop the `Upload macOS raw binary` step entirely.** Confirmed with the maintainer: shipping
  the `.app` is sufficient. This also removes a live hazard — that step publishes `dist/clipgen`,
  which under onedir silently becomes a *directory* rather than a binary. Remove the step and any
  reference to it in `README.md`.
- Windows artifact becomes a folder, so it must be zipped (or wrapped in an Inno Setup installer —
  worth considering while the packaging surface is already open; a folder of 3,000 files is a poor
  hand-off to a non-technical user).
- macOS DMG staging is unchanged: it already copies `dist/clipgen.app` wholesale.
- Codesigning gets *easier* — a conventional bundle is the well-trodden path, and a prerequisite
  for notarization later.

### 2d. Docs

`README.md` build section (the raw-binary CLI entry point line goes away), `CHANGELOG.md`, and the
`agents/ARCHITECTURE.md` `desktop.py` row if the layout description needs it. `feat:`/`build:` —
note that `build:` does **not** bump `build/VERSION` per `agents/CONTRIBUTING.md`, so decide the
commit type deliberately.

## Phase 2 verification

1. `/check` and `/ui-check` green.
2. Build; **`open dist/clipgen.app`** (Finder-equivalent), confirm the window opens and startup is
   visibly instant rather than a ~3 s wait.
3. `MacOS/clipgen --help` still prints; a real `--gif` run still works.
4. DMG round-trip preserves the exec bit and `codesign --verify --deep --strict` passes.
5. Assets resolve from the bundle: hit `/studio/`, a vendored font, and confirm the banner version.
6. **Windows via `workflow_dispatch`** — the part that cannot be checked locally: the zipped
   folder unpacks and runs, `clipgen.exe --help` prints in cmd, double-click opens the window, and
   `credentials.json` beside the app folder is found.

---

## Sequencing

1. **Phase 1 first, as its own PR.** The shipped `.app` is broken right now; it should not wait
   behind a build-system migration.
2. **Phase 2 after.** It needs a Windows CI run, so it will sit longer in review.

Phase 2 also unblocks the deferred ffmpeg bundling (see the Phase 2 note in the #637 plan): under
onedir a bundled ffmpeg is a real file at a stable path that can be signed individually and is
visibly user-replaceable — the strongest form of the arm's-length GPL argument. Under one-file it
would be extracted to temp on every launch, adding ~80 MB to a startup cost we are trying to
remove.
