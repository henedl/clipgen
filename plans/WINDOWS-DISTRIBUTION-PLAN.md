# Windows Distribution Plan

How the Windows build is shipped today, and the options for making it feel less like a
folder of loose parts. Nothing here is urgent; this is a decision record to argue with.

## Context

The Windows artifact is a one-dir PyInstaller build zipped whole:

```
clipgen/
  clipgen.exe          ~1 MB   the bootloader
  _internal/           ~1 GB   python312.dll, every wheel, assets/, bin/ffmpeg.exe
  INSTALL.txt
  THIRD-PARTY-LICENSES
```

Measured on run [31704920152](https://github.com/henedl/clipgen/actions/runs/31704920152):
the uploaded artifact is **440 MB** compressed. Unpacked it is roughly **1 GB**.

Where that goes, roughly (Windows wheel sizes from PyPI, expanded ≈ 2–2.5× compressed):

| Component | Compressed | Why it is there |
|---|---|---|
| `torch` | 122 MB | **Only** pulled in by `easyocr`. No clipgen module imports torch. |
| `ffmpeg.exe` + `ffprobe.exe` | ~130 MB uncompressed | Deliberate — the app must not depend on a user's ffmpeg |
| `scipy` | 37 MB | scikit-image |
| `ctranslate2` | 19 MB | faster-whisper |
| `onnxruntime`, `numpy`, `cv2`, `av`, `skimage`, `Pillow` | ~80 MB | core |

The `_internal` name is PyInstaller's default contents directory. Nothing in `source/`
depends on that literal string — `cli.get_runtime_working_dir()` and
`utils.get_bundled_assets_root()` both derive from `sys._MEIPASS` — so it is free to rename.

## What "more portable" could mean

Three different wishes hide behind the word, and they pull in opposite directions:

1. **One file** — the download and the thing you double-click are the same object.
2. **Hidden internals** — the user never sees a folder of DLLs, whatever the shape on disk.
3. **Smaller** — 440 MB is a lot to move around, whatever the shape.

Only (3) is a real constraint. (1) and (2) are presentation.

## Options

### A. Rename `_internal` — cosmetic, ~free

`COLLECT(..., contents_directory="lib")` (PyInstaller ≥ 6.0; we pin 6.19.0). Setting it to
`"."` restores the pre-6.0 flat layout, which is strictly worse — 800 loose DLLs beside the exe.

- **Pro:** one line; no runtime cost; verified safe against our `_MEIPASS`-derived path logic.
- **Con:** purely cosmetic. The folder is still there, just better named.

### B. Windows installer (Inno Setup) — recommended

Ship `clipgen-setup.exe` alongside the zip. It installs into `%LOCALAPPDATA%\Programs\clipgen`,
creates Start Menu and desktop shortcuts, and registers an uninstaller.

- **Pro:** the standard Windows shape. The user downloads one file, double-clicks, and
  launches from the Start Menu — `_internal` is never seen. Gives a real uninstall.
  Inno Setup is free and scriptable; one CI step (`choco install innosetup`) plus an `.iss`.
- **Pro:** composes with everything else here — it is a wrapper, not a rebuild.
- **Con:** SmartScreen still warns on an unsigned installer, exactly as it does today on the
  loose exe. Fixing *that* needs a code-signing certificate (Azure Trusted Signing is the
  cheap route), which is a separate decision.
- **Con:** an installed app is less portable in the USB-stick sense — so keep shipping the
  zip too. Two artifacts, one build.

### C. One-file exe (`--onefile`) — do not

- **Pro:** genuinely one file.
- **Con:** the bootloader unpacks the entire ~1 GB payload to `%TEMP%\_MEIxxxxxx` on **every
  launch** and deletes it on exit. We already measured this shape on macOS before abandoning
  it: **17.6 s vs 1.1 s** to first response. Windows would be worse, since Defender scans each
  extracted DLL as it lands.
- **Con:** PyInstaller deprecated one-file + windowed on macOS and makes it an error in v7.0,
  so we would be maintaining two different shapes per platform.

### D. Nuitka — not now

Compile to C instead of bundling an interpreter. `--onefile` there can cache its extraction
across launches, which is the one thing PyInstaller one-file cannot do.

- **Pro:** faster startup, single file, and the payload is real compiled code.
- **Con:** compiling a torch/cv2/easyocr dependency graph is slow (tens of minutes) and
  fragile, with its own per-package plugin story.
- **Con:** throws away everything encoded in `build/clipgen.spec` — the `freeze_support()`
  handling, the `_MEIPASS` path rules, the ffmpeg vendoring, the empty-bundle guards. That
  knowledge cost real incidents to acquire. Not worth spending for a cosmetic gain.

### E. Drop torch from the bundle — the only option that changes the number

`torch` is 122 MB compressed and is pulled in **solely** by `easyocr`, for the Screenspace
Text and Numbers tools. clipgen never imports torch itself, and `screenspace_ocr.py` already
imports easyocr lazily behind `utils.require_optional("easyocr", "text scan")` — the optional
seam exists.

Two ways to take it:

1. **Unbundle** — exclude easyocr/torch from the spec and fetch them on first use of an OCR
   tool, the way the app already offers to install Ollama.
   - *Pro:* the download drops by roughly a third, for a spec `excludes` entry and a
     first-use flow. *Con:* `require_optional`'s message (`uv add easyocr`) is written for a
     source checkout and would need to become a real download prompt in a frozen app.
2. **Replace** — swap easyocr for an ONNX-runtime OCR (RapidOCR and similar), removing torch
   from the dependency graph entirely.
   - *Pro:* biggest win, no runtime download, `onnxruntime` is already in the tree.
   - *Con:* a genuine engineering task. The glyph confusion-folding, numeric allowlists,
     and calibration paths in `screenspace_ocr.py` are tuned to easyocr's output shape and
     would all need re-tuning and re-verifying against real footage.

### F. MSIX / winget — later, if ever

Modern packaging with clean install/uninstall and auto-update.

- **Pro:** the most "proper" Windows answer; `winget install clipgen` is a nice story.
- **Con:** requires signing, more ceremony than a handful of researchers need, and it does
  not make anything smaller.

## Recommendation

1. **B (Inno Setup installer), keeping the zip** — the whole of the "loose folder" complaint
   disappears for the people who want an app, and the portable zip stays for those who don't.
2. **A** at the same time if the folder name still bothers you in the zip. It is one line.
3. **E** is the only thing that makes clipgen meaningfully smaller. Worth doing on its own
   schedule, starting with E1 (unbundle) since it is reversible and much cheaper than E2.
4. **Not C, not D.** Both trade measured startup time or hard-won packaging knowledge for
   presentation we can get from B.

## Status

- [ ] A — rename the contents directory
- [ ] B — Inno Setup installer in `build-binaries.yml`
- [ ] E1 — unbundle easyocr/torch behind a first-use download
- [ ] E2 — evaluate an ONNX OCR backend to drop torch entirely
- ~~C~~ — rejected, see above
- ~~D~~ — rejected, see above
