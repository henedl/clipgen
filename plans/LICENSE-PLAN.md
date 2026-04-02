# License Distribution Plan

## Context

clipgen bundles 20+ third-party libraries into a PyInstaller binary. All licenses (MIT, BSD, Apache 2.0, HPND) require attribution. The `build/THIRD-PARTY-LICENSES` file contains the full notices. This plan covers how to get that file to end users.

## Steps

- [ ] **1. PyInstaller `--onedir` builds**: The file will be in the output directory alongside the binary. No spec changes needed — just copy it to the dist folder.

- [ ] **2. PyInstaller `--onefile` builds** (current spec): The file cannot be embedded inside the binary itself. It should be distributed alongside the binary (e.g. in a zip/dmg/tar.gz archive). The spec's `datas` list could also include it so it extracts at runtime, but the standard approach is to ship it next to the executable.

- [ ] **3. GitHub Releases**: Include `THIRD-PARTY-LICENSES` as a release asset alongside the binary, or bundle both in an archive.

- [ ] **4. `--licenses` flag** (optional future enhancement): Add a `--licenses` CLI flag that prints the contents of the bundled THIRD-PARTY-LICENSES file to stdout. This requires adding the file to the spec's `datas` list and reading it at runtime via `utils.get_bundled_assets_root()`.
