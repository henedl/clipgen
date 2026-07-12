# Vendored third-party JS

The only third-party JavaScript in the project. Everything else under
`assets/web/` is hand-written vanilla JS (see AGENTS.md).

## three.min.js

- **Version:** r147 (npm `three@0.147.0`), `build/three.min.js`
- **Source:** https://unpkg.com/three@0.147.0/build/three.min.js
- **License:** MIT (banner retained at the top of the file)
- **Used by:** `map.js` (the Study Map page, `/map/`)

**Do not upgrade this file.** Three.js deprecated the UMD `three.min.js`
builds in r150 and removed them in r160 — newer releases ship ES modules
only, which are incompatible with this repo's no-build-tools rule (plain
`<script>` tags, global `THREE`). r147 is the last clean UMD build; it is
feature-complete for everything the Study Map needs (points, meshes,
raycasting, projection math).

Files in this directory are excluded from the frontend source-convention
tests on purpose (`tests/test_frontend_satellite_wiring.py` globs only the
`assets/web/` root); keep vendored code here, never in the web root.
