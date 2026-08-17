# Vendored Heroicons

Tailwind Labs' [Heroicons](https://heroicons.com), the project's primary icon set.
GitHub's Octicons live one level down in [octicon/](octicon/) so the two sets can
never collide on a filename.

- **Version:** 2.2.0 (npm `heroicons@2.2.0`), the `16/solid` ("micro") set
- **Source:** https://github.com/tailwindlabs/heroicons
- **License:** MIT, © Tailwind Labs, Inc. — full notice in
  [build/THIRD-PARTY-LICENSES](../../build/THIRD-PARTY-LICENSES)
- **Count:** 316 `*.svg`, flat in this directory, kebab-case upstream names

## Provenance caveat

These files are **not** byte-identical to the npm package. They were re-exported
at some point before the set was recorded here: arcs are flattened to cubic
béziers, the root carries `fill="none"` and each path a literal `fill="#0F172A"`
(Tailwind slate-900) instead of upstream's `fill="currentColor"`. The *geometry*
matches Heroicons 2.2.0 — verified by comparing against
`unpkg.com/heroicons@2.2.0/16/solid/`.

The hardcoded fill is harmless: every icon is consumed as a CSS `mask-image`
source and painted with `background-color: currentColor`, so the fill colour in
the file is never rendered. See the SVG icons section of [AGENTS.md](../../AGENTS.md).

**Never hand-edit these files.** To change an icon, pick a different one. To
re-vendor the set cleanly:

```bash
mkdir -p /tmp/heroicons assets/icons
curl -sL https://registry.npmjs.org/heroicons/-/heroicons-2.2.0.tgz | tar -xz -C /tmp/heroicons
cp /tmp/heroicons/package/16/solid/*.svg assets/icons/
```

Note that re-vendoring replaces the flattened exports with upstream's
`currentColor` markup, which is fine for the mask-image pipeline but will show up
as a large diff. Update this file and the notice's version if you do.
