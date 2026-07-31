# Vendored Octicons

GitHub's [Octicons](https://primer.style/octicons/), used alongside the Heroicons
in `assets/icons/`. Kept in this subdirectory so the two sets never collide on a
filename and so provenance stays obvious at a glance.

- **Version:** 19.32.0 (npm `@primer/octicons@19.32.0`), from `build/svg/`
- **Source:** https://registry.npmjs.org/@primer/octicons/-/octicons-19.32.0.tgz
- **License:** MIT, © GitHub Inc.
- **Used by:** `.ai-agent-badge` in `transcripts.css` (`dependabot-16.svg`) — the
  glyph marking a button that starts a local Ollama thinking agent.

Only the **380 `*-16.svg`** files are vendored (~215 KB). Upstream also ships
12/24/48/96 px variants, but every icon already in `assets/icons/` is authored at
a 16×16 viewBox, so the other sizes would only add noise. Filenames are kept
verbatim, size suffix included, so each maps 1:1 to its name on primer.style.

Reproduce the set with:

```bash
mkdir -p /tmp/octi assets/icons/octicon
curl -sL https://registry.npmjs.org/@primer/octicons/-/octicons-19.32.0.tgz | tar -xz -C /tmp/octi
cp /tmp/octi/package/build/svg/*-16.svg assets/icons/octicon/
```

**Never hand-edit these files.** They are consumed as CSS `mask-image` sources and
painted with `background-color: currentColor`, so their fill colors are irrelevant
— there is no reason to touch them. To change an icon, pick a different one; to
update the set, re-run the command above with a new version and update this file.

The blueprint icon routes are `@bp.route("/icons/<path:filename>")`
(`source/utils.py`), so these serve as `/<prefix>/icons/octicon/<name>-16.svg`
with no server change.
