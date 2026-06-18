# Frontend Redesign — Preparatory Work Plan (research-only)

## Context

Over the next couple of weeks the user is moving the clipgen web UIs from a
**central column layout (max-width: 1120px)** to a **fullscreen layout**.
Motivation: give existing tools room to breathe (Screenspace is the worst
offender today) and accommodate upcoming tools/interfaces. The redesign will
also touch button styles, component sizes, and text sizes throughout.

This document is **research and preparation only**. It captures:
1. What is and isn't safe to widen as-is.
2. Where current code makes column-width assumptions that will need to change.
3. A set of low-risk preparatory passes we can run *before* the redesign so
   that the redesign itself is less cumbersome.

## User decisions

1. **Width target**: ultimately *wider-but-bounded*, but the user wants to
   first try **true edge-to-edge** to discover where that breaks — the
   breakage is the signal for where bounds belong.
2. **Side-by-side panels** (Studio, Insights Builder, Viewer): **fixed-width
   sidebars + fluid main area (IDE-style)**.
3. **Scope**: **all 8 UIs at once**, layout + button/component-size refresh
   **bundled** in the same redesign sprint.
4. **Prep timing**: hold prep passes until the layout target is concrete —
   land prep *with the new target in mind*, not speculatively.
5. **Insights Viewer**: align with the wider/fullscreen treatment — no
   special narrow exception. Its current 900px cap goes away with the rest.
6. **Exported viewers** (`viewer.html`, `insights_viewer.html` already shipped
   to users): treat as ephemeral. No compat layer for old exports;
   regenerate as needed. Matches AGENTS.md.
7. **Sidebar width policy**: unify to a single default token (e.g.
   `--sidebar-width: 360px`) with **per-UI overrides where justified** (e.g.
   insights-builder's wider editing panel). One source of truth, controlled
   exceptions.
8. **Density target**: **looser** — more breathing room per element, fewer
   items visible per panel, larger hit targets. Implication: button/control
   sizes go *up*, list rows get more padding, and we should expect more
   scroll within panels (the fullscreen width compensates by allowing wider
   panels and side-by-side regions).

## Pass 0 — Live token-tweak widget (foundation)

Before the spike, build a small dev-only debug widget that lets us tune CSS
custom properties (design tokens) live in the browser. This is the
iteration loop the spike, the prep passes, and the redesign sprint will all
share. Doing it first removes the edit-save-reload cycle for every "what if
the bound were 1400 vs 1600?" or "what if list rows had 8px more padding?"
question that follows.

### Design

Mirror the theme-toggle precedent in `utils.js:567-626`.

- **Single shared file** `assets/web/dev-token-tweak.js` (with inline
  styles to keep it self-contained — single-use, per AGENTS.md).
- **Mount**: floating panel (top-right by default), collapsible, draggable,
  with an invisible-until-hovered trigger so the dev chrome stays quiet.
- **Token discovery is dynamic but scoped.** At boot, walk
  `document.styleSheets`, find the `:root` rule from `tokens.css`,
  enumerate every `--*` custom property, then filter to
  redesign-relevant prefixes:
  - **Layout & density**: `--space-`, `--text-`, `--radius-`, `--shadow-`,
    `--duration-`, plus the new tokens introduced in Pass 1 (`--layout-`,
    `--sidebar-`, `--bp-`, `--button-`, `--card-width-`, `--icon-size-`).
  - **Core theme colors**: `--color-bg`, `--color-surface`,
    `--color-surface-alt`, `--color-text`, `--color-text-dim`,
    `--color-accent` and its `*-hover/-highlight/-strong-hover` siblings,
    `--color-border`, `--color-selected`, `--color-grid`, the
    `--color-panel-*` cluster.
  - **Excluded** as categorical (not redesign-tunable): `--sev-*`,
    `--color-clip/screen/gif`, `--color-task-*`, `--region-color-*`,
    `--cat-*`, `--color-causes/behaviors/impacts*`.

  New tokens added in later passes appear automatically if they match an
  included prefix; no widget edits required.
- **Control inference from value shape**:
  - Hex / `rgb()` / `hsl()` → `<input type="color">` (text input fallback
    for `rgba()` with alpha and named colors).
  - Ends in `px`/`rem`/`em` → range slider + numeric text box (synced;
    range derived from the current value, e.g. ±2× with sensible step per
    unit).
  - Ends in `ms` → number input.
  - Otherwise → text input.
- **Apply** via `document.documentElement.style.setProperty(name, value)`.
  After every change, call the existing `refreshDetectorColors()` in
  `utils.js:549-561` (cheap and safe on every page).
- **Persistence**: `localStorage["clipgen-token-overrides"]` storing
  `{ tokenName: overrideValue }`. Load and apply on `DOMContentLoaded` so
  reloads keep trial values. Same persistence shape as
  `localStorage["clipgen-theme"]`.
- **Reset**: per-token "✕" button clears the inline property; global
  "Reset all" wipes the localStorage entry.
- **Export**: "Copy as `:root` snippet" button emits a paste-ready CSS
  block of just the overridden tokens, so once values feel right they can
  drop straight into `tokens.css`.
- **Theme-aware**: when the theme toggle flips light/dark, re-read
  defaults from `getComputedStyle(document.documentElement)` so the widget
  shows the right baseline per theme. Hook into the existing
  `initThemeToggle()` flow rather than reinventing the listener.
- **Feature-flag gated**: a `CLIPGEN_DEV_TOKEN_TWEAK` flag in `utils.js`
  (alongside the existing `CLIPGEN_ANIMATED_BG`) lets the widget be
  disabled without ripping out the script tag — flip to `false` for a
  build or when redesign work is paused. The widget bails at load if the
  flag is `false`.

### Dev-only delivery (export gate)

The widget must not ship in standalone exported HTML (`viewer.html`,
`gallery.html`, `insights_viewer.html`, `timeline_viewer.html` finalized
via `viewer.py`). One marker, applied by the HTML and honored by the
inliner, handles this by construction:

1. In dev-mode HTMLs, load via
   `<script src="dev-token-tweak.js" data-dev-only></script>`.
2. In `viewer.py:_generate_viewer_html()` (lines ~195-272), where script
   tags are matched and inlined: skip and strip any `<script>` carrying
   `data-dev-only`. Same treatment for any `<link ... data-dev-only>`.

No runtime "is this an export?" probing in JS — the widget simply isn't
present in exported HTML.

**HTMLs that include the widget script tag** (Flask-served dev UIs):
`studio.html`, `insights-builder.html`, `screenspace.html`,
`transcripts.html`, `gallery.html`. (Convergence is a sub-feature of
Studio — `convergence.css/.js` only, no standalone HTML — so it inherits
the widget through Studio.)

**HTMLs that do not** (already-exported viewers):
`viewer.html`, `insights-viewer.html`, `timeline-viewer.html`. They don't
load `tokens.css` at source either — tokens are inlined at export.

### Verification of Pass 0

- Open each dev UI; widget appears top-right.
- Drag a slider on `--space-3`; spacing changes immediately across the
  page.
- Reload; overrides persist via localStorage.
- Toggle dark mode; widget shows the dark-theme defaults.
- Run an export (`--gallery`, `--viewer`, or insights viewer export from
  Insights Builder); `grep dev-token-tweak` of the exported HTML returns
  nothing.

## The spike: edge-to-edge with the widget in hand

With Pass 0 in place, the next move is a **discovery spike** — strip every
`max-width: 1120px` site, walk all 8 UIs, and use the widget to A/B
candidate bounds and density values live rather than across separate
branches. Concretely:

- Branch off and remove (or set to `none`) every `max-width: 1120px` site
  identified below, in all 8 UIs at once. Touch nothing else.
- Open each UI on a wide display (the user's normal monitor + a 1440px
  laptop pass).
- Walk every screen state — Screenspace with 5+ regions and a multitool
  chain, Studio with a loaded sheet + drag/drop, Insights Builder mid-edit,
  the timeline viewer with filmstrip mode, etc.
- Capture **what looks wrong, what disappears, what stretches grotesquely**.
  That list *is* the input for the bound decision and for the redesign.
- Use the widget to try candidate `--layout-max-width` values (1400, 1600,
  1800, 2000) and tentative density adjustments in the moment, on the same
  loaded screen states, so the bound decision rests on direct comparison
  rather than memory.
- Hit "Copy as `:root` snippet" before closing the spike — the captured
  values become the seed for Pass 1's token defaults.
- Throw the spike branch away (the snippet survives in your clipboard /
  notes).

Predicted findings (from this audit) the spike will surface:
- Studio's `.queue-card`/`.stash-card`/`.preview-card` (140/160/200px) marooned
  in a sea of empty space — confirms we need either responsive cards or
  bounded inner regions.
- Screenspace `#bottomPanel`'s 2-column `1fr 1fr` grid stretching the
  param/multitool column to absurd width while the queue stays half-empty.
- Viewer/Transcripts text columns becoming unreadably long (>120ch) — likely
  candidates for keeping a reading-width bound.
- Convergence sticky 52px participant label looking like a thin stripe.
- Insights-builder popover (`:693-694`) anchoring at the far-left of a now-wide
  grid, with its `window.innerWidth - 170` clamp leaving it stuck at the wrong
  edge.
- Pill options popover in transcripts looking similarly stranded.

The spike answers the most expensive question — *where does the eye start
losing the layout?* — in maybe a half day, and tells us whether the eventual
bound is ~1400, ~1600, or ~2000, and whether it should be uniform or per-UI.

Once the bound is chosen, the prep passes (Section below) become mechanical.

---

## Codebase scope (what's actually there)

`assets/web/` contains 8 UIs (CLAUDE.md mentions 6 — `transcripts` and
`convergence` are also present and follow the same patterns):

| UI | max-width | Sidebars / panels | Notes |
|----|-----------|-------------------|-------|
| viewer | 1120px (`viewer.css:33,167,458`) | sidebar 340px (`:468`), detail flex:1 | Filmstrip + timeline |
| gallery | 1120px (`gallery.css:32,142,157`) | grid `auto-fill, minmax(220px,1fr)` (`:165`) | Most fluid already |
| studio | 1120px (`studio.css:63,239,520`) | flex split, fixed-width cards (140/160/200/520px) | Most hardcoded widths |
| insights-builder | 1120px (`insights-builder.css:32,196`) | sidebar `--sidebar-width:420px` (`:13`), resizable | Has resize handle already |
| insights-viewer | **900px** (`insights-viewer.css:29,129,332`) | grid `minmax(280px,1fr)` (`:142`) | Narrower by design |
| screenspace | 1120px (`screenspace.css:64,190,1744,1782`) | 2-col `1fr 1fr` bottom panel; `height:400px` | Most cramped (user-flagged) |
| transcripts | 1120px (`transcripts.css:39,350`) | similar to viewer | |
| convergence | sticky participant label `width:52px` (`convergence.css:215`) | grid with sticky left labels | |

Shared:
- `tokens.css` — design tokens (spacing, text, radius, shadow, duration, z, severity, content-type, screenspace task colors). No `--layout-max-width` exists yet.
- `utils.js` — `positionTooltipAnchored()` (`:275`), `debounce()` (`:289`), `getComputedStyle` token reads.
- `grid-bg.js` — decorative dot grid; uses `ResizeObserver` and is DPR-aware (`:111-114, 347-357`). Already fullscreen-safe.

Server side (`server.py`) registers blueprints per active mode and serves all
UIs from the same Flask process — no server changes needed for layout work.

---

## Findings: what's already safe to widen

These will simply expand correctly without touching their code:

- **All `@keyframes`** — opacity, rotate, background-position, color shifts. The few `translateY/X` values are ±4–8px UI polish (theme toggle, toast slide-in, pulse dot). No width-tied motion anywhere.
- **All `transition:` declarations** use token durations and animate color/opacity/transform — no width-bound animations.
- **Canvas / overlay rendering in Screenspace** — region coordinates are stored as **normalized 0–1 fractions** of the canvas (`screenspace.js:173-182, 1563-1565`). Drawing scales perfectly with any container width. `canvasCoords()` uses `scaleX = canvas.width / rect.width` (`screenspace.js:877-878`), so mouse interactions work at any size.
- **DPR-aware canvases** in studio (`studio.js:3363-3370, 4101`) and grid-bg (`grid-bg.js:111-114`) read `clientWidth` at draw time and ResizeObserve their parents.
- **Lazy-loading IntersectionObservers** (`viewer.js:297, 382`, `studio.js:3213`) compute against their own scroll roots, not the document. No viewport assumptions.
- **Modals/overlays** (`status`, `confirm`, `log` in studio; lightbox in gallery; settings) all use `position:fixed; inset:0` with flex centering — viewport-relative, fine at any size.
- **Tooltip helper** `positionTooltipAnchored()` (`utils.js:275-285`) clamps to `window.innerWidth` with 4px margins — already correct for any viewport.
- **Drag/drop** in studio uses native HTML5 drag events with DOM reorder, no coordinate math.
- **Studio panel resize** (`studio.js:1033-1108`) reads `window.innerHeight` and clamps to available space — fine for any viewport.
- **Gallery and Insights Viewer grids** already use `repeat(auto-fill, minmax(...))` — they'll add columns automatically at wider widths.

The structural good news: **no animation, drag, or canvas system has hardcoded "the column is N px wide" baked in**. The work is in CSS layout + a small handful of JS popover-positioning sites.

---

## Findings: risk areas (places that *will* need attention)

### A. Hard 1120px caps (the obvious ones)
Every UI has 3–5 sites pinning the main container to 1120px. These are the
trigger for the redesign and the easiest to identify, but they cannot just be
removed wholesale — many cards and panels inside *depend* on the column being
1120px to look right.

Locations: `viewer.css:33,167,458`, `gallery.css:32,142,157`,
`studio.css:63,239,520`, `insights-builder.css:32,196`,
`insights-viewer.css:29,129,332`, `screenspace.css:64,190,1744,1782`,
`transcripts.css:39,350`.

### B. Hardcoded component widths inside the column
These will look undersized or stranded once the parent widens:

- `studio.css:924` — `.queue-card { width: 140px }`
- `studio.css:1150` — `.stash-card { width: 160px }`
- `studio.css:796` — `.preview-card { width: 200px }`
- `studio.css:1903-1904` — `.log-panel { width: 520px; max-width: 92vw }` (this one is already good — `92vw` clamp)
- `insights-builder.css:13` — `--sidebar-width: 420px`
- `insights-builder.css:222` — `.collapsed { width: 48px }`
- `insights-viewer.css:280` — `.detail-artifact-card { width: 220px }`
- `viewer.css:468-469` — `#sidebar { width: 340px; min-width: 260px }`
- `viewer.css:1167` — `#playerPlaceholder { max-height: 480px }`
- `viewer.css:176` — `#timelineTrack { height: 56px }` (vertical, but worth noting)
- `screenspace.css:1796` — `#bottomPanel { height: 400px }` and JS init `panelHeight: 340` (`screenspace.js:81`)
- `screenspace.css:1451-1452, 1467` — param sliders 80px, number inputs 4.5rem (will look stranded in wide layout)
- `convergence.css:215` — sticky `.cv-participant-label { width: 52px; left: 0 }` (sticky-left in a wider layout still works, but visual ratio changes)

### C. JS popover positioning that inherits column-relative coordinates
Two popovers anchor to a trigger's `getBoundingClientRect().left` and clamp
against `window.innerWidth - <popover_width>`. This works today because the
trigger's left is "near the column edge" — once the column becomes the
viewport, the popover's left jumps far right and its width assumption (170px,
240px) may look misplaced even though it doesn't break.

- `insights-builder.js:693-694` — artifact grid popover, hardcoded 170/160 dims
- `transcripts.js:2380-2384` — pill options menu, mounts on `<body>`, uses `wrap.left` directly
- Repositions on resize/scroll: `transcripts.js:2443-2444`

These work but will look "stuck to the wrong edge" depending on where the
trigger sits in the new layout. Visit when the new layout is decided.

### D. Studio table colgroup
`studio.js:833-839` sets `<col>` widths to `3rem`, `3.5rem`, `auto` for the
sheet grid. Fine in fullscreen too (the `auto` column absorbs slack), but if
the new design places the grid in a sub-region, revisit.

### E. Breakpoints that may invert meaning
The only media queries in the project are `@media (max-width: 768px)` (tablet
collapse, in viewer/studio/insights-builder/screenspace) and one
`@media (min-width: 700px)` in insights-builder. **There are zero `min-width`
queries above the current desktop column** — once we go fullscreen, the
"desktop" baseline becomes the *widest* state, and we may want a *new* small-
desktop breakpoint that re-narrows or stacks panels. Worth thinking about
which existing rules become "the wide rule" vs "needs a narrow override".

### F. tokens.css gaps
- No `--layout-max-width`, `--layout-padding-x`, `--layout-content-width`.
- No tokens for the recurring card widths (`--card-width-queue`, `--card-width-stash`, `--card-width-preview`).
- No tokens for sidebar widths (`--sidebar-width-narrow`, `--sidebar-width-wide`).
- No icon-size tokens (icons are 12/14/16/20/24px ad hoc across files).

If we redesign without filling these in first, we'll touch the same dimensions
in ten files instead of one variable.

---

## Preparatory passes (run *after* the spike + bound decision)

These are passes that **don't change anything visible today** but consolidate
the surface area the redesign has to touch. Each pass is a small PR. The
order assumes: spike → choose bound + per-UI policy → run these passes →
redesign sprint (layout + button/size refresh, bundled).

### Pass 1 — introduce layout tokens in `tokens.css` (no consumers yet)
Add:
- `--layout-max-width: 1120px;` (and `1120px` overrides for insights-viewer's
  900px keep working since they're set per-file)
- `--layout-padding-x: var(--space-5);` matching current header padding
- `--card-width-queue: 140px;`
- `--card-width-stash: 160px;`
- `--card-width-preview: 200px;`
- `--sidebar-width-default: 340px;`
- `--sidebar-width-wide: 420px;`
- `--bottom-panel-height: 400px;`
- `--icon-size-xs/sm/md/lg`: 12/14/16/20px

Because the Pass 0 widget discovers tokens dynamically by prefix, every
token added here is immediately tunable in the browser — seed values are
first-pass guesses, not commitments. The seed values themselves come from
the snippet exported at the end of the spike.

Risk: zero. Tokens are unused.

### Pass 2 — migrate the 1120px sites to `var(--layout-max-width)`
Pure find-and-replace across all 8 UI CSS files. No visual change.
Result: redesign just changes one variable. ~25 sites total.

### Pass 3 — migrate hardcoded card/sidebar widths to tokens
`studio.css`, `viewer.css`, `insights-builder.css`, `insights-viewer.css`.
~12 sites. No visual change.

### Pass 4 — extract the screenspace bottom-panel height
`screenspace.css:1796` and `screenspace.js:81` currently both encode
"~340–400px". Pick one variable (CSS var consumed by both — JS reads it via
`getComputedStyle` like task colors do). Already the established pattern in
the project (per AGENTS.md screenspace task colors).

### Pass 5 — audit + add `min-width: 0` to flex-row children that hold scroll/text
This is a defensive pass. Many flex children that contain `overflow:auto` or
long labels (`#regionChips`, `#workflowTabs`, transcript pill rows) work today
because the column constrains them. In wider layouts they can refuse to
shrink, pushing siblings off-screen. Adding `min-width: 0` is invisible at
1120px but prevents nasty surprises at 2400px.

Likely sites (verify before changing each):
- `screenspace.css` — `#regionChips` parent, `#workflowTabs` parent
- `transcripts.css` — pill row (`.pill-options` parent)
- `viewer.css` — `#detailPane` (already `flex:1`, but check children)
- `studio.css` — sheet grid horizontal scroll wrapper

### Pass 6 — name the existing `@media (max-width: 768px)` breakpoint
Replace literal `768px` with `--bp-tablet`. When the redesign adds new
breakpoints (e.g. `--bp-desktop-narrow: 1200px`), they live alongside.
Trivial, but the redesign will introduce 2–3 new breakpoints and we want a
single source of truth.

### Pass 7 — popover-positioning helper
Today `insights-builder.js:693-694` and `transcripts.js:2380-2384` both
hand-roll the same "anchor to trigger, clamp to viewport" math. The shared
`positionTooltipAnchored()` in `utils.js:275-285` already does this for
tooltips with viewport clamps. Refactoring the two popover sites to call a
generalized helper now means the redesign can change positioning policy
("flip below if too low", "right-align if anchor is in right half") in one
place.

(AGENTS.md: "Don't extract helpers unless it is called more than once." —
this is the second use, so it qualifies.)

---

## Implications of the user decisions

The combined answers tighten the prep plan considerably. Concretely:

- **Pass 1 (tokens)** now has firm values to seed:
  - `--sidebar-width: 360px;` (default; per-UI overrides allowed)
  - `--sidebar-width-wide: 440px;` (insights-builder's editing panel)
  - Density tokens get a "looser" baseline — likely
    `--button-height-md` ~36–40px (up from current ~28–32px),
    list-row padding `--space-3` → `--space-4`, etc. Exact values get tuned
    after the spike when we see how looser controls fit at the chosen bound.
  - Insights-viewer's 900px cap is dropped; it now uses the same
    `--layout-max-width` as everyone else.

- **Pass 2 (max-width migration)** simplifies: every site becomes
  `var(--layout-max-width)` with no per-UI exceptions. Insights-viewer's
  900px in `insights-viewer.css:29,129,332` goes through the same swap.

- **Looser density + fullscreen width together** is a meaningful coupling.
  Looser controls eat horizontal space, but fullscreen gives it back —
  they're designed for each other. Worth verifying during the spike that
  the two together don't *also* cause excessive scroll within the
  vertical-bound panels (Screenspace results, transcript list).

- **No compat layer for exports** means we can change the inlined CSS shape
  freely. Worth a one-line check during the redesign that
  `viewer.py` / `insights_server.py` finalize functions still inject CSS
  correctly — but no migration code needed for old exports.

## Still-open (smaller) questions, deferrable until after the spike

1. **Per-UI sidebar overrides — which UIs justify the wide tier?** The
   default is 360px; insights-builder is the obvious candidate for 440px.
   Viewer (340px today) probably collapses to the 360px default. Decide
   when we see them side-by-side.

2. **Looser density numbers.** Approximate ranges are easy; exact
   button/control heights, list-row padding, and font-size bumps get tuned
   visually after the spike, not chosen in the abstract.

3. **Screenspace bottom-panel layout.** Today it's `1fr 1fr` 2-column at
   `height: 400px`. With fullscreen + looser density, options to evaluate
   during/after the spike: 3-column (params | queue | results),
   asymmetric (e.g. `2fr 1fr 1fr`), or vertically split with the timeline
   absorbing the freed top space. Defer the call.

---

## Files most likely to be touched by the prep passes

(For your reference, in case you want to scope/branch differently.)

- `assets/web/dev-token-tweak.js` — new shared widget (Pass 0)
- `assets/web/{studio,insights-builder,screenspace,transcripts,gallery}.html` — add `<script src="dev-token-tweak.js" data-dev-only></script>` (Pass 0)
- `viewer.py` (`_generate_viewer_html()`, lines ~195–272) — strip `data-dev-only` script/link tags during inlining (Pass 0)
- `assets/web/tokens.css` — receives new variables (Pass 1)
- `assets/web/{viewer,gallery,studio,insights-builder,insights-viewer,screenspace,transcripts,convergence}.css` — `1120px` → `var(--layout-max-width)` (Pass 2)
- `assets/web/{studio,viewer,insights-builder,insights-viewer}.css` — card/sidebar width tokens (Pass 3)
- `assets/web/screenspace.css` + `assets/web/screenspace.js` — bottom panel var (Pass 4)
- `assets/web/{screenspace,transcripts,viewer,studio}.css` — `min-width: 0` audit (Pass 5)
- `assets/web/{viewer,studio,insights-builder,screenspace}.css` — name the 768px breakpoint (Pass 6)
- `assets/web/utils.js` + `assets/web/{insights-builder,transcripts}.js` — popover helper (Pass 7)

## Verification (for the prep passes, when we get there)

- Open each UI at 1120px viewport — pixel-identical to before.
- Run `uv run --extra dev pytest -c tests/pytest.ini` — `tests/test_shared_constants.py` covers the JS/Python token mirror; nothing else asserts on CSS values directly.
- No tests cover CSS layout, so the verification is visual: screenshot each UI before and after each pass.

For Pass 0 specifically (see also "Verification of Pass 0" above):
- The widget appears on every dev UI, not on the three exported viewers.
- Exporting any artifact (gallery, viewer, insights viewer) produces HTML
  with no `dev-token-tweak` reference — `grep dev-token-tweak` on the
  output is empty.
- Overrides survive a reload, clear correctly via per-token reset and
  global reset, and the `:root` snippet output round-trips back into
  `tokens.css` without manual cleanup.
