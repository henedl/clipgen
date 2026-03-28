# CSS Design Token System — Implementation Plan

## Problem

Across 60+ Conductor sessions, different agents made slightly different spacing, font size, border radius, shadow, transition, and z-index choices. The color system is well-tokenized (CSS variables with light/dark theme support), but everything else is ad-hoc. This causes visual drift and makes it harder for agents to produce consistent CSS on the first try.

## Goal

Define a small, opinionated set of CSS design tokens that agents can reference by name instead of guessing raw values. Adopt incrementally — new code uses tokens, existing code converts opportunistically.

## Non-goals

- CSS framework or build step (stays vanilla CSS)
- Rewriting all 6500 lines of existing CSS in one pass
- Tokenizing one-off layout dimensions (sidebar widths, grid templates)
- Changing domain-specific semantic tokens that already work well (severity colors, task type colors, artifact type colors)

---

## Token Definitions

### Base

```css
:root {
  font-size: 16px;  /* changed from 15px for cleaner rem math */
}
```

**Migration note:** Changing from 15px to 16px will make everything ~6.7% smaller in rem terms. Every CSS file sets `html { font-size: 15px; }` — all six need updating to 16px. Existing rem values will render slightly smaller, but since we're also standardizing those values to the new token scale, this is absorbed by the migration. The visual difference at any single value is sub-pixel and won't be noticeable in isolation; the cumulative effect is a very slight tightening that gets corrected as token adoption proceeds.

### Spacing

8-point base scale. Covers all current usage patterns.

```css
--space-1: 0.25rem;   /*  4px — tight: inline icon gaps, badge padding */
--space-2: 0.5rem;    /*  8px — small: form element gaps, compact card padding */
--space-3: 0.75rem;   /* 12px — medium-small: card body padding, list gaps */
--space-4: 1rem;      /* 16px — medium: standard section padding, grid gaps */
--space-5: 1.5rem;    /* 24px — large: header padding, panel margins */
--space-6: 2rem;      /* 32px — extra-large: major section spacing */
--space-8: 3rem;      /* 48px — jumbo: empty states, hero areas */
```

The jump from 6→8 is intentional — 2.5rem (`--space-7`) barely appears in current code.

**Cluster mapping** (how current scattered values collapse):
| Current values | Token |
|---|---|
| 0.2rem, 0.25rem, 3px, 4px | `--space-1` |
| 0.3rem, 0.35rem, 0.4rem, 0.45rem, 0.5rem, 6px, 8px | `--space-2` |
| 0.6rem, 0.7rem, 0.75rem, 0.8rem | `--space-3` |
| 0.9rem, 1rem | `--space-4` |
| 1.2rem, 1.5rem | `--space-5` |
| 1.8rem, 2rem | `--space-6` |
| 2.5rem, 3rem | `--space-8` |

### Typography

Modular scale based on 16px root. Intentionally few steps — the current code has ~15 distinct font sizes that should collapse to 7.

```css
--text-xs: 0.75rem;    /* 12px — badges, tiny labels, counts */
--text-sm: 0.8125rem;  /* 13px — secondary text, metadata, timestamps */
--text-base: 0.875rem; /* 14px — body copy, form inputs, buttons */
--text-md: 1rem;       /* 16px — prominent body text, card titles */
--text-lg: 1.125rem;   /* 18px — section headings */
--text-xl: 1.25rem;    /* 20px — page titles, major headings */
--text-2xl: 1.5rem;    /* 24px — hero text (rare) */
```

**Cluster mapping:**
| Current values | Token |
|---|---|
| 0.6rem, 0.62rem, 0.65rem, 0.68rem, 0.7rem, 0.72rem, 0.75rem | `--text-xs` |
| 0.78rem, 0.8rem, 0.82rem, 0.85rem | `--text-sm` |
| 0.88rem, 0.9rem | `--text-base` |
| 0.95rem, 1rem | `--text-md` |
| 1.1rem, 1.15rem | `--text-lg` |
| 1.2rem, 1.25rem, 1.3rem | `--text-xl` |
| 1.4rem, 1.5rem | `--text-2xl` |

### Border Radius

```css
--radius-sm: 4px;   /* inputs, small badges, inline elements */
--radius-md: 8px;   /* cards, panels, buttons (current --radius value) */
--radius-lg: 14px;  /* major containers, modals, header bars */
--radius-full: 999px; /* pills, circular buttons, toggles */
```

The existing `--radius: 8px` variable becomes an alias for `--radius-md`. Keep `--radius` as-is for backwards compatibility — existing code that uses `var(--radius)` already gets the right value.

**Cluster mapping:**
| Current values | Token |
|---|---|
| 2px, 3px, 4px | `--radius-sm` |
| 6px, 8px, `var(--radius)`, `calc(var(--radius) - 2px)` | `--radius-md` |
| 10px, 12px, 14px, 18px | `--radius-lg` |
| 50%, 999px | `--radius-full` |

### Shadows / Elevation

Four elevation levels. Each defined as a full `box-shadow` value using existing `--color-panel-shadow`.

```css
--shadow-sm: 0 1px 3px var(--color-panel-shadow);       /* subtle lift: cards at rest */
--shadow-md: 0 4px 12px var(--color-panel-shadow);      /* moderate lift: cards on hover, dropdowns */
--shadow-lg: 0 8px 24px var(--color-panel-shadow);      /* high lift: popovers, floating panels */
--shadow-xl: 0 24px 60px var(--color-panel-shadow);     /* max lift: modals, lightbox overlays */
```

Plus a focus ring token:
```css
--shadow-focus: 0 0 0 2px var(--color-accent);          /* focus/selection ring */
```

### Transitions

Three durations covering all current usage. Named by purpose, not speed.

```css
--duration-fast: 150ms;    /* hover states, color changes, opacity — instant feedback */
--duration-normal: 250ms;  /* transforms, slides, panel transitions */
--duration-slow: 350ms;    /* height/width changes, complex layout shifts */
```

Standard easing: `ease` for most, `ease-in-out` for continuous/looping animations. Not tokenized — just documented as convention.

**Cluster mapping:**
| Current values | Token |
|---|---|
| 60ms, 0.1s, 0.12s, 0.15s, 0.18s | `--duration-fast` |
| 0.2s, 0.22s, 0.25s | `--duration-normal` |
| 0.3s, 0.35s, 0.5s | `--duration-slow` |

### Z-Index

Five named layers. Generous gaps for future insertion.

```css
--z-float: 10;       /* floating badges, sticky elements, resize handles */
--z-dropdown: 100;   /* dropdowns, popovers, context menus */
--z-modal: 1000;     /* modals, dialogs, settings panels */
--z-overlay: 2000;   /* lightbox overlays, full-screen takeovers */
--z-toast: 3000;     /* toast notifications — always on top */
```

**Current value mapping:**
| Current values | Token |
|---|---|
| 1, 2, 5, 10, 20 | `--z-float` |
| 100, 200, 300 | `--z-dropdown` |
| 999, 1000, 1001 | `--z-modal` |
| 2000 (tooltip in viewer) | `--z-overlay` |
| (new — currently toast uses 300) | `--z-toast` |

---

## File Structure

### New file: `assets/web/tokens.css`

Contains all token definitions in `:root`, with a cheat-sheet comment block at the top for agent reference. Also contains the dark-mode overrides for shadow tokens (since `--color-panel-shadow` already adapts via the existing theme toggle, shadows auto-adapt — but if any shadow token needs dark-specific values, they go here).

Structure:
```css
/*
 * clipgen design tokens — quick reference for agents
 *
 * SPACING:    --space-1 (4px) | --space-2 (8px) | --space-3 (12px) | --space-4 (16px) | --space-5 (24px) | --space-6 (32px) | --space-8 (48px)
 * TEXT:       --text-xs (12px) | --text-sm (13px) | --text-base (14px) | --text-md (16px) | --text-lg (18px) | --text-xl (20px) | --text-2xl (24px)
 * RADIUS:     --radius-sm (4px) | --radius-md (8px) | --radius-lg (14px) | --radius-full (999px)
 * SHADOW:     --shadow-sm | --shadow-md | --shadow-lg | --shadow-xl | --shadow-focus
 * DURATION:   --duration-fast (150ms) | --duration-normal (250ms) | --duration-slow (350ms)
 * Z-INDEX:    --z-float (10) | --z-dropdown (100) | --z-modal (1000) | --z-overlay (2000) | --z-toast (3000)
 */

:root {
  /* ... all token definitions ... */
}
```

### Existing files: `studio.css`, `screenspace.css`, `insights-builder.css`, `gallery.css`, `viewer.css`, `insights-viewer.css`

Each file keeps its own `:root` block for color variables and domain-specific variables. No changes to color definitions. The `html { font-size: 15px; }` rule in each file changes to `16px`.

Remove `--radius: 8px` from each file's `:root` — it moves to `tokens.css` (as `--radius-md`, with `--radius` kept as alias).

### HTML files: `studio.html`, `screenspace.html`, `insights-builder.html`, `gallery.html`

Add `<link rel="stylesheet" href="tokens.css">` before the page-specific CSS `<link>`. These are Flask-served, so this just works.

### Standalone inlined viewers: `viewer.html`, `insights-viewer.html`

These are generated by Python (`generate_timeline_viewer()`, `generate_insights_viewer()`) as self-contained HTML files with CSS inlined. The token values need to be duplicated into their `:root` blocks by the Python inlining logic. Since `viewer.py` already reads and inlines CSS content, it can read `tokens.css` and prepend its `:root` block to the inlined styles.

---

## Migration Strategy

### Phase 1: Foundation (one session)

1. Create `tokens.css` with all token definitions
2. Change `html { font-size: 15px }` → `16px` in all six CSS files
3. Add `<link>` to `tokens.css` in all four Flask-served HTML files
4. Update `viewer.py` inlining logic to include `tokens.css` content for standalone viewers
5. Add `--radius` as alias for `--radius-md` in `tokens.css` for backwards compat
6. Remove `--radius` definition from individual CSS files (now comes from `tokens.css`)
7. Run tests, visual spot-check

### Phase 2: Documentation (same session as Phase 1)

1. Add token usage rules to `AGENTS.md` under Learned User Preferences:
   - "Use design tokens from `tokens.css` for spacing (`--space-N`), font sizes (`--text-N`), border radius (`--radius-N`), shadows (`--shadow-N`), transitions (`--duration-N`), and z-index (`--z-N`). Never write raw `rem`/`px` values for these properties in new code."
2. Add a brief Tokens section to `CLAUDE.md` referencing `tokens.css` and listing the token names

### Phase 3: Incremental adoption (ongoing, across future sessions)

- **All new CSS** uses tokens exclusively for the tokenized properties
- **When editing existing CSS** for a feature or fix, convert values in the touched area to tokens
- No dedicated "convert everything" session — migration happens organically as code is touched

### Phase 4: Cleanup (optional, eventual)

Once most values are converted through organic adoption:
- Audit remaining raw values with a grep for hardcoded `rem`/`px` in spacing/font-size/radius/shadow/z-index properties
- Convert stragglers in a focused pass

---

## What This Doesn't Change

- **Color system**: Already well-tokenized with CSS variables and dark mode. No changes.
- **Severity/artifact/task type colors**: Domain-specific semantic tokens. No changes.
- **Font families**: Already using system fonts and `--font-mono`. No changes.
- **Animation keyframes**: Token the duration via `--duration-*`, but keyframe definitions stay per-component.
- **Layout dimensions**: Sidebar widths, grid column sizes, specific component dimensions stay as raw values — these are structural, not design tokens.
- **Existing `--radius` usages**: Still work because `--radius` is aliased to `--radius-md`.

---

## Verification

After Phase 1:
- `uv run pytest -c tests/pytest.ini` passes (viewer inlining tests especially)
- Visual spot-check: launch Studio, Insights, Screenspace and confirm no layout breakage from the 15px→16px base size change
- Confirm `tokens.css` loads in all four Flask-served interfaces
- Confirm standalone viewer HTML contains token values in its inlined styles

Ongoing:
- Agents reference the cheat-sheet comment in `tokens.css` and produce correct token usage without guidance
