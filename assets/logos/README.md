# clipgen — brand assets

## Files

### Mark
- `mark-dark.svg` / `mark-light.svg` — transparent bg
- `favicon.svg` — auto-switches via `prefers-color-scheme`

### Lockups
- `lockup-horizontal-dark.svg` / `lockup-horizontal-light.svg`

### Favicons (PNG, transparent)
- `favicon-{16,32,48,64,192,512}.png` — dark mark
- `favicon-light-{16,32,48,192,512}.png` — light mark

### Apple touch icons (PNG, solid bg)
- `apple-touch-icon-{120,152,167,180}.png` — black bg, light mark
- `apple-touch-icon-light-180.png` — light bg, dark mark

### Animation
- `clipgen-mark.jsx` — Reference implementation `<StackedF>` and `<StackedFAnimated>` (1.4s cascade)

#### Where to use

Once on page load: drop it in the header — no replay logic needed.
Onboarding / splash screens: render with key={mountedAt} so each visit replays.
Loading state: loop by adding animation-iteration-count: infinite and an idle gap (e.g. extend keyframes to 0% draw → 70% drawn → 100% drawn-and-hold, then animation-duration: 2.8s).

#### Animation notes for implementing agent

Don't change the path data or stroke-width — both affect visual weight. Use F.3 (stroke 13) at small sizes (<32px), F.2 (stroke 11) at large display sizes if you want a slightly lighter feel.
Color is currentColor — set the SVG's parent color (or the SVG's own color) to swap themes; no need to fork files.
Do not add easing variants per-line — uniform easing is the design; staggered timing is what creates the cascade.
Don't fade lines in (no opacity animation) — the draw-on stroke is the entire effect.
Min display size: 14px. Below that, swap to a static SVG — animations on tiny SVGs look broken.
Reference implementation: exports-f3/clipgen-mark.jsx — the React <StackedFAnimated> component is one valid implementation, but the CSS-only version above is lighter and preferred for production frontends.

## HTML usage

```html
<link rel="icon" type="image/svg+xml" href="/favicon.svg" />
<link rel="icon" type="image/png" sizes="32x32" href="/favicon-32.png" />
<link rel="icon" type="image/png" sizes="192x192" href="/favicon-192.png" />
<link rel="apple-touch-icon" sizes="180x180" href="/apple-touch-icon-180.png" />
<link rel="apple-touch-icon" sizes="167x167" href="/apple-touch-icon-167.png" />
<link rel="apple-touch-icon" sizes="152x152" href="/apple-touch-icon-152.png" />
<link rel="apple-touch-icon" sizes="120x120" href="/apple-touch-icon-120.png" />
```

## Mark specs

- viewBox: `0 0 100 100`
- Stroke: `13`
- Stroke caps: `square`, joins: `miter`
- Paths:
  - `M 18 18 L 82 18 L 82 82` (outer L)
  - `M 18 40 L 60 40 L 60 82` (middle L)
  - `M 18 62 L 38 62 L 38 82` (inner L)
- Min display size: 14px
- Clear space: 1× stroke width (13u)

## Color
- Dark: `#0a0a0a`
- Light: `#fafaf7`
