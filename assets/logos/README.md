# clipgen — brand assets

Three nested L-shapes cascading from top-left, suggesting layered frames / clip ranges.

## Files

### Mark (icon only)
- `mark-dark.svg` — black stroke, transparent bg
- `mark-light.svg` — light stroke, transparent bg
- `favicon.svg` — auto-switches via `prefers-color-scheme`

### Lockups
- `lockup-horizontal-dark.svg` — mark + wordmark, dark
- `lockup-horizontal-light.svg` — mark + wordmark, light

### Favicons (PNG, transparent bg)
- `favicon-{16,32,48,64,192,512}.png` — dark mark
- `favicon-light-{16,32,48,192,512}.png` — light mark for dark UIs

### Apple touch icons (PNG, solid bg)
- `apple-touch-icon-{120,152,167,180}.png` — black bg, light mark (default)
- `apple-touch-icon-light-180.png` — light bg, dark mark (alt)

### Animation
- `clipgen-mark.jsx` — exports `<StackedF>` and `<StackedFAnimated>` React components. The animated version cascades the three L-shapes in over ~1.4s.

## HTML usage

```html
<!-- Standard favicon set -->
<link rel="icon" type="image/svg+xml" href="/favicon.svg" />
<link rel="icon" type="image/png" sizes="32x32" href="/favicon-32.png" />
<link rel="icon" type="image/png" sizes="192x192" href="/favicon-192.png" />

<!-- Apple touch icon (iOS bookmark) -->
<link rel="apple-touch-icon" sizes="180x180" href="/apple-touch-icon-180.png" />
<link rel="apple-touch-icon" sizes="167x167" href="/apple-touch-icon-167.png" />
<link rel="apple-touch-icon" sizes="152x152" href="/apple-touch-icon-152.png" />
<link rel="apple-touch-icon" sizes="120x120" href="/apple-touch-icon-120.png" />
```

## Mark specs

- viewBox: `0 0 100 100`
- Stroke: `11` (relative to 100u viewBox)
- Stroke caps: `square`, joins: `miter`
- Paths:
  - `M 18 18 L 82 18 L 82 82` (outer L)
  - `M 18 40 L 60 40 L 60 82` (middle L)
  - `M 18 62 L 38 62 L 38 82` (inner L)
- Min display size: 14px
- Clear space: 1× the stroke width (11u in viewBox)

## Color

Mono only.
- Dark: `#0a0a0a`
- Light: `#fafaf7`
