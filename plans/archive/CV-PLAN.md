# Computer Vision Navigation & Summarization Exploration

## Context

Architectural exploration driven by user feedback from research managers outside the current user base. Clipgen's CV capabilities are meeting core needs, but the opportunity is in helping researchers *navigate* large video volumes when they don't know what they're looking for yet.

## Key Insight

Studio is batch-oriented with scrubbing and timeline infrastructure already in place. The gap isn't detection—it's intelligent summarization that feeds into existing workflows, turning raw video exploration into guided navigation.

## Recommended Primitives

### Near-term (Quick Wins)

- **Optical Flow Magnitude Over Time**: Generate temporal activity curve showing motion density. CPU-friendly, no training required. Visualize as density overlay on timeline. Researchers jump to activity spikes.
- **Keyframe Extraction**: Auto-mark scene boundaries using perceptual hashing or frame difference. Reduces linear scrubbing friction.

Together, these give researchers a skeleton of the video—entry points into longer sections without watching everything.

### Medium-term (Exploration)

- **Motion Summary Maps**: Heatmap-style temporal visualization of where activity concentrates (menu sections vs. frantic gameplay).
- **Perceptual Clustering**: Group visually similar scenes so navigation is thematic rather than purely temporal.
- **Attention-Guided Scrubbing**: Flag moments where on-screen elements changed significantly, or where visual focus likely shifted (saliency-based).

### Longer-play (Research Phase)

- **Contrastive Learning for Deviation Detection**: Researchers provide examples of "normal gameplay" and the model learns to spot friction or unusual patterns without explicit definition. Zero-shot approach keeps configuration light.

## Constraints to Preserve

- **CPU/GPU flexibility**: Solo researchers on laptops + studio servers running as service. Detectors must degrade gracefully.
- **On-demand recomputation**: Researchers adjust parameters and re-process when they ask — avoid shipping huge precomputed frame caches or blocking the UI on full-video pre-analysis. **This does not mean “never cache”:** reuse existing decode paths, manifest reads, and detector outputs where clipgen already does (see **Relationship to Screenspace detectors**).
- **Researcher agency**: Tool helps orient exploration, doesn't impose interpretation.

## Integration map

| Area | Role | Likely touchpoints |
| --- | --- | --- |
| **Studio** | Batch study view; spreadsheet-linked timeline and markers | [`server.py`](server.py) routes, [`assets/web/studio.js`](assets/web/studio.js) timeline / Screenspace intake density UI |
| **Screenspace** | Single-video deep analysis; per-task detectors and timeline | [`screenspace.py`](screenspace.py), [`screenspace_server.py`](screenspace_server.py), [`assets/web/screenspace.js`](assets/web/screenspace.js) (`amplitudeGraphEnabled` band on timeline — reference for Transcripts friction heatmap) |
| **Viewer** | Exported HTML timeline | [`viewer.py`](viewer.py), viewer timeline assets |
| **Manifest** | Persisted events remain source of truth for detectors | Screenspace manifest events; navigation curves are **additive** overlays, not replacements |

**Studio-first vs Screenspace-first (open):** Optical flow / keyframes could land in Studio first (orient across many participant videos) or Screenspace first (one long session, parameter iteration). Pick one primary surface per primitive before implementation; the other can follow.

## Relationship to Screenspace detectors

[`screenspace.py`](screenspace.py) already computes motion, template match, scene change, and related signals for task detectors. Navigation primitives should **prefer reusing or deriving from those outputs** (e.g. aggregate existing event timestamps into an activity curve) rather than always spawning a second full-video OpenCV pass. Where no detector covers the need (e.g. session-wide optical-flow magnitude without a task), add a focused pipeline — document cost and whether results are stored on the manifest or computed ephemerally per request.

## Primitive roadmap

| Primitive | Phase | UI surface (initial) | Dependencies | Effort (rough) |
| --- | --- | --- | --- | --- |
| Optical flow magnitude over time | Near-term | Studio or Screenspace timeline density band | OpenCV / ffmpeg (existing stack) | Small–medium |
| Keyframe / scene boundaries | Near-term | Timeline markers (Studio + Screenspace) | Frame diff or perceptual hash; no training | Small |
| Motion summary maps | Medium | Screenspace timeline or side panel | Aggregated flow or detector events | Medium |
| Perceptual clustering | Medium | Screenspace gallery or Studio batch | Embeddings or hash buckets; may add dep | Medium–large |
| Attention-guided scrubbing | Medium | Timeline flags | Saliency or change-detection; GPU optional | Medium |
| Contrastive deviation | Long | TBD (likely Screenspace task) | Example-driven model; research | Large |

## Integration Points

- Optical flow and keyframes feed into **existing timeline markers** in Studio and/or Screenspace (same marker hit-test and seek behavior as today).
- Screenspace manifest remains ground truth for detector-produced events; navigation layers read manifest + video metadata, write optional derived series (e.g. `navigation.flow_curve`) only if persistence is justified.
