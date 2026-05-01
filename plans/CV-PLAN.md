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
- **Real-time iteration**: No pre-caching. Researchers adjust parameters and re-process on demand. Keep latency manageable.
- **Researcher agency**: Tool helps orient exploration, doesn't impose interpretation.

## Integration Points

- Optical flow and keyframes feed into existing timeline markers in Studio.
- Screenspace manifest remains ground truth; navigation layer is additive, not prescriptive.
