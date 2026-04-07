# Convergence Browser

## Studio Tab 4

---

## Why We're Building This

Clipgen's Studio currently operates per-stream. Researchers work through spreadsheet data, ScreenSpace events, and transcript tags in separate intake tabs, manually cross-referencing by timestamp. This works for single-session analysis, but breaks down when working across multiple participants.

The gap: **cross-participant signal is invisible unless actively sought.** A researcher focused on participant 5 may not notice that participants 1, 3, and 4 all flagged the same moment. The data exists in memory — Studio already holds all session JSONs simultaneously for cross-referencing — but there is no surface that makes recurrence visible passively.

The Convergence Browser closes this gap. It answers two questions researchers ask constantly but currently answer by hand:

- **Did this happen for multiple participants?** (convergence)
- **When did it happen, and how much variance is there?** (distribution)

A tight temporal distribution (everyone hit this at minute 4) is a different finding than a wide one (spread across the session). Both are signal. Neither is currently legible without manual comparison.

---

## What It Is

An interactive, multi-participant timeline view — a 4th tab in Studio. It borrows the visual structure of the Timeline Viewer (already a generated artifact) but makes it interactive and adds filter and selection controls. The output of working in the Convergence Browser feeds directly into the existing artifact and reel fields.

It is not a replacement for the three intake tabs. Researchers who want to work stream-by-stream should continue to do so. The Convergence Browser is a higher-altitude entry point: use it when you want the data to surface candidates rather than hunting manually.

---

## How It Works

### Data Source

All session JSONs are already in memory when Studio is active. No new loading infrastructure is required. The Convergence Browser reads from the same in-memory data the cross-reference system already uses.

### Layout

- **Horizontal axis**: session-relative time (normalized across participants by study design — all participants start at minute 0 and proceed through the same tasks)
- **Rows**: one row per participant, each with sub-tracks per stream (spreadsheet, ScreenSpace, transcript)
- **Convergence overlay**: a density layer showing where N or more participants have flagged events within the same time window — visualized as tightness as well as count, so clustered flags are distinguishable from spread ones

### Filters

Populated dynamically from whatever event type strings exist in the loaded JSONs. No prepackaged taxonomy — the researcher named their detectors when configuring ScreenSpace, and those names are the vocabulary for this study.

Filters operate as **prerequisites for convergence calculation**, not post-filters. The researcher selects an event type first; convergence is then calculated for that subset. This is necessary because ScreenSpace event density varies wildly by detector type (a health bar detector may fire hundreds of times per minute; a loading screen detector may fire three times per session). Mixing them in a single density calculation produces noise.

Controls:

- Filter by stream (spreadsheet / ScreenSpace / transcript)
- Filter by event type (populated from loaded JSONs)
- Convergence threshold (minimum number of participants)
- Time range

### Display Normalisation

Per-track density is normalised for display so sparse tracks don't visually disappear alongside dense ones. Absolute event counts are preserved in the underlying data; this is a display concern only.

### Selection

Click or lasso a region or moment to select it. Selected artifacts are sent directly to the existing artifact or reel field in Studio. The Convergence Browser does not generate its own output format — it is a curation surface that feeds the existing generation pipeline.

---

## Key Design Decisions

**Query interface, not dashboard.** The browser's value is in answering specific cross-participant questions about specific event types — not displaying everything at once. The holistic view is useful for orientation, but the real work happens when the researcher narrows to a type and asks "did this happen for everyone, and when?"

**No interpretation imposed.** The browser surfaces convergence and distribution as evidence. What that evidence means is the researcher's call. This is consistent with clipgen's overall architecture.

**Study-relative taxonomy.** Event type labels are read from the loaded JSONs. The browser cannot and should not make assumptions about what those labels mean across studies. Clipgen must run on thousands of different games in various stages of development — no prepackaged definitions are possible.

**Additive, not disruptive.** The three existing intake tabs are unchanged. The Convergence Browser is an additional entry point for a different mode of working.

---

## Integration Points

- Reads from: all session JSONs currently in Studio memory
- Writes to: existing artifact field and reel field
- Relationship to Timeline Viewer: borrows its visual structure; the Timeline Viewer remains a separate static generation artifact for stakeholder output
- Relationship to Metadata Overview (Tab 5): the cross-stream collision data visible in the Metadata Overview provides useful context before entering the Convergence Browser

---

## Key Findings From Design Discussion

- Studio already holds all session JSONs in memory simultaneously — no new aggregation infrastructure is required as a prerequisite
- Session-relative timestamps are effectively normalised by study design (all participants start at minute 0)
- Fixed time windows for convergence calculation break down given ScreenSpace density variance; event-type-aware calculation is required
- The most common researcher query is temporal distribution ("when did this happen across participants, and is that spread tight or wide?") — threshold filtering alone doesn't capture this
- Task region as a concept is too rigid for freeform playtests; temporal distribution is the more honest and useful framing
- The Convergence Browser makes the intermediate curation step explicit: curate → query/filter → generate, rather than curate → generate
