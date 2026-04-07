# Metadata Overview & Export

## Studio Tab 5

---

## Why We're Building This

Clipgen produces structured artifacts, but currently has no surface for summarising the data that underlies them. Researchers can visually eyeball event distributions in the intake tabs, but there is no way to get quick aggregate statistics — how often did an event type occur, when did it first appear, how many participants were affected — without manually reviewing individual artifacts or session JSONs.

This creates two practical problems:

**1. No fast sanity check.** A misconfigured ScreenSpace detector — wrong threshold, wrong reference frame — shows up as anomalous event counts. Without a summary view, this isn't visible until the researcher is deep into artifact curation and wonders why one participant has 10x the events of everyone else.

**2. No metadata export.** Researchers working in different reporting contexts — stakeholder decks, research repositories, internal documentation — frequently need summary statistics alongside clip artifacts. Currently these must be compiled manually.

The Metadata Overview solves both. It is a computation and display problem, not an infrastructure problem — all the data required is already in memory when Studio is active.

---

## What It Is

A 5th tab in Studio that presents aggregate statistics across all loaded sessions and streams, with export to JSON and CSV. It is a read-only summary surface — no curation or selection happens here. Its output is statistics, not artifacts.

It also serves as a natural starting point before entering the Convergence Browser: a researcher can orient themselves to the overall shape of the data before drilling into cross-participant queries.

---

## What It Shows

### Per Event Type

- Total occurrence count across all participants
- First occurrence (earliest across all sessions)
- Last occurrence (latest across all sessions)
- Mean time of occurrence across participants
- Participant coverage: appeared in X of N loaded sessions

### Cross-Stream Collisions

How often a ScreenSpace event co-occurs with a spreadsheet flag within a configurable time window. This is a lightweight version of the convergence calculation — useful to see in aggregate before going into the Convergence Browser for deeper exploration.

### Session-Level Summary

Per participant, per stream: total flagged moments. Useful for spotting outliers at a glance — a participant with anomalous event counts may indicate a detector misconfiguration or a genuinely unusual session worth noting.

### Examples of Researcher Queries This Answers

- When was a loading screen first detected across all participants?
- How often did the health bar change event fire per session on average?
- How many participants had a spreadsheet flag within 5 seconds of a ScreenSpace event of type X?
- Which participants had no transcript tags in a given time window?

---

## Export

Two formats, both generated on demand:

- **JSON**: consistent with clipgen's existing artifact format, for programmatic use or archiving alongside other session outputs
- **CSV**: for researchers pulling data into their own reporting tools, stakeholder spreadsheets, or external analysis

Given that clipgen must fit into many different reporting contexts across many different organisations, minimal friction on export format is a design requirement, not a nice-to-have.

---

## Key Design Decisions

**Read-only.** No curation happens in this tab. It is a summary of what has been loaded, not a workspace. Keeping this boundary clean prevents the tab from becoming a parallel intake surface.

**Study-relative.** Like the Convergence Browser, all statistics are derived from whatever event types and labels exist in the loaded JSONs. No assumptions about taxonomy.

**QA function is first-class.** The session-level outlier view is explicitly useful for catching misconfigured detectors before committing to full artifact generation. This should be surfaced clearly, not buried.

**Configurable collision window.** The time window for cross-stream collision detection should be researcher-adjustable. The right window varies by what is being studied — a fast UI interaction has a different meaningful window than a slow narrative moment.

---

## Integration Points

- Reads from: all session JSONs currently in Studio memory
- Writes to: JSON and CSV export files
- Relationship to Convergence Browser (Tab 4): provides orienting context before the researcher enters the Convergence Browser; the cross-stream collision data here is a lighter-weight preview of what the browser explores interactively
- Relationship to existing artifacts: the metadata export is a companion to clip artifacts, not a replacement — intended to travel alongside them in reporting

---

## Key Findings From Design Discussion

- All data required for this tab is already in Studio memory — this is a computation and display problem only
- Researchers currently have no fast path to aggregate statistics; this must be compiled manually from individual session reviews
- The session outlier view (anomalous event counts per participant) doubles as a detector QA tool, catching misconfiguration early in the workflow
- Export format flexibility is a hard requirement given clipgen's need to fit into diverse organisational reporting contexts
- This tab is lower implementation complexity than the Convergence Browser and could be built first, with its collision data informing what queries to prioritise in the browser
