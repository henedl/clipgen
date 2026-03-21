# Insights Feature Plan

## Problem

clipgen serves UX researchers well as a workbench for generating artifacts from research sessions. However, there is no structured path from raw artifacts to a deliverable that stakeholders can consume. Researchers currently assemble findings outside of clipgen (slide decks, reports), losing the connection between the narrative and the underlying evidence.

## Design Principle

The value of research is the interpretation, not the data. Stakeholders should experience findings through the researcher's curated narrative, not by browsing raw clips. The tool should make the researcher's argument legible, not replace their analytical role.

The unit of persuasion is a **narrative arc**: cause → behavior → impact, supported by 1-3 examples per bucket. If a reader can see what caused a behavior, witness the behavior, and understand its impact — with clear connections between them — they are typically convinced enough to act.

**"The software works at the speed of the researcher."** Artifact generation and insight authoring are architecturally separate, but the boundary must be invisible to the researcher. The builder has full, instant access to all available artifacts — browsable, searchable, draggable — so the researcher never leaves the builder to find evidence. Two tools, one fluid experience.

## Concept: Insights

An **Insight** is a researcher-authored finding composed of:

- **Title** — a concise claim (e.g. "Players don't engage with social features")
- **Summary** - a researcher-written description of why this matters and summarizes the key points (including suggestions and references to previous findings)
- **Severity** — how critical the finding is
- **Three evidence buckets:**
  - **Causes** — what led to the observed behavior (artifacts + researcher narrative)
  - **Behaviors** — what users actually did (artifacts + researcher narrative)
  - **Impacts** — consequences of the behavior (artifacts + researcher narrative)
- **Narrative annotations** — researcher-written text in each bucket connecting the evidence and telling the story. Causes and impacts are where the researcher's analysis matters most.
- **Timeline context** — where the finding occurred in the general session flow

Insights reference artifacts by ID from the existing manifest. Artifacts don't need to know they're part of an insight — the insight layer is additive, not a modification of the existing system.

Insights can span multiple participants and sessions. For now, they are limited to a single study.

## Three-Tier Deliverable

The reader experience is structured as three tiers of increasing detail:

1. **Insights index** — the landing page. Lists all findings with title, severity, and summary. Answers: "what did you find?"
2. **Insight detail** — drill into one finding. Cause → behavior → impact structure with inline artifacts and researcher narrative. Answers: "why should I believe this and why does it matter?"
3. **Full artifact view** — the existing timeline/gallery viewer, accessible via tab or link. Answers: "show me everything."

Most stakeholders live in tiers 1 and 2. Curious readers can "back out" to tier 3 for the complete evidence set. This respects the researcher's interpretive authority while giving readers an escape valve for deeper exploration.

## Workflow

1. Researcher generates artifacts across sessions as they do today (clipgen always runs "after" observations are collected in the spreadsheet)
2. Researcher opens the insight builder, which **immediately loads all artifacts from the manifest** — browsable, searchable, and ready to drag into buckets. The researcher arranges evidence into cause → behavior → impact buckets and adds narrative annotations without leaving the builder.
3. Researcher exports/generates the Insight view
4. Reader receives the deliverable and navigates the three tiers

## Architecture

- **`insights_manifest.json`** — new persistence layer alongside `clipgen_manifest.json`. Each insight has an ID, title, severity, three buckets with artifact ID references and narrative text, and metadata (creation date, status such as draft/final).
- **Insight builder interface** — web view for constructing insights. Opens with full awareness of the artifact manifest — all artifacts are immediately browsable, searchable, and draggable. The builder is a **consumer** of artifacts, not a producer; artifact generation stays in clipgen's existing CLI/batch flow, keeping both sides simple. Should support multiple insights in parallel and across sessions. Must be faster than making a slide deck.
- **Insight viewer** — standalone HTML deliverable (inlined CSS/JS, embedded data) following the same pattern as the timeline and gallery viewers.
- **Deliverable format** — a folder/bundle (HTML files + media assets) that can be shared via zip, shared drive, or eventually a hosted URL.

## Implementation Sequence

1. Design the `insights_manifest.json` schema
2. Build the insight builder interface (web view working against the manifest)
3. Build the insight detail viewer (cause → behavior → impact page with inline artifacts)
4. Build the insights index page (landing/entry point)
5. Wire "back out to full view" links between insight detail and existing timeline viewer

## Insights builder details

- The Insights Builder should be built in a new `insights_server.py`
- The Insights Builder should auto-save changes as the researcher drafts (but make this a toggle-able feature, so that the user can disable autosave and have manual saving too)
- The Builder should led researchers add artifacts via drag-and-drop as well as clicks.
- The Builder should be scoped to a single study and single manifest.
- The **Timeline Context** should be a timeline visualization, where the researcher can "paint" the affected areas of a timeline.
- The builder **loads the full artifact manifest on startup** and presents all artifacts in a searchable, browsable sidebar. Artifacts are available for drag-and-drop and click-to-add without any file picking or tool switching.
- **Future enhancement**: If researchers frequently need clips that don't yet exist while authoring insights, a "request a clip" action could shell out to clipgen's generation pipeline without leaving the builder. Deferred until we validate whether full manifest access is sufficient.

## Insights viewer details

- The Insights Viewer is a standalone HTML deliverable, and comes as a single HTML file
- The Insights Viewer should embed media inline (video players, images) for referenced artifacts.
- Feature: Thumbnails for videos and static GIF previews (thumbnail generation should be done by generating thumbnails and adding them to video files via ffmpeg, needs to be built)

## Assumptions to Monitor

- **Researcher adoption**: The builder must feel like a single workspace, not a second tool requiring setup or context-switching. If researchers feel they're "switching tools," adoption will suffer regardless of feature quality. The builder must be faster than the slide-deck alternative.
- **Artifact availability gap**: If researchers frequently realize they need clips that don't exist yet while authoring insights, the "consumer only" builder design may need a lightweight clip-request mechanism. Monitor how often this happens before building it.
- **Three-bucket generality**: Cause/behavior/impact fits usability and playtest research well. May need configurable buckets for broader research types (attitudinal, competitive, survey).
- **Sharing friction**: The deliverable is a folder of files, not a single file. Sharing needs to be easy (zip, shared drive, or hosted).
