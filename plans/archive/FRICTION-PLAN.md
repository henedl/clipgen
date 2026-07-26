# Friction Detection — clipgen Transcripts

## Overview

This plan replaces the earlier "Transcript Cleanup" direction. Cleaning up filler words, hesitations, and self-corrections would actively destroy the signal UX researchers care about — those moments often *are* the friction. Instead, clipgen will surface them.

**Friction detection** is a hybrid pipeline that flags moments of likely interest in a per-participant transcript:

- A cheap, deterministic **programmatic scorer** runs over every segment, matching phrase patterns across six categories and producing a per-segment friction score plus session-level stats.
- A **local LLM agent** (Ollama, `qwen3.5:9b`) takes the top-N candidates plus the existing session summary as context and returns exactly five "moments" with category, rationale, and segment references.
- Results render as a **toggleable density heatmap** on the transcript timeline (mirroring the Screenspace per-task amplitude graph from PR #266) plus a new **Friction tab** in a tabbed summary panel.

The researcher still drives interpretation. The system *suggests* areas of interest; it does not auto-mark, auto-classify cross-participants, or modify transcript text.

---

## Goals (v1)

- Per-segment programmatic friction scoring with hardcoded phrase patterns across 6 categories
- LLM-refined top-5 "moments" per participant, with category and rationale
- Toggleable timeline heatmap (single weighted score, rolling-smoothed) with parity to the Screenspace amplitude graph control
- Inline segment background tint when the heatmap toggle is on
- Hover tooltip on hot segments showing programmatic markers and category breakdown
- New Friction tab in summary panel containing stats, top-5 moments, and a filter section
- Filter (score-threshold slider + category checkboxes) with a "Mark all matching" action that creates user Marks via the existing API
- Manual Run / Re-run via the participant pill dropdown (always available)
- Global config flag `OLLAMA_FRICTION_ENABLED`, off by default; when on, agent auto-runs after `summary` exists for any participant lacking a `friction` manifest field (see **Agent chain order** — not always a literal “third pass”)

## Non-goals (deferred to future)

- **Diarization.** Until we can separate facilitator from participant audio, group sessions (`G##`) and even participant sessions with vocal facilitators will produce noisier friction signal. Add a future-revisit note when diarization lands.
- **Cross-participant aggregate.** Surfacing patterns like "hesitation clusters around checkout across all participants" likely belongs in Studio's Convergence tool. Out of scope for v1.
- **Badging in Studio / Screenspace / Insights.** Friction is a Transcripts-UI feature first.
- **User-editable phrase patterns.** Hardcoded in Python for v1; structure leaves room for a per-study override hook later.
- **Per-participant normalization or sensitivity sliders.** Researchers mentally adjust for v1; the score-threshold filter doubles as the implicit calibration.
- **Pause-length scoring.** Whisper currently drops short fillers and hallucinates on silence; pause signal is unreliable until those issues are addressed separately.
- **Auto-update on segment edit.** Segment edits and corrections mark friction stale (UI indicator), but do not re-trigger the LLM pass. User explicitly re-runs.
- **Non-English transcripts.** Phrase patterns are English-centric in v1; multilingual pattern sets are future work.

---

## Architecture

### Categories and weights

Six fixed categories, hardcoded in `friction.py`:

| Category | Default weight |
| --- | --- |
| `hesitation` | 1.0 |
| `confusion` | 1.5 |
| `frustration` | 2.0 |
| `surprise` | 1.0 |
| `self_correction` | 1.5 |
| `help_seeking` | 1.5 |

Frustration and confusion are weighted higher because they are stronger UX-research signals than baseline hesitation. Weights are tunable in code; not user-facing in v1.

### Pipeline

```ascii
Whisper segments  (existing — transcripts.py)
        │
        ▼
[1] Programmatic scorer  (friction.py)
        │   - Match compiled phrase patterns per category
        │   - Score per segment: Σ(weight × match_count) / max(word_count, 1)
        │   - Compute smoothed series for heatmap (rolling window)
        │   - Aggregate stats (counts per category, markers/min)
        │   - Select top-15 candidates for LLM stage
        │
        ▼
[2] Friction agent  (thinking_agents.py — new entry in AGENTS)
        │   - depends_on: ["summary"] only (does not wait for citations)
        │   - Reads entry["segments"] + entry["summary"]; runs scorer in-thread
        │   - Inputs: session summary + top-15 candidates with ±1 segment of context
        │   - System prompt asks for EXACTLY 5 moments
        │   - Returns complete friction dict (programmatic + LLM fields)
        │   - ollama_client.generate(..., think=False) — same as citations
        │
        ▼
Manifest:  source_transcripts[participant].friction  (single commit when agent finishes)
```

### Agent chain order

`AgentOrchestrator` walks [`thinking_agents.AGENTS`](thinking_agents.py) **in list order**. Friction is appended **after** `citations`:

- If **`OLLAMA_CITATIONS_ENABLED` is True** and citations run: **summary → citations → friction** (third enabled pass).
- If citations are disabled or skipped: friction runs **immediately after summary** (second pass for that participant).

Friction **`depends_on: ["summary"]` only** — it does not require citations. The prompt uses the session summary text, not citation refs.

### Orchestrator invariants

These match existing summary/citations behavior in [`transcripts_server.py`](transcripts_server.py); friction must follow them:

1. **`next_eligible` skips when `entry[manifest_field]` is truthy.** Do not persist a partial `friction` object to the manifest before the agent thread completes, or auto-chain will never pick up friction for that participant.
2. **`run_agent` assigns `entry[field] = result` — no deep merge.** `_run_friction` must **return the full friction dict** (segments, stats, moments, computed_at, model, stale, …), not `{"moments": [...]}` alone.
3. **In-flight UX** uses `_orchestrator.is_generating(participant, "friction")` (and optional GET fields), not a truthy on-disk `friction` key.
4. **Manual regenerate** clears `entry.pop("friction", None)` then calls `run_agent("friction", participant, force=True)` — same pattern as summary regenerate clearing `summary` / `citations`.
5. **`model_config_key`** on the agent entry drives `_agent_model()` and delayed unload after Stop; required on the TypedDict.

### Programmatic scorer (`friction.py`, new module)

Functions (no class — single-purpose utilities, follows project preference). Segment types are **`list[dict[str, Any]]`** (same as transcript segments elsewhere); scored rows are plain dicts with `id`, `score`, `categories`, `markers`.

- `score_segments(segments: list[dict[str, Any]]) -> list[dict[str, Any]]` — runs phrase patterns, returns per-segment scores, matched markers, categories.
- `select_candidates(scored: list[dict[str, Any]], n: int = 15) -> list[dict[str, Any]]` — top-N for LLM input.
- `compute_stats(scored: list[dict[str, Any]], duration_seconds: float) -> dict` — `by_category` counts, `markers_per_minute`, total markers.
- `smooth_scores(scored: list[dict[str, Any]], window: int = 5) -> list[float]` — rolling-window mean for the heatmap render path. The raw per-segment score still drives the segment background tint.

Phrase patterns, not bare words. Examples (final list refined during implementation):

```python
FRICTION_PATTERNS = {
    "hesitation": [
        r"\bum+\b",
        r"\buh+\b",
        r"\berm+\b",
        r"\b(let me (see|think|try))\b",
        r"\b(i (think|guess|mean))\b",
        r"\b(kind of|sort of)\b",
        r"\b(\w+)\s+\1\b",  # word doubling — false-start signal
    ],
    "confusion": [
        r"\bwhere (is|are|do i|does)\b",
        r"\bhow (do i|does (this|it))\b",
        r"\bi (don't|can't) (see|find)\b",
        r"\bwait[,.]\s*what\b",
    ],
    "frustration": [
        r"\b(ugh|argh)\b",
        r"\bthis is (annoying|weird|broken|frustrating|stupid)\b",
        r"\bwhy (won't|isn't|can't|does)\b",
    ],
    "surprise": [
        r"\boh!?\b",
        r"\bhuh\b",
        r"\bwait what\b",
        r"\bno way\b",
        r"\bwhat the\b",
    ],
    "self_correction": [
        r"\bwait[,.]\s*(actually|no)\b",
        r"\bnever mind\b",
        r"\bscratch that\b",
        r"\blet me start over\b",
    ],
    "help_seeking": [
        r"\bcan you (help|tell|show)\b",
        r"\bhow should i\b",
        r"\bam i supposed to\b",
    ],
}
```

Patterns are compiled once at module load. Matching is case-insensitive. Word boundaries (`\b`) avoid the "like" overcounting problem the original plan identified.

### Friction agent (append to `AGENTS` in `thinking_agents.py`)

`Agent` is a **TypedDict** — append a dict literal (not a constructor call):

```python
(
    Agent(
        key="friction",
        enabled_config_key="OLLAMA_FRICTION_ENABLED",
        model_config_key="OLLAMA_FRICTION_MODEL",
        manifest_field="friction",
        depends_on=["summary"],
        thread_name_prefix="friction",
        run=_run_friction,
    ),
)
```

`_run_friction(entry, cancel_event)`:

1. Read `segments = entry.get("segments") or []` and `summary = entry.get("summary") or ""`. Segment IDs are already `{participant}:{idx}` (assigned in [`transcripts.py`](transcripts.py)).
2. Run `score_segments(segments)` → scored list; `compute_stats(...)`; `select_candidates(scored, n=config.FRICTION_CANDIDATE_LIMIT)`.
3. Resolve each candidate to its segment text plus ±1 context segment for readability.
4. Compose prompt: session summary, candidate segments (with IDs), task instructions.
5. Call `ollama_client.generate(..., model=config.OLLAMA_FRICTION_MODEL, think=False, cancel_event=cancel_event)`.
6. Parse JSON response (defensive: empty `moments` on failure, log warning). Extract first valid JSON array if wrapped in prose.
7. If `cancel_event` is set, return `None` (orchestrator skips persist).
8. **Return the complete friction dict** for orchestrator to assign to `entry["friction"]`:

```python
{
    "segments": [...],  # per-segment scores from step 2
    "stats": {...},
    "moments": [...],  # from LLM
    "computed_at": "...",
    "model": config.OLLAMA_FRICTION_MODEL,
    "stale": False,
}
```

System prompt sketch (refined during implementation):

```text
You are analyzing a UX research session transcript for moments of friction.

Friction categories:
- hesitation: filler words, false starts, hedges
- confusion: searching for UI, asking where things are
- frustration: irritation, complaints about the product
- surprise: unexpected reactions, "oh!", "wait what"
- self_correction: backtracking, "scratch that", "actually..."
- help_seeking: direct requests for help from the facilitator

You will receive:
- A summary of the session for context
- A list of candidate segments pre-filtered by automated heuristics, each with a segment ID

Return EXACTLY 5 moments where the participant most clearly shows friction.
Each moment may span 1-3 contiguous segment IDs.

Output JSON only, no prose, no <think> blocks:
[
  {"segment_ids": ["P01:7", "P01:8"], "category": "frustration",
   "rationale": "Participant repeatedly tried to find the save button",
   "score": 0.85}
]
```

### Manifest shape

`source_transcripts[participant].friction`:

```json
{
  "segments": [
    {"id": "P01:0", "score": 0.42, "categories": ["hesitation"],
     "markers": ["um", "uh"]}
  ],
  "moments": [
    {"segment_ids": ["P01:7", "P01:8"], "category": "frustration",
     "rationale": "Participant repeatedly tried to find save button", "score": 0.85}
  ],
  "stats": {
    "by_category": {"hesitation": 12, "confusion": 4, "frustration": 2,
                    "surprise": 1, "self_correction": 3, "help_seeking": 0},
    "markers_per_minute": 3.4,
    "total_markers": 22
  },
  "computed_at": "2026-05-02T14:32:00",
  "model": "qwen3.5:9b",
  "stale": false
}
```

The entire object is **persisted once** when `_run_friction` returns (orchestrator `entry["friction"] = result`). Nothing is written under `friction` while the agent thread is still running.

The `stale` flag is set to `true` when segment text, corrections, or transcript regeneration invalidate scores/moments; the UI shows a "Re-run friction" prompt. **Summary edits** also set `stale: true` (or `pop("friction")`) because the LLM prompt uses summary context — mirror summary-save invalidating citations. No auto-rerun.

### Trigger model

- `OLLAMA_FRICTION_ENABLED = False` in `config.py`. Mirrors existing `OLLAMA_SUMMARY_ENABLED` / `OLLAMA_CITATIONS_ENABLED`.
- When `True`: after transcription completes, `_on_task_complete` → `run_chain` eventually reaches friction when `summary` exists and `entry.get("friction")` is falsy (see **Agent chain order**).
- The pill-dropdown "Run friction" / "Re-run friction" action **bypasses the global flag** (`force=True`). It calls **`POST /api/friction/<participant>/regenerate`**, which:
  - Returns `200 {"ok": true, "generating": true}` if already in-flight (same as summary/citations — not 409)
  - Otherwise `entry.pop("friction", None)`, persist, then `_orchestrator.run_agent("friction", participant, force=True)`
- Cancellation: **`POST /api/friction/<participant>/stop`** → `_orchestrator.stop("friction", participant)` plus optional model unload (same as summary/citations).

---

## Backend changes

### `config.py`

- `OLLAMA_FRICTION_ENABLED: bool = False`
- `OLLAMA_FRICTION_MODEL: str = OLLAMA_SUMMARY_MODEL` (defaults to qwen3.5:9b; overridable for users who want a smaller/faster model)
- `FRICTION_CANDIDATE_LIMIT: int = 15`
- `FRICTION_MOMENT_LIMIT: int = 5`
- `FRICTION_HEATMAP_WINDOW: int = 5`
- Friction category labels for frontend mirroring (see Frontend config below)
- Add **`friction`** to `MARK_CATEGORIES` (single bucket for all friction-created marks): `{label: "Friction", color: ...}` — specificity lives in the mark **label**, not separate category keys per friction type

### `friction.py` (new module)

- Module-level `FRICTION_PATTERNS`, `CATEGORY_WEIGHTS`, compiled regex tables
- Functions: `score_segments`, `select_candidates`, `compute_stats`, `smooth_scores`
- `build_friction_result(segments, summary, *, cancel_event, model) -> dict | None` — internal helper used by `_run_friction`: programmatic scoring + LLM + assembled dict (not persisted here)

### `thinking_agents.py`

- Append `Agent` entry for `friction` to `AGENTS` list
- New `_run_friction(transcript_entry, cancel_event)` helper
- New `_FRICTION_PROMPT` constant

### `transcripts_server.py`

New endpoints, mirroring **summary/citations** conventions (not a separate DELETE cancel route):

- `GET /api/friction/<participant>` — return cached friction; if in-flight and absent, `{"ok": false, "generating": true}` (like summary); 404 when idle and absent
- `POST /api/friction/<participant>/regenerate` — manual trigger; `pop("friction")`, persist, `run_agent("friction", participant, force=True)`; `200 {"generating": true}` (or `generating: true` if already running)
- `POST /api/friction/<participant>/stop` — cancel in-flight run; schedule model unload via `_agent_model("friction")`

No orchestrator class edits beyond appending to `AGENTS` — `run_chain` / `run_agent` already chain dependents.

**Participant API:** extend responses that expose agent step state (e.g. `citations_generating` on summary GET, participant list payloads using `_step_state_agent`) with **`friction_generating`** / friction `done|idle|running` so the pill dropdown and Friction tab can poll without guessing.

Wire **invalidation** (not recompute) on segment edit, correction add, summary PUT, and transcript regeneration: set `friction["stale"] = true` or `pop("friction")` per the same rules as citations on summary edit.

### `data_export.py`

Add friction data to the analysis-ready JSON+CSV export so it flows through `--export`. Schema:

- One CSV row per moment: participant, segment_ids, category, rationale, score, computed_at
- One CSV row per segment-with-friction-score (for heatmap reproducibility)

---

## Frontend changes

### Prerequisite: tabbed summary panel rework

Today the summary area in `assets/web/transcripts.html` is hidden when no summary exists. Friction needs a sibling tab, so:

- Replace the current conditional summary block with an always-visible tabbed shell at the top of the transcripts page
- Tabs: `Summary` (existing content) and `Friction` (new)
- Each tab handles its own empty state with a CTA — `Run summary` / `Run friction analysis`
- Tab switching is purely client-side; tab data loads independently
- Persist last-viewed tab in `sessionStorage` per participant

This is listed as the first frontend implementation step. The summary tab must keep all current functionality (edit, regenerate, citations cross-link).

### Timeline heatmap

Reference implementation: the per-task amplitude graph toggle on the Screenspace timeline (PR #266 — `feat(screenspace): add per-task amplitude graph toggle on timeline`). The visual treatment, toggle UX, and event handling should mirror that pattern.

- Add a toggle button to the transcript timeline ruler (same control area as existing time-scrubbing)
- When on:
  - Render the smoothed friction score across the timeline as a horizontal density band (single color, alpha proportional to score)
  - Apply a background tint to every segment in the segment list, with intensity proportional to that segment's raw (unsmoothed) score
- Color: new `--color-friction` token in `tokens.css` (warm orange/red family — distinct from severity tokens). JS reads it via `getComputedStyle(document.documentElement).getPropertyValue("--color-friction")` per the project convention.
- Hover tooltip on hot segments: category badges (one per category present), markers list (max 5, "+N more" if longer), score formatted as `0.42`.

### Friction tab content

Three stacked sections inside the tab:

1. **Header strip**
   - Status text: `Computed 2 minutes ago · qwen3.5:9b`, or `Stale — segments edited since last run`
   - `Re-run friction` button (always available, bypasses global flag)
   - Cancel button while in-flight

2. **Stats panel**
   - Category chips with counts: `Hesitation 12 · Confusion 4 · Frustration 2 · Surprise 1 · Self-correction 3 · Help-seeking 0`
   - Single line: `3.4 markers/min · 22 total`
   - No charts in v1 — keep dense and scannable

3. **Filter + moments list**
   - Score threshold slider, range 0.0–1.0, default 0.5
   - Six category checkboxes (all checked by default)
   - `Mark all matching` button — creates Marks for every segment whose score ≥ threshold AND has at least one matching category
     - Mark **category** = `friction` (single `MARK_CATEGORIES` bucket for all friction marks)
     - Mark **label** = `Friction · <primary_category>` (e.g. `Friction · frustration`) — specificity in label only
     - Uses existing marks API (`POST /api/marks` with `segment_ids`, `category`, `label`)
   - Top-5 moments list (filtered by the same controls)
     - Each row: category badge, score, rationale text, timestamp range
     - Click → seek video to first segment's start AND scroll segment list to first segment_id

### Pill dropdown

Add `Run friction` / `Re-run friction` entry to the participant pill dropdown:

- Disabled with tooltip if `summary` doesn't exist (`depends_on`)
- In-flight indicator (spinner) on the entry while running
- Cancel option appears while in-flight
- Calls `POST /api/friction/<participant>/regenerate`

### Frontend config

Add to `utils.get_frontend_config()` so JS doesn't hardcode:

- `friction_categories`: ordered list of category keys + display labels
- `friction_color_token`: name of the CSS variable (`--color-friction`)
- `friction_moment_limit`: 5

Per the constants-mirroring rule, add a corresponding test in `tests/test_shared_constants.py` that asserts **`CLIPGEN_CONFIG` defaults in [`assets/web/utils.js`](assets/web/utils.js)** match Python after `clipgenApplyConfig()` overlays server config.

---

## Tests

### Unit

- `tests/test_friction_scorer.py`
  - Phrase patterns match expected strings per category
  - Word doubling pattern doesn't false-positive on common words
  - Score formula: matches × weight / word count, clamped sensibly
  - `smooth_scores` rolling window correctness at edges
  - `select_candidates` returns top-N by score
  - `compute_stats` aggregates category counts and markers/min correctly

- `tests/test_friction_agent.py`
  - Prompt is built from `entry["segments"]` + summary + candidates (not from pre-existing `friction` on entry)
  - JSON output parsing is defensive (handles wrapping prose, extra fields)
  - Cancellation event returns `None`; successful run returns full friction dict shape

### Integration

- `tests/test_transcripts_server.py`
  - `GET /api/friction/<p>` returns 404 then cached after run; `generating: true` while in-flight
  - `POST /api/friction/<p>/regenerate` triggers run; second call while in-flight returns `200 {"generating": true}` (summary convention)
  - `POST /api/friction/<p>/stop` cancels
  - Auto-chain skipped when `OLLAMA_FRICTION_ENABLED = False`
  - Auto-chain runs when flag is `True` and summary exists (with or without citations depending on `OLLAMA_CITATIONS_ENABLED`)
  - Manual `force=True` works regardless of flag
  - Segment edit sets `stale` or clears friction without rerunning LLM
  - Summary PUT invalidates friction like citations

### Smoke

- `tests/test_friction_smoke.py`
  - End-to-end with stub Whisper output → scored segments → moments via mocked `ollama_client.generate`

---

## Implementation order

1. **Prereq**: Rework summary panel into always-visible tabbed shell; existing Summary tab content moves into the first tab; second tab is a placeholder until step 6.
2. `friction.py` module (patterns, weights, scorer, smoothing, stats).
3. Friction agent entry in `thinking_agents.py` + Ollama prompt.
4. Endpoints in `transcripts_server.py` (GET, POST regenerate, POST stop) + participant `friction_generating` fields.
5. `config.py` flags + `utils.get_frontend_config()` mirroring + design token in `tokens.css`.
6. Friction tab content (stats, filter, top-5 moments) in `transcripts.{html,js,css}`.
7. Timeline toggle + heatmap render + segment tint (mirror PR #266 amplitude graph pattern).
8. Hover tooltip on hot segments.
9. Filter + Mark-all-matching action wired to existing marks API.
10. Pill dropdown Run/Re-run entry with progress + cancel.
11. `data_export.py` extension for friction CSV/JSON.
12. Tests (unit + integration + smoke + shared-constants assertion).
13. Update `agents/skills/transcribe/SKILL.md` to mention friction workflow.

---

## Open implementation-time questions (not blockers)

- **LLM latency on long sessions.** A 30-min session with 15 candidates fed to qwen3.5:9b is the latency reality check. If users find it sluggish, the path is a smaller model via `OLLAMA_FRICTION_MODEL` rather than truncating context.
- **JSON parse robustness.** Qwen sometimes wraps JSON in prose despite "JSON only" instructions. The parser should locate and extract the first valid JSON array, not require a clean response.
- **Smoothing window size.** Default 5 is a starting point. May need tuning once we see real heatmaps against real transcripts.
- **Stale detection granularity.** A single `stale` flag is coarse; future refinement might track which moments became stale, but v1 keeps it simple.
- **Group sessions (`G##`).** v1 runs friction with the same pipeline; UI may show a short “mixed voices — noisier signal” note. A config gate is optional if feedback warrants it.
- **Auto-run prerequisite.** Friction auto-chains when `summary` is non-empty on the entry; it does not require citations to succeed.

## Future work (revisit)

- **Diarization.** Friction analysis becomes meaningfully more accurate once we can separate participant from facilitator. Add a comment in `friction.py` near the scorer to revisit weights and candidate selection at that point.
- **Cross-participant aggregate.** Likely surfaces in Studio's Convergence tool; will need its own design pass.
- **Badging in Studio / Screenspace / Insights.** Once researcher value is established in Transcripts, propagate.
- **User-editable phrase patterns and weights.** A per-study override hook is the natural next step; keep `FRICTION_PATTERNS` shape stable so it can be loaded from a file.
- **Per-participant normalization / sensitivity sliders.** If raw scores prove hard to interpret across chatty vs taciturn participants.
- **Pause-length signal.** Defer until Whisper is reconfigured or replaced for word-level timestamps and silence-hallucination is addressed.
