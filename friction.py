# -*- coding: utf-8 -*-
"""Programmatic friction scorer for clipgen Transcripts.

A cheap, deterministic first pass over transcript segments that flags moments of
likely interest to a UX researcher. It matches compiled phrase patterns across
six categories and produces, per segment, a friction score plus the matched
markers, then aggregates session-level stats and selects the top candidates for
the LLM refinement stage (``thinking_agents._run_friction``).

This module is intentionally pure: no Ollama, no I/O, no ``config`` import. It is
the deterministic engine; the LLM/prompt/parse layer lives in
``thinking_agents.py`` per the project's module roles. Category keys here are
mirrored by ``config.FRICTION_CATEGORIES`` (display labels) and the equality of
the two key sets is asserted by ``tests/test_friction_scorer.py``.

Score formula (per segment):
    Σ(weight[cat] × match_count[cat]) / max(word_count, 1), clamped to [0, 1].

Segments are plain dicts (the same shape used elsewhere: ``id``, ``start``,
``end``, ``text``). Scored rows are plain dicts: ``id``, ``score``,
``categories`` (ordered, present-only), ``markers`` (matched substrings),
``counts`` (per-category match counts, non-zero only).

FUTURE (diarization): once facilitator and participant audio can be separated,
revisit the weights and candidate selection below — facilitator speech currently
inflates help_seeking/confusion signal in group and vocally-facilitated sessions.
"""

from __future__ import annotations

import re
from typing import Any

# Category weights. Frustration and confusion are stronger UX-research signals
# than baseline hesitation, so they are weighted higher. Tunable in code; not
# user-facing in v1.
CATEGORY_WEIGHTS: dict[str, float] = {
    "hesitation": 1.0,
    "confusion": 1.5,
    "frustration": 2.0,
    "surprise": 1.0,
    "self_correction": 1.5,
    "help_seeking": 1.5,
}

# Phrase patterns (not bare words). Word boundaries avoid overcounting common
# substrings. Compiled once at module load, case-insensitive.
_FRICTION_PATTERNS_RAW: dict[str, list[str]] = {
    "hesitation": [
        r"\bum+\b",
        r"\buh+\b",
        r"\berm+\b",
        r"\blet me (?:see|think|try)\b",
        r"\bi (?:think|guess|mean)\b",
        r"\b(?:kind of|sort of)\b",
        r"\b(\w+)\s+\1\b",  # word doubling — false-start signal
    ],
    "confusion": [
        r"\bwhere (?:is|are|do i|does)\b",
        r"\bhow (?:do i|does (?:this|it))\b",
        r"\bi (?:don't|can't) (?:see|find)\b",
        r"\bwait[,.]?\s*what\b",
        r"\bi'?m (?:not sure|confused)\b",
    ],
    "frustration": [
        r"\b(?:ugh|argh)\b",
        r"\bthis is (?:annoying|weird|broken|frustrating|stupid)\b",
        r"\bwhy (?:won't|isn't|can't|does|is)\b",
        r"\bcome on\b",
    ],
    "surprise": [
        r"\boh!?\b",
        r"\bhuh\b",
        r"\bwait what\b",
        r"\bno way\b",
        r"\bwhat the\b",
    ],
    "self_correction": [
        r"\bwait[,.]?\s*(?:actually|no)\b",
        r"\bnever ?mind\b",
        r"\bscratch that\b",
        r"\blet me start over\b",
        r"\bactually,?\s",
    ],
    "help_seeking": [
        r"\bcan you (?:help|tell|show)\b",
        r"\bhow should i\b",
        r"\bam i supposed to\b",
        r"\bwhat (?:should|do) i do\b",
    ],
}

# {category: [compiled regex, ...]} — order preserved from the raw table above.
FRICTION_PATTERNS: dict[str, list[re.Pattern[str]]] = {
    category: [re.compile(p, re.IGNORECASE) for p in patterns]
    for category, patterns in _FRICTION_PATTERNS_RAW.items()
}

# Ordered category keys (single source of order within this module).
CATEGORY_ORDER: tuple[str, ...] = tuple(FRICTION_PATTERNS.keys())

_MAX_MARKERS_PER_SEGMENT = 10  # cap stored markers to avoid manifest bloat


def _segment_score(
    text: str,
) -> tuple[float, list[str], list[str], dict[str, int]]:
    """Score a single segment's text.

    Returns ``(score, categories, markers, counts)`` where score is clamped to
    [0, 1], categories is the ordered list of categories with ≥1 match, markers
    is the deduped list of matched substrings, and counts maps each present
    category to its raw match count.
    """
    word_count = max(len(text.split()), 1)
    counts: dict[str, int] = {}
    markers: list[str] = []
    seen_markers: set[str] = set()
    weighted_total = 0.0

    for category in CATEGORY_ORDER:
        category_matches = 0
        for pattern in FRICTION_PATTERNS[category]:
            for match in pattern.finditer(text):
                category_matches += 1
                marker = match.group(0).strip().lower()
                if marker and marker not in seen_markers:
                    seen_markers.add(marker)
                    markers.append(marker)
        if category_matches:
            counts[category] = category_matches
            weighted_total += CATEGORY_WEIGHTS[category] * category_matches

    score = min(weighted_total / word_count, 1.0)
    categories = [c for c in CATEGORY_ORDER if c in counts]
    return round(score, 4), categories, markers[:_MAX_MARKERS_PER_SEGMENT], counts


def score_segments(segments: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Score every segment, returning per-segment friction rows.

    Each row: ``{"id", "score", "categories", "markers", "counts"}``. Segment
    order is preserved (the smoothing/heatmap path relies on positional order).
    """
    scored: list[dict[str, Any]] = []
    for idx, seg in enumerate(segments):
        text = (seg.get("text") or "").strip()
        score, categories, markers, counts = _segment_score(text)
        scored.append(
            {
                "id": seg.get("id") or str(idx),
                "score": score,
                "categories": categories,
                "markers": markers,
                "counts": counts,
            }
        )
    return scored


def select_candidates(
    scored: list[dict[str, Any]], n: int = 15
) -> list[dict[str, Any]]:
    """Return the top-*n* scored rows (score > 0) for the LLM stage.

    Sorted by score descending, breaking ties by total match count so denser
    segments win. Rows that scored zero (no category matches) are excluded.
    """
    candidates = [row for row in scored if row.get("score", 0) > 0]
    candidates.sort(
        key=lambda row: (row["score"], sum(row.get("counts", {}).values())),
        reverse=True,
    )
    return candidates[:n]


def compute_stats(
    scored: list[dict[str, Any]], duration_seconds: float
) -> dict[str, Any]:
    """Aggregate session-level friction stats.

    ``by_category`` includes every category (zeros shown) so the stats panel can
    render all six chips. ``markers_per_minute`` uses the transcript duration.
    """
    by_category: dict[str, int] = {c: 0 for c in CATEGORY_ORDER}
    total = 0
    for row in scored:
        for category, count in row.get("counts", {}).items():
            by_category[category] = by_category.get(category, 0) + count
            total += count
    minutes = duration_seconds / 60.0 if duration_seconds > 0 else 0.0
    markers_per_minute = round(total / minutes, 2) if minutes > 0 else 0.0
    return {
        "by_category": by_category,
        "markers_per_minute": markers_per_minute,
        "total_markers": total,
    }


def smooth_scores(scored: list[dict[str, Any]], window: int = 5) -> list[float]:
    """Rolling-mean of per-segment scores for the timeline heatmap.

    Centered window; at the edges the window shrinks to the available range.
    Returns one value per segment, in order. The raw per-segment score (not this
    smoothed series) still drives the inline segment background tint.
    """
    scores = [row.get("score", 0.0) for row in scored]
    if window < 1:
        window = 1
    half = window // 2
    smoothed: list[float] = []
    for i in range(len(scores)):
        lo = max(0, i - half)
        hi = min(len(scores), i + half + 1)
        window_vals = scores[lo:hi]
        smoothed.append(round(sum(window_vals) / len(window_vals), 4))
    return smoothed
