# -*- coding: utf-8 -*-
"""Thinking-agent registry for clipgen.

A "thinking agent" is a small, self-contained unit of Ollama-powered reasoning
over a transcript: summary generation, citation linking, and (in the future)
things like keyword extraction, pain-point tagging, or follow-up-question
generation.

This module owns:
  - The ``Agent`` shape (prompt building, model selection, response parsing,
    dependency metadata, manifest field).
  - The ``AGENTS`` registry, ordered so dependencies come before dependents.
  - The two built-in agents: ``summary`` (Pass 1) and ``citations`` (Pass 2).

It does *not* own HTTP transport — that stays in ``ollama_client.generate()``.
It does *not* own orchestration — that stays in ``transcripts_server._run_agent_chain()``.

Adding a new agent is a matter of writing a ``run`` callable, defining an
``Agent`` dict, and appending it to ``AGENTS``. No edits elsewhere required.
"""

from __future__ import annotations

import bisect
import json
import re
import threading
from datetime import datetime, timezone
from typing import Any, Callable, TypedDict

import config
import friction
import ollama_client
import utils


# ---------------------------------------------------------------------------
# Agent shape
# ---------------------------------------------------------------------------


class Agent(TypedDict):
    """Describes one thinking-agent pass over a transcript.

    Keys:
      key:                Stable identifier (e.g. "summary", "citations").
      enabled_config_key: Name of a ``config`` attribute (bool) that gates
                          this agent. None of the attributes currently
                          differ per-agent, but keeping them separate means
                          future agents can have their own toggle.
      model_config_key:   Name of the ``config`` attribute (str) holding the
                          Ollama model name this agent runs against. Read by
                          the orchestrator for cancel-after-stop unload
                          scheduling so future agents that use a different
                          model unload the right one.
      manifest_field:     Where the result lands in
                          ``source_transcripts[participant][manifest_field]``.
      depends_on:         Other agent keys whose ``manifest_field`` must be
                          present on the transcript entry before this agent
                          can run. Also used to skip Pass 2 when Pass 1 failed.
      thread_name_prefix: Prefix for the daemon thread name (useful for
                          debugging).
      run:                Callable invoked inside the daemon thread. Receives
                          the transcript entry (the ``source_transcripts[pid]``
                          dict) and an optional ``threading.Event`` that, when
                          set, signals the agent to abort its in-flight model
                          call. Returns the value to store under
                          ``manifest_field``, or ``None`` to skip storage.
    """

    key: str
    enabled_config_key: str
    model_config_key: str
    manifest_field: str
    depends_on: list[str]
    thread_name_prefix: str
    run: Callable[[dict[str, Any], threading.Event | None], Any]


# ---------------------------------------------------------------------------
# Summary agent (Pass 1)
# ---------------------------------------------------------------------------

_SUMMARIZE_PROMPT = """\
Summarize this user research session transcript. Write a concise paragraph \
(2-4 sentences) describing what happened in the session. Then list the key \
topics or themes as bullet points (prefix each with "- ").

Transcript:
{text}"""

_MIN_TEXT_LENGTH = 50  # skip summarization for very short transcripts
_MAX_TRANSCRIPT_CHARS = 6000  # truncate long transcripts to fit context window


def _truncate_middle(text: str, limit: int) -> str:
    """Keep the first and last ``limit // 2`` characters of *text* if it
    exceeds *limit*, separated by a ``[...]`` marker."""
    if len(text) <= limit:
        return text
    half = limit // 2
    return text[:half] + "\n[...]\n" + text[-half:]


def summarize_transcript(
    segments: list[dict[str, Any]],
    *,
    model: str | None = None,
    cancel_event: threading.Event | None = None,
) -> str | None:
    """Summarize transcript segments into a paragraph + bullet points.

    Uses ``config.OLLAMA_SUMMARY_MODEL`` unless an explicit model override is
    provided. If *cancel_event* is set during the model call, the request is
    aborted and ``None`` is returned.
    """
    text = " ".join(seg.get("text", "").strip() for seg in segments).strip()
    if len(text) < _MIN_TEXT_LENGTH:
        return None

    if model is None:
        model = config.OLLAMA_SUMMARY_MODEL

    text = _truncate_middle(text, _MAX_TRANSCRIPT_CHARS)

    utils.verbose_print(
        f"Summarizing transcript ({len(segments)} segments, "
        f"{len(text)} chars) with model {model}"
    )
    prompt = _SUMMARIZE_PROMPT.format(text=text)
    result = ollama_client.generate(prompt, model=model, cancel_event=cancel_event)
    if result:
        utils.verbose_print(f"Summary generated ({len(result)} chars)")
    return result


def _run_summary(
    entry: dict[str, Any], cancel_event: threading.Event | None
) -> str | None:
    segments = entry.get("segments") or []
    return summarize_transcript(segments, cancel_event=cancel_event)


# ---------------------------------------------------------------------------
# Citation agent (Pass 2)
# ---------------------------------------------------------------------------

_CITATION_SYSTEM = (
    "You match transcript segments to summary claims. "
    "For each claim, select only the 1-3 most relevant and representative "
    "segments. Prefer segments that most clearly and directly support the claim. "
    "Use the exact format shown."
)

_CITATION_PROMPT = """\
Claims:
{claims}

Transcript:
{transcript}

For each claim, pick the 1-3 BEST supporting timestamps — the clearest, \
most direct evidence. Do not list every vaguely related segment.
Format your response exactly as:
1: 0:45, 1:02
2: NONE
Write NONE if no segments clearly support a claim."""

_MAX_REFS_PER_CLAIM = 4  # hard cap enforced during parsing
_MAX_CITATION_TRANSCRIPT_CHARS = 12000  # generous context-window limit

_CITATION_LINE_RE = re.compile(r"^(\d+)\s*:\s*(.+)$", re.MULTILINE)
_TIMESTAMP_RE = re.compile(r"(\d{1,2}:\d{2}(?::\d{2})?)")


def _split_summary_sentences(summary: str) -> list[str]:
    """Split summary into individual claims (paragraph sentences + bullets)."""
    sentences: list[str] = []
    for line in summary.split("\n"):
        line = line.strip()
        if not line:
            continue
        if line.startswith("- ") or line.startswith("* "):
            sentences.append(line[2:].strip())
            continue
        parts = re.split(r"(?<=[.!?])\s+", line)
        for part in parts:
            part = part.strip()
            if part:
                sentences.append(part)
    return sentences


def _format_segment_chunk(segments: list[dict[str, Any]], offset: int) -> str:
    """Format a chunk of segments as ``[M:SS] text`` lines for the prompt."""
    lines: list[str] = []
    for seg in segments:
        start = seg.get("start", 0)
        ts = utils.seconds_to_timestamp(int(start))
        text = seg.get("text", "").strip()
        if text:
            lines.append(f"[{ts}] {text}")
    return "\n".join(lines)


def _timestamp_to_seconds(ts: str) -> float | None:
    """Parse ``M:SS`` or ``H:MM:SS`` to seconds."""
    parts = ts.split(":")
    try:
        if len(parts) == 2:
            return int(parts[0]) * 60 + int(parts[1])
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
    except ValueError:
        pass
    return None


def _find_closest_segment(
    target_seconds: float,
    sorted_starts: list[float],
    sorted_indices: list[int],
    tolerance: float = 5.0,
) -> int | None:
    """Find the original segment index whose start time is closest to *target_seconds*.

    *sorted_starts* must be ascending; *sorted_indices* maps each sorted position
    back to the segment's original index. The closest element in a sorted array
    is always a neighbour of the bisect insertion point, so only two candidates
    are checked. Returns None if the closest segment is more than *tolerance*
    seconds away.
    """
    if not sorted_starts:
        return None
    pos = bisect.bisect_left(sorted_starts, target_seconds)
    best_pos: int | None = None
    best_dist = tolerance + 1
    for candidate in (pos - 1, pos):
        if 0 <= candidate < len(sorted_starts):
            dist = abs(sorted_starts[candidate] - target_seconds)
            if dist < best_dist:
                best_dist = dist
                best_pos = candidate
    if best_pos is None:
        return None
    return sorted_indices[best_pos]


def _parse_citation_response(
    response: str,
    segments: list[dict[str, Any]],
) -> dict[int, list[dict[str, Any]]]:
    """Parse model output into ``{claim_index: [ref_dicts]}``."""
    sorted_pairs = sorted(
        (seg.get("start", 0.0), idx) for idx, seg in enumerate(segments)
    )
    sorted_starts = [start for start, _ in sorted_pairs]
    sorted_indices = [idx for _, idx in sorted_pairs]
    result: dict[int, list[dict[str, Any]]] = {}
    for match in _CITATION_LINE_RE.finditer(response):
        claim_num = int(match.group(1))
        if claim_num < 1:
            continue
        claim_idx = claim_num - 1
        body = match.group(2).strip()
        if body.upper() == "NONE":
            continue
        refs: list[dict[str, Any]] = []
        for ts_match in _TIMESTAMP_RE.finditer(body):
            ts_seconds = _timestamp_to_seconds(ts_match.group(1))
            if ts_seconds is None:
                continue
            best_idx = _find_closest_segment(ts_seconds, sorted_starts, sorted_indices)
            if best_idx is not None:
                seg = segments[best_idx]
                refs.append(
                    {
                        "start": seg.get("start", 0.0),
                        "end": seg.get("end", 0.0),
                        "segment_index": best_idx,
                    }
                )
        if refs:
            result[claim_idx] = refs
    return result


def find_citations(
    summary: str,
    segments: list[dict[str, Any]],
    *,
    model: str | None = None,
    cancel_event: threading.Event | None = None,
) -> list[dict[str, Any]] | None:
    """Find supporting transcript segments for each summary sentence.

    Sends the full transcript (truncated if very long) in a single model call
    to avoid multi-chunk latency. Uses ``config.OLLAMA_SUMMARY_MODEL`` unless an
    explicit *model* override is provided. If *cancel_event* is set during the
    model call, the request is aborted and ``None`` is returned. Returns an
    ordered list of ``{"sentence", "refs": [...]}`` dicts, or ``None`` if inputs
    are empty or the run was cancelled.
    """
    sentences = _split_summary_sentences(summary)
    if not sentences or not segments:
        return None

    if model is None:
        model = config.OLLAMA_SUMMARY_MODEL

    claims_text = "\n".join(f"{i + 1}. {s}" for i, s in enumerate(sentences))
    transcript_text = _truncate_middle(
        _format_segment_chunk(segments, 0), _MAX_CITATION_TRANSCRIPT_CHARS
    )

    utils.verbose_print(
        f"Finding citations for {len(sentences)} claims "
        f"({len(transcript_text)} chars transcript) with model {model}"
    )

    prompt = _CITATION_PROMPT.format(claims=claims_text, transcript=transcript_text)
    response = ollama_client.generate(
        prompt,
        model=model,
        system=_CITATION_SYSTEM,
        think=False,
        cancel_event=cancel_event,
    )

    parsed: dict[int, list[dict[str, Any]]] = {}
    if response:
        parsed = _parse_citation_response(response, segments)

    citations: list[dict[str, Any]] = []
    for i, sentence in enumerate(sentences):
        refs = sorted(parsed.get(i, []), key=lambda r: r["start"])
        citations.append({"sentence": sentence, "refs": refs[:_MAX_REFS_PER_CLAIM]})

    total_refs = sum(len(c["refs"]) for c in citations)
    utils.verbose_print(
        f"Citations complete: {total_refs} total refs across {len(citations)} claims"
    )
    return citations


def _run_citations(
    entry: dict[str, Any], cancel_event: threading.Event | None
) -> list[dict[str, Any]] | None:
    summary = entry.get("summary") or ""
    segments = entry.get("segments") or []
    return find_citations(summary, segments, cancel_event=cancel_event)


# ---------------------------------------------------------------------------
# Friction agent (Pass 3)
# ---------------------------------------------------------------------------

_FRICTION_SYSTEM = (
    "You analyze UX research session transcripts for moments of friction — "
    "points where the participant struggled, hesitated, got confused, or showed "
    "frustration. You respond with a JSON array only."
)

_FRICTION_PROMPT = """\
Session summary:
{summary}

Candidate segments (pre-filtered by automated heuristics; each line is \
"[segment_id] (timestamp) text"):
{segments}

Friction categories: hesitation, confusion, frustration, surprise, \
self_correction, help_seeking.

Return EXACTLY {limit} moments where the participant most clearly shows friction.
Each moment may span 1-3 contiguous segment IDs taken from the candidate list above.

Output a JSON array only — no prose, no markdown fences, no <think> blocks:
[
  {{"segment_ids": ["P01:7", "P01:8"], "category": "frustration",
    "rationale": "Participant repeatedly tried to find the save button", "score": 0.85}}
]"""

_MAX_FRICTION_SUMMARY_CHARS = 2000  # cap summary context fed to the friction prompt


def _extract_json_array(text: str) -> list[Any]:
    """Best-effort extraction of the first JSON array from a model response.

    Qwen sometimes wraps JSON in prose, ``<think>`` blocks, or markdown fences
    despite "JSON only" instructions. Strips think-blocks, tries a fenced block,
    then scans for the first ``[`` that ``raw_decode`` can parse into a list.
    Returns ``[]`` if nothing parses.
    """
    if not text:
        return []
    cleaned = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()

    fence = re.search(r"```(?:json)?\s*(\[.*?\])\s*```", cleaned, re.DOTALL)
    if fence:
        try:
            data = json.loads(fence.group(1))
            if isinstance(data, list):
                return data
        except json.JSONDecodeError:
            pass

    decoder = json.JSONDecoder()
    start = cleaned.find("[")
    while start != -1:
        try:
            data, _ = decoder.raw_decode(cleaned[start:])
            if isinstance(data, list):
                return data
        except json.JSONDecodeError:
            pass
        start = cleaned.find("[", start + 1)
    return []


def _format_friction_candidates(
    segments: list[dict[str, Any]], candidates: list[dict[str, Any]]
) -> str:
    """Render candidates plus ±1 context segment as ``[id] (M:SS) text`` lines.

    Context segments are merged into a single ordered, de-duplicated block so
    adjacent candidates don't repeat lines.
    """
    id_to_idx = {seg.get("id"): i for i, seg in enumerate(segments)}
    include: set[int] = set()
    for cand in candidates:
        idx = id_to_idx.get(cand.get("id"))
        if idx is None:
            continue
        for j in (idx - 1, idx, idx + 1):
            if 0 <= j < len(segments):
                include.add(j)

    lines: list[str] = []
    for j in sorted(include):
        seg = segments[j]
        text = (seg.get("text") or "").strip()
        if not text:
            continue
        ts = utils.seconds_to_timestamp(int(seg.get("start", 0)))
        sid = seg.get("id") or str(j)
        lines.append(f"[{sid}] ({ts}) {text}")
    return "\n".join(lines)


def _parse_friction_response(response: str) -> list[dict[str, Any]]:
    """Parse the model's JSON array into normalized moment dicts.

    Defensive: tolerates wrapping prose and extra fields, coerces a single
    ``segment_ids`` string into a list, normalizes the category key, and clamps
    score to [0, 1]. Drops entries without usable segment IDs.
    """
    moments: list[dict[str, Any]] = []
    for item in _extract_json_array(response):
        if not isinstance(item, dict):
            continue
        seg_ids = item.get("segment_ids")
        if isinstance(seg_ids, str):
            seg_ids = [seg_ids]
        if not isinstance(seg_ids, list):
            continue
        seg_ids = [str(s) for s in seg_ids if s]
        if not seg_ids:
            continue
        category = (
            str(item.get("category", ""))
            .strip()
            .lower()
            .replace("-", "_")
            .replace(" ", "_")
        )
        try:
            score = float(item.get("score", 0.0))
        except (TypeError, ValueError):
            score = 0.0
        moments.append(
            {
                "segment_ids": seg_ids,
                "category": category,
                "rationale": str(item.get("rationale", "")).strip(),
                "score": round(max(0.0, min(1.0, score)), 4),
            }
        )
    return moments


def friction_model() -> str:
    """Resolve the Ollama model the friction agent should use.

    Blank ``OLLAMA_FRICTION_MODEL`` means "follow the summary model", so a single
    AI-model setting drives all three thinking agents. Set the override to pin
    friction to a different (e.g. smaller/faster) model.
    """
    return config.OLLAMA_FRICTION_MODEL or config.OLLAMA_SUMMARY_MODEL


def find_friction_moments(
    summary: str,
    segments: list[dict[str, Any]],
    candidates: list[dict[str, Any]],
    *,
    model: str | None = None,
    cancel_event: threading.Event | None = None,
) -> list[dict[str, Any]] | None:
    """Refine programmatic candidates into a short list of friction moments.

    Sends the session summary plus candidate segments (with context) to Ollama
    and parses the JSON response. Returns:
      - a list of moments (possibly empty) when the model responded — empty means
        it ran but found nothing,
      - ``[]`` when there are no candidates to send,
      - ``None`` when the model call itself failed (unavailable / wrong model /
        cancelled), so the caller can distinguish "no moments" from "didn't run".
    """
    if not candidates:
        return []
    if model is None:
        model = friction_model()

    block = _format_friction_candidates(segments, candidates)
    if not block:
        return []

    prompt = _FRICTION_PROMPT.format(
        summary=_truncate_middle(summary, _MAX_FRICTION_SUMMARY_CHARS),
        segments=block,
        limit=config.FRICTION_MOMENT_LIMIT,
    )
    utils.verbose_print(
        f"Detecting friction over {len(candidates)} candidate segments "
        f"with model {model}"
    )
    response = ollama_client.generate(
        prompt,
        model=model,
        system=_FRICTION_SYSTEM,
        think=False,
        cancel_event=cancel_event,
    )
    if not response:
        utils.warning_print(
            f"Friction moment detection failed (no response from model {model})"
        )
        return None

    # Keep only moments that cite at least one real segment, trimming any
    # hallucinated IDs. An unsourced moment can't be seeked to or quoted, so it
    # must never reach the manifest.
    valid_ids = {seg.get("id") for seg in segments if seg.get("id")}
    moments: list[dict[str, Any]] = []
    for moment in _parse_friction_response(response):
        kept = [sid for sid in moment["segment_ids"] if sid in valid_ids]
        if not kept:
            continue
        moment["segment_ids"] = kept
        moments.append(moment)
    if not moments:
        utils.warning_print("Friction analysis returned no sourced moments")
    return moments[: config.FRICTION_MOMENT_LIMIT]


def _segments_duration(segments: list[dict[str, Any]]) -> float:
    """Return the transcript duration (largest segment end time), or 0.0."""
    end = 0.0
    for seg in segments:
        try:
            end = max(end, float(seg.get("end", 0.0)))
        except (TypeError, ValueError):
            continue
    return end


def _run_friction(
    entry: dict[str, Any], cancel_event: threading.Event | None
) -> dict[str, Any] | None:
    """Assemble the complete friction result: programmatic scores + LLM moments.

    Returns the full dict stored under ``source_transcripts[pid].friction`` (the
    orchestrator assigns it wholesale — no partial writes), or ``None`` when
    there is no transcript/summary or the run was cancelled.
    """
    segments = entry.get("segments") or []
    summary = entry.get("summary") or ""
    if not segments or not summary:
        return None

    model = friction_model()
    scored = friction.score_segments(segments)
    stats = friction.compute_stats(scored, _segments_duration(segments))
    candidates = friction.select_candidates(scored, config.FRICTION_CANDIDATE_LIMIT)

    if cancel_event is not None and cancel_event.is_set():
        return None

    moments = find_friction_moments(
        summary, segments, candidates, model=model, cancel_event=cancel_event
    )

    if cancel_event is not None and cancel_event.is_set():
        return None

    # moments is None only when the LLM call itself failed; persist the
    # programmatic scores anyway (heatmap/stats/marks still work) but record that
    # moment detection did not succeed so the UI can say so instead of pretending.
    llm_ok = moments is not None

    return {
        "segments": scored,
        "moments": moments or [],
        "stats": stats,
        "computed_at": datetime.now(timezone.utc).isoformat(),
        "model": model,
        "llm_ok": llm_ok,
        "stale": False,
    }


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

# Order matters: dependents must appear after their dependencies so the
# orchestrator can iterate this list to produce a valid run order.
AGENTS: list[Agent] = [
    Agent(
        key="summary",
        enabled_config_key="OLLAMA_SUMMARY_ENABLED",
        model_config_key="OLLAMA_SUMMARY_MODEL",
        manifest_field="summary",
        depends_on=[],
        thread_name_prefix="summary",
        run=_run_summary,
    ),
    Agent(
        key="citations",
        enabled_config_key="OLLAMA_CITATIONS_ENABLED",
        model_config_key="OLLAMA_SUMMARY_MODEL",
        manifest_field="citations",
        depends_on=["summary"],
        thread_name_prefix="citations",
        run=_run_citations,
    ),
    Agent(
        key="friction",
        enabled_config_key="OLLAMA_FRICTION_ENABLED",
        model_config_key="OLLAMA_FRICTION_MODEL",
        manifest_field="friction",
        depends_on=["summary"],
        thread_name_prefix="friction",
        run=_run_friction,
    ),
]


def get_agent(key: str) -> Agent | None:
    """Look up an agent by key, or None if unknown."""
    for agent in AGENTS:
        if agent["key"] == key:
            return agent
    return None


def resolve_model(agent: Agent) -> str:
    """Return the Ollama model name *agent* runs against.

    Reads the config attribute named by ``model_config_key``; a blank value
    means "inherit the summary model" (friction's default).
    """
    model = getattr(config, agent["model_config_key"], None)
    if not model:
        model = config.OLLAMA_SUMMARY_MODEL
    return model
