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

import re
import threading
from typing import Any, Callable, TypedDict

import config
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
    target_seconds: float, seg_starts: list[float], tolerance: float = 5.0
) -> int | None:
    """Find the segment index whose start time is closest to *target_seconds*.

    Returns None if the closest segment is more than *tolerance* seconds away.
    """
    best_idx: int | None = None
    best_dist = tolerance + 1
    for i, start in enumerate(seg_starts):
        dist = abs(start - target_seconds)
        if dist < best_dist:
            best_dist = dist
            best_idx = i
    return best_idx


def _parse_citation_response(
    response: str,
    segments: list[dict[str, Any]],
) -> dict[int, list[dict[str, Any]]]:
    """Parse model output into ``{claim_index: [ref_dicts]}``."""
    seg_starts = [seg.get("start", 0.0) for seg in segments]
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
            best_idx = _find_closest_segment(ts_seconds, seg_starts)
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
    cancel_event: threading.Event | None = None,
) -> list[dict[str, Any]] | None:
    """Find supporting transcript segments for each summary sentence.

    Sends the full transcript (truncated if very long) in a single model call
    to avoid multi-chunk latency. If *cancel_event* is set during the model
    call, the request is aborted and ``None`` is returned. Returns an ordered
    list of ``{"sentence", "refs": [...]}`` dicts, or ``None`` if inputs are
    empty or the run was cancelled.
    """
    sentences = _split_summary_sentences(summary)
    if not sentences or not segments:
        return None

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
]


def get_agent(key: str) -> Agent | None:
    """Look up an agent by key, or None if unknown."""
    for agent in AGENTS:
        if agent["key"] == key:
            return agent
    return None
