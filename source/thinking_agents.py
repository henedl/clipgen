"""Thinking-agent registry for clipgen.

A "thinking agent" is a small, self-contained unit of Ollama-powered reasoning
over a transcript: summary generation, citation linking, friction detection,
and mini-report writing.

Owns the ``Agent`` shape (prompt building, model selection, response parsing,
dependency metadata, manifest field) and the ``AGENTS`` registry, ordered so
dependencies precede dependents: ``summary``, ``citations``, ``friction``, and
``report`` — a per-participant mini-report over the summary, sheet observations
and transcript marks, the latter two arriving via the ``configure()`` injection
seam. ``report`` is disabled by default, so it only runs when triggered manually.

Owns neither HTTP transport (``ollama_client.generate()``) nor orchestration
(``transcripts_server._run_agent_chain()``). Adding an agent means writing a
``run`` callable, defining an ``Agent`` dict, and appending it to ``AGENTS``.
"""

from __future__ import annotations

import bisect
import json
import math
import re
import threading
from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, TypedDict, cast

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
      on_upstream_change: How this agent's result reacts when an upstream
                          dependency is regenerated: ``"clear"`` drops the
                          field (the default), ``"stale"`` keeps it but flags
                          it for a prompted re-run (only the friction shape
                          carries a ``stale`` flag today). Read generically by
                          the regenerate route's dependent-invalidation.
      run:                Callable invoked inside the daemon thread. Receives
                          the transcript entry (the ``source_transcripts[pid]``
                          dict), an optional ``threading.Event`` that, when
                          set, signals the agent to abort its in-flight model
                          call, and an optional ``on_token`` callback for
                          streaming partial text (the summary and report
                          agents use it — structured agents ignore it).
                          Returns the value
                          to store under ``manifest_field``, or ``None`` to skip
                          storage. Typed loosely (``Callable[..., Any]``) so the
                          orchestrator's 3-arg call and 2-arg test calls both
                          typecheck.
    """

    key: str
    enabled_config_key: str
    model_config_key: str
    manifest_field: str
    depends_on: list[str]
    thread_name_prefix: str
    on_upstream_change: str
    run: Callable[..., Any]


# ---------------------------------------------------------------------------
# Shared response helpers
# ---------------------------------------------------------------------------

_THINK_RE = re.compile(r"<think>[\s\S]*?</think>\s*", re.DOTALL)


def _strip_think(text: str) -> str:
    """Remove <think>...</think> reasoning blocks a model may emit despite think=False.

    Reasoning models sometimes emit chain-of-thought scaffolding even when asked
    not to; stripping it is response parsing, so every agent cleans its own raw
    completion here rather than relying on the transport layer.
    """
    return _THINK_RE.sub("", text).strip()


# ---------------------------------------------------------------------------
# Summary agent (Pass 1)
# ---------------------------------------------------------------------------

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
    on_token: Callable[[str], None] | None = None,
) -> str | None:
    """Summarize transcript segments into a paragraph + bullet points.

    Uses ``config.OLLAMA_SUMMARY_MODEL`` unless an explicit model override is
    provided. If *cancel_event* is set during the model call, the request is
    aborted and ``None`` is returned. When *on_token* is provided it is invoked
    with each streamed piece so callers can surface the summary as it forms.

    Thinking is disabled (``think=False``, matching the citations and friction
    agents): a reasoning model would otherwise spend its first chunk of time in
    a silent think phase that emits no ``response`` text — dead air with nothing
    to stream, and pure added latency for a task that doesn't need reasoning.
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
    prompt = config.OLLAMA_SUMMARY_PROMPT.format(text=text)
    result = ollama_client.generate(
        prompt,
        model=model,
        think=False,
        cancel_event=cancel_event,
        on_token=on_token,
    )
    if result:
        result = _strip_think(result)
    if not result:
        return None
    utils.verbose_print(f"Summary generated ({len(result)} chars)")
    return result


def _run_summary(
    entry: dict[str, Any],
    cancel_event: threading.Event | None,
    on_token: Callable[[str], None] | None = None,
) -> str | None:
    segments = entry.get("segments") or []
    return summarize_transcript(segments, cancel_event=cancel_event, on_token=on_token)


# ---------------------------------------------------------------------------
# Citation agent (Pass 2)
# ---------------------------------------------------------------------------

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
        if line.startswith(("- ", "* ")):
            sentences.append(line[2:].strip())
            continue
        parts = re.split(r"(?<=[.!?])\s+", line)
        for part in parts:
            part = part.strip()
            if part:
                sentences.append(part)
    return sentences


def _format_segment_chunk(segments: list[dict[str, Any]]) -> str:
    """Format a chunk of segments as ``[M:SS] text`` lines for the prompt."""
    lines: list[str] = []
    for seg in segments:
        start = seg.get("start", 0)
        ts = utils.seconds_to_timestamp(int(start))
        text = seg.get("text", "").strip()
        if text:
            lines.append(f"[{ts}] {text}")
    return "\n".join(lines)


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
    best_dist = math.inf
    for candidate in (pos - 1, pos):
        if 0 <= candidate < len(sorted_starts):
            dist = abs(sorted_starts[candidate] - target_seconds)
            if dist < best_dist:
                best_dist = dist
                best_pos = candidate
    if best_pos is None or best_dist > tolerance:
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
            ts_seconds = utils.timestamp_to_seconds(ts_match.group(1))
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
    ordered list of ``{"sentence", "refs": [...]}`` dicts — one entry per claim,
    with an empty ``refs`` for claims the model found no support for — or
    ``None`` if inputs are empty, the run was cancelled, or the model call
    itself failed.
    """
    sentences = _split_summary_sentences(summary)
    if not sentences or not segments:
        return None

    if model is None:
        model = config.OLLAMA_SUMMARY_MODEL

    claims_text = "\n".join(f"{i + 1}. {s}" for i, s in enumerate(sentences))
    transcript_text = _truncate_middle(
        _format_segment_chunk(segments), _MAX_CITATION_TRANSCRIPT_CHARS
    )

    utils.verbose_print(
        f"Finding citations for {len(sentences)} claims "
        f"({len(transcript_text)} chars transcript) with model {model}"
    )

    prompt = config.OLLAMA_CITATIONS_PROMPT.format(
        claims=claims_text, transcript=transcript_text
    )
    response = ollama_client.generate(
        prompt,
        model=model,
        system=config.OLLAMA_CITATIONS_SYSTEM,
        think=False,
        cancel_event=cancel_event,
    )

    if not response:
        # The model call itself failed (Ollama down, model missing, aborted
        # mid-request). Returning a full list of empty refs here would persist a
        # success-shaped result the UI cannot tell apart from "no sources
        # exist", so report the failure and let the caller retry.
        utils.warning_print("Citations: no response from the model")
        return None

    stripped = _strip_think(response)
    if not _CITATION_LINE_RE.search(stripped):
        # A non-empty response with zero lines in the expected "N: ts" shape
        # (markdown bolding, "1." numbering, prose) is a parse failure, not
        # "no sources exist" — committing a full list of empty refs would look
        # identical to the latter in the UI. A response that *does* match the
        # format but answers NONE everywhere still commits below.
        utils.warning_print("Citations: response did not match the expected format")
        return None
    parsed = _parse_citation_response(stripped, segments)

    citations: list[dict[str, Any]] = []
    for i, sentence in enumerate(sentences):
        refs = sorted(parsed.get(i, []), key=lambda r: r["start"])
        # The model may cite two timestamps that resolve to one segment;
        # duplicates would eat the per-claim ref budget.
        seen_segments: set[int] = set()
        deduped: list[dict[str, Any]] = []
        for ref in refs:
            if ref["segment_index"] in seen_segments:
                continue
            seen_segments.add(ref["segment_index"])
            deduped.append(ref)
        citations.append({"sentence": sentence, "refs": deduped[:_MAX_REFS_PER_CLAIM]})

    total_refs = sum(len(c["refs"]) for c in citations)
    utils.verbose_print(
        f"Citations complete: {total_refs} total refs across {len(citations)} claims"
    )
    return citations


def _run_citations(
    entry: dict[str, Any],
    cancel_event: threading.Event | None,
    on_token: Callable[[str], None] | None = None,
) -> list[dict[str, Any]] | None:
    # on_token is accepted for uniform dispatch but ignored: citations are
    # parsed line-by-line from the *complete* response, so streaming raw tokens
    # to the UI would be meaningless.
    summary = entry.get("summary") or ""
    segments = entry.get("segments") or []
    return find_citations(summary, segments, cancel_event=cancel_event)


# ---------------------------------------------------------------------------
# Friction agent (Pass 3)
# ---------------------------------------------------------------------------

_MAX_FRICTION_SUMMARY_CHARS = 2000  # cap summary context fed to the friction prompt


def _extract_json_objects(text: str) -> list[dict[str, Any]]:
    """Decode every top-level ``{...}`` chunk in *text*, skipping broken ones.

    The salvage path for a model that emits an array with one malformed entry:
    scanning brace-balanced spans recovers the entries that *are* valid instead
    of losing the whole response to a single stray character. Braces inside
    strings are ignored so a rationale containing one can't unbalance the scan.
    """
    decoder = json.JSONDecoder()
    objects: list[dict[str, Any]] = []
    depth = 0
    start = -1
    in_string = False
    escaped = False
    for i, ch in enumerate(text):
        if in_string:
            if escaped:
                escaped = False
            elif ch == "\\":
                escaped = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
        elif ch == "{":
            if depth == 0:
                start = i
            depth += 1
        elif ch == "}" and depth > 0:
            depth -= 1
            if depth == 0 and start >= 0:
                try:
                    data, _ = decoder.raw_decode(text[start : i + 1])
                    if isinstance(data, dict):
                        objects.append(cast(dict[str, Any], data))
                except json.JSONDecodeError:
                    pass
                start = -1
    return objects


def _extract_json_array(text: str) -> list[Any]:
    """Best-effort extraction of a model response's JSON array of objects.

    Qwen sometimes wraps JSON in prose, ``<think>`` blocks, or markdown fences
    despite "JSON only" instructions. Strips think-blocks, tries a fenced block,
    then scans for a ``[`` that ``raw_decode`` parses into a list.

    Only arrays *of objects* count. The callers all want a list of moment dicts,
    and a malformed outer array otherwise makes the scan settle on the first
    inner array it can decode — a ``segment_ids`` value — which parses cleanly
    into a list of strings and silently yields zero moments. When no array
    survives, fall back to salvaging individual objects so one stray character
    costs one entry rather than the whole response.
    """
    if not text:
        return []
    cleaned = _strip_think(text)

    def _is_object_array(data: Any) -> bool:
        return isinstance(data, list) and all(isinstance(x, dict) for x in data)

    # Only a *non-empty* object array wins the scan — `[]` matches trivially,
    # so an empty array anywhere before the payload ('{"moments": [], ...}')
    # would otherwise end the search with zero results. A genuine empty array
    # is remembered and returned only after nothing better turns up.
    saw_empty_array = False

    fence = re.search(r"```(?:json)?\s*(\[.*?\])\s*```", cleaned, re.DOTALL)
    if fence:
        try:
            data = json.loads(fence.group(1))
            if _is_object_array(data):
                if data:
                    return data
                saw_empty_array = True
        except json.JSONDecodeError:
            pass

    decoder = json.JSONDecoder()
    start = cleaned.find("[")
    while start != -1:
        try:
            data, _ = decoder.raw_decode(cleaned[start:])
            if _is_object_array(data):
                if data:
                    return data
                saw_empty_array = True
        except json.JSONDecodeError:
            pass
        start = cleaned.find("[", start + 1)

    if saw_empty_array:
        # The model explicitly answered with an empty array — that IS the
        # result; salvaging stray objects from surrounding prose would
        # fabricate entries the model never returned.
        return []
    return list(_extract_json_objects(cleaned))


def _format_friction_candidates(
    segments: list[dict[str, Any]], candidates: list[dict[str, Any]]
) -> str:
    """Render candidates plus ±1 context segment as ``[id] (M:SS) text`` lines.

    Context segments are merged into a single ordered, de-duplicated block so
    adjacent candidates don't repeat lines.
    """
    # Mirror friction.score_segments' id scheme exactly: it synthesizes
    # str(index) for id-less segments (the Workflows path, where segments are
    # never manifest-saved and so never get ids), and a lookup keyed on the
    # raw id would drop every candidate there.
    id_to_idx = {(seg.get("id") or str(i)): i for i, seg in enumerate(segments)}
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
      - ``None`` when no model call was made or it failed (unavailable / wrong
        model / cancelled / candidates that couldn't be rendered), so the
        caller can distinguish "no moments" from "didn't run".
    """
    if not candidates:
        return []
    if model is None:
        model = friction_model()

    block = _format_friction_candidates(segments, candidates)
    if not block:
        # Candidates exist but none rendered (ids drifted after an edit, or
        # every context segment is empty) — no model call was made, so this is
        # "didn't run", not "ran and found nothing".
        utils.warning_print("Friction: no candidate segments could be rendered")
        return None

    prompt = config.OLLAMA_FRICTION_PROMPT.format(
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
        system=config.OLLAMA_FRICTION_SYSTEM,
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
    valid_ids = {(seg.get("id") or str(i)) for i, seg in enumerate(segments)}
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
    entry: dict[str, Any],
    cancel_event: threading.Event | None,
    on_token: Callable[[str], None] | None = None,
) -> dict[str, Any] | None:
    """Assemble the complete friction result: programmatic scores + LLM moments.

    Returns the full dict stored under ``source_transcripts[pid].friction`` (the
    orchestrator assigns it wholesale — no partial writes), or ``None`` when
    there is no transcript/summary or the run was cancelled. *on_token* is
    accepted for uniform dispatch but ignored — friction parses a JSON array
    from the complete response.
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
        "computed_at": datetime.now(UTC).isoformat(),
        "model": model,
        "llm_ok": llm_ok,
        "stale": False,
    }


# ---------------------------------------------------------------------------
# Report agent (Pass 4) — per-participant mini-report
# ---------------------------------------------------------------------------

_MAX_REPORT_SUMMARY_CHARS = 2000
_MAX_REPORT_OBSERVATIONS_CHARS = 4000
_MAX_REPORT_MARKS_CHARS = 6000
_REPORT_EMPTY_SECTION = "(none recorded)"

# The report agent needs data that lives outside the transcript entry: sheet
# observation rows (studio-blueprint process state in server.py) and resolved
# transcript marks (transcripts_server manifest state). Both reach this module
# by injection because importing either owner from here would be an import
# cycle (transcripts_server imports this module). Either getter may be None (CLI runs, tests); the report then
# degrades to the sources that exist.
_observation_rows_getter: Any = None
_participant_marks_getter: Any = None


def configure(
    observation_rows_getter: Any = None,
    participant_marks_getter: Any = None,
) -> None:
    """Wire process-state providers into the report agent (called by server.py).

    *observation_rows_getter*: ``() -> list[dict]`` of per-(sheet row ×
    participant) records carrying at least ``participant``, ``text``,
    ``category``, and ``severity`` (server.py's ``_sheet_observation_rows``).
    *participant_marks_getter*: ``(pid) -> list[dict]`` of resolved marks
    carrying ``start``, ``category``, ``label``, and ``text``
    (``transcripts_server.marks_for_participant`` — it takes the transcripts
    manifest lock internally, which is safe here because the orchestrator
    invokes ``run`` outside that lock).
    """
    global _observation_rows_getter, _participant_marks_getter
    _observation_rows_getter = observation_rows_getter
    _participant_marks_getter = participant_marks_getter


def report_source_lines(participant: str) -> tuple[list[str], list[str]]:
    """Formatted observation + mark lines for *participant* via the configured getters.

    The seam ``configure()`` wires in server.py; unwired (CLI, tests, a
    sheet-less launch) both lists come back empty and a report simply covers
    the summary alone. Shared by ``_run_report`` and the Workflows report node.
    """
    obs_rows = _observation_rows_getter() if _observation_rows_getter else []
    marks = (
        _participant_marks_getter(participant)
        if (_participant_marks_getter and participant)
        else []
    )
    return (
        _format_report_observations(obs_rows, participant),
        _format_report_marks(marks),
    )


def report_model() -> str:
    """Resolve the Ollama model the report agent should use.

    Blank ``OLLAMA_REPORT_MODEL`` means "follow the summary model", matching
    the friction agent's inherit behavior.
    """
    return config.OLLAMA_REPORT_MODEL or config.OLLAMA_SUMMARY_MODEL


def _format_report_observations(
    rows: list[dict[str, Any]], participant: str
) -> list[str]:
    """Format *participant*'s sheet observations as ``- [M:SS] text (category, severity)`` lines.

    The leading timestamp (the cell's first parsed start time) is what lets the
    model cite video times the Reports tab can then link to generated clips;
    text-only cells get no bracket. Deduplicates on observation text: the
    getter emits one record per (row × cell), so a participant listed twice on
    a row would otherwise repeat the same note.
    """
    lines: list[str] = []
    seen: set[str] = set()
    for rec in rows:
        if rec.get("participant") != participant:
            continue
        text = str(rec.get("text") or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        category = str(rec.get("category") or "").strip() or "uncategorized"
        severity = str(rec.get("severity") or "").strip()
        qualifier = f"({category}, {severity})" if severity else f"({category})"
        seconds = rec.get("seconds") or []
        stamp = f"[{utils.seconds_to_timestamp(int(seconds[0]))}] " if seconds else ""
        lines.append(f"- {stamp}{text} {qualifier}")
    return lines


def _format_report_marks(marks: list[dict[str, Any]]) -> list[str]:
    """Format resolved marks as ``[M:SS] (Label) text`` lines.

    Prefers the mark's own label (e.g. "Friction · frustration") over the
    category bucket's display label; skips marks with no resolved text.
    """
    lines: list[str] = []
    for mark in marks:
        text = str(mark.get("text") or "").strip()
        if not text:
            continue
        try:
            start = int(float(mark.get("start", 0)))
        except (TypeError, ValueError):
            start = 0
        category = str(mark.get("category") or "")
        label = (
            str(mark.get("label") or "").strip()
            or config.MARK_CATEGORIES.get(category, {}).get("label")
            or "Mark"
        )
        lines.append(f"[{utils.seconds_to_timestamp(start)}] ({label}) {text}")
    return lines


def build_report(
    summary: str,
    observations_text: str,
    marks_text: str,
    *,
    participant: str,
    model: str | None = None,
    cancel_event: threading.Event | None = None,
    on_token: Callable[[str], None] | None = None,
) -> str | None:
    """Generate the mini-report text from the three formatted sources.

    Empty sources are sent as a "(none recorded)" placeholder rather than
    omitted, so the prompt's structure is stable and the model is told
    explicitly that a source has no data. Streams tokens through *on_token*
    (surfaced as ``partial`` by the generic agent GET). Returns the cleaned
    report text, or ``None`` when the model call failed.
    """
    if model is None:
        model = report_model()
    prompt = config.OLLAMA_REPORT_PROMPT.format(
        participant=participant or "unknown",
        summary=_truncate_middle(summary, _MAX_REPORT_SUMMARY_CHARS)
        or _REPORT_EMPTY_SECTION,
        observations=_truncate_middle(observations_text, _MAX_REPORT_OBSERVATIONS_CHARS)
        or _REPORT_EMPTY_SECTION,
        bookmarks=_truncate_middle(marks_text, _MAX_REPORT_MARKS_CHARS)
        or _REPORT_EMPTY_SECTION,
    )
    utils.verbose_print(f"Generating mini-report for {participant} with model {model}")
    response = ollama_client.generate(
        prompt,
        model=model,
        system=config.OLLAMA_REPORT_SYSTEM,
        think=False,
        cancel_event=cancel_event,
        on_token=on_token,
    )
    if not response:
        utils.warning_print(
            f"Mini-report generation failed (no response from model {model})"
        )
        return None
    return _strip_think(response)


def _run_report(
    entry: dict[str, Any],
    cancel_event: threading.Event | None,
    on_token: Callable[[str], None] | None = None,
) -> dict[str, Any] | None:
    """Assemble the mini-report from summary + sheet observations + marks.

    Returns the dict stored under ``source_transcripts[pid].report``, or
    ``None`` when there is no summary yet or the model call failed/was
    cancelled. ``entry["participant"]`` is injected by the orchestrator's
    snapshot (manifest entries do not carry their own id).
    """
    summary = entry.get("summary") or ""
    if not summary:
        return None
    participant = str(entry.get("participant") or "")

    observation_lines, mark_lines = report_source_lines(participant)

    if cancel_event is not None and cancel_event.is_set():
        return None

    model = report_model()
    text = build_report(
        summary,
        "\n".join(observation_lines),
        "\n".join(mark_lines),
        participant=participant,
        model=model,
        cancel_event=cancel_event,
        on_token=on_token,
    )
    if not text:
        return None

    return {
        "text": text,
        "model": model,
        "generated_at": datetime.now(UTC).isoformat(),
        "sources": {
            "observations": len(observation_lines),
            "bookmarks": len(mark_lines),
        },
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
        on_upstream_change="clear",
        run=_run_summary,
    ),
    Agent(
        key="citations",
        enabled_config_key="OLLAMA_CITATIONS_ENABLED",
        model_config_key="OLLAMA_SUMMARY_MODEL",
        manifest_field="citations",
        depends_on=["summary"],
        thread_name_prefix="citations",
        on_upstream_change="clear",
        run=_run_citations,
    ),
    Agent(
        key="friction",
        enabled_config_key="OLLAMA_FRICTION_ENABLED",
        model_config_key="OLLAMA_FRICTION_MODEL",
        manifest_field="friction",
        depends_on=["summary"],
        thread_name_prefix="friction",
        on_upstream_change="stale",
        run=_run_friction,
    ),
    Agent(
        key="report",
        enabled_config_key="OLLAMA_REPORT_ENABLED",
        model_config_key="OLLAMA_REPORT_MODEL",
        manifest_field="report",
        depends_on=["summary"],
        thread_name_prefix="report",
        on_upstream_change="clear",
        run=_run_report,
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
