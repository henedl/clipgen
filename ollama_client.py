# -*- coding: utf-8 -*-
"""Ollama local LLM client for clipgen.

Provides a reusable interface to the Ollama REST API for text generation.
All functions fail gracefully (return None / False) and never raise on
network errors. On connection refused, generate() auto-starts ``ollama serve``
and retries once.

Key functions:
  is_available()            - check Ollama server connectivity
  generate()                - send a prompt and get a text response
  summarize_transcript()    - summarize transcript segments into paragraph + bullets
  find_citations()          - find supporting transcript segments for each summary sentence
"""

import json
import re
import subprocess
import time
import urllib.error
import urllib.request
from typing import Any

import config
import utils

_HEALTH_TIMEOUT = 5  # seconds for connectivity check
_GENERATE_TIMEOUT = (
    300  # seconds — generous to allow cold model loading + long transcripts
)
_START_POLL_INTERVAL = 0.5  # seconds between health-check polls after starting server
_START_TIMEOUT = 10  # seconds to wait for server to become available after starting
_MAX_TRANSCRIPT_CHARS = 6000  # truncate long transcripts to fit model context window
_THINK_RE = re.compile(r"<think>[\s\S]*?</think>\s*", re.DOTALL)


def is_available() -> bool:
    """Check whether the Ollama server is reachable.

    Returns True if the server responds to GET /api/tags, False otherwise.
    Does not log on failure — callers decide how to handle unavailability.
    """
    try:
        req = urllib.request.Request(f"{config.OLLAMA_BASE_URL}/api/tags")
        with urllib.request.urlopen(req, timeout=_HEALTH_TIMEOUT):
            return True
    except (urllib.error.URLError, OSError, ValueError):
        return False


def _is_connection_refused(exc: Exception) -> bool:
    """Check whether an exception indicates connection refused (server not running)."""
    if isinstance(exc, urllib.error.URLError) and isinstance(
        exc.reason, ConnectionRefusedError
    ):
        return True
    if isinstance(exc, ConnectionRefusedError):
        return True
    return False


def _start_server() -> bool:
    """Attempt to start ``ollama serve`` and wait for it to become available.

    Returns True if the server is responding after startup, False otherwise.
    """
    utils.info_print("Starting Ollama server...")
    try:
        subprocess.Popen(
            ["ollama", "serve"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except FileNotFoundError:
        utils.warning_print("Ollama binary not found — is Ollama installed?")
        return False
    except OSError as exc:
        utils.warning_print(f"Failed to start Ollama: {exc}")
        return False

    deadline = time.monotonic() + _START_TIMEOUT
    while time.monotonic() < deadline:
        if is_available():
            utils.info_print("Ollama server started.")
            return True
        time.sleep(_START_POLL_INTERVAL)

    utils.warning_print("Ollama server did not start within timeout.")
    return False


def _do_generate(
    body: dict[str, Any],
) -> str | None:
    """Execute a single Ollama /api/generate request. Returns text or None."""
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        f"{config.OLLAMA_BASE_URL}/api/generate",
        data=data,
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=_GENERATE_TIMEOUT) as resp:
        result = json.loads(resp.read().decode("utf-8"))
        text = result.get("response", "")
        # Strip <think>...</think> blocks produced by reasoning models (e.g. qwen3.5)
        text = _THINK_RE.sub("", text).strip()
        if not text:
            utils.warning_print(
                f"Ollama returned empty response (model: {body.get('model')})"
            )
            return None
        return text


def generate(
    prompt: str,
    *,
    model: str | None = None,
    system: str | None = None,
    think: bool | None = None,
) -> str | None:
    """Send a prompt to Ollama and return the generated text.

    On connection refused, automatically attempts to start ``ollama serve``
    and retries the request once.

    Args:
        prompt: The user prompt to send.
        model: Ollama model name. Defaults to config.OLLAMA_SUMMARY_MODEL.
        system: Optional system prompt.
        think: Explicitly enable/disable thinking mode for reasoning models.
            When False, the model skips chain-of-thought and responds directly.

    Returns the generated text string, or None on any failure.
    """
    resolved_model = model or config.OLLAMA_SUMMARY_MODEL
    body: dict[str, Any] = {
        "model": resolved_model,
        "prompt": prompt,
        "stream": False,
    }
    if system:
        body["system"] = system
    if think is not None:
        body["think"] = think

    try:
        return _do_generate(body)
    except urllib.error.HTTPError as exc:
        utils.warning_print(f"Ollama generate failed (HTTP {exc.code}): {exc.reason}")
        return None
    except (urllib.error.URLError, OSError) as exc:
        if not _is_connection_refused(exc):
            utils.warning_print(f"Ollama generate failed (connection): {exc}")
            return None
        # Connection refused — try to start the server and retry once
        if not _start_server():
            return None
        try:
            return _do_generate(body)
        except (urllib.error.URLError, urllib.error.HTTPError, OSError) as retry_exc:
            utils.warning_print(f"Ollama generate failed after retry: {retry_exc}")
            return None
        except (json.JSONDecodeError, KeyError, ValueError) as retry_exc:
            utils.warning_print(f"Ollama generate failed (response): {retry_exc}")
            return None
    except (json.JSONDecodeError, KeyError, ValueError) as exc:
        utils.warning_print(f"Ollama generate failed (response): {exc}")
        return None


_SUMMARIZE_PROMPT = """\
Summarize this user research session transcript. Write a concise paragraph \
(2-4 sentences) describing what happened in the session. Then list the key \
topics or themes as bullet points (prefix each with "- ").

Transcript:
{text}"""

_MIN_TEXT_LENGTH = 50  # skip summarization for very short transcripts
_LARGE_MODEL_THRESHOLD = 8000  # chars — use the larger model above this


def summarize_transcript(
    segments: list[dict[str, Any]],
    *,
    model: str | None = None,
) -> str | None:
    """Summarize transcript segments into a paragraph + bullet points.

    Automatically selects between OLLAMA_SUMMARY_MODEL (short transcripts) and
    OLLAMA_SUMMARY_MODEL_LARGE (longer transcripts) unless an explicit model
    override is provided.

    Args:
        segments: List of transcript segment dicts with "text" keys.
        model: Ollama model override. Skips auto-selection when set.

    Returns the summary string, or None if segments are empty/too short
    or if generation fails.
    """
    text = " ".join(seg.get("text", "").strip() for seg in segments).strip()
    if len(text) < _MIN_TEXT_LENGTH:
        return None

    # Auto-select model based on transcript length
    if model is None:
        if len(text) > _LARGE_MODEL_THRESHOLD:
            model = config.OLLAMA_SUMMARY_MODEL_LARGE
        else:
            model = config.OLLAMA_SUMMARY_MODEL

    # Truncate to fit model context window — keep beginning and end for
    # a representative overview of the session
    if len(text) > _MAX_TRANSCRIPT_CHARS:
        half = _MAX_TRANSCRIPT_CHARS // 2
        text = text[:half] + "\n[...]\n" + text[-half:]

    utils.verbose_print(
        f"Summarizing transcript ({len(segments)} segments, "
        f"{len(text)} chars) with model {model}"
    )
    prompt = _SUMMARIZE_PROMPT.format(text=text)
    result = generate(prompt, model=model)
    if result:
        utils.verbose_print(f"Summary generated ({len(result)} chars)")
    return result


# ---- Citation linking (Pass 2) ----

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
        # Split paragraph text on sentence boundaries
        parts = re.split(r"(?<=[.!?])\s+", line)
        for part in parts:
            part = part.strip()
            if part:
                sentences.append(part)
    return sentences


def _format_segment_chunk(segments: list[dict[str, Any]], offset: int) -> str:
    """Format a chunk of segments as ``[M:SS] text`` lines for the prompt."""
    lines: list[str] = []
    for i, seg in enumerate(segments):
        start = seg.get("start", 0)
        ts = utils.seconds_to_timestamp(int(start))
        text = seg.get("text", "").strip()
        if text:
            lines.append(f"[{ts}] {text}")
    return "\n".join(lines)


def _parse_citation_response(
    response: str,
    segments: list[dict[str, Any]],
) -> dict[int, list[dict[str, Any]]]:
    """Parse model output into ``{claim_index: [ref_dicts]}``.

    Each ref dict has keys ``start``, ``end``, ``segment_index``.
    Timestamps are matched to the nearest segment by start time.
    """
    # Build a lookup of segment start times for matching
    seg_starts = [seg.get("start", 0.0) for seg in segments]

    result: dict[int, list[dict[str, Any]]] = {}
    for match in _CITATION_LINE_RE.finditer(response):
        claim_num = int(match.group(1))
        claim_idx = claim_num - 1  # 0-based
        body = match.group(2).strip()
        if body.upper() == "NONE":
            continue
        refs: list[dict[str, Any]] = []
        for ts_match in _TIMESTAMP_RE.finditer(body):
            ts_str = ts_match.group(1)
            ts_seconds = _timestamp_to_seconds(ts_str)
            if ts_seconds is None:
                continue
            # Find closest segment by start time
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


_MAX_CITATION_TRANSCRIPT_CHARS = 12000  # generous limit for 9B context window


def find_citations(
    summary: str,
    segments: list[dict[str, Any]],
) -> list[dict[str, Any]] | None:
    """Find supporting transcript segments for each summary sentence (Pass 2).

    Sends the full transcript (truncated if very long) in a single model call
    to avoid multi-chunk latency.

    Args:
        summary: The summary text (paragraph + bullets).
        segments: Full list of transcript segment dicts with start/end/text.

    Returns a list of citation dicts ``[{"sentence", "refs": [...]}]`` ordered
    by sentence position, or None on failure.
    """
    sentences = _split_summary_sentences(summary)
    if not sentences or not segments:
        return None

    model = config.OLLAMA_SUMMARY_MODEL_LARGE

    claims_text = "\n".join(f"{i + 1}. {s}" for i, s in enumerate(sentences))
    transcript_text = _format_segment_chunk(segments, 0)

    # Truncate long transcripts — keep beginning and end
    if len(transcript_text) > _MAX_CITATION_TRANSCRIPT_CHARS:
        half = _MAX_CITATION_TRANSCRIPT_CHARS // 2
        transcript_text = transcript_text[:half] + "\n[...]\n" + transcript_text[-half:]

    utils.verbose_print(
        f"Finding citations for {len(sentences)} claims "
        f"({len(transcript_text)} chars transcript) with model {model}"
    )

    prompt = _CITATION_PROMPT.format(claims=claims_text, transcript=transcript_text)
    response = generate(prompt, model=model, system=_CITATION_SYSTEM, think=False)

    # Parse refs from the single response
    parsed: dict[int, list[dict[str, Any]]] = {}
    if response:
        parsed = _parse_citation_response(response, segments)

    # Build final citations list, capping refs per claim
    citations: list[dict[str, Any]] = []
    for i, sentence in enumerate(sentences):
        refs = sorted(parsed.get(i, []), key=lambda r: r["start"])
        citations.append({"sentence": sentence, "refs": refs[:_MAX_REFS_PER_CLAIM]})

    total_refs = sum(len(c["refs"]) for c in citations)
    utils.verbose_print(
        f"Citations complete: {total_refs} total refs across {len(citations)} claims"
    )
    return citations
