# -*- coding: utf-8 -*-
"""Transcription support for clipgen using faster-whisper."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, List, Optional, TypedDict

import config
import utils

# ---------------------------------------------------------------------------
# Data types
# ---------------------------------------------------------------------------


class TranscriptSegment(TypedDict):
    start: float  # seconds
    end: float  # seconds
    text: str


class TranscriptResult(TypedDict):
    segments: list[TranscriptSegment]
    language: str
    source_file: str
    model: str


# ---------------------------------------------------------------------------
# Module-level model cache
# ---------------------------------------------------------------------------

_cached_model: Any = None
_cached_model_name: Optional[str] = None


# ---------------------------------------------------------------------------
# Core transcription
# ---------------------------------------------------------------------------


def _load_model() -> Any:
    """Lazy-load the WhisperModel, caching it for reuse."""
    global _cached_model, _cached_model_name  # noqa: PLW0603

    model_name = config.TRANSCRIBE_MODEL
    if _cached_model is not None and _cached_model_name == model_name:
        return _cached_model

    try:
        from faster_whisper import WhisperModel
    except ImportError:
        utils.error_print(
            "faster-whisper is not installed.",
            details=["Install it with: uv add faster-whisper"],
        )
        return None

    def _do_load() -> Any:
        return WhisperModel(model_name, compute_type=config.TRANSCRIBE_COMPUTE_TYPE)

    utils.info_print(f"Loading transcription model '{model_name}'...")
    _cached_model = utils.run_with_spinner("Loading transcription model...", _do_load)
    _cached_model_name = model_name
    return _cached_model


def transcribe_video(
    video_path: str,
    *,
    language: Optional[str] = None,
    initial_prompt: Optional[str] = None,
    context_keywords: Optional[List[str]] = None,
) -> Optional[TranscriptResult]:
    """Transcribe a video file and return timestamped segments.

    Args:
        video_path: Path to the video file.
        language: Language code (e.g. "en"). None = auto-detect.
        initial_prompt: Override the default initial prompt.
        context_keywords: Extra keywords to append to the prompt.

    Returns:
        TranscriptResult with segments, or None on failure.
    """
    if config.DEBUGGING:
        return TranscriptResult(
            segments=[],
            language=language or "en",
            source_file=str(video_path),
            model=config.TRANSCRIBE_MODEL,
        )

    model = _load_model()
    if model is None:
        return None

    prompt = initial_prompt or config.TRANSCRIBE_INITIAL_PROMPT
    if context_keywords:
        prompt = f"{prompt} Key terms: {', '.join(context_keywords)}."

    lang = language or config.TRANSCRIBE_LANGUAGE

    try:
        segments_iter, info = model.transcribe(
            str(video_path),
            beam_size=config.TRANSCRIBE_BEAM_SIZE,
            language=lang,
            initial_prompt=prompt,
        )
        segments: list[TranscriptSegment] = [
            TranscriptSegment(start=seg.start, end=seg.end, text=seg.text.strip())
            for seg in segments_iter
            if seg.text.strip()
        ]
        detected_lang = (
            info.language if hasattr(info, "language") else (lang or "unknown")
        )
        return TranscriptResult(
            segments=segments,
            language=detected_lang,
            source_file=str(video_path),
            model=config.TRANSCRIBE_MODEL,
        )
    except Exception as exc:
        utils.warning_print(f"Transcription failed for {Path(video_path).name}: {exc}")
        return None


# ---------------------------------------------------------------------------
# Segment filtering
# ---------------------------------------------------------------------------


def filter_segments(
    result: TranscriptResult,
    start_sec: float,
    end_sec: float,
    *,
    offset_to_zero: bool = False,
) -> TranscriptResult:
    """Return a copy of *result* with only segments overlapping [start_sec, end_sec].

    When *offset_to_zero* is True, segment times are shifted so the clip
    starts at 0:00 (useful for per-clip transcript files).
    """
    filtered: list[TranscriptSegment] = [
        seg
        for seg in result["segments"]
        if seg["end"] > start_sec and seg["start"] < end_sec
    ]
    if offset_to_zero and filtered:
        filtered = [
            TranscriptSegment(
                start=max(0.0, seg["start"] - start_sec),
                end=seg["end"] - start_sec,
                text=seg["text"],
            )
            for seg in filtered
        ]
    return TranscriptResult(
        segments=filtered,
        language=result["language"],
        source_file=result["source_file"],
        model=result["model"],
    )


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------


def _fmt_display(seconds: float) -> str:
    """Format seconds as M:SS or H:MM:SS for Markdown display."""
    total = int(seconds)
    h, remainder = divmod(total, 3600)
    m, s = divmod(remainder, 60)
    if h > 0:
        return f"{h}:{m:02d}:{s:02d}"
    return f"{m}:{s:02d}"


def _fmt_srt(seconds: float) -> str:
    """Format seconds as HH:MM:SS,mmm for SRT."""
    total = int(seconds)
    ms = int((seconds - total) * 1000)
    h, remainder = divmod(total, 3600)
    m, s = divmod(remainder, 60)
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


def _fmt_vtt(seconds: float) -> str:
    """Format seconds as MM:SS.mmm for VTT."""
    total = int(seconds)
    ms = int((seconds - total) * 1000)
    h, remainder = divmod(total, 3600)
    m, s = divmod(remainder, 60)
    if h > 0:
        return f"{h:02d}:{m:02d}:{s:02d}.{ms:03d}"
    return f"{m:02d}:{s:02d}.{ms:03d}"


# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------


def _format_markdown(result: TranscriptResult) -> str:
    source_name = Path(result["source_file"]).name
    lines = [
        f"# Transcript: {source_name}",
        "",
        f"- **Source:** {source_name}",
        f"- **Model:** {result['model']}",
        f"- **Language:** {result['language']}",
        f"- **Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "---",
        "",
    ]
    for seg in result["segments"]:
        lines.append(f"**[{_fmt_display(seg['start'])} - {_fmt_display(seg['end'])}]**")
        lines.append(seg["text"])
        lines.append("")
    return "\n".join(lines)


def _format_srt(result: TranscriptResult) -> str:
    blocks: list[str] = []
    for i, seg in enumerate(result["segments"], start=1):
        blocks.append(
            f"{i}\n{_fmt_srt(seg['start'])} --> {_fmt_srt(seg['end'])}\n{seg['text']}"
        )
    return "\n\n".join(blocks) + "\n" if blocks else ""


def _format_vtt(result: TranscriptResult) -> str:
    lines = ["WEBVTT", ""]
    for seg in result["segments"]:
        lines.append(f"{_fmt_vtt(seg['start'])} --> {_fmt_vtt(seg['end'])}")
        lines.append(seg["text"])
        lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Write / read transcript files
# ---------------------------------------------------------------------------

_FORMAT_EXT = {"md": ".md", "srt": ".srt", "vtt": ".vtt"}


def get_transcript_extension(fmt: Optional[str] = None) -> str:
    """Return the file extension for the given transcript format."""
    return _FORMAT_EXT.get(fmt or config.TRANSCRIBE_FORMAT, ".md")


def write_transcript(
    result: TranscriptResult,
    output_path: str,
    *,
    fmt: Optional[str] = None,
) -> bool:
    """Write a formatted transcript file to *output_path*.

    Returns True on success, False on failure.
    """
    fmt = fmt or config.TRANSCRIBE_FORMAT
    formatters = {"md": _format_markdown, "srt": _format_srt, "vtt": _format_vtt}
    formatter = formatters.get(fmt, _format_markdown)

    try:
        text = formatter(result)
        Path(output_path).write_text(text, encoding="utf-8")
        utils.verbose_print(f"  Transcript written: {Path(output_path).name}")
        return True
    except OSError as exc:
        utils.error_print(f"Failed to write transcript: {exc}")
        return False


def read_transcript(filepath: str) -> Optional[TranscriptResult]:
    """Parse a transcript file back into a TranscriptResult.

    Detects format from file extension (.md, .srt, .vtt).
    Returns None if the file cannot be read or parsed.
    """
    path = Path(filepath)
    if not path.is_file():
        return None

    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return None

    ext = path.suffix.lower()
    if ext == ".srt":
        return _parse_srt(text, filepath)
    if ext == ".vtt":
        return _parse_vtt(text, filepath)
    return _parse_markdown(text, filepath)


# ---------------------------------------------------------------------------
# Parsers for read-back
# ---------------------------------------------------------------------------

_SRT_BLOCK = re.compile(
    r"(\d+)\s*\n"
    r"(\d{2}:\d{2}:\d{2},\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2},\d{3})\s*\n"
    r"(.+?)(?=\n\n|\n\d+\s*\n|\Z)",
    re.DOTALL,
)

_VTT_CUE = re.compile(
    r"(\d{2}:\d{2}[:\.]?\d{0,2}\.?\d{0,3})\s*-->\s*(\d{2}:\d{2}[:\.]?\d{0,2}\.?\d{0,3})\s*\n"
    r"(.+?)(?=\n\n|\Z)",
    re.DOTALL,
)

_MD_SEGMENT = re.compile(
    r"\*\*\[(.+?)\s*-\s*(.+?)\]\*\*\s*\n(.+?)(?=\n\*\*\[|\n---|\Z)",
    re.DOTALL,
)


def _srt_time_to_seconds(ts: str) -> float:
    h, m, rest = ts.split(":")
    s, ms = rest.split(",")
    return int(h) * 3600 + int(m) * 60 + int(s) + int(ms) / 1000


def _vtt_time_to_seconds(ts: str) -> float:
    parts = ts.replace(".", ":").split(":")
    if len(parts) == 3:
        m, s, ms = parts
        return int(m) * 60 + int(s) + int(ms) / 1000
    if len(parts) == 4:
        h, m, s, ms = parts
        return int(h) * 3600 + int(m) * 60 + int(s) + int(ms) / 1000
    return 0.0


def _md_time_to_seconds(ts: str) -> float:
    parts = ts.strip().split(":")
    if len(parts) == 2:
        return int(parts[0]) * 60 + int(parts[1])
    if len(parts) == 3:
        return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
    return 0.0


def _parse_srt(text: str, filepath: str) -> TranscriptResult:
    segments: list[TranscriptSegment] = []
    for match in _SRT_BLOCK.finditer(text):
        segments.append(
            TranscriptSegment(
                start=_srt_time_to_seconds(match.group(2)),
                end=_srt_time_to_seconds(match.group(3)),
                text=match.group(4).strip(),
            )
        )
    return TranscriptResult(
        segments=segments, language="", source_file=filepath, model=""
    )


def _parse_vtt(text: str, filepath: str) -> TranscriptResult:
    segments: list[TranscriptSegment] = []
    for match in _VTT_CUE.finditer(text):
        segments.append(
            TranscriptSegment(
                start=_vtt_time_to_seconds(match.group(1)),
                end=_vtt_time_to_seconds(match.group(2)),
                text=match.group(3).strip(),
            )
        )
    return TranscriptResult(
        segments=segments, language="", source_file=filepath, model=""
    )


def _parse_markdown(text: str, filepath: str) -> TranscriptResult:
    segments: list[TranscriptSegment] = []
    # Extract metadata
    language = ""
    model = ""
    lang_match = re.search(r"\*\*Language:\*\*\s*(\S+)", text)
    if lang_match:
        language = lang_match.group(1)
    model_match = re.search(r"\*\*Model:\*\*\s*(\S+)", text)
    if model_match:
        model = model_match.group(1)

    for match in _MD_SEGMENT.finditer(text):
        segments.append(
            TranscriptSegment(
                start=_md_time_to_seconds(match.group(1)),
                end=_md_time_to_seconds(match.group(2)),
                text=match.group(3).strip(),
            )
        )
    return TranscriptResult(
        segments=segments, language=language, source_file=filepath, model=model
    )
