# -*- coding: utf-8 -*-
"""Transcription support for clipgen via faster-whisper (CTranslate2-based Whisper).

The Whisper model is lazy-loaded on first use and cached at module level for the session.

Data types (defined below):
  TranscriptSegment – TypedDict: start (float), end (float), text (str)
  TranscriptResult  – TypedDict: segments, language, source_file, model
  ManifestSegment   – TypedDict: id (str), start, end, text — enriched segment for manifest storage

Key functions:
  transcribe_video(path, *, language, initial_prompt, context_keywords)
    → TranscriptResult; context_keywords are appended to the initial prompt
  filter_segments(result, start_sec, end_sec, *, offset_to_zero)
    → TranscriptResult for a clip's time range; offset_to_zero=True shifts to clip-relative times
  write_transcript(result, output_path, *, fmt)
    → writes .md (Markdown), .srt (SRT), or .vtt (WebVTT)
  read_transcript(filepath) → TranscriptResult (parses any supported format)
  get_transcript_extension(fmt) → file extension string for a format

Manifest I/O:
  load_transcripts_manifest() → dict with source_transcripts, corrections, and marks keys
  save_transcripts_manifest(source_transcripts, corrections, marks=None) → assigns segment IDs, writes JSON

Corrections:
  apply_corrections(segments, corrections) → new segment list with from→to substitutions applied
  get_corrections_keywords(corrections) → unique "to" values for Whisper context_keywords

Pipeline integration: clipgen.process_clips() calls _transcribe_segments() which checks the
transcripts manifest for pre-existing source transcripts, then falls back to live Whisper.
Transcript segments are embedded on clip/reel artifact records as a ``transcript`` field.
Standalone transcript file output (type "transcript" artifacts) is opt-in via --transcribe.
"""

import copy

import queue
import re
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, TypedDict

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


class ManifestSegment(TypedDict):
    """Segment stored in transcripts_manifest.json — includes an ID for provenance tracking."""

    id: str  # e.g. "P01:42"
    start: float
    end: float
    text: str


# ---------------------------------------------------------------------------
# Mark categories (shared constant — mirrored in transcripts.js)
# ---------------------------------------------------------------------------

MARK_CATEGORIES: dict[str, dict[str, str]] = {
    "pain_point": {"label": "Pain Point", "color": "#dc2626"},
    "delight": {"label": "Delight", "color": "#16a34a"},
    "quote": {"label": "Quote", "color": "#2563eb"},
    "insight": {"label": "Insight", "color": "#f97316"},
    "task": {"label": "Task Issue", "color": "#8b5cf6"},
    "bookmark": {"label": "Bookmark", "color": "#0891b2"},
}


# ---------------------------------------------------------------------------
# Module-level model cache
# ---------------------------------------------------------------------------

_cached_model: Any = None
_cached_model_name: str | None = None


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
    language: str | None = None,
    initial_prompt: str | None = None,
    context_keywords: list[str] | None = None,
    on_segment: Callable[[float, "TranscriptSegment"], None] | None = None,
) -> TranscriptResult | None:
    """Transcribe a video file and return timestamped segments.

    Args:
        video_path: Path to the video file.
        language: Language code (e.g. "en"). None = auto-detect.
        initial_prompt: Override the default initial prompt.
        context_keywords: Extra keywords to append to the prompt.
        on_segment: Optional callback invoked after each segment with the
            segment's end time (seconds) and the TranscriptSegment.
            Useful for progress tracking and streaming partial results.

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
        segments: list[TranscriptSegment] = []
        for seg in segments_iter:
            text = seg.text.strip()
            if not text:
                continue
            segments.append(TranscriptSegment(start=seg.start, end=seg.end, text=text))
            if on_segment is not None:
                on_segment(seg.end, segments[-1])
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
# Transcripts manifest (source-video transcripts + corrections dictionary)
# ---------------------------------------------------------------------------


def _empty_transcripts_manifest() -> dict[str, Any]:
    return {"source_transcripts": {}, "corrections": [], "marks": []}


def load_transcripts_manifest() -> dict[str, Any]:
    """Load the transcripts manifest from the output directory.

    Returns a dict with ``source_transcripts`` and ``corrections`` keys.
    Missing or corrupt files return an empty manifest.
    """
    data = utils.load_json_manifest(config.TRANSCRIPTS_MANIFEST_FILENAME)
    if not isinstance(data, dict):
        return _empty_transcripts_manifest()
    return {
        "source_transcripts": data.get("source_transcripts") or {},
        "corrections": data.get("corrections") or [],
        "marks": data.get("marks") or [],
    }


def save_transcripts_manifest(
    source_transcripts: dict[str, Any],
    corrections: list[dict[str, Any]],
    marks: list[dict[str, Any]] | None = None,
) -> Path | None:
    """Write the transcripts manifest to disk.

    Assigns segment IDs (``"{participant}:{index}"``) at storage time.
    *marks* defaults to ``None`` which preserves whatever marks are already on
    disk (load → merge → save).  Pass an explicit list to overwrite.
    Returns the manifest path on success, or ``None`` on failure.
    """
    for participant_id, entry in source_transcripts.items():
        for idx, seg in enumerate(entry.get("segments", [])):
            seg["id"] = f"{participant_id}:{idx}"

    # When marks is None, preserve existing marks from disk
    if marks is None:
        existing = utils.load_json_manifest(
            config.TRANSCRIPTS_MANIFEST_FILENAME, default={}
        )
        marks = (existing.get("marks") or []) if isinstance(existing, dict) else []

    data = {
        "source_transcripts": source_transcripts,
        "corrections": corrections,
        "marks": marks,
    }
    return utils.save_json_manifest(
        config.TRANSCRIPTS_MANIFEST_FILENAME,
        data,
        warn_label="transcripts manifest",
    )


# ---------------------------------------------------------------------------
# Corrections
# ---------------------------------------------------------------------------


def apply_corrections(
    segments: list[TranscriptSegment],
    corrections: list[dict[str, Any]],
) -> list[TranscriptSegment]:
    """Apply corrections to transcript segments as post-processing.

    Returns a new list of segments with ``from -> to`` substitutions applied.
    Never mutates the input. Case-insensitive, word-boundary matching.
    """
    if not corrections:
        return list(segments)

    pairs = [(c["from"], c["to"]) for c in corrections if c.get("from") and c.get("to")]
    if not pairs:
        return list(segments)

    corrected: list[TranscriptSegment] = []
    total_applied = 0
    for seg in segments:
        text = seg["text"]
        seg_applied = 0
        for from_text, to_text in pairs:
            pattern = re.compile(re.escape(from_text), re.IGNORECASE)
            new_text, count = pattern.subn(to_text, text)
            if count > 0:
                text = new_text
                seg_applied += count
        total_applied += seg_applied
        corrected.append(
            TranscriptSegment(start=seg["start"], end=seg["end"], text=text)
        )

    if total_applied > 0:
        utils.verbose_print(
            f"  Auto-applied {total_applied} correction(s) across {len(segments)} segment(s)."
        )
    return corrected


def get_corrections_keywords(corrections: list[dict[str, Any]]) -> list[str]:
    """Extract unique ``to`` values from corrections for use as Whisper context keywords."""
    seen: set[str] = set()
    keywords: list[str] = []
    for c in corrections:
        to_val = c.get("to", "").strip()
        if to_val and to_val not in seen:
            seen.add(to_val)
            keywords.append(to_val)
    return keywords


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
        f"- **Generated:** {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')}",
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


def get_transcript_extension(fmt: str | None = None) -> str:
    """Return the file extension for the given transcript format."""
    return _FORMAT_EXT.get(fmt or config.TRANSCRIBE_FORMAT, ".md")


def write_transcript(
    result: TranscriptResult,
    output_path: str,
    *,
    fmt: str | None = None,
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


def read_transcript(filepath: str) -> TranscriptResult | None:
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


# ---------------------------------------------------------------------------
# TranscriptWorker — background transcription task queue
# ---------------------------------------------------------------------------

TASK_STATUS_QUEUED = "queued"
TASK_STATUS_RUNNING = "running"
TASK_STATUS_COMPLETED = "completed"
TASK_STATUS_FAILED = "failed"
TASK_STATUS_CANCELLED = "cancelled"

_TRANSCRIPT_SENTINEL = object()


def create_transcript_task(participant: str, video_path: str) -> dict[str, Any]:
    """Create a new transcription task dict ready to enqueue."""
    return {
        "id": f"tr_{uuid.uuid4().hex[:8]}",
        "participant": participant,
        "video_path": video_path,
        "status": TASK_STATUS_QUEUED,
        "progress": 0.0,
        "partial_segments": [],
        "result": None,
        "error": None,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "completed_at": None,
    }


class TranscriptWorker:
    """Background thread that processes transcription tasks sequentially.

    Simplified version of ``ScreenspaceWorker`` — no pause/resume, no
    reordering, no concurrent execution, no event generation.
    """

    def __init__(self) -> None:
        self._queue: queue.PriorityQueue[Any] = queue.PriorityQueue()
        self._tasks: dict[str, dict[str, Any]] = {}
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._running = False
        self.on_task_complete: Callable[[], None] | None = None

    def start(self) -> None:
        """Start the worker thread."""
        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        """Signal the worker thread to stop."""
        self._running = False
        self._queue.put((0, _TRANSCRIPT_SENTINEL))
        if self._thread is not None:
            self._thread.join(timeout=15)

    def restore_tasks(self, tasks: list[dict[str, Any]]) -> None:
        """Load historical tasks (completed/failed/cancelled) for display."""
        with self._lock:
            for t in tasks:
                if t.get("id"):
                    self._tasks[t["id"]] = copy.deepcopy(t)

    def enqueue(self, task: dict[str, Any]) -> str:
        """Add a task to the queue. Returns the task ID."""
        task_id = task["id"]
        with self._lock:
            self._tasks[task_id] = task
        self._queue.put((100, task_id))
        return task_id

    def cancel(self, task_id: str) -> bool:
        """Cancel a queued task. Returns True if cancelled."""
        with self._lock:
            task = self._tasks.get(task_id)
            if task and task["status"] == TASK_STATUS_QUEUED:
                task["status"] = TASK_STATUS_CANCELLED
                return True
        return False

    def get_task(self, task_id: str) -> dict[str, Any] | None:
        """Return a task dict by ID (thread-safe copy)."""
        with self._lock:
            t = self._tasks.get(task_id)
            return copy.deepcopy(t) if t else None

    def get_all_tasks(self) -> list[dict[str, Any]]:
        """Return all tasks (thread-safe copies)."""
        with self._lock:
            return [copy.deepcopy(t) for t in self._tasks.values()]

    @property
    def is_alive(self) -> bool:
        """Return whether the worker thread is alive."""
        return self._thread is not None and self._thread.is_alive()

    def _run(self) -> None:
        """Worker loop: dequeue and execute tasks sequentially."""
        while self._running:
            try:
                item = self._queue.get(timeout=0.5)
            except queue.Empty:
                continue

            _priority, task_id = item
            if task_id is _TRANSCRIPT_SENTINEL:
                break

            with self._lock:
                task = self._tasks.get(task_id)
                if task is None or task["status"] != TASK_STATUS_QUEUED:
                    continue
                task["status"] = TASK_STATUS_RUNNING

            self._execute_task(task)

            if self.on_task_complete:
                try:
                    self.on_task_complete()
                except Exception:
                    pass

    def _execute_task(self, task: dict[str, Any]) -> None:
        """Run a single transcription task."""
        import video as video_mod

        video_path = task["video_path"]

        # Probe video duration for progress estimation
        duration = 0.0
        props = video_mod.probe_video_properties(video_path)
        if props:
            duration = props.get("duration", 0.0)

        # Load corrections for context keywords
        manifest = load_transcripts_manifest()
        corrections = manifest.get("corrections", [])
        context_kw = get_corrections_keywords(corrections) or None

        def _on_seg(end_time: float, segment: TranscriptSegment) -> None:
            with self._lock:
                if duration > 0:
                    task["progress"] = min(end_time / duration, 0.99)
                task["partial_segments"].append(segment)

        try:
            result = transcribe_video(
                video_path,
                context_keywords=context_kw,
                on_segment=_on_seg,
            )
            if result is None:
                with self._lock:
                    task["status"] = TASK_STATUS_FAILED
                    task["error"] = "Transcription returned None"
                    task["partial_segments"] = []
                    task["completed_at"] = datetime.now(timezone.utc).isoformat()
                return

            with self._lock:
                task["status"] = TASK_STATUS_COMPLETED
                task["progress"] = 1.0
                task["partial_segments"] = []
                task["result"] = {
                    "segments": result["segments"],
                    "language": result["language"],
                    "model": result["model"],
                    "source_file": result["source_file"],
                    "transcribed_at": datetime.now(timezone.utc).isoformat(),
                }
                task["completed_at"] = datetime.now(timezone.utc).isoformat()

        except Exception as exc:
            with self._lock:
                task["status"] = TASK_STATUS_FAILED
                task["error"] = str(exc)
                task["partial_segments"] = []
                task["completed_at"] = datetime.now(timezone.utc).isoformat()
