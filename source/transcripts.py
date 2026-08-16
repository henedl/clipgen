"""Transcription support for clipgen via faster-whisper (CTranslate2-based Whisper).

The Whisper model is lazy-loaded on first use and cached at module level for the session.
Loads are serialized with a lock so background warm-up and transcription cannot race.

Data types (defined below):
  TranscriptSegment – TypedDict: start (float), end (float), text (str)
  TranscriptResult  – TypedDict: segments, language, source_file, model
  ManifestSegment   – TypedDict: id (str), start, end, text — enriched segment for manifest storage

Key functions:
  transcribe_video(path, *, model_name, language, initial_prompt, context_keywords)
    → TranscriptResult; context_keywords are appended to the initial prompt;
      model_name overrides config.TRANSCRIBE_MODEL for one call
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

Anti-hallucination knobs (``config.TRANSCRIBE_*``) are passed through to faster-whisper:
VAD pre-filter, no-speech / log-probability / compression-ratio thresholds, optional
hallucination silence skip (requires word timestamps when > 0), and condition-on-previous-text.
"""

import copy

import os
import queue
import re
import sys
import threading
import time
import uuid
from collections.abc import Callable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, Literal, TypedDict

import config
import profiling
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


# Known faster-whisper model variants with approximate download sizes.
WHISPER_MODELS: list[dict[str, Any]] = [
    {"name": "tiny", "size_mb": 40, "description": "Fastest, least accurate"},
    {"name": "base", "size_mb": 140, "description": "Fast, good for short segments"},
    {"name": "small", "size_mb": 500, "description": "Balanced speed and accuracy"},
    {"name": "medium", "size_mb": 1500, "description": "Slower, more accurate"},
    {
        "name": "large-v3",
        "size_mb": 2900,
        "description": "Best accuracy, requires significant RAM",
    },
]


# ---------------------------------------------------------------------------
# Module-level model cache
# ---------------------------------------------------------------------------

_cached_model: Any = None
_cached_model_name: str | None = None
# The full construction signature the cached model was built from. The name alone
# is not enough: device, compute type and thread count are all read at
# WhisperModel() time and all user-editable, so name-keying let a switch to
# TRANSCRIBE_DEVICE=cuda save and display while transcription silently kept
# running on the model already loaded for cpu.
_cached_model_key: tuple[Any, ...] | None = None
_model_load_lock = threading.Lock()
# True while _load_model is actually constructing a WhisperModel — the ~10s a
# cold load takes. Lets the model-status endpoint report "warming" for
# on-demand loads too, not just the explicit warmup path.
_model_loading = False


def is_transcription_model_loading() -> bool:
    """True while a WhisperModel construction is in flight (any trigger)."""
    return _model_loading


# ---------------------------------------------------------------------------
# Core transcription
# ---------------------------------------------------------------------------


def _model_load_key(model_name: str) -> tuple[Any, ...]:
    """The full signature ``_do_load`` would construct a WhisperModel from.

    Every element is read at construction time and is user-editable, so any of
    them changing means the cached model is no longer the one the settings
    describe. Resolved rather than raw: ``TRANSCRIBE_DEVICE="auto"`` and
    ``"cpu"`` are the same model in a frozen build, and reloading between them
    would cost ~10s for nothing.
    """
    threads = config.TRANSCRIBE_CPU_THREADS or (os.cpu_count() or 0)
    return (
        model_name,
        _resolve_transcribe_device(),
        config.TRANSCRIBE_COMPUTE_TYPE,
        threads,
    )


def is_transcription_model_loaded() -> bool:
    """Return True if a usable model for the current settings is warm.

    Keyed on the whole load signature, not just the name: after a device change
    the loaded model no longer matches the settings, and reporting it as loaded
    would hide the reload the next transcription pays for.
    """
    if config.DEBUGGING:
        return True
    return _cached_model is not None and _cached_model_key == _model_load_key(
        config.TRANSCRIBE_MODEL
    )


# What faster_whisper.utils.download_model does for a bare size name, mirrored
# here so the cache check can call huggingface_hub directly: importing
# faster_whisper costs ~600 ms (it pulls ctranslate2/tokenizers/av at package
# import), which used to land on the first /api/models request of every server
# session. huggingface_hub imports in ~1 ms. The repo prefix and allow_patterns
# are pinned against faster_whisper.utils in tests/test_transcripts.py.
_WHISPER_REPO_PREFIX = "Systran/faster-whisper-"
_WHISPER_ALLOW_PATTERNS = [
    "config.json",
    "preprocessor_config.json",
    "model.bin",
    "tokenizer.json",
    "vocabulary.*",
]


def is_whisper_model_cached(model_name: str | None = None) -> bool:
    """Return True if *model_name* is already downloaded to the local HF cache.

    Used to gate silent downloads: a non-cached model requires explicit user
    confirmation before transcription pulls it (40 MB-2.9 GB). Mirrors
    faster-whisper's ``download_model(local_files_only=True)`` — the same
    ``snapshot_download`` call on the same repo id and file patterns — without
    importing faster_whisper itself; it only inspects the cache and does
    **not** load model weights.
    """
    if config.DEBUGGING:
        return True
    model_name = model_name or config.TRANSCRIBE_MODEL
    if _cached_model is not None and _cached_model_name == model_name:
        return True
    try:
        from huggingface_hub import snapshot_download
    except ImportError:
        return False
    # download_model's rule: anything with a "/" is already a full repo id.
    repo_id = model_name if "/" in model_name else _WHISPER_REPO_PREFIX + model_name
    try:
        snapshot_download(
            repo_id,
            local_files_only=True,
            allow_patterns=_WHISPER_ALLOW_PATTERNS,
        )
        return True
    except Exception:
        # huggingface_hub raises (LocalEntryNotFoundError/FileNotFoundError)
        # when the snapshot is absent or incomplete in the cache.
        return False


def warmup_transcription_model() -> bool:
    """Ensure the Whisper model is loaded into the module cache.

    Returns True when a model is available for transcription (including DEBUGGING
    stub mode). Returns False if loading failed (e.g. missing faster-whisper).
    """
    if config.DEBUGGING:
        return True
    return _load_model() is not None


def _stdin_is_interactive() -> bool:
    """True when there is a terminal on stdin to prompt against.

    A closed stdin has no ``isatty``, and pytest's capture stub answers False —
    both mean the same thing here: do not call ``input()``.
    """
    try:
        return bool(sys.stdin) and sys.stdin.isatty()
    except (AttributeError, ValueError, OSError):
        return False


def _confirm_model_download(model_name: str) -> bool:
    """Gate a first-time Whisper model download on the terminal.

    The web UI has had an explicit consent dialog and an authoritative
    server-side ``allow_download`` gate since transcription shipped; the CLI had
    neither, so ``--transcribe`` with a large model silently pulled up to 2.9 GB
    behind a "Loading transcription model..." spinner. Returns True to proceed.

    Non-interactive runs announce the download and continue rather than blocking
    on a prompt nobody can answer. That covers ``--no-input`` and every
    server-side call (which forces ``NO_INPUT_MODE`` — and where the browser has
    already asked), but also a piped or closed stdin: ``input()`` raises there,
    so asking would turn a scripted ``--transcribe`` into a crash. Consent is a
    guard against a surprise 2.9 GB download, not a reason to fail the run.
    """
    if is_whisper_model_cached(model_name):
        return True
    size_mb = next((m["size_mb"] for m in WHISPER_MODELS if m["name"] == model_name), 0)
    size = f" (~{size_mb / 1000:.1f} GB)" if size_mb >= 1000 else f" (~{size_mb} MB)"
    if not size_mb:
        size = ""
    if utils.NO_INPUT_MODE or not _stdin_is_interactive():
        utils.info_print(f"Downloading transcription model '{model_name}'{size}...")
        return True
    utils.warning_print(
        f"The '{model_name}' transcription model{size} is not downloaded yet.",
        details=["It will be downloaded once and stored locally."],
    )
    answer = utils.read_user_input("Download it now? [y/n]\n>> ")
    if answer.strip().lower() in ("y", "yes"):
        return True
    utils.info_print(
        "Skipped. Pick a smaller model with TRANSCRIBE_MODEL (see --settings)."
    )
    return False


def _resolve_transcribe_device() -> str:
    """Return the device string to hand ``WhisperModel``.

    faster-whisper defaults to ``device="auto"``, which makes CTranslate2 pick
    CUDA whenever it can see an NVIDIA device — and CTranslate2 ships its own
    CUDA support independent of torch, needing a matching cuBLAS/cuDNN beside
    it. The desktop bundle has neither: nothing in the dependency tree pulls a
    CUDA runtime, so no ``nvidia-*`` wheel is installed or collected. On any machine
    with an NVIDIA GPU the frozen app therefore selected CUDA and then died at
    the first inference with ``Library cublas64_12.dll is not found or cannot
    be loaded`` — long after the model had finished downloading and loading.

    So "auto" resolves to CPU when frozen. Explicit "cpu"/"cuda" are passed
    through untouched: a user who installed the CUDA runtime themselves is
    entitled to ask for it, and to see CTranslate2's own error if it is wrong.
    """
    device = (config.TRANSCRIBE_DEVICE or "auto").strip().lower()
    if device not in ("auto", "cpu", "cuda"):
        utils.warning_print(
            f"Unknown TRANSCRIBE_DEVICE {config.TRANSCRIBE_DEVICE!r}; using auto."
        )
        device = "auto"
    if device == "auto" and getattr(sys, "frozen", False):
        return "cpu"
    return device


def _load_model(model_name: str | None = None) -> Any:
    """Lazy-load the WhisperModel, caching it for reuse (thread-safe).

    When *model_name* is None, uses ``config.TRANSCRIBE_MODEL``. Passing a
    different name swaps the cached model — as does changing any other setting
    the construction reads (device, compute type, thread count); see
    :func:`_model_load_key`.
    """
    global _cached_model, _cached_model_name, _cached_model_key

    model_name = model_name or config.TRANSCRIBE_MODEL
    load_key = _model_load_key(model_name)
    if _cached_model is not None and _cached_model_key == load_key:
        profiling.count("transcribe.model_cache.hit")
        return _cached_model

    # Timed separately from the load itself: n here is the number of callers
    # that missed the fast path, so a model_lock_wait n=14 across a 14-file
    # batch means the cache is being thrashed by a settings change — which is
    # exactly what _model_load_key exists to surface.
    _t_lock = time.perf_counter() if config.PROFILING else 0.0
    with _model_load_lock:
        if _t_lock:
            profiling.add("transcribe.model_lock_wait", time.perf_counter() - _t_lock)
        if _cached_model is not None and _cached_model_key == load_key:
            # Another thread finished the load while we waited. Still a hit —
            # counting only the pre-lock check would report this as a miss.
            profiling.count("transcribe.model_cache.hit")
            return _cached_model
        profiling.count("transcribe.model_cache.miss")

        try:
            from faster_whisper import WhisperModel
        except ImportError:
            utils.error_print(
                "faster-whisper is not installed.",
                details=["Install it with: uv add faster-whisper"],
            )
            return None

        # Drop the previous model first, or ~1-2 GB of weights is double-held
        # until the next GC cycle; gc.collect() also prods CUDA/MPS cleanup.
        if _cached_model is not None:
            import gc

            _cached_model = None
            _cached_model_name = None
            _cached_model_key = None
            gc.collect()

        def _do_load() -> Any:
            # Built from load_key, not re-read from config, so what is cached
            # is exactly what was constructed even if a setting changes while
            # this load is in flight.
            #
            # cpu_threads 0 = auto: CTranslate2's own heuristic under-uses
            # many-core CPUs, so resolve to os.cpu_count(). num_workers stays
            # default — the poller runs one job at a time, so >1 only multiplies
            # model memory.
            _name, device, compute_type, threads = load_key
            load_kwargs: dict[str, Any] = {
                "compute_type": compute_type,
                "device": device,
            }
            if threads > 0:
                load_kwargs["cpu_threads"] = threads
            # Scoped to the construction only — deliberately inside _do_load
            # rather than around the run_with_spinner call below, so it excludes
            # _confirm_model_download, which can block on interactive input. A
            # span that includes a human deciding whether to fetch 3 GB is not a
            # measurement. On a cold HF cache this *does* include the download,
            # so a first run is orders of magnitude larger and is not a
            # regression.
            with profiling.span("transcribe.model_load"):
                return WhisperModel(model_name, **load_kwargs)

        global _model_loading
        _model_loading = True
        try:
            if not _confirm_model_download(model_name):
                return None

            utils.info_print(f"Loading transcription model '{model_name}'...")
            _cached_model = utils.run_with_spinner(
                "Loading transcription model...", _do_load
            )
            _cached_model_name = model_name
            _cached_model_key = load_key
            return _cached_model
        finally:
            _model_loading = False


def _build_transcribe_kwargs(
    *,
    language: str | None,
    initial_prompt: str,
) -> dict[str, Any]:
    """Return keyword arguments for ``WhisperModel.transcribe`` from config."""
    kwargs: dict[str, Any] = {
        "beam_size": config.TRANSCRIBE_BEAM_SIZE,
        "language": language,
        "initial_prompt": initial_prompt,
        "vad_filter": config.TRANSCRIBE_VAD_FILTER,
        "no_speech_threshold": config.TRANSCRIBE_NO_SPEECH_THRESHOLD,
        "log_prob_threshold": config.TRANSCRIBE_LOG_PROB_THRESHOLD,
        "compression_ratio_threshold": config.TRANSCRIBE_COMPRESSION_RATIO_THRESHOLD,
        "condition_on_previous_text": config.TRANSCRIBE_CONDITION_ON_PREVIOUS_TEXT,
    }
    # Recall-safe VAD tuning: a low threshold plus boundary padding so quiet speech
    # and word edges aren't clipped. Sent only when VAD is on; faster-whisper merges
    # this partial dict with its VadOptions defaults.
    if config.TRANSCRIBE_VAD_FILTER:
        kwargs["vad_parameters"] = {
            "threshold": config.TRANSCRIBE_VAD_THRESHOLD,
            "speech_pad_ms": config.TRANSCRIBE_VAD_SPEECH_PAD_MS,
            "min_silence_duration_ms": config.TRANSCRIBE_VAD_MIN_SILENCE_MS,
        }
    hall_silence = config.TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD
    if hall_silence > 0:
        kwargs["hallucination_silence_threshold"] = hall_silence
        kwargs["word_timestamps"] = True
    return kwargs


def _resolve_audio_index(video_path: str, requested: int | None) -> int:
    """Resolve which audio stream to transcribe: explicit index, else auto-detect.

    An explicit *requested* index (including 0) always wins. ``None`` runs
    ``video.pick_speech_audio_track`` over the file's track names, so a session
    whose participant mic landed on track 2 transcribes the mic rather than the
    system audio. The probe is cached by (path, mtime_ns) in video.py, so callers
    that already probed pay a dict lookup, not a second ffprobe.
    """
    if requested is not None:
        return requested
    import video as video_mod

    props = video_mod.probe_video_properties(video_path)
    tracks = (props or {}).get("audio_tracks") or []
    return video_mod.pick_speech_audio_track(tracks)


def _flush_transcribe_profile(
    name: str,
    *,
    prepare_s: float,
    decode_s: float,
    decode_max: float,
    n_segments: int,
    callback_s: float,
    callback_max: float,
    n_callbacks: int,
    audio_s: float,
    info: Any,
) -> None:
    """Flush one transcription's accumulated timings plus a per-run summary line.

    The realtime factor is computed against ``prepare + decode``, never decode
    alone. VAD shrinks decode and hides its own cost in prepare, so a
    decode-only ratio *improves* whenever VAD is enabled — it would endorse
    ``TRANSCRIBE_VAD_FILTER`` no matter what the setting actually did.

    The numerator is the last segment's ``end``, not ``info.duration``, because
    it stays correct on the cancelled path: a run stopped ten minutes into an
    hour-long file reports a true 10-min/wall ratio, where ``info.duration``
    would report a wildly optimistic 60-min/wall one. ``file=`` is printed
    alongside so a truncated run is visible (``audio`` << ``file``), and
    ``vad=`` gives the VAD win directly as ``vad/file``.
    """
    profiling.add("transcribe.decode", decode_s, n_segments, peak=decode_max)
    profiling.add("transcribe.callback", callback_s, n_callbacks, peak=callback_max)

    file_s = float(getattr(info, "duration", 0.0) or 0.0)
    vad_s = float(getattr(info, "duration_after_vad", 0.0) or 0.0)
    extra = f"audio={audio_s:.1f}s  file={file_s:.1f}s"
    if vad_s:
        extra += f"  vad={vad_s:.1f}s"
    wall = prepare_s + decode_s
    if wall > 0 and audio_s > 0:
        extra += f"  xrt={audio_s / wall:.1f}x"
    profiling.scan_summary(
        name,
        [
            ("prepare", prepare_s, 1),
            ("decode", decode_s, n_segments),
            ("callback", callback_s, n_callbacks),
        ],
        kind="whisper",
        extra=extra,
    )


def transcribe_video(
    video_path: str,
    *,
    model_name: str | None = None,
    language: str | None = None,
    audio_index: int | None = None,
    initial_prompt: str | None = None,
    context_keywords: list[str] | None = None,
    on_segment: Callable[[float, "TranscriptSegment"], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
) -> TranscriptResult | None:
    """Transcribe a video file and return timestamped segments.

    Args:
        video_path: Path to the video file.
        model_name: Whisper model size override (e.g. "tiny", "large-v3").
            None uses ``config.TRANSCRIBE_MODEL``.
        language: Language code (e.g. "en"). None = auto-detect.
        audio_index: Which audio stream (``0:a:N``) to transcribe. None
            auto-detects a speech-looking track from the stream names.
        initial_prompt: Override the default initial prompt.
        context_keywords: Extra keywords to append to the prompt.
        on_segment: Optional callback invoked after each segment with the
            segment's end time (seconds) and the TranscriptSegment.
            Useful for progress tracking and streaming partial results.
        cancel_flag: Optional callable checked at each safe boundary —
            before/after the (slow) model load and between segments. When it
            returns True, transcription aborts with ``_TranscriptionCancelled``
            instead of waiting for the next ``on_segment`` callback. faster-
            whisper exposes no hard abort, so cancellation is cooperative.

    Returns:
        TranscriptResult with segments, or None on failure.
    """
    resolved_model = model_name or config.TRANSCRIBE_MODEL
    if config.DEBUGGING:
        return TranscriptResult(
            segments=[],
            language=language or "en",
            source_file=str(video_path),
            model=resolved_model,
        )

    def _is_cancelled() -> bool:
        return cancel_flag is not None and cancel_flag()

    if _is_cancelled():
        raise _TranscriptionCancelled

    # Short-circuit when the file has no audio stream — faster-whisper raises
    # an opaque "tuple index out of range" from its decoder in that case.
    import video as video_mod

    props = video_mod.probe_video_properties(video_path)
    tracks: list[dict[str, Any]] = (props or {}).get("audio_tracks") or []
    if props is not None and not props.get("audio_codec") and not tracks:
        utils.warning_print(
            f"No audio stream in {Path(video_path).name} — skipping transcription."
        )
        return None

    idx = _resolve_audio_index(video_path, audio_index)
    if tracks and not 0 <= idx < len(tracks):
        utils.warning_print(
            f"{Path(video_path).name} has {len(tracks)} audio track(s); "
            f"track {idx + 1} was requested — skipping transcription."
        )
        return None
    if audio_index is None and idx > 0:
        utils.verbose_print(
            f"Auto-selected audio track {idx + 1} "
            f"({tracks[idx].get('label') or 'unnamed'}) in {Path(video_path).name}"
        )

    model = _load_model(resolved_model)
    if model is None:
        return None
    if _is_cancelled():
        raise _TranscriptionCancelled

    prompt = initial_prompt or config.TRANSCRIBE_INITIAL_PROMPT
    if context_keywords:
        prompt = f"{prompt} Key terms: {', '.join(context_keywords)}."

    lang = language or config.TRANSCRIBE_LANGUAGE

    # faster-whisper's decoder always reads the container's first audio stream —
    # there is no index parameter — so a non-default track has to be demuxed to
    # its own file first. Track 0 keeps the raw video path (byte-identical to the
    # pre-track-picker behaviour); the extraction is the same cached, locked one
    # the browser's per-track volume mixer uses.
    audio_source = str(video_path)
    if idx > 0:
        extracted = video_mod.extract_audio_track(video_path, idx)
        if extracted is None:
            # Never fall back to track 0 — that would silently transcribe the
            # wrong audio, which reads as "clipgen is broken", not "it failed".
            utils.warning_print(
                f"Could not extract audio track {idx + 1} from "
                f"{Path(video_path).name} — skipping transcription."
            )
            return None
        audio_source = str(extracted)

    try:
        transcribe_kwargs = _build_transcribe_kwargs(
            language=lang, initial_prompt=prompt
        )
        # Not a free call: faster-whisper loads the audio, extracts features and
        # runs VAD + language detection *eagerly* here, then hands back a lazy
        # generator. That is why info.duration_after_vad is already populated
        # below. Without this span the TRANSCRIBE_VAD_* knobs — the ones
        # PERFORMANCE.md calls "the big win" — are the one thing the transcribe
        # labels cannot see, because their cost is here and not in the pull.
        with profiling.span("transcribe.prepare"):
            _t_prepare = time.perf_counter()
            segments_iter, info = model.transcribe(audio_source, **transcribe_kwargs)
            _prepare_s = time.perf_counter() - _t_prepare
        if _is_cancelled():
            raise _TranscriptionCancelled
        segments: list[TranscriptSegment] = []
        # Profiling accumulates into locals and flushes once after the loop, so
        # the off-path per-segment cost is a single boolean check (see
        # profiling.py). The flush lives in a finally because a cancelled
        # 20-minute run is exactly the one whose numbers you want.
        _prof = config.PROFILING
        _dec_s = _cb_s = _dec_max = _cb_max = 0.0
        _n_seg = _n_cb = 0
        _audio_s = 0.0
        _t_last = time.perf_counter() if _prof else 0.0
        try:
            for seg in segments_iter:
                if _prof:
                    _t_pull = time.perf_counter()
                    _dt = _t_pull - _t_last
                    _dec_s += _dt
                    _dec_max = max(_dec_max, _dt)
                    _n_seg += 1
                    _audio_s = seg.end  # decode watermark, incl. empty segments
                if _is_cancelled():
                    raise _TranscriptionCancelled
                text = seg.text.strip()
                # Inverted from `if not text: continue` so the re-stamp below is
                # never skipped — an empty segment would otherwise charge its
                # (tiny) processing to the *next* pull.
                if text:
                    segments.append(
                        TranscriptSegment(start=seg.start, end=seg.end, text=text)
                    )
                    if on_segment is not None:
                        if _prof:
                            _t_cb = time.perf_counter()
                        on_segment(seg.end, segments[-1])
                        if _prof:
                            _dt = time.perf_counter() - _t_cb
                            _cb_s += _dt
                            _cb_max = max(_cb_max, _dt)
                            _n_cb += 1
                if _prof:
                    _t_last = time.perf_counter()
        finally:
            if _prof:
                _flush_transcribe_profile(
                    Path(video_path).name,
                    prepare_s=_prepare_s,
                    decode_s=_dec_s,
                    decode_max=_dec_max,
                    n_segments=_n_seg,
                    callback_s=_cb_s,
                    callback_max=_cb_max,
                    n_callbacks=_n_cb,
                    audio_s=_audio_s,
                    info=info,
                )
        detected_lang = (
            info.language if hasattr(info, "language") else (lang or "unknown")
        )
        return TranscriptResult(
            segments=segments,
            language=detected_lang,
            source_file=str(video_path),
            model=resolved_model,
        )
    except _TranscriptionCancelled:
        raise
    except Exception as exc:
        import traceback

        utils.warning_print(
            f"Transcription failed for {Path(video_path).name}: {exc}",
            details=traceback.format_exc().rstrip().splitlines(),
        )
        return None


def transcribe_timeline(
    timeline: list[tuple[str, int, int]],
    *,
    model_name: str | None = None,
    language: str | None = None,
    audio_index: int | None = None,
    context_keywords: list[str] | None = None,
    on_segment: Callable[[float, "TranscriptSegment"], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
) -> TranscriptResult | None:
    """Transcribe an ordered set of source-video parts as one continuous timeline.

    For a participant whose session spans several files (see
    ``video.build_source_timeline``): each part is transcribed independently,
    then its segment times are shifted by the part's cumulative start so the
    merged result lands on the participant's global timeline (matching clip
    artifact times). Returns ``None`` if any part fails. A single-element
    timeline is just a normal transcription at offset 0.

    ``on_segment``/``cancel_flag`` are forwarded to each part's transcription so
    live progress streaming and cancellation work for multi-part jobs; the
    reported end time and segment times are shifted to the global timeline.
    """
    merged: list[TranscriptSegment] = []
    out_language = ""
    out_model = ""
    # Resolve the audio track ONCE, from the first part, and pass the concrete
    # index to every part. Auto-detecting per part would splice two different
    # microphones into one transcript if the recorder's track order shifted
    # between files.
    resolved_audio_index = (
        _resolve_audio_index(timeline[0][0], audio_index) if timeline else 0
    )
    for path, _duration, cumulative in timeline:
        part_on_segment: Callable[[float, TranscriptSegment], None] | None = None
        if on_segment is not None:
            inner = on_segment

            # Small per-part shifter: rebases this part's segment times onto the
            # stitched timeline via the default-arg captures below.
            def part_on_segment(
                end_time: float,
                segment: "TranscriptSegment",
                _cum: int = cumulative,
                _cb: Callable[[float, "TranscriptSegment"], None] = inner,
            ) -> None:
                _cb(
                    end_time + _cum,
                    {
                        "start": segment["start"] + _cum,
                        "end": segment["end"] + _cum,
                        "text": segment["text"],
                    },
                )

        result = transcribe_video(
            path,
            model_name=model_name,
            language=language,
            audio_index=resolved_audio_index,
            context_keywords=context_keywords,
            on_segment=part_on_segment,
            cancel_flag=cancel_flag,
        )
        if result is None:
            return None
        for seg in result["segments"]:
            merged.append(
                {
                    "start": seg["start"] + cumulative,
                    "end": seg["end"] + cumulative,
                    "text": seg["text"],
                }
            )
        out_language = out_language or result["language"]
        out_model = out_model or result["model"]
    return {
        "segments": merged,
        "language": out_language,
        "source_file": " + ".join(path for path, _d, _c in timeline),
        "model": out_model,
    }


# ---------------------------------------------------------------------------
# Transcripts manifest (source-video transcripts + corrections dictionary)
# ---------------------------------------------------------------------------


def _empty_transcripts_manifest() -> dict[str, Any]:
    return {"source_transcripts": {}, "corrections": [], "marks": []}


def _is_empty_transcripts_manifest(data: dict[str, Any]) -> bool:
    """True when no transcripts, corrections, or marks exist — nothing to persist."""
    return not (
        data.get("source_transcripts") or data.get("corrections") or data.get("marks")
    )


# Module-level cache for the parsed transcripts manifest, keyed on the file's
# path and mtime_ns. Per-reel builds and the CLI hit load_transcripts_manifest()
# repeatedly, and the manifest (every participant's full segment list) is easily
# multi-MB, so re-reading and re-parsing the JSON each time is pure overhead. The
# cache invalidates automatically whenever the file is rewritten (the atomic
# save_transcripts_manifest bumps mtime); _reset_transcripts_manifest_cache()
# exists only as a guard against coarse filesystem mtime resolution eliding a
# same-tick save->load, and for test fixtures that reuse one output dir.
_TRANSCRIPTS_MANIFEST_CACHE_LOCK = threading.Lock()
_transcripts_manifest_cache: dict[str, Any] = {
    "path": None,
    "mtime_ns": None,
    "data": None,
}


def _reset_transcripts_manifest_cache() -> None:
    """Drop the in-memory transcripts-manifest cache. For save() + test fixtures."""
    with _TRANSCRIPTS_MANIFEST_CACHE_LOCK:
        _transcripts_manifest_cache["path"] = None
        _transcripts_manifest_cache["mtime_ns"] = None
        _transcripts_manifest_cache["data"] = None


def load_transcripts_manifest() -> dict[str, Any]:
    """Load the transcripts manifest from the output directory.

    Returns a dict with ``source_transcripts``, ``corrections``, and ``marks`` keys.

    Memoizes the parsed result keyed on the manifest's path + mtime_ns so repeated
    calls from the same process share a single read/parse until the file is
    rewritten. Returns a deep copy so callers that mutate the returned structure in
    place (e.g. ``--summarize``/``--citations`` set fields on ``source_transcripts``
    entries before saving) cannot corrupt the cached state.
    """
    path = Path(utils.get_effective_output_dir()) / config.TRANSCRIPTS_MANIFEST_FILENAME
    path_str = str(path)
    try:
        mtime_ns: int | None = path.stat().st_mtime_ns if path.is_file() else None
    except OSError:
        mtime_ns = None

    with _TRANSCRIPTS_MANIFEST_CACHE_LOCK:
        if (
            mtime_ns is not None
            and _transcripts_manifest_cache["path"] == path_str
            and _transcripts_manifest_cache["mtime_ns"] == mtime_ns
        ):
            return copy.deepcopy(_transcripts_manifest_cache["data"])

    data = utils.load_json_manifest(
        config.TRANSCRIPTS_MANIFEST_FILENAME, default=_empty_transcripts_manifest()
    )

    with _TRANSCRIPTS_MANIFEST_CACHE_LOCK:
        _transcripts_manifest_cache["path"] = path_str
        _transcripts_manifest_cache["mtime_ns"] = mtime_ns
        _transcripts_manifest_cache["data"] = data

    return copy.deepcopy(data)


def save_transcripts_manifest(
    source_transcripts: dict[str, Any],
    corrections: list[dict[str, Any]],
    marks: list[dict[str, Any]] | None = None,
) -> Path | None:
    """Write the transcripts manifest to disk.

    Assigns a segment ID (``"{participant}:{index}"``) when a segment doesn't
    already have one, but never overwrites an existing id — so marks and
    corrections that reference ``seg["id"]`` keep pointing at the same
    segment even when later segments are added, edited, or reordered.
    *marks* defaults to ``None`` which preserves whatever marks are already on
    disk (load → merge → save).  Pass an explicit list to overwrite.
    Returns the manifest path on success, or ``None`` on failure.
    """
    for participant_id, entry in source_transcripts.items():
        for idx, seg in enumerate(entry.get("segments", [])):
            if not seg.get("id"):
                seg["id"] = f"{participant_id}:{idx}"

    # When marks is None, preserve existing marks from disk
    if marks is None:
        existing = utils.load_json_manifest(
            config.TRANSCRIPTS_MANIFEST_FILENAME,
            default=_empty_transcripts_manifest(),
        )
        marks = existing["marks"]

    data = {
        "source_transcripts": source_transcripts,
        "corrections": corrections,
        "marks": marks,
    }
    if _is_empty_transcripts_manifest(data):
        utils.remove_json_manifest(config.TRANSCRIPTS_MANIFEST_FILENAME)
        # Invalidate the cache so the next load reflects the removal even if the
        # filesystem's mtime resolution would elide the change.
        _reset_transcripts_manifest_cache()
        return None
    result = utils.save_json_manifest(
        config.TRANSCRIPTS_MANIFEST_FILENAME,
        data,
        warn_label="transcripts manifest",
    )
    # Invalidate the cache so the next load picks up what we just wrote, even if
    # the filesystem's mtime resolution would elide the change.
    if result is not None:
        _reset_transcripts_manifest_cache()
    return result


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

    pairs = [
        (re.compile(re.escape(c["from"]), re.IGNORECASE), c["to"])
        for c in corrections
        if c.get("from") and c.get("to")
    ]
    if not pairs:
        return list(segments)

    corrected: list[TranscriptSegment] = []
    total_applied = 0
    for seg in segments:
        text = seg["text"]
        seg_applied = 0
        for pattern, to_text in pairs:
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


def _format_timestamp(seconds: float, fmt: Literal["display", "srt", "vtt"]) -> str:
    """Format a seconds value for transcript output.

    ``display`` is M:SS / H:MM:SS for Markdown. ``srt`` is HH:MM:SS,mmm.
    ``vtt`` is MM:SS.mmm or HH:MM:SS.mmm.
    """
    total = int(seconds)
    h, remainder = divmod(total, 3600)
    m, s = divmod(remainder, 60)
    if fmt == "display":
        if h > 0:
            return f"{h}:{m:02d}:{s:02d}"
        return f"{m}:{s:02d}"
    ms = int((seconds - total) * 1000)
    if fmt == "srt":
        return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"
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
        f"- **Generated:** {datetime.now(UTC).strftime('%Y-%m-%d %H:%M:%S UTC')}",
        "",
        "---",
        "",
    ]
    for seg in result["segments"]:
        start = _format_timestamp(seg["start"], "display")
        end = _format_timestamp(seg["end"], "display")
        lines.append(f"**[{start} - {end}]**")
        lines.append(seg["text"])
        lines.append("")
    return "\n".join(lines)


def _format_srt(result: TranscriptResult) -> str:
    blocks: list[str] = []
    for i, seg in enumerate(result["segments"], start=1):
        start = _format_timestamp(seg["start"], "srt")
        end = _format_timestamp(seg["end"], "srt")
        blocks.append(f"{i}\n{start} --> {end}\n{seg['text']}")
    return "\n\n".join(blocks) + "\n" if blocks else ""


def _format_vtt(result: TranscriptResult) -> str:
    lines = ["WEBVTT", ""]
    for seg in result["segments"]:
        start = _format_timestamp(seg["start"], "vtt")
        end = _format_timestamp(seg["end"], "vtt")
        lines.append(f"{start} --> {end}")
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
    return utils.timestamp_to_seconds(ts) or 0.0


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


class _TranscriptionCancelled(Exception):
    """Raised inside the on_segment callback to abort a running transcription."""


def create_transcript_task(
    participant: str,
    video_paths: list[str],
    *,
    model: str | None = None,
    language: str | None = None,
    audio_index: int | None = None,
) -> dict[str, Any]:
    """Create a new transcription task dict ready to enqueue.

    *video_paths* is the participant's ordered source video(s) — one entry for a
    normal participant, several for a multi-video participant whose session spans
    files (transcribed as one continuous timeline). *model*, *language* and
    *audio_index* are optional per-participant overrides; when None, the worker
    falls back to ``config.TRANSCRIBE_MODEL``, whisper auto-detect, and
    speech-track auto-detection respectively.
    """
    return {
        "id": f"tr_{uuid.uuid4().hex[:8]}",
        "participant": participant,
        "video_paths": video_paths,
        "model": model,
        "language": language,
        "audio_index": audio_index,
        "status": TASK_STATUS_QUEUED,
        # Sub-state of "running": "loading_model" while the Whisper model is
        # constructed (~10s cold, invisible to the progress float), then
        # "transcribing". Lets the frontend say what the 0% wait actually is.
        "phase": "queued",
        "progress": 0.0,
        "partial_segments": [],
        "result": None,
        "error": None,
        "created_at": datetime.now(UTC).isoformat(),
        "completed_at": None,
        "_cancelled": False,
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
        """Cancel a queued or running task. Returns True if cancelled."""
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return False
            if task["status"] in (TASK_STATUS_QUEUED, TASK_STATUS_CANCELLED):
                task["status"] = TASK_STATUS_CANCELLED
                return True
            if task["status"] == TASK_STATUS_RUNNING:
                task["_cancelled"] = True
                return True
        return False

    def remove_task(self, task_id: str) -> bool:
        """Cancel (if active) and fully remove a task."""
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return False
            if task["status"] == TASK_STATUS_RUNNING:
                task["_cancelled"] = True
            self._tasks.pop(task_id, None)
            return True

    def get_task(self, task_id: str) -> dict[str, Any] | None:
        """Return a task dict by ID (thread-safe copy)."""
        with self._lock:
            t = self._tasks.get(task_id)
            return copy.deepcopy(t) if t else None

    def get_all_tasks(self, include_partials: bool = True) -> list[dict[str, Any]]:
        """Return all tasks (thread-safe copies).

        ``include_partials=False`` omits each task's ever-growing
        ``partial_segments`` list (reporting ``partial_count`` instead), so the
        3 s status poll no longer deep-copies the whole segment tail on every
        tick; clients pull new segments via :meth:`get_partial_segments`.
        """
        with self._lock:
            if include_partials:
                return [copy.deepcopy(t) for t in self._tasks.values()]
            slim: list[dict[str, Any]] = []
            for t in self._tasks.values():
                segs = t.get("partial_segments") or []
                light = {k: v for k, v in t.items() if k != "partial_segments"}
                copied = copy.deepcopy(light)
                copied["partial_count"] = len(segs)
                slim.append(copied)
            return slim

    def get_partial_segments(
        self, task_id: str, since: int = 0
    ) -> tuple[list[TranscriptSegment], int]:
        """Thread-safe copy of a task's partial-segment tail from index ``since``.

        ``partial_segments`` is append-only during transcription (see
        :meth:`_on_seg`), so a count cursor yields exactly the new segments.
        Returns ``(tail, total)`` where ``total`` is the current segment count.
        """
        with self._lock:
            t = self._tasks.get(task_id)
            segs = (t.get("partial_segments") if t else None) or []
            total = len(segs)
            start = max(since, 0)
            tail = copy.deepcopy(segs[start:]) if start < total else []
            return tail, total

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

            # The worker must outlive any single task. _execute_task owns its
            # own try/except for the transcription itself, but work outside it
            # (the model preload, ffprobe/timeline setup) can still raise, and
            # there is no restart path for this thread — is_alive is reported
            # to the frontend but never acted on, so a death here wedges every
            # later transcribe request in a queue nothing drains.
            try:
                self._execute_task(task)
            except Exception as exc:
                utils.error_print(f"Transcription task failed: {exc}")
                with self._lock:
                    if task.get("status") == TASK_STATUS_RUNNING:
                        task["status"] = TASK_STATUS_FAILED
                        task["error"] = str(exc)
                        task.setdefault("partial_segments", [])
                        task["completed_at"] = datetime.now(UTC).isoformat()

            if self.on_task_complete:
                try:
                    self.on_task_complete()
                except Exception as exc:
                    utils.warning_print(f"on_task_complete callback failed: {exc}")

    def _execute_task(self, task: dict[str, Any]) -> None:
        """Run a single transcription task."""
        import video as video_mod

        video_paths = task["video_paths"]
        if not video_paths:
            with self._lock:
                task["status"] = TASK_STATUS_FAILED
                task["error"] = "No video files — nothing to transcribe."
                task["partial_segments"] = []
                task["completed_at"] = datetime.now(UTC).isoformat()
            return

        # Multi-video participants form one continuous timeline; transcribe each
        # part and merge with global-shifted times. Single video → fast path,
        # no extra duration probe beyond the audio guard below.
        timeline = video_mod.timeline_or_none(video_paths)

        # Probe the first part for the audio guard; derive the progress
        # denominator from the whole timeline for multi-video.
        props = video_mod.probe_video_properties(video_paths[0])
        tracks: list[dict[str, Any]] = (props or {}).get("audio_tracks") or []
        if props is not None and not props.get("audio_codec") and not tracks:
            with self._lock:
                task["status"] = TASK_STATUS_FAILED
                task["error"] = (
                    "No audio stream — this video has no audio track to transcribe."
                )
                task["partial_segments"] = []
                task["completed_at"] = datetime.now(UTC).isoformat()
            return

        # Resolve here rather than inside transcribe_video so the task can record
        # which track was actually used (the inner guard would only surface a
        # generic "returned None").
        audio_index = _resolve_audio_index(video_paths[0], task.get("audio_index"))
        if tracks and not 0 <= audio_index < len(tracks):
            with self._lock:
                task["status"] = TASK_STATUS_FAILED
                task["error"] = (
                    f"Audio track {audio_index + 1} does not exist — "
                    f"this video has {len(tracks)}."
                )
                task["partial_segments"] = []
                task["completed_at"] = datetime.now(UTC).isoformat()
            return

        if timeline is not None:
            duration = float(timeline[-1][1] + timeline[-1][2])
        else:
            duration = float(props.get("duration", 0.0)) if props else 0.0

        # Load corrections for context keywords
        manifest = load_transcripts_manifest()
        corrections = manifest.get("corrections", [])
        context_kw = get_corrections_keywords(corrections) or None

        # Load the model up front (a no-op when already cached) so the ~10s
        # cold construction is visible as its own phase instead of an opaque
        # "running, 0%". The DEBUGGING guard is load-bearing: debug mode
        # returns stub results without ever touching Whisper, and worker tests
        # depend on that. transcribe_video's own _load_model call then hits
        # the cache. A task cancelled while queued must not pay for a model
        # load either — transcribe_video aborts it right below.
        if not config.DEBUGGING and not task.get("_cancelled"):
            with self._lock:
                task["phase"] = "loading_model"
            # _load_model does not only return None on failure — it raises.
            # run_with_spinner calls its callback bare, and the WhisperModel
            # construction inside has a finally but no except, so a CTranslate2
            # "Library cublas64_12.dll is not found", an unsupported
            # compute_type, or a huggingface download OSError all propagate.
            # This sits outside the try below, so without this handler the
            # exception escapes _run and kills the worker thread for good.
            try:
                loaded = _load_model(task.get("model"))
            except Exception as exc:
                loaded = None
                load_error = f"Transcription model failed to load: {exc}"
            else:
                load_error = "Transcription model failed to load."
            if loaded is None:
                with self._lock:
                    task["status"] = TASK_STATUS_FAILED
                    task["error"] = load_error
                    task["partial_segments"] = []
                    task["completed_at"] = datetime.now(UTC).isoformat()
                return
        with self._lock:
            task["phase"] = "transcribing"
            # Progress/ETA baseline: excludes queue-wait and the model load.
            task["transcribe_started_at"] = datetime.now(UTC).isoformat()

        def _on_seg(end_time: float, segment: TranscriptSegment) -> None:
            # Check cancel flag outside the lock to avoid deadlock — the
            # except handler also acquires self._lock.
            if task.get("_cancelled"):
                raise _TranscriptionCancelled
            with self._lock:
                if duration > 0:
                    task["progress"] = min(end_time / duration, 0.99)
                task["partial_segments"].append(segment)

        try:
            if timeline is not None:
                result = transcribe_timeline(
                    timeline,
                    model_name=task.get("model"),
                    language=task.get("language"),
                    audio_index=audio_index,
                    context_keywords=context_kw,
                    on_segment=_on_seg,
                    cancel_flag=lambda: bool(task.get("_cancelled")),
                )
            else:
                result = transcribe_video(
                    video_paths[0],
                    model_name=task.get("model"),
                    language=task.get("language"),
                    audio_index=audio_index,
                    context_keywords=context_kw,
                    on_segment=_on_seg,
                    cancel_flag=lambda: bool(task.get("_cancelled")),
                )
            if result is None:
                with self._lock:
                    task["status"] = TASK_STATUS_FAILED
                    task["error"] = "Transcription returned None"
                    task["partial_segments"] = []
                    task["completed_at"] = datetime.now(UTC).isoformat()
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
                    # Recorded so a transcript that changed after an auto-detect
                    # deviation can be explained (surfaced as the pill's
                    # "Last run: Track 2 · Interview" hint).
                    "audio_index": audio_index,
                    "audio_track_label": (
                        tracks[audio_index].get("label", "")
                        if audio_index < len(tracks)
                        else ""
                    ),
                    "transcribed_at": datetime.now(UTC).isoformat(),
                }
                task["completed_at"] = datetime.now(UTC).isoformat()

        except _TranscriptionCancelled:
            with self._lock:
                task["status"] = TASK_STATUS_CANCELLED
                task["partial_segments"] = []
                task["completed_at"] = datetime.now(UTC).isoformat()

        except Exception as exc:
            with self._lock:
                task["status"] = TASK_STATUS_FAILED
                task["error"] = str(exc)
                task["partial_segments"] = []
                task["completed_at"] = datetime.now(UTC).isoformat()
