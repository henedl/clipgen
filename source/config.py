"""Configuration constants for clipgen.

Sections
--------
  Core Runtime         Feature toggles, version, verbosity, debugging
  Directories          Input/output path overrides
  Spreadsheet          Column headers, participant prefixes, severity maps, annotations
  Files & Durations    Clip limits, gallery, manifests, server port, workers
  Screenspace          Analysis thresholds, sampling, matching, performance
  Highlights Reel      Duration budget and scoring weights
  Browse Mode          Interactive table display settings
  Spreadsheet Commands Special command keywords for document selection
  Display / Preview    Text truncation limits
  Timestamps           Format constants for timestamp parsing
  FFmpeg               Encoding parameters, compression
  Source Video          Filename pattern for participant videos
  Transcription        faster-whisper model and format settings
  Rich Output          Terminal color/panel/progress settings
  Settings Metadata    Descriptions and Studio UI metadata
"""

import importlib
from typing import Any

# ── Core Runtime ─────────────────────────────────────────────────────
REENCODING: bool = False
AUDIO_NORMALIZE: bool = False
FILEFORMAT: str = ".mp4"
TITLECARDS_ENABLED: bool = (
    False  # use --titlecards / --no-titlecards to override per run
)
FILMSTRIP_ENABLED: bool = False  # use --filmstrip / --no-filmstrip to override per run
TITLECARD_DURATION_SECONDS: int = (
    2  # duration in seconds; falls back to color fill when no source frame available
)
# Selected card backgrounds (Studio picker). Empty = bundled default asset; the
# sentinels below select a solid-color fill or (endcard only) no card; any other
# value is an uploaded filename under <output>/TITLECARD_IMAGES_DIRNAME.
TITLECARD_IMAGE: str = ""
ENDCARD_IMAGE: str = ""
CARD_IMAGE_COLOR: str = "__color__"  # solid color fill, no background image
CARD_IMAGE_NONE: str = "__none__"  # endcard only: append no endcard
TITLECARD_IMAGES_DIRNAME: str = "titlecard_images"  # upload pool under the output dir
# Fill colors (#rrggbb) used when the corresponding card is set to a solid color.
TITLECARD_COLOR: str = "#000000"
ENDCARD_COLOR: str = "#000000"
WORKSHEET_PRIORITY: list[str] = [  # tried in order before falling back to first sheet
    "Sheet1",
    "Data",
    "data",
    "Observations",
    "Data set",
    "data set",
    "dataset",
    "Dataset",
]
DEBUGGING: bool = (
    False  # enables icecream output, skips ffmpeg execution, returns stub transcripts
)
QUIET: int = 0
STANDARD: int = 1
VERBOSE: int = 2
VERBOSITY: int = STANDARD

_ICECREAM_IC: Any | None = None


def debug_ic(*args: Any, **kwargs: Any) -> Any:
    """Lazy icecream debug printer; no import cost unless DEBUGGING is enabled."""
    if not DEBUGGING:
        return None
    global _ICECREAM_IC
    if _ICECREAM_IC is None:
        _ICECREAM_IC = importlib.import_module("icecream").ic
        _ICECREAM_IC.configureOutput(prefix="! DEBUG ic| ", includeContext=False)
    return _ICECREAM_IC(*args, **kwargs)


# ── Directories ──────────────────────────────────────────────────────
# When left empty, clipgen will use the current working directory for
# both input (source videos) and output (generated artifacts), matching
# the existing default behavior.
INPUT_DIR: str = ""
OUTPUT_DIR: str = ""

# ── Spreadsheet ──────────────────────────────────────────────────────
ID_HEADER: str = "ID"
OBSERVATION_HEADER: str = "Observation"
CATEGORY_HEADER: str = "Category"
FILENAME_HEADER: str = "Filename"
SEVERITY_HEADER: str = (
    "Severity"  # optional column; when present adds severity metadata to clips
)
PARTICIPANT_PREFIXES: tuple[str, ...] = ("P", "G")  # 'P' for individual, 'G' for group
SEVERITY_NUMERIC_TO_LABEL: dict[str, str] = {
    "-4": "Critical",
    "-3": "High",
    "-2": "Medium",
    "-1": "Low",
    "0": "N/A",
    "1": "Positive",
    "2": "Very Positive",
}
SEVERITY_LABEL_TO_NUMERIC: dict[str, int] = {
    "critical": -4,
    "high": -3,
    "medium": -2,
    "low": -1,
    "n/a": 0,
    "positive": 1,
    "very positive": 2,
}
ANNOTATION_KEYPHRASES: dict[
    str, str
] = {  # cell token → annotation name; stripped before timestamp parsing
    "!key": "key",
}
IGNORED_TIMESTAMP_TOKENS: set[str] = {
    "x"
}  # tokens skipped silently during timestamp parsing

# ── Files & Durations ────────────────────────────────────────────────
MAX_FILENAME_LENGTH: int = 255
MAX_CLIP_DURATION_SECONDS: int = (
    600  # 10 min; prompts user for confirmation before generating longer clips
)
DEFAULT_DURATION_SECONDS: int = 60  # clip length when only a start time is given
DEFAULT_GIF_DURATION_SECONDS: int = (
    5  # GIF extraction length when only a start time is given
)
GALLERY_INTERVAL_SECONDS: int = 10  # Default interval between gallery captures
GALLERY_GIF_DURATION_SECONDS: int = 3  # Default per-GIF duration in gallery mode
GALLERY_PARALLEL_WORKERS: int = (
    4  # Max concurrent ffmpeg processes for gallery GIF extraction
)
GALLERY_BUNDLE_ENABLED: bool = False  # embed images as base64 data URIs in gallery HTML
CLIP_PARALLEL_WORKERS: int = 0  # Max concurrent ffmpeg processes for clip/screenshot/GIF generation; 0 = auto (min(4, cpu_count))
MAX_FILESIZE_MB: int = 0  # Maximum output file size in MB (0 = disabled)
MIN_SOURCE_VIDEO_SIZE_MB: int = 100  # Minimum file size (MB) to consider as a source video candidate during fuzzy matching
MANIFEST_FILENAME: str = (
    "clipgen_manifest.json"  # cumulative artifact manifest; consumed by --regenerate
)
MANIFEST_ENABLED: bool = (
    False  # use --manifest CLI flag or set True to write manifest alongside artifacts
)
SERVER_PORT: int = (
    8089  # port for the combined Studio/Screenspace/Transcripts Flask server
)
# Desktop-window chrome (see desktop_chrome.py). The native title bar is hidden and
# the traffic lights float inside the topnav, so AppKit and CSS have to agree on the
# same two numbers: these flow to the frontend as CSS custom properties via
# utils.render_index_html(), rather than being written twice.
DESKTOP_CHROME_BAR_HEIGHT: int = 48  # titlebar band height; drives --topnav-height
# Left gutter the three buttons need. They are inset from the window edge by the
# same margin that centering leaves above them ((bar - 16) / 2 = 16) and pitched
# 20px apart, so the row ends at 16 + 40 + 14 = 70; the rest is breathing room
# before the brand.
DESKTOP_TRAFFIC_LIGHT_INSET: int = 87
STASHES_MANIFEST_FILENAME: str = "reel_stashes.json"
ARTIFACT_STASHES_MANIFEST_FILENAME: str = "artifact_stashes.json"
STUDIO_SETTINGS_FILENAME: str = "studio_settings.json"
WORKFLOWS_MANIFEST_FILENAME: str = (
    "workflows_manifest.json"  # node-canvas blueprints, stashes, run history
)
# Prefix for our own scratch temp-files written into the output dir (currently the
# reel-builder's mkstemp clips). Lets sweep_stale_temp_artifacts() reclaim orphans
# left by a hard kill without ever touching user files.
TEMP_ARTIFACT_PREFIX: str = "clipgen_tmp_"
# Watch-dir trigger (P6): poll interval for the daemon that auto-runs an armed
# blueprint when a new participant video lands in the input dir. The partial-copy
# stability window is 2x this (a file must stat identically across two polls).
# Server-only — never mirrored to the frontend.
WORKFLOWS_WATCH_POLL_SECONDS: float = 5.0
# Concurrent participants per Workflows whole-study batch. 1 = sequential (the
# default). Values >1 run child runs in a thread pool, multiplying peak
# ffmpeg/Whisper/OCR/Ollama load — heavy graphs (Transcribe, Detect) serialize
# on those shared resources anyway and rarely benefit past 2. The server clamps
# to [1, 4]. Server-only — never mirrored to the frontend.
WORKFLOWS_BATCH_WORKERS: int = 1
CONVERGENCE_OFFSETS_FILENAME: str = "convergence_offsets.json"
COMPOSER_MANIFEST_FILENAME: str = (
    "composer_manifest.json"  # Composer cut pairs, trims, annotations, UI state
)
# Composer annotation defaults. Geometry is normalized to the video frame
# (stroke width to frame width, font size to frame height) so the browser
# preview and the PIL/ffmpeg burn-in agree at any resolution. Mirrored to the
# frontend via utils.get_frontend_config() — do not hardcode these in JS.
COMPOSER_ANNOTATION_COLOR: str = "#f05a3c"
COMPOSER_ANNOTATION_STROKE_WIDTH: float = 0.004  # fraction of frame width
COMPOSER_ANNOTATION_STROKE_STYLE: str = "solid"  # solid | dashed | dotted
COMPOSER_ANNOTATION_FONT_SIZE: float = 0.035  # fraction of frame height
COMPOSER_ANNOTATION_SPAN_SECONDS: float = 10.0  # default visibility span
# Double-click the Composer timeline to set the pending in point, then again
# to commit the out point. Mirrored to the frontend via
# utils.get_frontend_config() — do not hardcode this in JS.
COMPOSER_DOUBLE_CLICK_CUTS: bool = True
# Warn on pages that play source video when a recording is a fragmented MP4 the
# browser cannot seek (OBS "fragmented recording"), and offer the one-click
# remux that fixes it. Mirrored to the frontend via utils.get_frontend_config()
# — do not hardcode this in JS.
MEDIA_CONTAINER_WARNING: bool = True
# Cap on the WAV extracted for Composer's marker/cut audio scrub — markers can
# span minutes (unlike Studio clips). Mirrored to the frontend via
# utils.get_frontend_config() so the client skips scrubbing longer spans and
# hover fraction ↔ audio buffer stay aligned; do not hardcode this in JS.
COMPOSER_SCRUB_MAX_AUDIO_SECONDS: float = 180.0
# Data-source lanes shown (in order) per participant in the Convergence Browser.
# Mirrored to the frontend via utils.get_frontend_config() so the swim-lane
# layout and per-lane offset keys stay in sync; do not hardcode this list in JS.
CONVERGENCE_SOURCES: tuple[str, ...] = (
    "sheet",
    "screenspace",
    "transcript",
    "composer",
)
STUDIO_CELL_EXPAND_HOVER: bool = True
GOOGLE_API_MAX_RETRIES: int = 3  # Retries for transient Google API errors (429, 5xx)

STUDIO_THUMBNAIL_WIDTH: int = 200

# Card scrubber: hover a queue card to scrub frames (sprite sheet) + hear audio
# with a waveform overlay. Opt-in. The sprite grid dims are mirrored to the
# frontend via utils.get_frontend_config() (the scrubber computes frameCount /
# per-frame interval from them), so they must not be hardcoded in JS.
STUDIO_CARD_SCRUBBER: bool = False
STUDIO_SCRUBBER_SPRITE_COLS: int = 5
STUDIO_SCRUBBER_SPRITE_ROWS: int = 5

# Metadata tab: count Screenspace as time-adjacent clusters instead of raw
# events, so a dense scan (e.g. 10k events) reads as a handful of blocks and
# doesn't overshadow the sheet/transcript streams in the tab's tables/charts.
# Boot-embedded into /api/sheet (like STUDIO_CARD_SCRUBBER); not mirrored via
# get_frontend_config.
STUDIO_METADATA_CLUSTER_SCREENSPACE: bool = True

# ── Screenspace ──────────────────────────────────────────────────────
SCREENSPACE_MANIFEST_FILENAME: str = "screenspace_manifest.json"
TRANSCRIPTS_MANIFEST_FILENAME: str = "transcripts_manifest.json"
SCREENSPACE_DEFAULT_INTERVAL: float = (
    1.0  # default frame sampling interval (seconds) for analysis tasks
)
SCREENSPACE_NOISE_THRESHOLD: int = 30  # image preprocessing tuning for change detection
SCREENSPACE_BLUR_KERNEL: int = 5  # image preprocessing tuning for change detection
SCREENSPACE_SSIM_THRESHOLD: float = 0.90  # min SSIM score for Similarity tool matches
SCREENSPACE_CHANGE_RATIO_THRESHOLD: float = (
    0.03  # min changed-pixel fraction for Change tool matches
)
SCREENSPACE_MORPH_KERNEL: int = 3  # image preprocessing tuning for change detection
SCREENSPACE_OCR_FUZZY_THRESHOLD: float = (
    0.75  # min fuzzy match score for Text/Numbers tool matches
)
SCREENSPACE_OCR_MIN_CONFIDENCE: float = (
    0.6  # min EasyOCR per-detection confidence for Text/Numbers; gates noisy OCR
)
SCREENSPACE_OCR_MIN_HEIGHT: int = (
    60  # target px height for upscaling small ROIs in opt-in OCR preprocessing
)
SCREENSPACE_OCR_GPU: bool | str = (
    True  # EasyOCR device: True = CUDA-if-available else CPU (preserves EasyOCR's default); False = force CPU; or a device string ("cuda:0", "mps")
)
SCREENSPACE_OCR_POOL_SIZE: int = 0  # max concurrent EasyOCR Readers per language set (0 = auto = SCREENSPACE_PARALLEL_WORKERS). Each Reader holds its own model copy, so raising this multiplies OCR RAM/VRAM
SCREENSPACE_MASK_FALLBACK_TOOLS: tuple[str, ...] = (
    "similarity",
    "inactivity",
    "boundary",
    "timelapse",
    "attention",
)  # tools that analyze a shaped region's bounding rect instead of its polygon (SSIM/phash are global, boundary and attention are full-frame by design, timelapse is a pure ffmpeg crop)
SCREENSPACE_PHASH_THRESHOLD: int = 15
SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD: float = 2.0  # mean-abs-diff cutoff for skipping near-identical frames (Similarity/Text/Numbers/Scene scans)
SCREENSPACE_TEMPLATE_MATCH_THRESHOLD: float = 0.70
SCREENSPACE_TEMPLATE_NMS_OVERLAP: float = 0.50
SCREENSPACE_TEMPLATE_SCALE_MIN: float = 0.25
SCREENSPACE_TEMPLATE_SCALE_MAX: float = 2.0
SCREENSPACE_FLOW_MAGNITUDE_THRESHOLD: float = 2.0
SCREENSPACE_FLOW_PYR_SCALE: float = 0.5
SCREENSPACE_SCENE_SIMILARITY_THRESHOLD: float = 0.75
SCREENSPACE_SCENE_HISTOGRAM_BINS: int = 64
SCREENSPACE_FLOW_GRID_SIZE: int = 8
SCREENSPACE_FLOW_GRID_MIN_MAG: float = 0.5
SCREENSPACE_HEATMAP_ROLLING_WINDOW: int = (
    6  # GIF buckets (of 24) inside the rolling-window heatmap's sliding window
)
SCREENSPACE_HEATMAP_SPRITE_COLS: int = (
    6  # columns in the hover-scrub sprite sheet written alongside each heatmap GIF
)
SCREENSPACE_HEATMAP_SPRITE_FRAME_WIDTH: int = (
    320  # per-frame width in that sprite sheet (downscale only; never upscales)
)
SCREENSPACE_CHANGE_HEATMAP_GRID: int = (
    16  # cells per axis for the Change heatmap's downsampled per-frame change grid
)
SCREENSPACE_CHANGE_HEATMAP_MIN_FRAC: float = (
    0.1  # min fraction of changed pixels for a change-grid cell to be recorded
)
SCREENSPACE_FAST_SCAN_INTERVAL_MULTIPLIER: float = 3.0
SCREENSPACE_FAST_SCAN_PHASH_THRESHOLD: int = 12  # tighter than general 15
SCREENSPACE_FAST_SCAN_SKIP_NONKEY: bool = True  # fast-scan: decode only keyframes (H.264/HEVC) for GOP-sized decode savings; auto-disabled per-video when the probed worst-case keyframe gap is too long, and the sample grid is tightened by that gap so coverage is never coarser than the interval
SCREENSPACE_KEYFRAME_PROBE_SECONDS: float = 20.0  # window (seconds from start) the keyframe-gap probe inspects to find the worst-case (max) GOP length
SCREENSPACE_KEYFRAME_SKIP_MARGIN: float = 1.0  # enable keyframe-only decode only when the max keyframe gap <= sampling_interval * this margin (lower = only very dense keyframes qualify)
SCREENSPACE_PARALLEL_WORKERS: int = (
    2  # max concurrent analysis tasks in ScreenspaceWorker
)
SCREENSPACE_INACTIVITY_PHASH_THRESHOLD: int = (
    10  # max Hamming distance for "same frame" in Inactivity tool
)
SCREENSPACE_INACTIVITY_MIN_DURATION: float = (
    2.0  # min seconds to report an inactivity span
)
SCREENSPACE_BOUNDARY_PHASH_THRESHOLD: int = (
    14  # min Hamming distance to call a scene boundary (above inactivity/static-skip)
)
SCREENSPACE_BOUNDARY_MIN_GAP_SECONDS: float = (
    3.0  # debounce: suppress further boundaries this long after one fires
)
SCREENSPACE_BOUNDARY_INTERVAL: float = (
    1.0  # default frame sampling interval for boundaries
)
SCREENSPACE_BOUNDARY_HASH_DIM: int = (
    64  # downscale max-dim for cheap full-frame hashing (pushed to the ffmpeg pipe)
)
SCREENSPACE_BOUNDARY_CONFIDENCE_EPSILON: float = (
    0.05  # confidence floor for a boundary that just crosses threshold
)
# Phase 4: scene-aware period segmentation. "phash" is the v1 consecutive-frame
# spike detector; "scene" measures a content fingerprint against the current
# period's reference (robust to motion); "hybrid" fires only when both agree.
SCREENSPACE_BOUNDARY_METRIC: str = "hybrid"
SCREENSPACE_BOUNDARY_SCENE_THRESHOLD: float = 0.25  # fingerprint distance (1 − similarity) to call a scene shift; mirrors scene's 0.75 sim
SCREENSPACE_BOUNDARY_CONFIRM_WINDOW: int = 2  # samples a scene shift must persist before it counts (suppresses one-frame blips)
SCREENSPACE_BOUNDARY_SCENE_HASH_DIM: int = 128  # pipe downscale for fingerprinting (coarser 64px phash dim is too sparse for the HSV histogram)
SCREENSPACE_BOUNDARY_MERGE_THRESHOLD: float = 0.15  # post-run: merge adjacent periods whose fingerprints are at least this similar (below firing threshold)
SCREENSPACE_BOUNDARY_TYPE_THRESHOLD: float = 0.35  # scene labeling: group distinct scenes this close into one "type" (Scene A1/A2); looser than merge
SCREENSPACE_BOUNDARY_SHORT_PERIOD_SECONDS: float = (
    3.0  # post-run: only periods shorter than this are transient-dissolve candidates
)
SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_ENABLED: bool = True  # post-run: drop boundaries far below the session-median strength (threshold-portability mitigation)
SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_FACTOR: float = 0.5  # prune boundaries with entry distance below this fraction of the session median
# Attention tool: classic bottom-up saliency composite (spectral residual +
# Lab contrast + frame-diff motion [+ optional Haar faces], center-biased,
# EMA-smoothed). Full-frame only; predicts where visual attention goes.
SCREENSPACE_ATTENTION_INTERVAL: float = (
    0.5  # default frame sampling interval for attention; also the heatmap's dwell unit
)
SCREENSPACE_ATTENTION_WORKING_DIM: int = (
    256  # pipe downscale max-dim for saliency math (pushed to the ffmpeg pipe)
)
SCREENSPACE_ATTENTION_GRID: int = (
    16  # cells per axis for the per-frame saliency_grid (mirrors CHANGE_HEATMAP_GRID)
)
SCREENSPACE_ATTENTION_GRID_MIN_MAG: float = (
    0.15  # min normalized saliency for a grid cell to be recorded
)
SCREENSPACE_ATTENTION_SHIFT_THRESHOLD: float = (
    0.15  # normalized distance the attention peak must jump to count as a shift
)
SCREENSPACE_ATTENTION_SHIFT_CONFIRM: int = (
    2  # consecutive samples a jumped peak must persist before a shift event fires
)
SCREENSPACE_ATTENTION_EMA_ALPHA: float = (
    0.6  # temporal smoothing of the saliency map (1.0 = no smoothing)
)
SCREENSPACE_ATTENTION_WEIGHT_SPECTRAL: float = 1.0  # spectral-residual channel weight
SCREENSPACE_ATTENTION_WEIGHT_CONTRAST: float = (
    0.7  # Lab center-surround contrast weight
)
SCREENSPACE_ATTENTION_WEIGHT_MOTION: float = (
    1.2  # frame-diff motion weight (motion dominates attention on screens)
)
SCREENSPACE_ATTENTION_WEIGHT_FACE: float = (
    0.8  # face-blob channel weight (when enabled)
)
SCREENSPACE_ATTENTION_FACE_CHANNEL: bool = False  # opt-in Haar face channel: false-positives on UI avatars/icons; slowest channel
SCREENSPACE_ATTENTION_CENTER_BIAS: float = 0.25  # strength of the center-weighted prior (screens are less center-biased than photos)
SCREENSPACE_CV_RESOLUTION_SCALE: float = (
    1.0  # multiplier applied to extracted region frames before CV analysis
    # (1.0 = no change; >1 sharper but slower and more memory; <1 faster but coarser)
)
SCREENSPACE_RESTORE_MARKERS_ON_EDIT: bool = (
    True  # restore In/Out timeline markers when editing a task
)
SCREENSPACE_SHOW_CONFIDENCE_HISTOGRAM: bool = False  # show the per-detection confidence-distribution histogram in the Results panel
SCREENSPACE_GROUPED_TOOL_NAV: bool = (
    True  # group the analysis tools into category dropdowns instead of a flat tab row
)
SCREENSPACE_GENERATE_TEMPLATE_HEATMAP: bool = (
    True  # generate detection heatmaps for Template tasks
)
SCREENSPACE_GENERATE_FLOW_HEATMAP: bool = (
    True  # generate motion heatmaps for Flow tasks
)
SCREENSPACE_GENERATE_CHANGE_HEATMAP: bool = (
    True  # generate change heatmaps for Change tasks
)
SCREENSPACE_GENERATE_ATTENTION_HEATMAP: bool = (
    True  # generate attention heatmaps for Attention tasks
)
SCREENSPACE_MAX_PINS: int = 12  # soft cap on calibration pins per participant — keeps synchronous calibration interactive
SCREENSPACE_MULTITOOL_MAX_OFFSET_SECONDS: float = 30.0  # bound (±) for a multitool step's offset window relative to the previous step's matched frame

# ── Highlights Reel ──────────────────────────────────────────────────
HIGHLIGHTS_REEL_DURATION_SECONDS: int = 180  # 3-minute budget for highlights reel
HIGHLIGHTS_WEIGHT_SEVERITY: float = (
    1.0  # scoring weight for severity in highlights reel ranking
)
HIGHLIGHTS_WEIGHT_UNIQUENESS: float = 0.5  # scoring weight for participant uniqueness
HIGHLIGHTS_WEIGHT_KEYWORD: float = 0.3  # scoring weight for !key annotation

# ── Browse Mode ──────────────────────────────────────────────────────
BROWSE_LINES_TO_DISPLAY: int = 5  # Number of rows to show at once when browsing
BROWSE_LINES_TO_SCROLL: int = 5  # Number of rows to move when scrolling up/down
BROWSE_DESCRIPTION_MAX_WIDTH: int = 40  # Max width for description column in table
BROWSE_TIMESTAMP_MAX_WIDTH: int = 15  # Max width for each timestamp column

# ── Spreadsheet Commands ─────────────────────────────────────────────
COMMAND_LIST_ALL: str = "all"
COMMAND_LIST_NEW: str = "new"
COMMAND_OPEN_LAST: str = "last"
COMMAND_SETTINGS: str = "settings"
COMMAND_HTTP_PREFIX: str = "http"
COMMAND_EXCEL: str = "excel"
NUM_NEWEST_DOCS_TO_SHOW: int = (
    3  # Number of newest documents to show when using 'new' command
)

# ── Display / Preview ────────────────────────────────────────────────
DESCRIPTION_PREVIEW_LENGTH: int = 50  # Max chars for description in previews
PROGRESS_DESCRIPTION_LENGTH: int = 30  # Max chars for description in progress bar
REEL_PREVIEW_CLIP_COUNT: int = 10  # Number of clips to show in reel mode preview
MAX_SKIPPED_TIMESTAMPS_TO_SHOW: int = (
    3  # Max skipped timestamps to list in parse_timestamps warning
)

# ── Timestamps ───────────────────────────────────────────────────────
SECONDS_PER_HOUR: int = 3600
SECONDS_PER_MINUTE: int = 60

# ── FFmpeg ────────────────────────────────────────────────────────────
FFMPEG_LOGLEVEL: str = "16"  # ffmpeg -loglevel value (16 = error)
# H.264 encoder for the re-encode paths (reel concat, clip re-encode, timelapse,
# Composer burn). "auto" uses Apple's VideoToolbox hardware encoder when ffmpeg
# lists it and it hasn't failed this session, else libx264. compress_to_size is
# deliberately excluded — see its docstring.
# See video.resolve_video_encoder / video.video_encoder_args.
FFMPEG_VIDEO_ENCODER: str = "auto"  # "auto" | "libx264" | "h264_videotoolbox"
FFMPEG_SCREENSHOT_QUALITY: str = "2"  # -q:v value for screenshots (1=best, 31=worst)
# x264 settings for title/endcard generation and the card-wrap re-encode. The
# wrap stream-copies a copy-safe clip body (titlecards._body_is_copy_safe) and
# re-encodes only the ~2s cards, so the preset dominates titlecard time only on
# the fallback re-encode path (non-h264/exotic bodies, or a failed copy concat);
# "veryfast" is several times quicker than libx264's "medium" default at
# negligible quality cost for short research clips. Raise quality with a lower
# CRF, or trade quality for speed with "superfast"/"ultrafast".
TITLECARD_ENCODE_PRESET: str = "veryfast"
TITLECARD_ENCODE_CRF: int = 20
SCREENSHOT_FORMAT: str = ".png"  # ".png" | ".jpg" | ".webp"
GIF_FORMAT: str = ".gif"  # ".gif" | ".webp" | ".webm" (WebM uses VP9 silent-loop video)
WEBP_QUALITY: int = 80  # 0-100, used by libwebp when output is .webp
GIF_FPS: int = 10
GIF_SCALE_WIDTH: int = 480
AUDIO_BITRATE_KBPS: int = 128
COMPRESSION_SIZE_FACTOR: float = 0.95  # Target 95% of max to leave headroom
MIN_VIDEO_BITRATE_KBPS: int = 100

# ── Source Video ──────────────────────────────────────────────────────
# Matches source-video filenames so discover_clips() can exclude them from the
# generated-clip list. The optional ``-N`` group matches numbered parts of a
# multi-video participant (e.g. study_P01-1.mp4, study_P01-2.mp4).
SOURCE_VIDEO_PATTERN: str = r"_[PG]\d+(-\d+)?\.mp4$"
# Trailing numbered suffix used to auto-detect a participant's source-video parts
# on disk (one continuous timeline): study_P01-1.mp4, study_P01-2.mp4. The
# capture group is the integer order; parts are sorted numerically, not lexically.
NUMBERED_SOURCE_VIDEO_SUFFIX_PATTERN: str = r"-(\d+)\.mp4$"

# ── Transcription ────────────────────────────────────────────────────
TRANSCRIBE_ENABLED: bool = False  # use --transcribe CLI flag to enable per run
TRANSCRIBE_MODEL: str = "base"  # tiny, base, small, medium, large-v3
TRANSCRIBE_LANGUAGE: str | None = None  # None = auto-detect
TRANSCRIBE_COMPUTE_TYPE: str = "int8"  # int8 (fastest), float16, float32
# Compute device for the Whisper model: "auto", "cpu", or "cuda". See
# transcripts._resolve_transcribe_device() for why "auto" means CPU in a frozen
# build — the desktop bundle ships no CUDA runtime, so letting CTranslate2
# auto-select a GPU only produces a "cublas64_12.dll is not found" crash.
TRANSCRIBE_DEVICE: str = "auto"
TRANSCRIBE_FORMAT: str = "md"  # md, srt, vtt
TRANSCRIBE_INITIAL_PROMPT: str = "This is a recorded user experience research session."  # biases Whisper toward UX research terminology
TRANSCRIBE_BEAM_SIZE: int = (
    2  # beam search width (1 = greedy; higher = slower, marginally better)
)
# CTranslate2 CPU threads for the Whisper model; 0 = auto (os.cpu_count()).
TRANSCRIBE_CPU_THREADS: int = 0
TRANSCRIBE_VAD_FILTER: bool = True  # Silero VAD: transcribe speech spans only
# VAD tuning (only applied when TRANSCRIBE_VAD_FILTER is on). Defaults are chosen
# recall-safe: a lower-than-Silero-default threshold plus boundary padding so quiet
# speech and word onsets/offsets aren't clipped.
TRANSCRIBE_VAD_THRESHOLD: float = 0.2  # Silero speech-probability cutoff (lower = more permissive; 0.3 missed quiet speech in testing)
TRANSCRIBE_VAD_SPEECH_PAD_MS: int = (
    400  # padding added to each speech span so word edges aren't clipped
)
TRANSCRIBE_VAD_MIN_SILENCE_MS: int = 2000  # minimum silence gap before splitting a span
TRANSCRIBE_NO_SPEECH_THRESHOLD: float = (
    0.6  # drop segments with high no-speech probability
)
TRANSCRIBE_LOG_PROB_THRESHOLD: float = -1.0  # drop low-confidence segments
TRANSCRIBE_COMPRESSION_RATIO_THRESHOLD: float = 2.4  # drop repetitive / looped text
# Seconds of surrounding silence for hallucination skip logic; 0 = off (requires word_timestamps when > 0)
TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD: float = 0.0
TRANSCRIBE_CONDITION_ON_PREVIOUS_TEXT: bool = (
    True  # False reduces chained hallucinations
)
# When to pre-load faster-whisper in the Transcripts web UI: off, queue_open (open Queue panel), page_load (after participants load).
TRANSCRIBE_PREWARM: str = "queue_open"
# Mark categories shown in the Transcripts mark popover. Each value is {label, color}.
# "friction" is a single bucket for all friction-detection marks; the specific
# friction type lives in each mark's label (e.g. "Friction · frustration").
MARK_CATEGORIES: dict[str, dict[str, str]] = {
    "pain_point": {"label": "Pain Point", "color": "#dc2626"},
    "delight": {"label": "Delight", "color": "#16a34a"},
    "quote": {"label": "Quote", "color": "#2563eb"},
    "insight": {"label": "Insight", "color": "#f97316"},
    "task": {"label": "Task Issue", "color": "#8b5cf6"},
    "bookmark": {"label": "Bookmark", "color": "#0891b2"},
    "friction": {"label": "Friction", "color": "#ea580c"},
}

# ── Hotkeys ─────────────────────────────────────────────────────────
# User overrides for web-frontend keyboard shortcuts, keyed by action id
# (the catalog of ids and default combos lives in assets/web/hotkeys.js).
# Values are space-separated combo strings ("Mod+Shift+Z Mod+Y"); an empty
# string disables the shortcut. Edited via Settings → Hotkeys; the server
# only persists and structurally validates this dict, never interprets it.
HOTKEY_OVERRIDES: dict[str, str] = {}

# ── Ollama (Local AI) ───────────────────────────────────────────────
OLLAMA_SUMMARY_ENABLED: bool = (
    True  # auto-generate transcript summaries via Ollama after transcription completes
)
OLLAMA_CITATIONS_ENABLED: bool = (
    True  # auto-generate citation links via Ollama after the summary completes
)
OLLAMA_FRICTION_ENABLED: bool = (
    False  # auto-detect friction moments via Ollama after the summary completes
)
OLLAMA_SUMMARY_MODEL: str = (
    "qwen3.5:9b"  # model for transcript summaries, citations, and friction
)
OLLAMA_FRICTION_MODEL: str = (
    ""  # friction agent model; blank → use OLLAMA_SUMMARY_MODEL, set to override
)
OLLAMA_REPORT_ENABLED: bool = False  # mini-report is manual-only (Overview → Summary); True adds it to the auto-chain
OLLAMA_REPORT_MODEL: str = (
    ""  # report agent model; blank → use OLLAMA_SUMMARY_MODEL, set to override
)
OLLAMA_BASE_URL: str = "http://localhost:11434"  # Ollama server address
OLLAMA_UNLOAD_DELAY_SECONDS: float = 15.0  # after Stop, evict the model from memory if no new run starts within this delay

# ── Thinking-agent prompts ───────────────────────────────────────────
# Editable via Settings → Summaries → "Agent prompts". thinking_agents.py reads
# these at call time, so an edit takes effect on the next agent run. The user
# prompts are .format()-ed with the placeholders noted below; the *_SYSTEM
# prompts are sent verbatim (never formatted), so braces in them are literal.
OLLAMA_SUMMARY_PROMPT: str = """\
Summarize this user research session transcript. Write a concise paragraph \
(2-4 sentences) describing what happened in the session. Then list the key \
topics or themes as bullet points (prefix each with "- ").

Transcript:
{text}"""
OLLAMA_CITATIONS_SYSTEM: str = (
    "You match transcript segments to summary claims. "
    "For each claim, select only the 1-3 most relevant and representative "
    "segments. Prefer segments that most clearly and directly support the claim. "
    "Use the exact format shown."
)
OLLAMA_CITATIONS_PROMPT: str = """\
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
OLLAMA_FRICTION_SYSTEM: str = (
    "You analyze UX research session transcripts for moments of friction: "
    "points where the participant struggled, hesitated, got confused, or showed "
    "frustration. You respond with a JSON array only."
)
OLLAMA_FRICTION_PROMPT: str = """\
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
OLLAMA_REPORT_SYSTEM: str = (
    "You are a UX research assistant writing a short per-participant session "
    "report. Ground every statement in the provided data. Never invent quotes, "
    "observations, or timestamps; if the data does not support a claim, leave "
    "it out."
)
OLLAMA_REPORT_PROMPT: str = """\
Write a concise research mini-report for participant {participant}, using only \
the data below.

Session summary:
{summary}

Sheet observations (researcher notes, "- text (category, severity)"):
{observations}

Marked transcript moments ("[M:SS] (Category) text"):
{bookmarks}

Structure the report as short markdown sections:
## Overview
2-3 sentences on what happened in the session.
## Key findings
3-5 bullet points, most important first.
## What went well
Positives: things that worked smoothly, delights, favorable reactions. Omit \
this section entirely if the data shows none.
## Pain points & friction
Bullet points grounded in the observations and marked moments above; include \
the timestamp when one is given.
## Notable moments
Quotes or delight moments worth sharing, with timestamps.

Keep the whole report under 300 words. If a section has no supporting data, \
write "No data." under it instead of inventing content."""

# ── Friction detection ───────────────────────────────────────────────
# Ordered category keys → display labels. Single source of truth, mirrored to
# the frontend via utils.get_frontend_config(). The programmatic scorer in
# friction.py keys its patterns/weights by the same keys (enforced by tests).
FRICTION_CATEGORIES: dict[str, str] = {
    "hesitation": "Hesitation",
    "confusion": "Confusion",
    "frustration": "Frustration",
    "surprise": "Surprise",
    "self_correction": "Self-correction",
    "help_seeking": "Help-seeking",
}
FRICTION_CANDIDATE_LIMIT: int = 15  # top-N scored segments fed to the LLM stage
FRICTION_MOMENT_LIMIT: int = 5  # number of moments the LLM returns

# ── Rich Output ──────────────────────────────────────────────────────
RICH_COLORS: bool = True  # Enable/disable colored output (set False for piped output)
RICH_PANELS: bool = True  # Use bordered panels for errors/warnings/success messages
RICH_PROGRESS: bool = True  # Show progress bars during batch/reel processing

# ── Settings Metadata ────────────────────────────────────────────────
SETTINGS_DESCRIPTIONS: dict[str, str] = {
    "REENCODING": "Re-encode clips via ffmpeg instead of stream-copying. Slower but fixes some codec issues.",
    "AUDIO_NORMALIZE": "Normalize audio levels across generated clips for consistent volume.",
    "FILEFORMAT": "Output container format for generated video clips.",
    "FFMPEG_VIDEO_ENCODER": "H.264 encoder used whenever clipgen re-encodes (reel concat, clip re-encode, timelapse, Composer burn-in). auto picks Apple's VideoToolbox hardware encoder on Macs that have it — several times faster, at the cost of somewhat larger files for the same visual quality — and falls back to libx264 if it is missing or fails. Pick libx264 to always encode in software. Size-capped compression (Max filesize) always uses libx264 regardless: hardware encoders cannot hit a bitrate target accurately.",
    "MAX_FILESIZE_MB": "Compress output to stay under this size limit. Set to 0 to disable.",
    "DEFAULT_DURATION_SECONDS": "Clip length when only a start time is provided.",
    "MAX_CLIP_DURATION_SECONDS": "Prompt for confirmation before generating clips longer than this.",
    "DEBUGGING": "Enable debug output via icecream and skip ffmpeg execution.",
    "VERBOSITY": "Verbosity level: 0=quiet (CLI default), 1=standard (interactive default), 2=verbose (most detail).",
    "RICH_COLORS": "Use colored terminal output via Rich library.",
    "RICH_PANELS": "Show bordered panels for error, warning, and success messages.",
    "RICH_PROGRESS": "Display progress bars during batch and reel processing.",
    "TITLECARDS_ENABLED": "Prepend a generated titlecard to each video clip.",
    "TITLECARD_DURATION_SECONDS": "Duration in seconds for the intro titlecard frame.",
    "TITLECARD_IMAGE": "Background for the intro titlecard. Pick the default, a solid color, or upload your own image. The clip description is overlaid as text.",
    "ENDCARD_IMAGE": "Background for the outro endcard. Pick the default, no endcard, a solid color, or upload your own image.",
    "TITLECARD_COLOR": "Fill color used when the titlecard is set to a solid color.",
    "ENDCARD_COLOR": "Fill color used when the endcard is set to a solid color.",
    "TRANSCRIBE_ENABLED": "Generate transcripts alongside clips using faster-whisper.",
    "TRANSCRIBE_MODEL": "Whisper model size: tiny, base, small, medium, large-v3.",
    "TRANSCRIBE_FORMAT": "Transcript output format: md (Markdown), srt, or vtt.",
    "TRANSCRIBE_PREWARM": "When the Transcripts page pre-loads the Whisper model: off, queue_open (opening a pill's options pane or hovering a pill that needs transcription), or page_load (after listing participants).",
    "TRANSCRIBE_BEAM_SIZE": "Beam-search width. 1 = greedy (fastest); higher is slower for marginally better accuracy.",
    "TRANSCRIBE_CPU_THREADS": "CTranslate2 CPU threads for the Whisper model. 0 = auto (all cores).",
    "TRANSCRIBE_DEVICE": "Compute device for the Whisper model. auto = GPU only where the CUDA runtime is actually available (always CPU in the desktop app, which ships no CUDA). Set cuda only if you installed a matching cuBLAS/cuDNN yourself.",
    "TRANSCRIBE_VAD_FILTER": "Use Silero VAD to transcribe speech spans only, skipping long silence (reduces silence hallucinations).",
    "TRANSCRIBE_VAD_THRESHOLD": "Silero speech-probability cutoff when VAD is on. Lower = more permissive (catches quieter speech, fewer dropped words).",
    "TRANSCRIBE_VAD_SPEECH_PAD_MS": "Padding (ms) added to each detected speech span so word onsets/offsets aren't clipped by VAD.",
    "TRANSCRIBE_VAD_MIN_SILENCE_MS": "Minimum silence gap (ms) before VAD splits a speech span; higher avoids over-fragmenting.",
    "TRANSCRIBE_NO_SPEECH_THRESHOLD": "Drop segments when no-speech probability exceeds this value (higher = stricter).",
    "TRANSCRIBE_LOG_PROB_THRESHOLD": "Drop segments below this average log-probability (higher = stricter).",
    "TRANSCRIBE_COMPRESSION_RATIO_THRESHOLD": "Drop segments whose gzip compression ratio exceeds this (catches repetitive loops).",
    "TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD": "Seconds of surrounding silence for hallucination skip logic; 0 = off. When > 0, enables word-level timestamps (slower).",
    "TRANSCRIBE_CONDITION_ON_PREVIOUS_TEXT": "Use prior segment text as context for the next decode; disable to reduce chained hallucinations.",
    "MARK_CATEGORIES": "Categories available when marking transcript segments. Each entry has a label and a color swatch.",
    "HOTKEY_OVERRIDES": "Custom keyboard-shortcut bindings, keyed by action id. Click a shortcut to rebind it; an empty value disables the shortcut.",
    "HIGHLIGHTS_REEL_DURATION_SECONDS": "Maximum duration in seconds for the highlights reel time budget.",
    "MANIFEST_ENABLED": "Write a manifest JSON file alongside generated artifacts for session tracking.",
    "STUDIO_CELL_EXPAND_HOVER": "Expand overflowing timestamp cells on hover in the Sheet Preview.",
    "STUDIO_CARD_SCRUBBER": "Hover a queue card's thumbnail to scrub through frames, hear the clip's audio, and see a waveform overlay.",
    "COMPOSER_DOUBLE_CLICK_CUTS": "Double-click the Composer timeline to set the in point, then double-click again to commit the out point.",
    "MEDIA_CONTAINER_WARNING": "Warn when a source recording is a fragmented MP4 that browsers cannot seek (OBS 'fragmented recording'), and offer a one-click remux to fix it.",
    "STUDIO_METADATA_CLUSTER_SCREENSPACE": "In the Metadata tab, count Screenspace data as time-adjacent clusters instead of raw events, so a dense scan doesn't overshadow the spreadsheet and transcript streams. On by default.",
    "FILMSTRIP_ENABLED": "Show thumbnail images on timeline markers instead of solid colors (in the HTML viewer).",
    "GALLERY_BUNDLE_ENABLED": "Embed gallery images as base64 data URIs in the HTML file, making it fully self-contained.",
    "CLIP_PARALLEL_WORKERS": "Number of concurrent ffmpeg processes for clip generation. 0 = auto, 1 = sequential.",
    "OLLAMA_SUMMARY_ENABLED": "Auto-generate an AI summary of each transcript after transcription completes. Disable to keep summaries manual-only (the per-participant Regenerate Summary button still works).",
    "OLLAMA_CITATIONS_ENABLED": "Auto-generate citation links between summary claims and transcript segments after the summary completes. Disable to keep citations manual-only (the per-participant Regenerate Citations button still works).",
    "OLLAMA_FRICTION_ENABLED": "Auto-detect friction moments after the summary completes. Disable to keep friction manual-only (the per-participant Run/Re-run friction button still works). Uses the AI summary model.",
    "OLLAMA_SUMMARY_MODEL": "Ollama model used for transcript summaries, citation linking, and friction detection.",
    "OLLAMA_FRICTION_MODEL": "Ollama model for friction-moment detection. Leave as 'Same as summary model' to reuse the summary model, or pick a different installed model.",
    "OLLAMA_BASE_URL": "Base URL of the local Ollama server.",
    "OLLAMA_SUMMARY_PROMPT": "Prompt that generates each session summary. Keep the {text} placeholder. The transcript is inserted there.",
    "OLLAMA_CITATIONS_SYSTEM": "System instruction that frames the citation agent's behavior. Sent verbatim; no placeholders.",
    "OLLAMA_CITATIONS_PROMPT": "Prompt that links summary claims to transcript segments. Keep the {claims} and {transcript} placeholders.",
    "OLLAMA_FRICTION_SYSTEM": "System instruction that frames the friction agent's behavior. Sent verbatim; no placeholders.",
    "OLLAMA_FRICTION_PROMPT": "Prompt that detects friction moments. Keep the {summary}, {segments}, and {limit} placeholders.",
    "OLLAMA_REPORT_ENABLED": "Auto-generate a per-participant mini-report after the summary completes. Off by default: generate reports from the Overview page's Reports tab instead.",
    "OLLAMA_REPORT_MODEL": "Ollama model for mini-report generation. Leave as 'Same as summary model' to reuse the summary model, or pick a different installed model.",
    "OLLAMA_REPORT_SYSTEM": "System instruction that frames the report agent's behavior. Sent verbatim; no placeholders.",
    "OLLAMA_REPORT_PROMPT": "Prompt that writes the per-participant mini-report. Keep the {participant}, {summary}, {observations}, and {bookmarks} placeholders.",
    "SCREENSHOT_FORMAT": "File format for screenshot artifacts. WebP is smaller but requires modern browsers (Safari 16+).",
    "GIF_FORMAT": "File format for animated artifacts. WebM (VP9) is the smallest and most-compatible modern option; animated WebP is also small but requires Safari 16+; GIF works everywhere but is large.",
    "WEBP_QUALITY": "WebP encoding quality (0-100). Higher values mean better quality and larger files.",
    "SCREENSPACE_CV_RESOLUTION_SCALE": "Scale extracted region frames before CV analysis. Higher (e.g. 2.0) gives the models more signal on noisy/compressed video at the cost of speed and memory; lower speeds up scans on large footage. 1.0 = unchanged.",
    "SCREENSPACE_FAST_SCAN_SKIP_NONKEY": "Fast scans decode only keyframes on H.264/HEVC when the source's keyframe interval is short enough that no samples are lost (auto-probed per video), giving large decode savings. Turn off to always full-decode. Only affects fast scans; the precise scan path is never changed.",
    "SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD": "Skip frames whose average pixel difference from the previous sampled frame is below this value (Similarity/Text/Numbers/Scene scans). Lower = process more frames (catch subtle changes); higher = skip more aggressively on noisy footage. Default 2.0.",
    "SCREENSPACE_OCR_MIN_CONFIDENCE": "Default minimum EasyOCR per-detection confidence for Text/Numbers tasks. Raise to suppress noisy OCR misreads; lower if real hits are being dropped. Per-task slider overrides this default.",
    "SCREENSPACE_RESTORE_MARKERS_ON_EDIT": "When editing a task, restore the In/Out timeline markers to the range it was originally run with. Disable to keep your current markers in place when iterating across different parts of the timeline.",
    "SCREENSPACE_SHOW_CONFIDENCE_HISTOGRAM": "Show a confidence-distribution histogram above the Results list (for tools that have confidence scores). Lets you see where detections cluster before moving the certainty cutoff. Off by default.",
    "SCREENSPACE_GROUPED_TOOL_NAV": "Group the analysis tools into category dropdowns (Difference, Detection, Classification, Attention, Utility) with a standalone Multitool chip, instead of a flat row of tool tabs. Easier to scan when picking a tool. On by default; turn off for the classic flat tab row.",
    "SCREENSPACE_GENERATE_TEMPLATE_HEATMAP": "Generate detection heatmaps (static image plus accumulation and rolling-window animations) for Template tasks. Disable to skip heatmap generation when you don't need it. Useful on long videos where it adds processing time.",
    "SCREENSPACE_GENERATE_FLOW_HEATMAP": "Generate motion heatmaps (static image plus accumulation animation) for Flow tasks. Disable to skip heatmap generation when you don't need it.",
    "SCREENSPACE_GENERATE_CHANGE_HEATMAP": "Generate change heatmaps (static image plus accumulation and rolling-window animations) for Change tasks. Disable to skip heatmap generation when you don't need it. Useful on long videos where it adds processing time.",
    "SCREENSPACE_GENERATE_ATTENTION_HEATMAP": "Generate attention heatmaps (static image plus accumulation and rolling-window animations) for Attention tasks. The rolling-window animation is the closest analog to an eye-tracking gaze replay. Disable to skip heatmap generation when you don't need it.",
    "SCREENSPACE_BOUNDARY_MERGE_THRESHOLD": "Boundary post-processing (Scene/Hybrid metrics): merge two periods whose content is at least this similar, removing the boundary between them. Higher = merge more aggressively (fewer boundaries); lower = keep more. Below the firing sensitivity by design.",
    "SCREENSPACE_BOUNDARY_TYPE_THRESHOLD": "Boundary scene labeling (Scene/Hybrid metrics): distinct scenes closer than this are grouped into one 'type', labeled Scene A1, A2, … (same letter, different number). Looser than the merge threshold. Higher = group more scenes per type; lower = more distinct type letters.",
    "SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_ENABLED": "Boundary post-processing (Scene/Hybrid metrics): after merging, drop boundaries far weaker than the session's typical scene change. Adapts to each recording instead of a fixed threshold; disable to keep every detected boundary.",
}

# Studio-exposed settings with UI metadata (tab, group, type, constraints).
STUDIO_SETTINGS: dict[str, dict[str, Any]] = {
    "MANIFEST_ENABLED": {"tab": "General", "group": "Manifest", "type": "bool"},
    "CLIP_PARALLEL_WORKERS": {
        "tab": "General",
        "group": "Workers",
        "type": "int",
        "min": 0,
        "step": 1,
    },
    "STUDIO_CELL_EXPAND_HOVER": {
        "tab": "General",
        "group": "Sheet Preview",
        "type": "bool",
    },
    "STUDIO_CARD_SCRUBBER": {
        "tab": "General",
        "group": "Sheet Preview",
        "type": "bool",
    },
    "STUDIO_METADATA_CLUSTER_SCREENSPACE": {
        "tab": "General",
        "group": "Metadata Overview",
        "type": "bool",
    },
    "SCREENSHOT_FORMAT": {
        "tab": "General",
        "group": "Image Output",
        "type": "select",
        "options": [".png", ".jpg", ".webp"],
    },
    "GIF_FORMAT": {
        "tab": "General",
        "group": "Image Output",
        "type": "select",
        "options": [".gif", ".webp", ".webm"],
    },
    "WEBP_QUALITY": {
        "tab": "General",
        "group": "Image Output",
        "type": "int",
        "min": 0,
        "max": 100,
        "step": 1,
    },
    "REENCODING": {"tab": "Video & Clips", "group": "Video Output", "type": "bool"},
    "AUDIO_NORMALIZE": {
        "tab": "Video & Clips",
        "group": "Video Output",
        "type": "bool",
    },
    "FILEFORMAT": {
        "tab": "Video & Clips",
        "group": "Video Output",
        "type": "select",
        "options": [".mp4", ".webm", ".mkv"],
    },
    "FFMPEG_VIDEO_ENCODER": {
        "tab": "Video & Clips",
        "group": "Video Output",
        "type": "select",
        "options": ["auto", "libx264", "h264_videotoolbox"],
    },
    "MAX_FILESIZE_MB": {
        "tab": "Video & Clips",
        "group": "Video Output",
        "type": "int",
        "min": 0,
        "step": 1,
    },
    "DEFAULT_DURATION_SECONDS": {
        "tab": "Video & Clips",
        "group": "Clip Behavior",
        "type": "int",
        "min": 1,
        "step": 1,
    },
    "MAX_CLIP_DURATION_SECONDS": {
        "tab": "Video & Clips",
        "group": "Clip Behavior",
        "type": "int",
        "min": 1,
        "step": 1,
    },
    "TITLECARDS_ENABLED": {
        "tab": "Video & Clips",
        "group": "Titlecards",
        "type": "bool",
    },
    "TITLECARD_DURATION_SECONDS": {
        "tab": "Video & Clips",
        "group": "Titlecards",
        "type": "int",
        "min": 1,
        "step": 1,
    },
    "TITLECARD_IMAGE": {
        "tab": "Video & Clips",
        "group": "Titlecards",
        "type": "card_picker",
        "kind": "title",
    },
    "ENDCARD_IMAGE": {
        "tab": "Video & Clips",
        "group": "Titlecards",
        "type": "card_picker",
        "kind": "end",
    },
    # Persisted + sent to the frontend, but not rendered as their own rows —
    # the card_picker widget edits them via its inline color box.
    "TITLECARD_COLOR": {
        "tab": "Video & Clips",
        "group": "Titlecards",
        "type": "hidden",
    },
    "ENDCARD_COLOR": {
        "tab": "Video & Clips",
        "group": "Titlecards",
        "type": "hidden",
    },
    "HIGHLIGHTS_REEL_DURATION_SECONDS": {
        "tab": "Video & Clips",
        "group": "Highlights",
        "type": "int",
        "min": 10,
        "step": 10,
    },
    "TRANSCRIBE_ENABLED": {
        "tab": "Transcription",
        "group": "Transcription",
        "type": "bool",
    },
    "TRANSCRIBE_MODEL": {
        "tab": "Transcription",
        "group": "Transcription",
        "type": "model_select",
        "provider": "whisper",
    },
    "TRANSCRIBE_FORMAT": {
        "tab": "Transcription",
        "group": "Transcription",
        "type": "select",
        "options": ["md", "srt", "vtt"],
    },
    "TRANSCRIBE_PREWARM": {
        "tab": "Transcription",
        "group": "Transcription",
        "type": "select",
        "options": ["off", "queue_open", "page_load"],
    },
    "TRANSCRIBE_CPU_THREADS": {
        "tab": "Transcription",
        "group": "Transcription",
        "type": "int",
        "min": 0,
        "max": 64,
        "step": 1,
    },
    "TRANSCRIBE_DEVICE": {
        "tab": "Transcription",
        "group": "Transcription",
        "type": "select",
        "options": ["auto", "cpu", "cuda"],
    },
    "TRANSCRIBE_BEAM_SIZE": {
        "tab": "Transcription",
        "group": "Transcription quality",
        "type": "int",
        "min": 1,
        "max": 5,
        "step": 1,
    },
    "TRANSCRIBE_VAD_FILTER": {
        "tab": "Transcription",
        "group": "Transcription quality",
        "type": "bool",
    },
    "TRANSCRIBE_VAD_THRESHOLD": {
        "tab": "Transcription",
        "group": "Transcription quality",
        "type": "float",
        "min": 0.0,
        "max": 1.0,
        "step": 0.05,
    },
    "TRANSCRIBE_VAD_SPEECH_PAD_MS": {
        "tab": "Transcription",
        "group": "Transcription quality",
        "type": "int",
        "min": 0,
        "max": 1000,
        "step": 50,
    },
    "TRANSCRIBE_VAD_MIN_SILENCE_MS": {
        "tab": "Transcription",
        "group": "Transcription quality",
        "type": "int",
        "min": 0,
        "max": 10000,
        "step": 250,
    },
    "TRANSCRIBE_NO_SPEECH_THRESHOLD": {
        "tab": "Transcription",
        "group": "Transcription quality",
        "type": "float",
        "min": 0.0,
        "max": 1.0,
        "step": 0.05,
    },
    "TRANSCRIBE_LOG_PROB_THRESHOLD": {
        "tab": "Transcription",
        "group": "Transcription quality",
        "type": "float",
        "min": -2.0,
        "max": 0.0,
        "step": 0.1,
    },
    "TRANSCRIBE_COMPRESSION_RATIO_THRESHOLD": {
        "tab": "Transcription",
        "group": "Transcription quality",
        "type": "float",
        "min": 1.0,
        "max": 10.0,
        "step": 0.1,
    },
    "TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD": {
        "tab": "Transcription",
        "group": "Transcription quality",
        "type": "float",
        "min": 0.0,
        "max": 10.0,
        "step": 0.5,
    },
    "TRANSCRIBE_CONDITION_ON_PREVIOUS_TEXT": {
        "tab": "Transcription",
        "group": "Transcription quality",
        "type": "bool",
    },
    "MARK_CATEGORIES": {
        "tab": "Transcription",
        "group": "Markers",
        "type": "mark_categories",
    },
    "HOTKEY_OVERRIDES": {
        "tab": "Hotkeys",
        "group": "",
        "type": "hotkeys",
    },
    "OLLAMA_SUMMARY_ENABLED": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "bool",
    },
    "OLLAMA_CITATIONS_ENABLED": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "bool",
    },
    "OLLAMA_FRICTION_ENABLED": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "bool",
    },
    "OLLAMA_REPORT_ENABLED": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "bool",
    },
    "OLLAMA_SUMMARY_MODEL": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "model_select",
        "provider": "ollama",
    },
    "OLLAMA_FRICTION_MODEL": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "model_select",
        "provider": "ollama",
        # Blank value inherits OLLAMA_SUMMARY_MODEL; surfaced as this option.
        "emptyLabel": "Same as summary model",
    },
    "OLLAMA_REPORT_MODEL": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "model_select",
        "provider": "ollama",
        # Blank value inherits OLLAMA_SUMMARY_MODEL; surfaced as this option.
        "emptyLabel": "Same as summary model",
    },
    "OLLAMA_BASE_URL": {"tab": "Summaries", "group": "AI Summary", "type": "str"},
    "OLLAMA_SUMMARY_PROMPT": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": ["text"],
    },
    "OLLAMA_CITATIONS_SYSTEM": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": [],
    },
    "OLLAMA_CITATIONS_PROMPT": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": ["claims", "transcript"],
    },
    "OLLAMA_FRICTION_SYSTEM": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": [],
    },
    "OLLAMA_FRICTION_PROMPT": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": ["summary", "segments", "limit"],
    },
    "OLLAMA_REPORT_SYSTEM": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": [],
    },
    "OLLAMA_REPORT_PROMPT": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": ["participant", "summary", "observations", "bookmarks"],
    },
    "SCREENSPACE_CV_RESOLUTION_SCALE": {
        "tab": "Screenspace",
        "group": "Analysis Quality",
        "type": "float",
        "min": 0.25,
        "max": 4.0,
        "step": 0.25,
    },
    "SCREENSPACE_FAST_SCAN_SKIP_NONKEY": {
        "tab": "Screenspace",
        "group": "Analysis Quality",
        "type": "bool",
    },
    "SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD": {
        "tab": "Screenspace",
        "group": "Analysis Quality",
        "type": "float",
        "min": 0.5,
        "max": 10.0,
        "step": 0.5,
    },
    "SCREENSPACE_RESTORE_MARKERS_ON_EDIT": {
        "tab": "Screenspace",
        "group": "Task Editing",
        "type": "bool",
    },
    "SCREENSPACE_SHOW_CONFIDENCE_HISTOGRAM": {
        "tab": "Screenspace",
        "group": "Results Display",
        "type": "bool",
    },
    "SCREENSPACE_GROUPED_TOOL_NAV": {
        "tab": "Screenspace",
        "group": "Tool Selection",
        "type": "bool",
    },
    "SCREENSPACE_GENERATE_TEMPLATE_HEATMAP": {
        "tab": "Screenspace",
        "group": "Heatmaps",
        "type": "bool",
    },
    "SCREENSPACE_GENERATE_FLOW_HEATMAP": {
        "tab": "Screenspace",
        "group": "Heatmaps",
        "type": "bool",
    },
    "SCREENSPACE_GENERATE_CHANGE_HEATMAP": {
        "tab": "Screenspace",
        "group": "Heatmaps",
        "type": "bool",
    },
    "SCREENSPACE_GENERATE_ATTENTION_HEATMAP": {
        "tab": "Screenspace",
        "group": "Heatmaps",
        "type": "bool",
    },
    "SCREENSPACE_BOUNDARY_MERGE_THRESHOLD": {
        "tab": "Screenspace",
        "group": "Boundaries",
        "type": "float",
        "min": 0.0,
        "max": 1.0,
        "step": 0.01,
    },
    "SCREENSPACE_BOUNDARY_TYPE_THRESHOLD": {
        "tab": "Screenspace",
        "group": "Boundaries",
        "type": "float",
        "min": 0.0,
        "max": 1.0,
        "step": 0.01,
    },
    "SCREENSPACE_BOUNDARY_RELATIVE_PRUNE_ENABLED": {
        "tab": "Screenspace",
        "group": "Boundaries",
        "type": "bool",
    },
    "COMPOSER_DOUBLE_CLICK_CUTS": {
        "tab": "Composer",
        "group": "Timeline",
        "type": "bool",
    },
    "MEDIA_CONTAINER_WARNING": {
        "tab": "General",
        "group": "Media",
        "type": "bool",
    },
    "RICH_COLORS": {"tab": "CLI", "group": "Terminal Output", "type": "bool"},
    "RICH_PANELS": {"tab": "CLI", "group": "Terminal Output", "type": "bool"},
    "RICH_PROGRESS": {"tab": "CLI", "group": "Terminal Output", "type": "bool"},
}
