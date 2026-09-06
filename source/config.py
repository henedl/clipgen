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
# Project home, shown in /api/status and credited in every exported viewer.
REPO_URL: str = "https://github.com/henedl/clipgen"
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
# Empty = bundled default; sentinels below = solid fill or no card; else uploaded filename.
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
PROFILING: bool = False  # opt-in perf instrumentation (--profile); never changes behavior, only measures
PROFILE_DEEP: str = (
    ""  # --profile-deep LABEL: cProfile spans whose label contains this substring
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
# Empty means the current working directory.
INPUT_DIR: str = ""
OUTPUT_DIR: str = ""

# ── Spreadsheet ──────────────────────────────────────────────────────
ID_HEADER: str = "ID"
OBSERVATION_HEADER: str = "Observation"
CATEGORY_HEADER: str = "Category"
FILENAME_HEADER: str = "Filename"
# Runtime state, not a setting: per-participant filename overrides seeded from start.json;
# beat the Filename row.
FILENAME_OVERRIDES: dict[str, str] = {}
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
MANIFEST_FILENAME: str = "clipgen.json"  # one sectioned state file per output dir; see utils.load_manifest_section
MANIFEST_ENABLED: bool = (
    False  # use --manifest CLI flag or set True to write manifest alongside artifacts
)
SERVER_PORT: int = (
    8089  # port for the combined Studio/Screenspace/Transcripts Flask server
)
# Desktop chrome (desktop_chrome.py): AppKit and CSS share these via render_index_html() custom properties.
DESKTOP_CHROME_BAR_HEIGHT: int = 48  # titlebar band height; drives --topnav-height
# Traffic lights span 16 + 40 + 14 = 70px; remainder is gap before brand.
DESKTOP_TRAFFIC_LIGHT_INSET: int = 87
# Lives in the per-user config dir beside start.json: preferences, not project data.
STUDIO_SETTINGS_FILENAME: str = "studio_settings.json"
# Prefix for scratch temp files so sweep_stale_temp_artifacts() reclaims orphans without touching user files.
TEMP_ARTIFACT_PREFIX: str = "clipgen_tmp_"
# New-video trigger poll interval; a file must stat identically across two polls. Server-only.
WORKFLOWS_WATCH_POLL_SECONDS: float = 5.0
# Parallel participants per whole-study batch; clamped to [1, 4]. Rarely helps past 2. Server-only.
WORKFLOWS_BATCH_WORKERS: int = 1
# Composer annotation defaults, frame-normalized so preview and PIL burn-in agree. Mirrored to JS.
COMPOSER_ANNOTATION_COLOR: str = "#f05a3c"
# Inactive half of the primary/secondary swap (X); never written onto an annotation record.
COMPOSER_ANNOTATION_COLOR_SECONDARY: str = "#f8fafc"
COMPOSER_ANNOTATION_STROKE_WIDTH: float = 0.004  # fraction of frame width
COMPOSER_ANNOTATION_STROKE_STYLE: str = "solid"  # solid | dashed | dotted
COMPOSER_ANNOTATION_FONT_SIZE: float = 0.035  # fraction of frame height
COMPOSER_ANNOTATION_SPAN_SECONDS: float = 10.0  # default visibility span
# Timeline double-click sets the in point, then the out point. Mirrored to JS.
COMPOSER_DOUBLE_CLICK_CUTS: bool = True
# Warn about fragmented MP4s the browser cannot seek and offer the remux. Mirrored to JS.
MEDIA_CONTAINER_WARNING: bool = True
# Desktop app asks GitHub Releases for a newer build once per launch. Server-only.
UPDATE_CHECK_ON_LAUNCH: bool = True
# Cap on Composer's scrub WAV; the client skips longer spans. Mirrored to JS.
COMPOSER_SCRUB_MAX_AUDIO_SECONDS: float = 180.0
# Convergence Browser lanes, in order; mirrored to JS so lane layout and offset keys agree.
CONVERGENCE_SOURCES: tuple[str, ...] = (
    "sheet",
    "screenspace",
    "transcript",
    "composer",
)
STUDIO_CELL_EXPAND_HOVER: bool = True
GOOGLE_API_MAX_RETRIES: int = 3  # Retries for transient Google API errors (429, 5xx)

STUDIO_THUMBNAIL_WIDTH: int = 200

# Opt-in hover-to-scrub on queue cards; JS derives frameCount and interval from the grid dims.
STUDIO_CARD_SCRUBBER: bool = False
STUDIO_SCRUBBER_SPRITE_COLS: int = 5
STUDIO_SCRUBBER_SPRITE_ROWS: int = 5

# Cross-reference badges on Studio, Transcripts, and Overview.
# Mirrored to JS via get_frontend_config() (never hardcode).
CROSS_REFERENCES_ENABLED: bool = True

# Metadata tab counts Screenspace clusters, not raw events. Boot-embedded into
# /api/sheet, not get_frontend_config.
STUDIO_METADATA_CLUSTER_SCREENSPACE: bool = True

# ── Screenspace ──────────────────────────────────────────────────────
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
SCREENSPACE_OCR_MIN_CONFIDENCE: float = 0.6  # min OCR confidence for Text/Numbers; calibrated on easyocr, unchecked against RapidOCR
SCREENSPACE_OCR_MIN_HEIGHT: int = (
    60  # target px height for upscaling small ROIs in opt-in OCR preprocessing
)
SCREENSPACE_OCR_POOL_SIZE: int = 0  # RapidOCR engines per rec model (0 = SCREENSPACE_PARALLEL_WORKERS); each holds ONNX sessions, so RAM scales
SCREENSPACE_MASK_FALLBACK_TOOLS: tuple[str, ...] = (
    "similarity",
    "inactivity",
    "boundary",
    "timelapse",
    "attention",
)  # tools that use a shaped region's bounding rect, not its polygon
SCREENSPACE_PHASH_THRESHOLD: int = 15
SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD: float = 2.0  # mean-abs-diff cutoff for skipping near-identical frames (Similarity/Text/Numbers/Scene scans)
SCREENSPACE_TEMPLATE_MATCH_THRESHOLD: float = 0.70
SCREENSPACE_TEMPLATE_NMS_OVERLAP: float = 0.50
SCREENSPACE_TEMPLATE_SCALE_MIN: float = 0.25
SCREENSPACE_TEMPLATE_SCALE_MAX: float = 2.0
SCREENSPACE_SHAPE_MATCH_THRESHOLD: float = (
    0.55  # edge correlations peak lower than intensity correlations
)
SCREENSPACE_SHAPE_SCALE_MIN: float = 0.5
SCREENSPACE_SHAPE_SCALE_MAX: float = 2.0
SCREENSPACE_SHAPE_SCALE_STEPS: int = 7
SCREENSPACE_EDGE_CANNY_LOW: int = 100
SCREENSPACE_EDGE_CANNY_HIGH: int = 200
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
SCREENSPACE_FAST_SCAN_SKIP_NONKEY: bool = True  # fast-scan decodes keyframes only; auto-off per video when the probed keyframe gap is too long
SCREENSPACE_KEYFRAME_PROBE_SECONDS: float = 20.0  # window (seconds from start) the keyframe-gap probe inspects to find the worst-case (max) GOP length
SCREENSPACE_KEYFRAME_SKIP_MARGIN: float = 1.0  # keyframe-only decode needs max gap <= interval * this; lower = denser keyframes required
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
# phash: consecutive-frame spike; scene: fingerprint vs period reference (motion-robust);
# hybrid: both must agree.
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
# Saliency composite (spectral residual, Lab contrast, motion, optional Haar faces),
# center-biased, EMA-smoothed. Full-frame only.
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
SCREENSPACE_GENERATE_SHAPE_HEATMAP: bool = (
    True  # generate detection heatmaps for Shape tasks
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
# "auto" = VideoToolbox when listed and unfailed this session, else libx264.
# See video.resolve_video_encoder.
FFMPEG_VIDEO_ENCODER: str = "auto"  # "auto" | "libx264" | "h264_videotoolbox"
FFMPEG_SCREENSHOT_QUALITY: str = "2"  # -q:v value for screenshots (1=best, 31=worst)
# x264 settings for the ~2s cards; copy-safe bodies stream-copy
# (titlecards._body_is_copy_safe). Lower CRF raises quality.
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
# Stem template; see utils.format_source_video_stem() and utils.compile_source_video_regex().
# {study} is optional.
SOURCE_FILENAME_PATTERN: str = "{study}_{participant}"

# ── Transcription ────────────────────────────────────────────────────
TRANSCRIBE_ENABLED: bool = False  # use --transcribe CLI flag to enable per run
TRANSCRIBE_MODEL: str = "base"  # tiny, base, small, medium, large-v3
TRANSCRIBE_LANGUAGE: str | None = None  # None = auto-detect
TRANSCRIBE_COMPUTE_TYPE: str = "int8"  # int8 (fastest), float16, float32
# Device: auto, cpu, cuda. Frozen builds ship no CUDA, so auto means CPU
# (transcripts._resolve_transcribe_device).
TRANSCRIBE_DEVICE: str = "auto"
TRANSCRIBE_FORMAT: str = "md"  # md, srt, vtt
TRANSCRIBE_INITIAL_PROMPT: str = "This is a recorded user experience research session."  # biases Whisper toward UX research terminology
TRANSCRIBE_BEAM_SIZE: int = (
    2  # beam search width (1 = greedy; higher = slower, marginally better)
)
# CTranslate2 CPU threads for the Whisper model; 0 = auto (os.cpu_count()).
TRANSCRIBE_CPU_THREADS: int = 0
TRANSCRIBE_VAD_FILTER: bool = True  # Silero VAD: transcribe speech spans only
# VAD tuning, applied only when TRANSCRIBE_VAD_FILTER is on; recall-safe defaults keep quiet speech.
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
# Surrounding silence for hallucination skip; 0 = off, > 0 requires word_timestamps
TRANSCRIBE_HALLUCINATION_SILENCE_THRESHOLD: float = 0.0
# DTW per-word timing: tightens segment bounds to first/last word and powers word highlight; slower.
TRANSCRIBE_WORD_TIMESTAMPS: bool = True
# Snap segment edges to speech energy, trimming Whisper's silence overshoot. One envelope pass per file.
TRANSCRIBE_EDGE_SNAP: bool = True
TRANSCRIBE_CONDITION_ON_PREVIOUS_TEXT: bool = (
    True  # False reduces chained hallucinations
)
# Pre-load faster-whisper in Transcripts UI: off | queue_open | page_load.
TRANSCRIBE_PREWARM: str = "queue_open"
# Mark popover categories, {label, color}. "friction" is one bucket; type lives in each label.
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
# Per-action overrides (ids in assets/web/hotkeys.js); space-separated combos, empty disables.
# Server only validates.
HOTKEY_OVERRIDES: dict[str, str] = {}

# ── Local AI (llama.cpp) ────────────────────────────────────────────
LLM_SUMMARY_ENABLED: bool = (
    True  # auto-generate transcript summaries after transcription completes
)
LLM_CITATIONS_ENABLED: bool = (
    True  # auto-generate citation links after the summary completes
)
LLM_FRICTION_ENABLED: bool = (
    False  # auto-detect friction moments after the summary completes
)
LLM_SUMMARY_MODEL: str = (
    # HF ref (user/repo:QUANT) or a local GGUF stem; see llm_client.model_name().
    "unsloth/Qwen3.5-9B-GGUF:Q4_K_M"  # model for summaries, citations, and friction
)
LLM_FRICTION_MODEL: str = (
    ""  # friction agent model; blank → use LLM_SUMMARY_MODEL, set to override
)
LLM_REPORT_ENABLED: bool = False  # mini-report is manual-only (Overview → Summary); True adds it to the auto-chain
LLM_REPORT_MODEL: str = (
    ""  # report agent model; blank → use LLM_SUMMARY_MODEL, set to override
)
LLM_BASE_URL: str = "http://127.0.0.1:8790"  # llama-server router address
LLM_UNLOAD_DELAY_SECONDS: float = 15.0  # after Stop, evict the model from memory if no new run starts within this delay

# ── Thinking-agent prompts ───────────────────────────────────────────
# thinking_agents.py reads these per run. User prompts .format(); *_SYSTEM sent verbatim.
LLM_SUMMARY_PROMPT: str = """\
Summarize this user research session transcript. Write a concise paragraph \
(2-4 sentences) describing what happened in the session. Then list the key \
topics or themes as bullet points (prefix each with "- ").

Transcript:
{text}"""
LLM_CITATIONS_SYSTEM: str = (
    "You match transcript segments to summary claims. "
    "For each claim, select only the 1-3 most relevant and representative "
    "segments. Prefer segments that most clearly and directly support the claim. "
    "Use the exact format shown."
)
LLM_CITATIONS_PROMPT: str = """\
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
LLM_FRICTION_SYSTEM: str = (
    "You analyze UX research session transcripts for moments of friction: "
    "points where the participant struggled, hesitated, got confused, or showed "
    "frustration. You respond with a JSON array only."
)
LLM_FRICTION_PROMPT: str = """\
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
LLM_REPORT_SYSTEM: str = (
    "You are a UX research assistant writing a short per-participant session "
    "report. Ground every statement in the provided data. Never invent quotes, "
    "observations, or timestamps; if the data does not support a claim, leave "
    "it out."
)
LLM_REPORT_PROMPT: str = """\
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
# Keys → labels, mirrored to JS; friction.py patterns share the keys.
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
    "SOURCE_FILENAME_PATTERN": "Filename pattern for source session videos in the input folder. Placeholders: {study}, {participant} (required). Extension comes from the video file format; multi-part recordings append -1, -2, … Example: {study}_{participant} matches mystudy_P01.mp4.",
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
    "TRANSCRIBE_WORD_TIMESTAMPS": "Per-word timing via alignment. Tightens segment boundaries to the spoken words and enables the word-level playback highlight; adds decode time (roughly 10–30%).",
    "TRANSCRIBE_EDGE_SNAP": "Snap segment boundaries to measured speech energy in the decoded audio, trimming silence overshoot at segment edges. Effectively free.",
    "MARK_CATEGORIES": "Categories available when marking transcript segments. Each entry has a label and a color swatch.",
    "HOTKEY_OVERRIDES": "Custom keyboard-shortcut bindings, keyed by action id. Click a shortcut to rebind it; an empty value disables the shortcut.",
    "HIGHLIGHTS_REEL_DURATION_SECONDS": "Maximum duration in seconds for the highlights reel time budget.",
    "MANIFEST_ENABLED": "Write a manifest JSON file alongside generated artifacts for session tracking.",
    "UPDATE_CHECK_ON_LAUNCH": "Check GitHub for a newer clipgen release when the desktop app starts. You can always check manually from the Start panel's About tab.",
    "STUDIO_CELL_EXPAND_HOVER": "Expand overflowing timestamp cells on hover in the Sheet Preview.",
    "STUDIO_CARD_SCRUBBER": "Hover a queue card's thumbnail to scrub through frames, hear the clip's audio, and see a waveform overlay.",
    "CROSS_REFERENCES_ENABLED": "Show cross-reference badges linking spreadsheet, Screenspace, transcript, and Composer data across pages.",
    "COMPOSER_DOUBLE_CLICK_CUTS": "Double-click the Composer timeline to set the in point, then double-click again to commit the out point.",
    "MEDIA_CONTAINER_WARNING": "Warn when a source recording is a fragmented MP4 that browsers cannot seek (OBS 'fragmented recording'), and offer a one-click remux to fix it.",
    "STUDIO_METADATA_CLUSTER_SCREENSPACE": "In the Metadata tab, count Screenspace data as time-adjacent clusters instead of raw events, so a dense scan doesn't overshadow the spreadsheet and transcript streams. On by default.",
    "FILMSTRIP_ENABLED": "Show thumbnail images on timeline markers instead of solid colors (in the HTML viewer).",
    "GALLERY_BUNDLE_ENABLED": "Embed gallery images as base64 data URIs in the HTML file, making it fully self-contained.",
    "CLIP_PARALLEL_WORKERS": "Number of concurrent ffmpeg processes for clip generation. 0 = auto, 1 = sequential.",
    "LLM_SUMMARY_ENABLED": "Auto-generate an AI summary of each transcript after transcription completes. Disable to keep summaries manual-only (the per-participant Regenerate Summary button still works).",
    "LLM_CITATIONS_ENABLED": "Auto-generate citation links between summary claims and transcript segments after the summary completes. Disable to keep citations manual-only (the per-participant Regenerate Citations button still works).",
    "LLM_FRICTION_ENABLED": "Auto-detect friction moments after the summary completes. Disable to keep friction manual-only (the per-participant Run/Re-run friction button still works). Uses the AI summary model.",
    "LLM_SUMMARY_MODEL": "AI model for transcript summaries, citation linking, and friction detection. A Hugging Face ref (user/repo:QUANT) or a downloaded model name. Models already in the llama.cpp cache, HF cache, or Ollama are reused automatically.",
    "LLM_FRICTION_MODEL": "AI model for friction-moment detection. Leave as 'Same as summary model' to reuse the summary model, or pick a different installed model.",
    "LLM_BASE_URL": "Base URL of the local llama-server router.",
    "LLM_SUMMARY_PROMPT": "Prompt that generates each session summary. Keep the {text} placeholder. The transcript is inserted there.",
    "LLM_CITATIONS_SYSTEM": "System instruction that frames the citation agent's behavior. Sent verbatim; no placeholders.",
    "LLM_CITATIONS_PROMPT": "Prompt that links summary claims to transcript segments. Keep the {claims} and {transcript} placeholders.",
    "LLM_FRICTION_SYSTEM": "System instruction that frames the friction agent's behavior. Sent verbatim; no placeholders.",
    "LLM_FRICTION_PROMPT": "Prompt that detects friction moments. Keep the {summary}, {segments}, and {limit} placeholders.",
    "LLM_REPORT_ENABLED": "Auto-generate a per-participant mini-report after the summary completes. Off by default: generate reports from the Overview page's Reports tab instead.",
    "LLM_REPORT_MODEL": "AI model for mini-report generation. Leave as 'Same as summary model' to reuse the summary model, or pick a different installed model.",
    "LLM_REPORT_SYSTEM": "System instruction that frames the report agent's behavior. Sent verbatim; no placeholders.",
    "LLM_REPORT_PROMPT": "Prompt that writes the per-participant mini-report. Keep the {participant}, {summary}, {observations}, and {bookmarks} placeholders.",
    "SCREENSHOT_FORMAT": "File format for screenshot artifacts. WebP is smaller but requires modern browsers (Safari 16+).",
    "GIF_FORMAT": "File format for animated artifacts. WebM (VP9) is the smallest and most-compatible modern option; animated WebP is also small but requires Safari 16+; GIF works everywhere but is large.",
    "WEBP_QUALITY": "WebP encoding quality (0-100). Higher values mean better quality and larger files.",
    "SCREENSPACE_CV_RESOLUTION_SCALE": "Scale extracted region frames before CV analysis. Higher (e.g. 2.0) gives the models more signal on noisy/compressed video at the cost of speed and memory; lower speeds up scans on large footage. 1.0 = unchanged.",
    "SCREENSPACE_FAST_SCAN_SKIP_NONKEY": "Fast scans decode only keyframes on H.264/HEVC when the source's keyframe interval is short enough that no samples are lost (auto-probed per video), giving large decode savings. Turn off to always full-decode. Only affects fast scans; the precise scan path is never changed.",
    "SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD": "Skip frames whose average pixel difference from the previous sampled frame is below this value (Similarity/Text/Numbers/Scene scans). Lower = process more frames (catch subtle changes); higher = skip more aggressively on noisy footage. Default 2.0.",
    "SCREENSPACE_OCR_MIN_CONFIDENCE": "Default minimum OCR per-detection confidence for Text/Numbers tasks. Raise to suppress noisy OCR misreads; lower if real hits are being dropped. Per-task slider overrides this default.",
    "SCREENSPACE_RESTORE_MARKERS_ON_EDIT": "When editing a task, restore the In/Out timeline markers to the range it was originally run with. Disable to keep your current markers in place when iterating across different parts of the timeline.",
    "SCREENSPACE_SHOW_CONFIDENCE_HISTOGRAM": "Show a confidence-distribution histogram above the Results list (for tools that have confidence scores). Lets you see where detections cluster before moving the certainty cutoff. Off by default.",
    "SCREENSPACE_GROUPED_TOOL_NAV": "Group the analysis tools into category dropdowns (Difference, Detection, Classification, Attention, Utility) with a standalone Multitool chip, instead of a flat row of tool tabs. Easier to scan when picking a tool. On by default; turn off for the classic flat tab row.",
    "SCREENSPACE_GENERATE_TEMPLATE_HEATMAP": "Generate detection heatmaps (static image plus accumulation and rolling-window animations) for Template tasks. Disable to skip heatmap generation when you don't need it. Useful on long videos where it adds processing time.",
    "SCREENSPACE_GENERATE_SHAPE_HEATMAP": "Generate detection heatmaps (static image plus accumulation and rolling-window animations) for Shape tasks. Disable to skip heatmap generation when you don't need it. Useful on long videos where it adds processing time.",
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
    "UPDATE_CHECK_ON_LAUNCH": {"tab": "General", "group": "Updates", "type": "bool"},
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
    "CROSS_REFERENCES_ENABLED": {
        "tab": "General",
        "group": "Cross-References",
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
    "SOURCE_FILENAME_PATTERN": {
        "tab": "Video & Clips",
        "group": "Source Videos",
        "type": "str",
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
    # Persisted and sent to the frontend; edited inline by card_picker, not as rows.
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
    "TRANSCRIBE_WORD_TIMESTAMPS": {
        "tab": "Transcription",
        "group": "Transcription quality",
        "type": "bool",
    },
    "TRANSCRIBE_EDGE_SNAP": {
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
    "LLM_SUMMARY_ENABLED": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "bool",
    },
    "LLM_CITATIONS_ENABLED": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "bool",
    },
    "LLM_FRICTION_ENABLED": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "bool",
    },
    "LLM_REPORT_ENABLED": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "bool",
    },
    "LLM_SUMMARY_MODEL": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "model_select",
        "provider": "llm",
    },
    "LLM_FRICTION_MODEL": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "model_select",
        "provider": "llm",
        # Blank value inherits LLM_SUMMARY_MODEL; surfaced as this option.
        "emptyLabel": "Same as summary model",
    },
    "LLM_REPORT_MODEL": {
        "tab": "Summaries",
        "group": "AI Summary",
        "type": "model_select",
        "provider": "llm",
        # Blank value inherits LLM_SUMMARY_MODEL; surfaced as this option.
        "emptyLabel": "Same as summary model",
    },
    "LLM_BASE_URL": {"tab": "Summaries", "group": "AI Summary", "type": "str"},
    "LLM_SUMMARY_PROMPT": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": ["text"],
    },
    "LLM_CITATIONS_SYSTEM": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": [],
    },
    "LLM_CITATIONS_PROMPT": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": ["claims", "transcript"],
    },
    "LLM_FRICTION_SYSTEM": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": [],
    },
    "LLM_FRICTION_PROMPT": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": ["summary", "segments", "limit"],
    },
    "LLM_REPORT_SYSTEM": {
        "tab": "Summaries",
        "group": "Agent prompts",
        "type": "prompt",
        "placeholders": [],
    },
    "LLM_REPORT_PROMPT": {
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
    "SCREENSPACE_GENERATE_SHAPE_HEATMAP": {
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
