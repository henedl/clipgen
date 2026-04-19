# -*- coding: utf-8 -*-
"""Configuration constants for clipgen.

Sections
--------
  Core Runtime         Feature toggles, version, verbosity, debugging
  Directories          Input/output path overrides
  Spreadsheet          Column headers, participant prefixes, severity maps, annotations
  Files & Durations    Clip limits, gallery, manifests, server port, workers
  Sprite Sheets        Insights builder hover-to-scrub thumbnails
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

from typing import Any

from icecream import ic

# ── Core Runtime ─────────────────────────────────────────────────────
REENCODING: bool = False
AUDIO_NORMALIZE: bool = False
FILEFORMAT: str = ".mp4"
VERSIONNUM: str = "0.10.68"
TITLECARDS_ENABLED: bool = (
    False  # use --titlecards / --no-titlecards to override per run
)
FILMSTRIP_ENABLED: bool = False  # use --filmstrip / --no-filmstrip to override per run
TITLECARD_DURATION_SECONDS: int = (
    2  # duration in seconds; falls back to color fill when no source frame available
)
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

# ── Directories ──────────────────────────────────────────────────────
# When left empty, clipgen will use the current working directory for
# both input (source videos) and output (generated artifacts), matching
# the existing default behavior.
INPUT_DIR: str = ""
OUTPUT_DIR: str = ""

# Configure Icecream debugging
if DEBUGGING:
    ic.configureOutput(prefix="! DEBUG ic| ", includeContext=False)
else:
    ic.disable()

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
MANIFEST_FILENAME: str = "clipgen_manifest.json"  # cumulative artifact manifest; consumed by Insights and --regenerate
MANIFEST_ENABLED: bool = (
    False  # use --manifest CLI flag or set True to write manifest alongside artifacts
)
SERVER_PORT: int = (
    8089  # port for the combined Studio/Insights/Screenspace Flask server
)
INSIGHTS_MANIFEST_FILENAME: str = (
    "insights_manifest.json"  # insights data manifest; read/written by Insights
)
STASHES_MANIFEST_FILENAME: str = "reel_stashes.json"
ARTIFACT_STASHES_MANIFEST_FILENAME: str = "artifact_stashes.json"
STUDIO_SETTINGS_FILENAME: str = "studio_settings.json"
STUDIO_CELL_EXPAND_HOVER: bool = True
GOOGLE_API_MAX_RETRIES: int = 3  # Retries for transient Google API errors (5xx)

# ── Sprite Sheets ────────────────────────────────────────────────────
SPRITE_SHEET_FRAME_COUNT: int = 20
SPRITE_SHEET_THUMB_WIDTH: int = 160
SPRITE_SHEET_MIN_INTERVAL: int = 1
STUDIO_THUMBNAIL_WIDTH: int = 200

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
    0.8  # min fuzzy match score for Text/Numbers tool matches
)
SCREENSPACE_PHASH_THRESHOLD: int = 15
SCREENSPACE_SEQUENTIAL_READ_MAX_INTERVAL: float = 3.0
SCREENSPACE_INTAKE_CLUSTER_SECONDS: int = 5
SCREENSPACE_TEMPLATE_MATCH_THRESHOLD: float = 0.70
SCREENSPACE_TEMPLATE_NMS_OVERLAP: float = 0.50
SCREENSPACE_TEMPLATE_SCALE_MIN: float = 0.25
SCREENSPACE_TEMPLATE_SCALE_MAX: float = 2.0
SCREENSPACE_TEMPLATE_SCALE_STEP: float = 0.05
SCREENSPACE_FLOW_MAGNITUDE_THRESHOLD: float = 2.0
SCREENSPACE_FLOW_PYR_SCALE: float = 0.5
SCREENSPACE_SCENE_SIMILARITY_THRESHOLD: float = 0.75
SCREENSPACE_SCENE_HISTOGRAM_BINS: int = 64
SCREENSPACE_FLOW_GRID_SIZE: int = 8
SCREENSPACE_FLOW_GRID_MIN_MAG: float = 0.5
SCREENSPACE_FAST_SCAN_INTERVAL_MULTIPLIER: float = 3.0
SCREENSPACE_FAST_SCAN_PHASH_THRESHOLD: int = 12  # tighter than general 15
SCREENSPACE_BATCH_EXTRACT: bool = (
    True  # use ffmpeg pipe for frame extraction; falls back to cv2
)
SCREENSPACE_PARALLEL_WORKERS: int = (
    2  # max concurrent analysis tasks in ScreenspaceWorker
)
SCREENSPACE_INACTIVITY_PHASH_THRESHOLD: int = (
    10  # max Hamming distance for "same frame" in Inactivity tool
)
SCREENSPACE_INACTIVITY_MIN_DURATION: float = (
    2.0  # min seconds to report an inactivity span
)

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
MAX_MMSS_LENGTH: int = 5  # Max length of an MM:SS timestamp string

# ── FFmpeg ────────────────────────────────────────────────────────────
FFMPEG_LOGLEVEL: str = "16"  # ffmpeg -loglevel value (16 = error)
FFMPEG_SCREENSHOT_QUALITY: str = "2"  # -q:v value for screenshots (1=best, 31=worst)
GIF_FPS: int = 10
GIF_SCALE_WIDTH: int = 480
AUDIO_BITRATE_KBPS: int = 128
COMPRESSION_SIZE_FACTOR: float = 0.95  # Target 95% of max to leave headroom
MIN_VIDEO_BITRATE_KBPS: int = 100

# ── Source Video ──────────────────────────────────────────────────────
SOURCE_VIDEO_PATTERN: str = r"_[PG]\d+\.mp4$"

# ── Transcription ────────────────────────────────────────────────────
TRANSCRIBE_ENABLED: bool = False  # use --transcribe CLI flag to enable per run
TRANSCRIBE_MODEL: str = "base"  # tiny, base, small, medium, large-v3
TRANSCRIBE_LANGUAGE: str | None = None  # None = auto-detect
TRANSCRIBE_COMPUTE_TYPE: str = "int8"  # int8 (fastest), float16, float32
TRANSCRIBE_FORMAT: str = "md"  # md, srt, vtt
TRANSCRIBE_INITIAL_PROMPT: str = "This is a recorded user experience research session."  # biases Whisper toward UX research terminology
TRANSCRIBE_BEAM_SIZE: int = 5  # beam search width
# When to pre-load faster-whisper in the Transcripts web UI: off, queue_open (open Queue panel), page_load (after participants load).
TRANSCRIBE_PREWARM: str = "queue_open"

# ── Ollama (Local AI) ───────────────────────────────────────────────
OLLAMA_SUMMARY_ENABLED: bool = True  # auto-generate transcript summaries via Ollama
OLLAMA_SUMMARY_MODEL: str = "qwen3.5:9b"  # model for transcript summaries and citations
OLLAMA_BASE_URL: str = "http://localhost:11434"  # Ollama server address

# ── Rich Output ──────────────────────────────────────────────────────
RICH_COLORS: bool = True  # Enable/disable colored output (set False for piped output)
RICH_PANELS: bool = True  # Use bordered panels for errors/warnings/success messages
RICH_PROGRESS: bool = True  # Show progress bars during batch/reel processing

# ── Settings Metadata ────────────────────────────────────────────────
SETTINGS_DESCRIPTIONS: dict[str, str] = {
    "REENCODING": "Re-encode clips via ffmpeg instead of stream-copying. Slower but fixes some codec issues.",
    "AUDIO_NORMALIZE": "Normalize audio levels across generated clips for consistent volume.",
    "FILEFORMAT": "Output container format for generated video clips.",
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
    "TRANSCRIBE_ENABLED": "Generate transcripts alongside clips using faster-whisper.",
    "TRANSCRIBE_MODEL": "Whisper model size: tiny, base, small, medium, large-v3.",
    "TRANSCRIBE_FORMAT": "Transcript output format: md (Markdown), srt, or vtt.",
    "TRANSCRIBE_PREWARM": "When the Transcripts page pre-loads the Whisper model: off, queue_open (opening Queue), or page_load (after listing participants).",
    "HIGHLIGHTS_REEL_DURATION_SECONDS": "Maximum duration in seconds for the highlights reel time budget.",
    "MANIFEST_ENABLED": "Write a manifest JSON file alongside generated artifacts for session tracking.",
    "STUDIO_CELL_EXPAND_HOVER": "Expand overflowing timestamp cells on hover in the Sheet Preview.",
    "FILMSTRIP_ENABLED": "Show thumbnail images on timeline markers instead of solid colors (in the HTML viewer).",
    "GALLERY_BUNDLE_ENABLED": "Embed gallery images as base64 data URIs in the HTML file, making it fully self-contained.",
    "CLIP_PARALLEL_WORKERS": "Number of concurrent ffmpeg processes for clip generation. 0 = auto, 1 = sequential.",
    "OLLAMA_SUMMARY_ENABLED": "Auto-generate AI summaries of transcripts using a local Ollama model.",
    "OLLAMA_SUMMARY_MODEL": "Ollama model used for transcript summaries and citation linking.",
    "OLLAMA_BASE_URL": "Base URL of the local Ollama server.",
}

# Studio-exposed settings with UI metadata (group, type, constraints).
STUDIO_SETTINGS: dict[str, dict[str, Any]] = {
    "REENCODING": {"group": "Video Output", "type": "bool"},
    "AUDIO_NORMALIZE": {"group": "Video Output", "type": "bool"},
    "FILEFORMAT": {
        "group": "Video Output",
        "type": "select",
        "options": [".mp4", ".webm", ".mkv"],
    },
    "MAX_FILESIZE_MB": {"group": "Video Output", "type": "int", "min": 0, "step": 1},
    "DEFAULT_DURATION_SECONDS": {
        "group": "Clip Behavior",
        "type": "int",
        "min": 1,
        "step": 1,
    },
    "MAX_CLIP_DURATION_SECONDS": {
        "group": "Clip Behavior",
        "type": "int",
        "min": 1,
        "step": 1,
    },
    "TITLECARDS_ENABLED": {"group": "Titlecards", "type": "bool"},
    "TITLECARD_DURATION_SECONDS": {
        "group": "Titlecards",
        "type": "int",
        "min": 1,
        "step": 1,
    },
    "HIGHLIGHTS_REEL_DURATION_SECONDS": {
        "group": "Generation",
        "type": "int",
        "min": 10,
        "step": 10,
    },
    "MANIFEST_ENABLED": {"group": "Generation", "type": "bool"},
    "CLIP_PARALLEL_WORKERS": {
        "group": "Generation",
        "type": "int",
        "min": 0,
        "step": 1,
    },
    "STUDIO_CELL_EXPAND_HOVER": {"group": "Sheet Preview", "type": "bool"},
    "TRANSCRIBE_ENABLED": {"group": "Transcription", "type": "bool"},
    "TRANSCRIBE_MODEL": {
        "group": "Transcription",
        "type": "model_select",
        "provider": "whisper",
    },
    "TRANSCRIBE_FORMAT": {
        "group": "Transcription",
        "type": "select",
        "options": ["md", "srt", "vtt"],
    },
    "TRANSCRIBE_PREWARM": {
        "group": "Transcription",
        "type": "select",
        "options": ["off", "queue_open", "page_load"],
    },
    "OLLAMA_SUMMARY_ENABLED": {"group": "AI Summary", "type": "bool"},
    "OLLAMA_SUMMARY_MODEL": {
        "group": "AI Summary",
        "type": "model_select",
        "provider": "ollama",
    },
    "OLLAMA_BASE_URL": {"group": "AI Summary", "type": "str"},
}
