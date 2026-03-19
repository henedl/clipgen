# -*- coding: utf-8 -*-
"""Configuration constants for clipgen."""

from typing import Dict, List, Optional, Set, Tuple

from icecream import ic

# Configuration Constants
REENCODING: bool = False
AUDIO_NORMALIZE: bool = False
FILEFORMAT: str = ".mp4"
VERSIONNUM: str = "0.9.2"
TITLECARDS_ENABLED: bool = False
TITLECARD_DURATION_SECONDS: int = 2
WORKSHEET_PRIORITY: List[str] = [
    "Sheet1",
    "Data",
    "data",
    "Observations",
    "Data set",
    "data set",
    "dataset",
    "Dataset",
]
DEBUGGING: bool = False
QUIET: int = 0
STANDARD: int = 1
VERBOSE: int = 2
VERBOSITY: int = STANDARD

# Directory configuration
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

# Spreadsheet Structure Constants
ID_HEADER: str = "ID"
OBSERVATION_HEADER: str = "Observation"
CATEGORY_HEADER: str = "Category"
FILENAME_HEADER: str = "Filename"
PARTICIPANT_PREFIXES: Tuple[str, ...] = ("P", "G")  # 'P' for individual, 'G' for group
ANNOTATION_KEYPHRASES: Dict[str, str] = {
    "!key": "key",
}
IGNORED_TIMESTAMP_TOKENS: Set[str] = {"x"}

# File and Duration Constants
MAX_FILENAME_LENGTH: int = 255
MAX_CLIP_DURATION_SECONDS: int = 600  # 10 minutes
DEFAULT_DURATION_SECONDS: int = 60
DEFAULT_GIF_DURATION_SECONDS: int = 5
MAX_FILESIZE_MB: int = 0  # Maximum output file size in MB (0 = disabled)
MIN_SOURCE_VIDEO_SIZE_MB: int = 100  # Minimum file size (MB) to consider as a source video candidate during fuzzy matching
MANIFEST_FILENAME: str = "clipgen_manifest.json"
MANIFEST_ENABLED: bool = False

# Browse Mode Constants
BROWSE_LINES_TO_DISPLAY: int = 5  # Number of rows to show at once when browsing
BROWSE_LINES_TO_SCROLL: int = 5  # Number of rows to move when scrolling up/down
BROWSE_DESCRIPTION_MAX_WIDTH: int = 40  # Max width for description column in table
BROWSE_TIMESTAMP_MAX_WIDTH: int = 15  # Max width for each timestamp column

# Spreadsheet Selection Commands
COMMAND_LIST_ALL: str = "all"
COMMAND_LIST_NEW: str = "new"
COMMAND_OPEN_LAST: str = "last"
COMMAND_SETTINGS: str = "settings"
COMMAND_HTTP_PREFIX: str = "http"
COMMAND_EXCEL: str = "excel"
NUM_NEWEST_DOCS_TO_SHOW: int = (
    3  # Number of newest documents to show when using 'new' command
)

# Display / preview constants
DESCRIPTION_PREVIEW_LENGTH: int = 50  # Max chars for description in previews
PROGRESS_DESCRIPTION_LENGTH: int = 30  # Max chars for description in progress bar
REEL_PREVIEW_CLIP_COUNT: int = 10  # Number of clips to show in reel mode preview
MAX_SKIPPED_TIMESTAMPS_TO_SHOW: int = (
    3  # Max skipped timestamps to list in parse_timestamps warning
)

# Timestamp format constants
SECONDS_PER_HOUR: int = 3600
SECONDS_PER_MINUTE: int = 60
MAX_MMSS_LENGTH: int = 5  # Max length of an MM:SS timestamp string

# ffmpeg constants
FFMPEG_LOGLEVEL: str = "16"  # ffmpeg -loglevel value (16 = error)
FFMPEG_SCREENSHOT_QUALITY: str = "2"  # -q:v value for screenshots (1=best, 31=worst)
GIF_FPS: int = 10
GIF_SCALE_WIDTH: int = 480
AUDIO_BITRATE_KBPS: int = 128
COMPRESSION_SIZE_FACTOR: float = 0.95  # Target 95% of max to leave headroom
MIN_VIDEO_BITRATE_KBPS: int = 100

# Source video pattern
SOURCE_VIDEO_PATTERN: str = r"_[PG]\d+\.mp4$"

# Transcription constants (faster-whisper)
TRANSCRIBE_ENABLED: bool = False
TRANSCRIBE_MODEL: str = "base"  # tiny, base, small, medium, large-v3
TRANSCRIBE_LANGUAGE: Optional[str] = None  # None = auto-detect
TRANSCRIBE_COMPUTE_TYPE: str = "int8"  # int8, float16, float32
TRANSCRIBE_FORMAT: str = "md"  # md, srt, vtt
TRANSCRIBE_INITIAL_PROMPT: str = "This is a recorded user experience research session."
TRANSCRIBE_BEAM_SIZE: int = 5

# Rich output settings
RICH_COLORS: bool = True  # Enable/disable colored output (set False for piped output)
RICH_PANELS: bool = True  # Use bordered panels for errors/warnings/success messages
RICH_PROGRESS: bool = True  # Show progress bars during batch/reel processing

# Settings descriptions (used by interactive settings helpers)
SETTINGS_DESCRIPTIONS: Dict[str, str] = {
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
}
