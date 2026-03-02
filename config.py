# -*- coding: utf-8 -*-
"""Configuration constants for clipgen."""

from typing import Dict, List, Set, Tuple

from icecream import ic

# Configuration Constants
REENCODING: bool = False
AUDIO_NORMALIZE: bool = False
FILEFORMAT: str = '.mp4'
VERSIONNUM: str = '0.7.18'
WORKSHEET_PRIORITY: List[str] = ['Sheet1', 'Data', 'data', 'Observations', 'Data set', 'data set', 'dataset', 'Dataset']
DEBUGGING: bool = False
VERBOSE: bool = True  # Set to False in CLI mode unless -v flag is used

# Configure Icecream debugging
if DEBUGGING:
    ic.configureOutput(prefix='! DEBUG ic| ', includeContext=False)
else:
    ic.disable()

# Spreadsheet Structure Constants
ID_HEADER: str = 'ID'
OBSERVATION_HEADER: str = 'Observation'
CATEGORY_HEADER: str = 'Category'
PARTICIPANT_PREFIXES: Tuple[str, ...] = ('P', 'G')  # 'P' for individual, 'G' for group
ANNOTATION_KEYPHRASES: Dict[str, str] = {
    '!key': 'key',
}
IGNORED_TIMESTAMP_TOKENS: Set[str] = {'x'}

# File and Duration Constants
MAX_FILENAME_LENGTH: int = 255
MAX_CLIP_DURATION_SECONDS: int = 600  # 10 minutes
DEFAULT_DURATION_SECONDS: int = 60
DEFAULT_GIF_DURATION_SECONDS: int = 5
MAX_FILESIZE_MB: int = 0  # Maximum output file size in MB (0 = disabled)

# Browse Mode Constants
BROWSE_LINES_TO_DISPLAY: int = 5  # Number of rows to show at once when browsing
BROWSE_DESCRIPTION_MAX_WIDTH: int = 40  # Max width for description column in table
BROWSE_TIMESTAMP_MAX_WIDTH: int = 15    # Max width for each timestamp column

# Spreadsheet Selection Commands
COMMAND_LIST_ALL: str = 'all'
COMMAND_LIST_NEW: str = 'new'
COMMAND_OPEN_LAST: str = 'last'
COMMAND_SETTINGS: str = 'settings'
COMMAND_HTTP_PREFIX: str = 'http'
COMMAND_EXCEL: str = 'excel'
NUM_NEWEST_DOCS_TO_SHOW: int = 3  # Number of newest documents to show when using 'new' command

# Display / preview constants
DESCRIPTION_PREVIEW_LENGTH: int = 50  # Max chars for description in previews
PROGRESS_DESCRIPTION_LENGTH: int = 30  # Max chars for description in progress bar
REEL_PREVIEW_CLIP_COUNT: int = 10  # Number of clips to show in reel mode preview
MAX_SKIPPED_TIMESTAMPS_TO_SHOW: int = 3  # Max skipped timestamps to list in parse_timestamps warning

# Timestamp format constants
SECONDS_PER_HOUR: int = 3600
SECONDS_PER_MINUTE: int = 60
MAX_MMSS_LENGTH: int = 5  # Max length of an MM:SS timestamp string

# ffmpeg constants
FFMPEG_LOGLEVEL: str = '16'  # ffmpeg -loglevel value (16 = error)
FFMPEG_SCREENSHOT_QUALITY: str = '2'  # -q:v value for screenshots (1=best, 31=worst)
GIF_FPS: int = 10
GIF_SCALE_WIDTH: int = 480
AUDIO_BITRATE_KBPS: int = 128
COMPRESSION_SIZE_FACTOR: float = 0.95  # Target 95% of max to leave headroom
MIN_VIDEO_BITRATE_KBPS: int = 100

# Source video pattern
SOURCE_VIDEO_PATTERN: str = r'_[PG]\d+\.mp4$'

# Rich output settings
RICH_COLORS: bool = True    # Enable/disable colored output (set False for piped output)
RICH_PANELS: bool = True    # Use bordered panels for errors/warnings/success messages
RICH_PROGRESS: bool = True  # Show progress bars during batch/reel processing

# Textual TUI settings
TEXTUAL_TUI: bool = False    # Use Textual interactive screens

# Settings descriptions (shown in TUI settings screen)
SETTINGS_DESCRIPTIONS: Dict[str, str] = {
    'REENCODING': 'Re-encode clips via ffmpeg instead of stream-copying. Slower but fixes some codec issues.',
    'AUDIO_NORMALIZE': 'Normalize audio levels across generated clips for consistent volume.',
    'FILEFORMAT': 'Output container format for generated video clips.',
    'MAX_FILESIZE_MB': 'Compress output to stay under this size limit. Set to 0 to disable.',
    'DEFAULT_DURATION_SECONDS': 'Clip length when only a start time is provided.',
    'MAX_CLIP_DURATION_SECONDS': 'Prompt for confirmation before generating clips longer than this.',
    'DEBUGGING': 'Enable debug output via icecream and skip ffmpeg execution.',
    'RICH_COLORS': 'Use colored terminal output via Rich library.',
    'RICH_PANELS': 'Show bordered panels for error, warning, and success messages.',
    'RICH_PROGRESS': 'Display progress bars during batch and reel processing.',
    'TEXTUAL_TUI': 'Use Textual interactive screens for settings, browse, and category selection.',
}
