# -*- coding: utf-8 -*-
"""Configuration constants for clipgen."""

from icecream import ic

# Configuration Constants
REENCODING = False
AUDIO_NORMALIZE = False
FILEFORMAT = '.mp4'
VERSIONNUM = '0.7.17'
WORKSHEET_PRIORITY = ['Sheet1', 'Data', 'data', 'Observations', 'Data set', 'data set', 'dataset', 'Dataset']
DEBUGGING = False
VERBOSE = True  # Set to False in CLI mode unless -v flag is used

# Configure Icecream debugging
if DEBUGGING:
    ic.configureOutput(prefix='! DEBUG ic| ', includeContext=False)
else:
    ic.disable()

# Spreadsheet Structure Constants
ID_HEADER = 'ID'
OBSERVATION_HEADER = 'Observation'
CATEGORY_HEADER = 'Category'
PARTICIPANT_PREFIXES = ('P', 'G')  # 'P' for individual, 'G' for group
ANNOTATION_KEYPHRASES = {
    '!key': 'key',
}
IGNORED_TIMESTAMP_TOKENS = {'x'}

# File and Duration Constants
MAX_FILENAME_LENGTH = 255
MAX_CLIP_DURATION_SECONDS = 600  # 10 minutes
DEFAULT_DURATION_SECONDS = 60
DEFAULT_GIF_DURATION_SECONDS = 5
MAX_FILESIZE_MB = 0  # Maximum output file size in MB (0 = disabled)

# Browse Mode Constants
BROWSE_LINES_TO_DISPLAY = 5  # Number of rows to show at once when browsing
BROWSE_DESCRIPTION_MAX_WIDTH = 40  # Max width for description column in table
BROWSE_TIMESTAMP_MAX_WIDTH = 15    # Max width for each timestamp column

# Spreadsheet Selection Commands
COMMAND_LIST_ALL = 'all'
COMMAND_LIST_NEW = 'new'
COMMAND_OPEN_LAST = 'last'
COMMAND_SETTINGS = 'settings'
COMMAND_HTTP_PREFIX = 'http'
COMMAND_EXCEL = 'excel'
NUM_NEWEST_DOCS_TO_SHOW = 3  # Number of newest documents to show when using 'new' command

# Display / preview constants
DESCRIPTION_PREVIEW_LENGTH = 50  # Max chars for description in previews
REEL_PREVIEW_CLIP_COUNT = 10  # Number of clips to show in reel mode preview
MAX_SKIPPED_TIMESTAMPS_TO_SHOW = 3  # Max skipped timestamps to list in parse_timestamps warning

# Rich output settings
RICH_COLORS = True    # Enable/disable colored output (set False for piped output)
RICH_PANELS = True    # Use bordered panels for errors/warnings/success messages
RICH_PROGRESS = True  # Show progress bars during batch/reel processing

# Textual TUI settings
TEXTUAL_TUI = False    # Use Textual interactive screens

# Settings descriptions (shown in TUI settings screen)
SETTINGS_DESCRIPTIONS = {
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
