# -*- coding: utf-8 -*-
"""Configuration constants for clipgen."""

from icecream import ic

# Configuration Constants
REENCODING = False
AUDIO_NORMALIZE = False
FILEFORMAT = '.mp4'
VERSIONNUM = '0.6.4'
SHEET_NAME = 'Sheet1'  # Deprecated: use get_worksheet() instead
WORKSHEET_PRIORITY = ['Sheet1', 'Data', 'data', 'Observations', 'Data set', 'data set']
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
NOTES_COLUMN = 'Notes'

# File and Duration Constants
MAX_FILENAME_LENGTH = 255
MAX_CLIP_DURATION_SECONDS = 600  # 10 minutes
DEFAULT_DURATION_SECONDS = 60

# Browse Mode Constants
BROWSE_LINES_TO_DISPLAY = 5  # Number of rows to show at once when browsing

# Spreadsheet Selection Commands
COMMAND_LIST_ALL = 'all'
COMMAND_LIST_NEW = 'new'
COMMAND_OPEN_LAST = 'last'
COMMAND_SETTINGS = 'settings'
COMMAND_HTTP_PREFIX = 'http'
NUM_NEWEST_DOCS_TO_SHOW = 3  # Number of newest documents to show when using 'new' command
