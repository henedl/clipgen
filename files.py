# -*- coding: utf-8 -*-
"""File and filename operations for clipgen."""

import os

import gspread
from icecream import ic

import config
import utils


def double_digits(number):
	"""Takes a string, returns a double digit number."""
	try:
		if int(number) < 10:
			return '0' + number
		return number
	except TypeError:
		return number

def format_filesize(size_bytes, precision=2):
	"""Format byte size as human-readable string."""
	suffixes = ['B', 'KB', 'MB', 'GB', 'TB']
	suffix_index = 0
	while size_bytes > 1024 and suffix_index < 4:
		suffix_index += 1
		size_bytes = size_bytes / 1024
	return f'{size_bytes:.{precision}f}{suffixes[suffix_index]}'

def get_unique_filename(filename):
	"""Appends an incremented number to prevent overwriting existing files."""
	step = 1
	while True:
		if os.path.isfile(filename):
			if step < 2:
				suffix_pos = filename.find(config.FILEFORMAT)
				filename = filename[0:suffix_pos] + '-' + str(step) + config.FILEFORMAT
			else:
				dash_pos = filename.rfind('-')
				filename = filename[0:dash_pos] + '-' + str(step) + config.FILEFORMAT
			step += 1
		else:
			filename = truncate_filename(filename, step)
			break
	return filename

def truncate_filename(filename, step=1):
	"""Truncate filenames that exceed maximum length (255 chars on Windows)."""
	if len(filename) > config.MAX_FILENAME_LENGTH:
		if step > 1:
			utils.debug_print(f'Filename was longer than {config.MAX_FILENAME_LENGTH} chars ({filename}, length {len(filename)})')
			filename = filename[0:config.MAX_FILENAME_LENGTH-(1+len(str(step))+len(config.FILEFORMAT))] + '-' + str(step) + config.FILEFORMAT
		else:
			filename = filename[0:config.MAX_FILENAME_LENGTH-(len(config.FILEFORMAT))] + config.FILEFORMAT
	return filename

def clean_issue(issue):
	"""Parse timestamps and sanitize description/category for filename use."""
	ic(issue)
	utils.debug_print(f"clean_issue() received issue with cell contents {issue['cell'].value}")
	utils.debug_print('Will attempt to split the cell contents')
	
	# Get cell reference for error messages
	cell_ref = gspread.utils.rowcol_to_a1(issue['cell'].row, issue['cell'].col)
	
	# Parse timestamps from cell value
	issue['times'] = utils.parse_timestamps(issue['cell'].value, cell_ref=cell_ref)
	ic(issue['times'])
	
	# Warn if no valid timestamps were parsed
	if not issue['times']:
		print(f"! WARNING No valid timestamps found in cell {cell_ref}")
		print(f"  Cell contents: '{issue['cell'].value}'")
		print(f"  Participant: {issue['participant']}, Description: {issue['desc'][:50]}...")

	# Clean description: remove bracketed prefix and sanitize
	# Handle case where description doesn't contain ']'
	bracket_pos = issue['desc'].rfind(']')
	if bracket_pos >= 0:
		desc = issue['desc'][bracket_pos+1:].strip()
	else:
		desc = issue['desc'].strip()
	issue['desc'] = utils.sanitize_filename(desc)
	ic(issue['desc'])
	
	# Sanitize category (handle None/empty)
	if issue['category']:
		issue['category'] = utils.sanitize_filename(issue['category'])
	else:
		issue['category'] = 'uncategorized'
	ic(issue['category'])

	ic(issue)
	return issue
