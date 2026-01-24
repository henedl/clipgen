# -*- coding: utf-8 -*-
"""clipgen - Video clip generator from Google Sheets timestamps.

This program will help quickly create video snippets from longer video files, based on timestamps in a spreadsheet!
Check out README.md for more detailed information about setting up and using clipgen.

This script supports full unicode/UTF-8 for international characters in:
- Study names
- Participant IDs  
- Category names
- Descriptions
- File paths
"""
import argparse
import os
import sys
import subprocess
from datetime import datetime, timedelta

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configuration Constants
REENCODING = False
FILEFORMAT = '.mp4'
VERSIONNUM = '0.5.1'
SHEET_NAME = 'Sheet1'
DEBUGGING  = False
VERBOSE    = True  # Set to False in CLI mode unless -v flag is used

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

# ============================================================================
# Utility Functions
# ============================================================================

def parse_arguments():
	"""Parse command-line arguments for non-interactive mode."""
	parser = argparse.ArgumentParser(
		description='clipgen - Video clip generator from Google Sheets timestamps.',
		formatter_class=argparse.RawDescriptionHelpFormatter,
		epilog='''
Examples:
  python clipgen.py                    Interactive mode (default)
  python clipgen.py -b                 Batch mode - generate all clips
  python clipgen.py -l 5               Single line mode - line 5
  python clipgen.py -l 1+4+5           Multi-line mode - lines 1, 4, and 5
  python clipgen.py -l 1,4,5           Multi-line mode (comma separator)
  python clipgen.py -r 1-10            Range mode - lines 1 through 10
  python clipgen.py -b -s "Study Name" Batch mode with specific spreadsheet
  python clipgen.py -l 5 -y            Line mode, skip confirmation prompts
  python clipgen.py -b -v              Batch mode with verbose output

Note: Non-interactive mode (using -b, -l, or -r) is silent by default,
      only showing errors and the final summary. Use -v for full output.
'''
	)
	
	# Mode arguments (mutually exclusive)
	mode_group = parser.add_mutually_exclusive_group()
	mode_group.add_argument('-b', '--batch', action='store_true',
		help='Batch mode: generate all possible clips')
	mode_group.add_argument('-l', '--lines', type=str, metavar='LINES',
		help='Line mode: specify line numbers separated by + or , (e.g., 1+4+5 or 1,4,5)')
	mode_group.add_argument('-r', '--range', type=str, metavar='RANGE',
		help='Range mode: specify start-end line range (e.g., 1-10)')
	
	# Optional arguments
	parser.add_argument('-s', '--spreadsheet', type=str, metavar='NAME',
		help='Spreadsheet name, URL, or index number')
	parser.add_argument('-y', '--yes', action='store_true',
		help='Skip confirmation prompts (auto-confirm)')
	parser.add_argument('-v', '--verbose', action='store_true',
		help='Enable verbose output in non-interactive mode (shows all messages)')
	
	return parser.parse_args()

def debug_print(message):
	"""Print debug messages when DEBUGGING is enabled."""
	if DEBUGGING:
		print(f'! DEBUG {message}')

def verbose_print(message):
	"""Print informational messages when VERBOSE is enabled.
	
	In interactive mode, VERBOSE is always True.
	In CLI mode, VERBOSE is False unless -v flag is used.
	"""
	if VERBOSE:
		print(message)

def normalize_study_name(raw_name):
	"""Convert study name to a filesystem-safe format.
	Preserves unicode characters for international study names."""
	# Ensure we're working with a string
	name = str(raw_name)
	name = name.lower()
	name = name.replace('study ', 'study')
	name = name.replace(' ', '_')
	return name

def sanitize_filename(text):
	"""Remove or replace characters that are unsafe for filenames.
	Preserves unicode characters to support international filenames."""
	# Ensure text is a string and handle unicode properly
	text = str(text)
	
	# Characters that need special replacement
	text = text.replace('\\', '-')
	text = text.replace('/', '-')
	text = text.replace('?', '_')
	# Characters to remove entirely (filesystem-unsafe characters)
	for char in ['\'', '\"', '.', '>', '<', '|', ':']:
		text = text.replace(char, '')
	return text

def parse_timestamps(cell_value, cell_ref=None):
	"""Parse timestamp pairs from a cell value string.
	
	Args:
		cell_value: The raw cell value containing timestamps
		cell_ref: Optional cell reference (e.g., 'B5') for error messages
	
	Returns a list of (start_time, end_time) tuples.
	"""
	parsed_timestamps = []
	skipped_timestamps = []
	raw_times = cell_value.lower().replace('+', ' ').replace(';', ' ').replace(',', ' ').split()
	debug_print(f'raw_times content after split is {raw_times}')
	debug_print(f'Timestamp list raw_times is {len(raw_times)} entries long')

	for i in range(len(raw_times)):
		debug_print(f'Cleaning timestamp {raw_times[i]}')
		# Remove trailing commas and dashes.
		raw_times[i] = raw_times[i].strip().rstrip(',').rstrip('-')

		# Change . to : for the timestamp.
		raw_times[i] = raw_times[i].replace('.', ':')
		
		if raw_times[i] == '':
			# We don't need to do anything with blank timestamps.
			debug_print(f'Found blank timestamp {raw_times[i]}')
		elif '-' in raw_times[i]:
			# We have a dash which should mean we have two timestamps, so we need to split it into two timestamps.
			dash_pos = raw_times[i].find('-')
			if dash_pos > 0 and raw_times[i][dash_pos-1].isdigit():
				# Slice the timestamp until the dash, and then from after the dash
				time_pair = (raw_times[i][:dash_pos], raw_times[i][dash_pos+1:])
				parsed_timestamps.append(time_pair)
			else:
				skipped_timestamps.append(raw_times[i])
		elif ':' in raw_times[i]:
			colon_pos = raw_times[i].find(':')
			if colon_pos > 0 and raw_times[i][colon_pos-1].isdigit():
				# Single timestamp - add default end time
				end_time = add_duration(raw_times[i])
				if end_time != -1:
					time_pair = (raw_times[i], end_time)
					parsed_timestamps.append(time_pair)
				# If -1, warning already printed by add_duration()
			else:
				skipped_timestamps.append(raw_times[i])
		elif raw_times[i]:
			# Non-empty but doesn't match expected patterns
			skipped_timestamps.append(raw_times[i])
	
	# Report skipped timestamps if any
	if skipped_timestamps:
		cell_info = f" in cell {cell_ref}" if cell_ref else ""
		print(f"! WARNING Skipped {len(skipped_timestamps)} unparseable timestamp(s){cell_info}:")
		for ts in skipped_timestamps[:3]:  # Show first 3
			print(f"    '{ts}'")
		if len(skipped_timestamps) > 3:
			print(f"    ... and {len(skipped_timestamps) - 3} more")
		print("  Expected formats: MM:SS-MM:SS, HH:MM:SS-HH:MM:SS, or single timestamps like MM:SS")

	return parsed_timestamps

def set_program_settings():
	SETTINGSLIST = ['REENCODING', 'FILEFORMAT', 'DEBUGGING']

	print('\nWhich setting? Available:\n')
	print(', '.join(SETTINGSLIST))
	setting_to_change = input('\n>> ')

	print(f"* Current value for '{setting_to_change}' is '{globals()[setting_to_change]}'")

	new_value = input('\nWhich new value?\n>> ')

	print(f"* '{setting_to_change}' SET TO '{new_value}'")

	if setting_to_change != '':
		globals()[setting_to_change] = new_value
		return True
	return False

# ============================================================================
# Spreadsheet Functions
# ============================================================================

def generate_list(sheet, mode, line_numbers=None, range_start=None, range_end=None, skip_prompts=False):
	"""Goes through a sheet, bundles values from timestamp columns and descriptions columns into tuples.
	
	Args:
		sheet: The gspread worksheet object
		mode: One of 'batch', 'line', 'range', 'category', 'select'
		line_numbers: Optional list of line numbers for 'line' mode (CLI)
		range_start: Optional start line for 'range' mode (CLI)
		range_end: Optional end line for 'range' mode (CLI)
		skip_prompts: If True, skip confirmation prompts (CLI -y flag)
	"""
	# Find required headers
	id_cell = sheet.find(ID_HEADER)
	observation_cell = sheet.find(OBSERVATION_HEADER)
	category_cell = sheet.find(CATEGORY_HEADER)
	timestamps = []
	
	# Validate required headers exist
	missing_headers = []
	if id_cell is None:
		missing_headers.append(f"'{ID_HEADER}'")
	if observation_cell is None:
		missing_headers.append(f"'{OBSERVATION_HEADER}'")
	if category_cell is None:
		missing_headers.append(f"'{CATEGORY_HEADER}'")
	
	if missing_headers:
		print(f"! ERROR Required header(s) not found in spreadsheet: {', '.join(missing_headers)}")
		print(f"  The spreadsheet must contain columns with these exact headers: {ID_HEADER}, {OBSERVATION_HEADER}, {CATEGORY_HEADER}")
		print(f"  Please check your spreadsheet structure.")
		return []

	# Sheet data is a list of lists, which forms a matrix
	# - sheet_data[row][col] where indices start at 0 (real spreadsheet starts at 1)
	sheet_data = sheet.get_all_values()
	debug_print(f'Sheet dumped into memory at {get_current_time()}')
	
	# Check if sheet is empty or has only headers
	if len(sheet_data) <= 1:
		print("! ERROR Spreadsheet appears to be empty (no data rows found).")
		print(f"  The spreadsheet only has {len(sheet_data)} row(s).")
		return []

	# Determine the study name.
	study_name = sheet_data[0][0]
	if study_name == '':
		study_name = sheet.spreadsheet.title
	verbose_print(f'\nBeginning work on {study_name}.')

	# Normalize study name for filesystem use
	study_name = normalize_study_name(study_name)

	# Get number of participants needed to loop through the worksheet
	num_participants = get_num_participants(sheet.row_values(id_cell.row), id_cell, sheet.col_count)
	
	# Warn if no participants found
	if num_participants == 0:
		print(f"! WARNING No participant columns found in the spreadsheet.")
		print(f"  Looking for columns starting with: {', '.join(PARTICIPANT_PREFIXES)}")
		print(f"  Check that participant column headers start with 'P' or 'G' (e.g., P01, P02, G01).")
		return []

	# Generate the timestamps, according to the selected mode.
	if mode == 'batch':
		if skip_prompts:
			verbose_print('Batch mode: generating all possible clips...')
			timestamps = generate_batch_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name)
		else:
			yn = input('\nWarning: This will generate all possible clips. Do you want to proceed? y/n\n>> ')
			if yn == 'y':
				timestamps = generate_batch_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name)
	elif mode == 'category':
		# Collect all unique categories from the sheet
		categories = collect_categories(sheet_data, id_cell, category_cell)
		
		if not categories:
			print('\nNo categories found in the spreadsheet.')
			return []
		
		# Display categories with numbered options
		print('\nAvailable categories:')
		for i, cat in enumerate(categories, 1):
			print(f'  {i}. {cat}')
		
		# Get user selection
		while True:
			selection = input('\nEnter category numbers (comma-separated, e.g., "1,3,5") or "all":\n>> ')
			
			if selection.lower() == 'all':
				selected_categories = categories
				break
			
			try:
				indices = [int(x.strip()) for x in selection.split(',')]
				selected_categories = []
				invalid_indices = []
				
				for idx in indices:
					if 1 <= idx <= len(categories):
						if categories[idx-1] not in selected_categories:
							selected_categories.append(categories[idx-1])
					else:
						invalid_indices.append(idx)
				
				if invalid_indices:
					print(f'  Invalid index(es): {", ".join(str(i) for i in invalid_indices)}')
				
				if selected_categories:
					print('\nSelected categories:')
					for cat in selected_categories:
						print(f'  - {cat}')
					yn = input('\nIs this correct? y/n\n>> ')
					if yn == 'y':
						break
				else:
					print('No valid categories selected. Please try again.')
			except ValueError:
				print('Please enter valid numbers separated by commas.')
		
		timestamps = generate_category_timestamps(sheet_data, id_cell, observation_cell, category_cell, num_participants, study_name, selected_categories)
	elif mode == 'line':
		timestamps = generate_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name, line_numbers, skip_prompts)
	elif mode == 'range':
		if range_start is not None and range_end is not None:
			# CLI mode - use provided range with bounds validation
			max_row = len(sheet_data)
			if range_start < 1 or range_end < 1:
				print(f"! ERROR Line numbers must be positive. Got start={range_start}, end={range_end}")
				return []
			if range_start > max_row or range_end > max_row:
				print(f"! ERROR Line number(s) out of range. Spreadsheet has {max_row} rows.")
				print(f"  Requested: lines {range_start} to {range_end}")
				return []
			if range_start > range_end:
				print(f"! ERROR Start line ({range_start}) must be less than or equal to end line ({range_end}).")
				return []
			verbose_print(f'Range mode: lines {range_start} to {range_end}')
			verbose_print(f'Lines selected: {sheet_data[range_start-1][observation_cell.col-1]} to {sheet_data[range_end-1][observation_cell.col-1]}')
			timestamps = generate_range_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name, range_start, range_end)
		else:
			# Interactive mode
			max_row = len(sheet_data)
			while True:
				try:
					start_line = int(input('\nWhich starting line (row number only)?\n>> '))
					end_line = int(input('\nWhich ending line (row number only)?\n>> '))
				except ValueError:
					print('\nInvalid input. Please enter row numbers as integers.')
					continue
				
				# Validate bounds
				if start_line < 1 or end_line < 1:
					print(f'\n! ERROR Line numbers must be positive (got {start_line} and {end_line}).')
					continue
				if start_line > max_row or end_line > max_row:
					print(f'\n! ERROR Line number out of range. Spreadsheet has {max_row} rows.')
					continue
				if start_line > end_line:
					print(f'\n! ERROR Start line ({start_line}) must be less than or equal to end line ({end_line}).')
					continue
				
				print(f'Lines selected: {sheet_data[start_line-1][observation_cell.col-1]} to {sheet_data[end_line-1][observation_cell.col-1]}')
				yn = input('Is this correct? y/n\n>> ')
				if yn == 'y':
					break
			timestamps = generate_range_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name, start_line, end_line)
	elif mode == 'select':
		pass

	return timestamps

def get_num_participants(header_row, id_cell, col_count):
	"""Returns the number of participant columns in the worksheet."""
	num_participants = 0
	for j in range(0, col_count):
		if len(header_row[j]) > 0:
			if header_row[j][0] in PARTICIPANT_PREFIXES:
				num_participants += 1
			elif header_row[j] == NOTES_COLUMN:
				break
	verbose_print(f'Found {num_participants} participants in total, spanning columns {id_cell.col+1} to {num_participants+id_cell.col+1}.')
	return num_participants

def get_current_time():
	return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def generate_batch_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name):
	debug_print('Running method generate_batch_timestamps()')
	timestamps = []
	for i in range(id_cell.row+1, len(sheet_data)):
		debug_print(f'Batching on line {i} (real sheet line {i+1})')
		timestamps.extend(get_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, i, study_name))
	return timestamps

def collect_categories(sheet_data, id_cell, category_cell):
	"""Scan sheet and return unique categories in order of first appearance."""
	categories = []
	category_col = category_cell.col - 1  # Convert from 1-indexed to 0-indexed
	
	# Start from the row after the category header
	for i in range(category_cell.row, len(sheet_data)):		
		category = sheet_data[i][category_col].strip()
		if category and category not in categories:
			categories.append(category)
	
	return categories

def generate_category_timestamps(sheet_data, id_cell, observation_cell, category_cell, num_participants, study_name, selected_categories):
	"""Generate timestamps for all rows matching any of the selected categories."""
	debug_print('Starting method generate_category_timestamps()')
	timestamps = []
	category_col = category_cell.col - 1  # Convert from 1-indexed to 0-indexed
	
	# Start from the row after the category header
	for i in range(category_cell.row, len(sheet_data)):
		row_category = sheet_data[i][category_col].strip()
		if row_category in selected_categories:
			debug_print(f"Row {i+1} matches category '{row_category}'")
			timestamps.extend(get_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, i, study_name))
	
	return timestamps

def generate_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name, cli_line_numbers=None, skip_prompts=False):
	"""Generate videos for one or more line/row numbers.
	
	Args:
		sheet_data: The sheet data matrix
		id_cell: The ID header cell
		observation_cell: The observation header cell
		num_participants: Number of participant columns
		study_name: Normalized study name
		cli_line_numbers: Optional list of line numbers from CLI (skips interactive input)
		skip_prompts: If True, skip confirmation prompts
	"""
	valid_lines = []
	
	if cli_line_numbers is not None:
		# CLI mode - use provided line numbers
		verbose_print(f'\nLine mode: processing lines {", ".join(str(n) for n in cli_line_numbers)}')
		verbose_print('\nSelected issues:')
		for line_num in cli_line_numbers:
			if line_num < 1 or line_num > len(sheet_data):
				verbose_print(f'  Line {line_num}: [INVALID - out of range]')
			else:
				desc = sheet_data[line_num-1][observation_cell.col-1]
				verbose_print(f'  Line {line_num}: {desc}')
				valid_lines.append(line_num)
		
		if not valid_lines:
			print('\nNo valid lines found. Exiting.')
			return []
	else:
		# Interactive mode
		while True:
			try:
				line_input = input('\nWhich issue(s)? Enter row number(s), comma-separated for multiple.\n>> ')
				# Parse comma-separated line numbers
				line_numbers = [int(num.strip()) for num in line_input.split(',')]
			except ValueError:
				print('\nTry again. Enter row numbers as integers, separated by commas.')
				continue
			
			# Preview all selected lines
			print('\nSelected issues:')
			valid_lines = []
			for line_num in line_numbers:
				if line_num < 1 or line_num > len(sheet_data):
					print(f'  Line {line_num}: [INVALID - out of range]')
				else:
					desc = sheet_data[line_num-1][observation_cell.col-1]
					print(f'  Line {line_num}: {desc}')
					valid_lines.append(line_num)
			
			if not valid_lines:
				print('\nNo valid lines selected. Please try again.')
				continue
			
			print()
			yn = input('Are these the correct issues? y/n\n>> ')
			if yn == 'y':
				break

	# Collect timestamps from all valid lines
	timestamps = []
	for line_num in valid_lines:
		debug_print(f'Calling get_line_timestamps() from generate_line_timestamps() for line {line_num}')
		line_timestamps = get_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, line_num-1, study_name)
		timestamps.extend(line_timestamps)
	
	debug_print(f'Printing return of get_line_timestamps() in generate_line_timestamps(): {len(timestamps)} total timestamps')
	debug_print(str(timestamps))

	return timestamps

def get_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, line_index, study_name):
	debug_print(f'Running method get_line_timestamps, starting line index {line_index} (real sheet line {line_index+1})')
	
	# Bounds checking
	if line_index < 0 or line_index >= len(sheet_data):
		print(f"! ERROR Line index {line_index} (row {line_index+1}) is out of bounds.")
		print(f"  Spreadsheet has {len(sheet_data)} rows.")
		return []

	timestamps = []
	try:
		for col_index, value in enumerate(sheet_data[line_index]):
			debug_print(f"Item {col_index} with value '{value}' being processed.")
			if col_index < id_cell.col:
				debug_print(f"Skipping item {col_index} with value '{value}'")
			elif col_index == id_cell.col + num_participants:
				debug_print(f'Exit for-loop, reached final column {col_index} (real sheet column {col_index+1}).')
				break
			elif value is None or value == '':
				pass
			else:
				cell = gspread.cell.Cell(line_index+1, col_index+1, value)
				debug_print(f'Found something at step {col_index}')
				debug_print(f'study_name is {study_name}')
				
				# Safely access array indices with bounds checking
				desc_col = observation_cell.col - 1
				category_col = observation_cell.col - 2
				participant_row = id_cell.row - 1
				
				# Get description with bounds check
				desc = ''
				if desc_col >= 0 and desc_col < len(sheet_data[line_index]):
					desc = sheet_data[line_index][desc_col]
				else:
					print(f"! WARNING Could not read description at row {line_index+1}, column {desc_col+1}")
				
				# Get participant with bounds check
				participant = ''
				if participant_row >= 0 and participant_row < len(sheet_data) and col_index < len(sheet_data[participant_row]):
					participant = sheet_data[participant_row][col_index]
				else:
					print(f"! WARNING Could not read participant ID at row {participant_row+1}, column {col_index+1}")
				
				# Get category with bounds check
				category = ''
				if category_col >= 0 and category_col < len(sheet_data[line_index]):
					category = sheet_data[line_index][category_col]
				
				issue = {
					'cell': cell,
					'desc': desc,
					'study': study_name,
					'participant': participant,
					'category': category
				}
				debug_print(f"Participant ID at R{id_cell.row},C{col_index} -> '{participant}'")
				debug_print(f"Description at R{line_index},C{observation_cell.col-1} -> '{desc}'")
				debug_print(f"Timestamp at R{cell.row-1},C{cell.col-1} -> '{cell.value}'")
				debug_print(f'Actual cell {cell} at actual address {gspread.utils.rowcol_to_a1(cell.row, cell.col)}')
				timestamps.append(issue)
				verbose_print(f"+ Found timestamp: {value.replace(chr(10), ' ')}")
	except IndexError as e:
		print(f"! ERROR Index error while reading row {line_index+1}: {e}")
		print("  The spreadsheet structure may be malformed.")

	debug_print(f'Line completed, returning list of {len(timestamps)} potential timestamps.')
	return timestamps

def generate_range_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name, start_line, end_line):
	timestamps = []
	for i in range(start_line-1, end_line):
		debug_print(f'Batching on line {i}')
		timestamps.extend(get_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, i, study_name))
	return timestamps

# ============================================================================
# File Operations
# ============================================================================

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
				suffix_pos = filename.find(FILEFORMAT)
				filename = filename[0:suffix_pos] + '-' + str(step) + FILEFORMAT
			else:
				dash_pos = filename.rfind('-')
				filename = filename[0:dash_pos] + '-' + str(step) + FILEFORMAT
			step += 1
		else:
			filename = truncate_filename(filename, step)
			break
	return filename

def truncate_filename(filename, step=1):
	"""Truncate filenames that exceed maximum length (255 chars on Windows)."""
	if len(filename) > MAX_FILENAME_LENGTH:
		if step > 1:
			debug_print(f'Filename was longer than {MAX_FILENAME_LENGTH} chars ({filename}, length {len(filename)})')
			filename = filename[0:MAX_FILENAME_LENGTH-(1+len(str(step))+len(FILEFORMAT))] + '-' + str(step) + FILEFORMAT
		else:
			filename = filename[0:MAX_FILENAME_LENGTH-(len(FILEFORMAT))] + FILEFORMAT
	return filename

def clean_issue(issue):
	"""Parse timestamps and sanitize description/category for filename use."""
	debug_print(f"clean_issue() received issue with cell contents {issue['cell'].value}")
	debug_print('Will attempt to split the cell contents')
	
	# Get cell reference for error messages
	cell_ref = gspread.utils.rowcol_to_a1(issue['cell'].row, issue['cell'].col)
	
	# Parse timestamps from cell value
	issue['times'] = parse_timestamps(issue['cell'].value, cell_ref=cell_ref)
	
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
	issue['desc'] = sanitize_filename(desc)
	
	# Sanitize category (handle None/empty)
	if issue['category']:
		issue['category'] = sanitize_filename(issue['category'])
	else:
		issue['category'] = 'uncategorized'

	return issue

# ============================================================================
# Video Operations
# ============================================================================

def run_ffmpeg(input_file, output_file, start_pos, end_pos, reencode):
	"""Calls ffmpeg to cut a video clip. Requires ffmpeg in system PATH.
	
	Returns True if video was generated successfully, False otherwise.
	"""
	# Check if input file exists before processing
	if not os.path.isfile(input_file):
		print(f"! ERROR Input video file not found: '{input_file}'")
		print(f"  Expected location: {os.path.join(os.getcwd(), input_file)}")
		print("  Skipping this clip.")
		return False
	
	duration = get_duration(start_pos, end_pos)
	if duration is None:
		# Error already printed by get_duration
		return False
	
	file_length = get_file_duration(input_file)
	if file_length is None:
		# Error already printed by get_file_duration
		return False

	if duration < 0:
		print(f"! ERROR Negative duration calculated for video clip. Skipping.")
		print(f"  Start: {start_pos}, End: {end_pos}, Duration: {duration}s")
		print("  The end timestamp must be after the start timestamp.")
		return False
	if duration > file_length:
		print(f"! ERROR Timestamp duration ({duration}s) exceeds video file length ({file_length}s). Skipping.")
		print(f"  Start: {start_pos}, End: {end_pos}")
		print(f"  Video file: '{input_file}'")
		return False
	if duration > MAX_CLIP_DURATION_SECONDS:
		yn = input(f'The generated video will be {duration}s ({duration//60}m {duration%60}s), over 10 minutes long. Generate anyway? (y/n)\n>> ')
		if yn != 'y':
			return False

	verbose_print(f'Cutting {input_file} from {start_pos} to {end_pos}.')
	if DEBUGGING:
		debug_print(f'Debugging enabled, not calling ffmpeg.\n  input_file: {input_file},\n  output_file: {output_file}')
		return False

	try:
		if not reencode:
			# Use list form to properly handle unicode in filenames
			ffmpeg_command = ['ffmpeg', '-y', '-loglevel', '16', '-ss', start_pos, '-i', input_file, '-t', str(duration), '-c', 'copy', '-avoid_negative_ts', '1', output_file]
			debug_print(f"ffmpeg_command is '{' '.join(ffmpeg_command)}'")
			result = subprocess.run(ffmpeg_command, encoding='utf-8', capture_output=True)
		else:
			result = subprocess.run(['ffmpeg', '-y', '-loglevel', '16', '-ss', start_pos, '-i', input_file, '-t', str(duration), output_file], encoding='utf-8', capture_output=True)
		
		# Check if ffmpeg succeeded
		if result.returncode != 0:
			print(f"! ERROR ffmpeg failed with exit code {result.returncode}")
			print(f"  Input: '{input_file}', Output: '{output_file}'")
			print(f"  Timestamps: {start_pos} to {end_pos}")
			if result.stderr:
				print(f"  ffmpeg error: {result.stderr.strip()}")
			return False
		
		# Verify output file was created
		if not os.path.isfile(output_file):
			print(f"! ERROR ffmpeg completed but output file was not created: '{output_file}'")
			return False
		
		verbose_print(f"+ Generated video '{output_file}' successfully.\n File size: {format_filesize(os.path.getsize(output_file))}\n Expected duration: {duration} s\n")
		return True
	except FileNotFoundError:
		print("! ERROR ffmpeg is not installed or not found in system PATH.")
		print("  Please install ffmpeg and ensure it's in your PATH.")
		print("  Download from: https://www.ffmpeg.org/download.html")
		return False
	except OSError as e:
		print(f"! ERROR ffmpeg could not successfully run.")
		print(f"  Error: {e}")
		print(f"  Working directory: '{os.getcwd()}'")
		print(f"  Input file: '{input_file}'")
		print(f"  Output file: '{output_file}'")
		return False

def get_file_duration(filepath):
	"""Calls ffprobe, returns duration of video container in seconds.
	
	Returns the duration in seconds, or None if the file cannot be probed.
	"""
	# Check if file exists before attempting to probe
	if not os.path.isfile(filepath):
		print(f"! ERROR Video file not found: '{filepath}'")
		print(f"  Expected location: {os.path.join(os.getcwd(), filepath)}")
		print("  Please ensure the video file exists in the working directory.")
		return None
	
	probe_command = ['ffprobe', '-v', 'error', '-show_entries', 'format=duration', '-of', 'default=noprint_wrappers=1:nokey=1', filepath]
	debug_print(f"probe_command is {' '.join(probe_command)}")
	
	try:
		file_length = float(subprocess.check_output(probe_command, encoding='utf-8'))
		return int(file_length)
	except FileNotFoundError:
		print("! ERROR ffprobe is not installed or not found in system PATH.")
		print("  Please install ffmpeg (which includes ffprobe) and ensure it's in your PATH.")
		print("  Download from: https://www.ffmpeg.org/download.html")
		return None
	except subprocess.CalledProcessError as e:
		print(f"! ERROR ffprobe failed to read video file: '{filepath}'")
		print(f"  ffprobe exit code: {e.returncode}")
		print("  The file may be corrupted, not a valid video, or in an unsupported format.")
		return None
	except ValueError as e:
		print(f"! ERROR Could not parse duration from video file: '{filepath}'")
		print(f"  ffprobe returned unexpected output. Error: {e}")
		return None

def get_duration(start_time, end_time):
	"""Returns the duration of a clip as seconds, or None if timestamps are invalid."""
	debug_print(f'start_time is {start_time} with length {len(start_time)}, end_time is {end_time}')
	
	# Handle case where add_duration() returned -1 (error)
	if end_time == -1:
		print(f"! ERROR Invalid end timestamp (derived from start: '{start_time}')")
		print("  Could not calculate end time. Check the timestamp format.")
		return None
	
	formats = ['%M:%S', '%H:%M:%S'] if len(str(start_time)) <= 5 else ['%H:%M:%S', '%M:%S']
	
	for fmt in formats:
		try:
			start_datetime = datetime.strptime(str(start_time), fmt)
			end_datetime = datetime.strptime(str(end_time), fmt)
			return int((end_datetime - start_datetime).total_seconds())
		except ValueError:
			continue
	
	print(f"! ERROR Timestamp formatting error in get_duration().")
	print(f"  Start time: '{start_time}', End time: '{end_time}'")
	print("  Accepted formats: HH:MM:SS, MM:SS, or M:SS (e.g., 1:23:45, 12:34, 1:23)")
	return None

def add_duration(start_time):
	"""Adds one minute to the given timestamp.
	
	Returns the new timestamp string, or -1 if the timestamp format is invalid.
	"""
	try:
		if len(start_time) <= 5:
			start_datetime = datetime.strptime(str(start_time), '%M:%S')
			new_time = start_datetime + timedelta(seconds=DEFAULT_DURATION_SECONDS)
			return new_time.strftime('%M:%S')
		else:
			start_datetime = datetime.strptime(start_time, '%H:%M:%S')
			new_time = start_datetime + timedelta(seconds=DEFAULT_DURATION_SECONDS)
			return new_time.strftime('%H:%M:%S')
	except ValueError:
		print(f"! WARNING Could not parse single timestamp '{start_time}' to add default duration.")
		print(f"  Expected format: MM:SS or HH:MM:SS (e.g., 12:34 or 1:23:45)")
		print("  This timestamp will be skipped.")
		return -1

# ============================================================================
# Google Sheets API
# ============================================================================

def get_all_spreadsheets(connection):
	"""Returns comma-separated list of all accessible Google Spreadsheets."""
	docs = []
	for doc in connection.list_spreadsheet_files():
		debug_print(str(doc))
		docs.append(doc['name'])
	return ', '.join(docs)

def find_spreadsheet_by_name(search_name, doc_list):
	"""Find a matching Google Sheet name from doc_list.
	Returns the index of matching sheet, or -1 if not found."""
	debug_print('Running method find_spreadsheet_by_name()')
	search_name = search_name.strip().lower()
	search_name_guess = search_name + ' data set'
	debug_print(f"Using search_name '{search_name}', search_name_guess '{search_name_guess}'")
	
	for i, doc in enumerate(doc_list):
		doc_name = doc.strip().lower()
		debug_print(f"Attempting match with '{doc}', formatted as '{doc_name}'")
		if doc_name == search_name:
			debug_print(f"Matched sheet '{doc_name}' with input '{search_name}'")
			return i
		elif doc_name == search_name_guess:
			debug_print(f"Matched sheet '{doc_name}' with guess '{search_name_guess}'")
			return i
		else:
			debug_print(f'Found nothing at step {i}')
	return -1

def connect_to_google_service_account():
	scopes = ['https://spreadsheets.google.com/feeds',
	 		 'https://www.googleapis.com/auth/drive']
	credentials_path = os.path.join(os.getcwd(), 'credentials.json')
	try:
		credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scopes=scopes)  # pyright: ignore[reportArgumentType]
	except IOError as e:
		print(f"! ERROR Could not find credentials file.")
		print(f"  Expected location: {credentials_path}")
		print("  Please ensure 'credentials.json' is in the same directory as clipgen.py.")
		print("  See Google's documentation for creating service account credentials:")
		print("  https://docs.gspread.org/en/latest/oauth2.html#service-account")
		sys.exit(1)
	return credentials

# ============================================================================
# Main Program
# ============================================================================

def select_spreadsheet(gc, doc_list):
	"""Interactive spreadsheet selection. Returns the selected worksheet."""
	input_file_fails = 0
	
	while True:
		input_name = input("\nPlease enter the index, name, URL or key of the spreadsheet ('all' for list, 'new' for list of newest, 'last' to immediately open latest, 'settings' to change settings):\n>> ")
		try:
			if input_name[:4] == 'http':
				return gc.open_by_url(input_name).worksheet(SHEET_NAME)
			elif input_name[:3] == 'all':
				print('\nAvailable documents:')
				for i, doc in enumerate(doc_list):
					print(f'{i+1}. {doc.strip()}')
			elif input_name[:3] == 'new':
				print('\nNewest documents: (modified or opened most recently)')
				for i in range(min(3, len(doc_list))):
					print(f'{i+1}. {doc_list[i].strip()}')
			elif input_name[:4] == 'last':
				latest = get_all_spreadsheets(gc).split(',')[0]
				return gc.open(latest).worksheet(SHEET_NAME)
			elif input_name[0].isdigit():
				chosen_index = int(input_name) - 1
				print(f'Opening document: {doc_list[chosen_index].strip()}')
				return gc.open(doc_list[chosen_index].strip()).worksheet(SHEET_NAME)
			elif input_name[:8] == 'settings':
				set_program_settings()
			else:
				chosen_index = find_spreadsheet_by_name(input_name, doc_list)
				if chosen_index >= 0:
					return gc.open(get_all_spreadsheets(gc).split(',')[chosen_index].strip().lstrip()).worksheet(SHEET_NAME)
		except (gspread.SpreadsheetNotFound, gspread.exceptions.APIError, gspread.exceptions.GSpreadException) as e:
			input_file_fails += 1
			if input_file_fails == 1:
				print(f"\n! ERROR Could not access spreadsheet: {e}")
				print("  Please try again. Type 'all' to see available documents.")
			elif input_file_fails == 2:
				print("\n! ERROR Spreadsheet not found or not accessible.")
				print("  Common causes:")
				print("  - The spreadsheet name is misspelled")
				print("  - The spreadsheet hasn't been shared with your service account")
				print("    (Share it with the email in credentials.json 'client_email' field)")
				print("  - The spreadsheet doesn't have a worksheet named 'Sheet1'")
				print("\n  Type 'all' to see accessible documents, or 'new' for recent ones.")
			else:
				print(f"\n! ERROR {e}")
				print("  Tip: Use the document index number (1, 2, 3...) from the 'all' list.")

def select_mode_and_generate(worksheet):
	"""Interactive mode selection. Returns the clips list for processing."""
	mode_map = {
		'b': 'batch', 'batch': 'batch',
		'l': 'line', 'line': 'line',
		'r': 'range', 'range': 'range',
		'c': 'category', 'cat': 'category', 'category': 'category',
		'test': 'test'
	}
	
	while True:
		input_mode = input('\nSelect mode: (b)atch, (r)ange, (c)ategory or (l)ine\n>> ').strip().lower()
		
		if not input_mode:
			print("  Please enter a mode (b, r, c, or l).")
			continue
		
		try:
			# Check first character or full word
			mode = mode_map.get(input_mode[0]) or mode_map.get(input_mode)
			if mode:
				return generate_list(worksheet, mode)
			else:
				print(f"  Unknown mode '{input_mode}'. Available modes:")
				print("    b or batch   - Generate all clips in the spreadsheet")
				print("    r or range   - Generate clips from a range of rows")
				print("    c or category - Generate clips by category")
				print("    l or line    - Generate clips from specific line(s)")
		except gspread.exceptions.GSpreadException as e:
			print(f"! ERROR Google Sheets API error: {e}")
			debug_print(f"ERROR Message '{e}', Attempting reconnect")

def process_clips(clips_list):
	"""Process and generate video clips from the clips list. Returns count of videos generated."""
	# Check if clips_list is empty
	if not clips_list:
		print("! WARNING No clips to process. No timestamps were found or selected.")
		return 0
	
	verbose_print('\n* ffmpeg is set to never prompt for input and will always overwrite.\n  Only warns if close to crashing.\n')
	videos_generated = 0
	videos_skipped = 0
	missing_videos = set()  # Track unique missing video files

	for clip in clips_list:
		clip = clean_issue(clip)
		
		# Skip if no valid timestamps were parsed
		if not clip['times']:
			videos_skipped += 1
			continue
		
		base_video = f"{clip['study']}_{clip['participant']}{FILEFORMAT}"
		
		# Check if base video exists (only warn once per unique file)
		if not os.path.isfile(base_video) and base_video not in missing_videos:
			missing_videos.add(base_video)
			print(f"! ERROR Source video file not found: '{base_video}'")
			print(f"  Expected location: {os.path.join(os.getcwd(), base_video)}")
			print(f"  Clips for participant '{clip['participant']}' in study '{clip['study']}' will be skipped.")
		
		for vid_in, vid_out in clip['times']:
			try:
				# Ensure all components are strings and handle unicode properly
				vid_name = get_unique_filename(
					f"[{clip['category']}] {clip['study']} {clip['participant']} {clip['desc']}{FILEFORMAT}"
				)
			except (TypeError, UnicodeEncodeError, UnicodeDecodeError) as e:
				print(f'! ERROR Character encoding issue occurred:\n  {e}')
				print(f"  Category: '{clip['category']}', Study: '{clip['study']}', Participant: '{clip['participant']}'")
				print("  Try simplifying the description or category names to use only ASCII characters.")
				videos_skipped += 1
				break

			completed = run_ffmpeg(
				input_file=base_video,
				output_file=vid_name,
				start_pos=vid_in,
				end_pos=vid_out,
				reencode=REENCODING
			)
			if completed:
				videos_generated += 1
			else:
				videos_skipped += 1
	
	# Report summary if any videos were skipped
	if videos_skipped > 0:
		verbose_print(f"\n* Summary: {videos_generated} video(s) generated, {videos_skipped} skipped due to errors.")
	if missing_videos:
		verbose_print(f"* Missing source video files: {len(missing_videos)}")

	return videos_generated

def main():
	# Ensure UTF-8 encoding for stdout/stderr to handle unicode properly
	if sys.stdout.encoding.lower() != 'utf-8':
		import io
		sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
		sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
	
	# Parse command-line arguments
	args = parse_arguments()
	
	# Determine if running in CLI mode (any mode argument provided)
	cli_mode = args.batch or args.lines or args.range
	
	# Set verbose mode: silent by default in CLI mode, verbose in interactive mode
	global VERBOSE
	VERBOSE = not cli_mode or args.verbose
	
	# Parse CLI arguments for line and range modes
	cli_line_numbers = None
	cli_range_start = None
	cli_range_end = None
	
	if args.lines:
		try:
			# Support both + and , as separators
			line_str = args.lines.replace(',', '+')
			cli_line_numbers = [int(num.strip()) for num in line_str.split('+')]
		except ValueError:
			print(f'Error: Invalid line numbers "{args.lines}". Use format: 1+4+5 or 1,4,5')
			sys.exit(1)
	
	if args.range:
		try:
			parts = args.range.split('-')
			if len(parts) != 2:
				raise ValueError('Range must have exactly two parts')
			cli_range_start = int(parts[0].strip())
			cli_range_end = int(parts[1].strip())
			if cli_range_start > cli_range_end:
				print(f'Error: Range start ({cli_range_start}) must be less than or equal to end ({cli_range_end})')
				sys.exit(1)
		except ValueError as e:
			print(f'Error: Invalid range "{args.range}". Use format: 1-10')
			sys.exit(1)
	
	# Change working directory to place of python script
	os.chdir(os.path.dirname(os.path.abspath(__file__)))
	verbose_print('-------------------------------------------------------------------------------')
	verbose_print(f'Welcome to clipgen v{VERSIONNUM}\nWorking directory: {os.getcwd()}\nPlace video files and the credentials.json file in this directory.')
	debug_print('Debug mode is ON. Several limitations apply and more things will be printed.')
	
	# Authenticate with Google
	try:
		debug_print('Attempting login...')
		gc = gspread.oauth(credentials_filename='credentials.json')
		debug_print('Login successful!')
	except gspread.exceptions.GSpreadException as e:
		print(f"! ERROR Could not authenticate with Google.")
		print(f"  Error details: {e}")
		print(f"  Credentials file location: {os.path.join(os.getcwd(), 'credentials.json')}")
		print("\n  Troubleshooting steps:")
		print("  1. Ensure 'credentials.json' exists in the working directory")
		print("  2. Verify the credentials file is valid JSON")
		print("  3. Check that the service account has access to Google Sheets API")
		print("  4. For OAuth flow, delete any existing token files and re-authenticate")
		sys.exit(1)

	# Get document list and select spreadsheet
	doc_list = get_all_spreadsheets(gc).split(',')

	# Spreadsheet selection
	worksheet = None
	if args.spreadsheet:
		# CLI-specified spreadsheet
		try:
			if args.spreadsheet.startswith('http'):
				worksheet = gc.open_by_url(args.spreadsheet).worksheet(SHEET_NAME)
			elif args.spreadsheet.isdigit():
				chosen_index = int(args.spreadsheet) - 1
				verbose_print(f'Opening document: {doc_list[chosen_index].strip()}')
				worksheet = gc.open(doc_list[chosen_index].strip()).worksheet(SHEET_NAME)
			else:
				chosen_index = find_spreadsheet_by_name(args.spreadsheet, doc_list)
				if chosen_index >= 0:
					matched_name = doc_list[chosen_index].strip()
					verbose_print(f'Opening document: {matched_name}')
					worksheet = gc.open(matched_name).worksheet(SHEET_NAME)
				else:
					print(f'Error: Could not find spreadsheet "{args.spreadsheet}"')
					sys.exit(1)
		except (gspread.SpreadsheetNotFound, gspread.exceptions.APIError, gspread.exceptions.GSpreadException) as e:
			print(f'Error: Could not open spreadsheet "{args.spreadsheet}": {e}')
			sys.exit(1)
	else:
		# Auto-connect if working directory name matches a spreadsheet
		cwd_name = os.path.basename(os.getcwd())
		auto_match_index = find_spreadsheet_by_name(cwd_name, doc_list)
		if auto_match_index >= 0:
			matched_name = doc_list[auto_match_index].strip()
			verbose_print(f'\nAuto-connecting to spreadsheet: {matched_name}')
			worksheet = gc.open(matched_name).worksheet(SHEET_NAME)
		elif cli_mode:
			# CLI mode requires a spreadsheet - can't prompt interactively
			print('Error: No spreadsheet found matching working directory name.')
			print('Use -s to specify a spreadsheet name, URL, or index.')
			sys.exit(1)
		else:
			worksheet = select_spreadsheet(gc, doc_list)
	
	verbose_print('\nConnected to Google Drive!')

	if cli_mode:
		# CLI mode - run once and exit
		skip_prompts = args.yes
		
		if args.batch:
			clips_list = generate_list(worksheet, 'batch', skip_prompts=skip_prompts)
		elif args.lines:
			clips_list = generate_list(worksheet, 'line', line_numbers=cli_line_numbers, skip_prompts=skip_prompts)
		elif args.range:
			clips_list = generate_list(worksheet, 'range', range_start=cli_range_start, range_end=cli_range_end, skip_prompts=skip_prompts)
		
		videos_generated = process_clips(clips_list)
		
		if not REENCODING:
			verbose_print('* No re-encoding done, expect:\n- inaccurate start and end timings\n- lossy frames until first keyframe\n- bad timecodes at the end\n')
		print(f'All done, created {videos_generated} videos!\nFiles are in {os.getcwd()}\n')
	else:
		# Interactive mode - main processing loop
		while True:
			clips_list = select_mode_and_generate(worksheet)
			videos_generated = process_clips(clips_list)

			if not REENCODING:
				print('* No re-encoding done, expect:\n- inaccurate start and end timings\n- lossy frames until first keyframe\n- bad timecodes at the end\n')
			print(f'All done, created {videos_generated} videos!\nFiles are in {os.getcwd()}\n')
			
			yn = input('Continue working (y) or quit the program (n)? y/n\n>> ')
			if yn == 'n':
				break

if __name__ == '__main__':
	try:
		main()
	except KeyboardInterrupt:
		print('\nInterrupted by user')
		sys.exit(0)
