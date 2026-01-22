# -*- coding: utf-8 -*-
"""clipgen - Video clip generator from Google Sheets timestamps.

This script supports full unicode/UTF-8 for international characters in:
- Study names
- Participant IDs  
- Category names
- Descriptions
- File paths
"""
import os
import sys
import subprocess
from datetime import datetime, timedelta

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configuration Constants
REENCODING = False
FILEFORMAT = '.mp4'
VERSIONNUM = '0.4.4'
SHEET_NAME = 'Sheet1'
DEBUGGING  = False

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

# What is this program?
# This script will help quickly create video snippets from longer video files, based on timestamps in a spreadsheet!
# Check out README.md for more detailed information about clipgen.

# ============================================================================
# Utility Functions
# ============================================================================

def debug_print(message):
	"""Print debug messages when DEBUGGING is enabled."""
	if DEBUGGING:
		print(f'! DEBUG {message}')

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

def parse_timestamps(cell_value):
	"""Parse timestamp pairs from a cell value string.
	
	Returns a list of (start_time, end_time) tuples.
	"""
	parsed_timestamps = []
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
		elif ':' in raw_times[i]:
			colon_pos = raw_times[i].find(':')
			if colon_pos > 0 and raw_times[i][colon_pos-1].isdigit():
				# Single timestamp - add default end time to trigger add_duration later
				time_pair = (raw_times[i], add_duration(raw_times[i]))
				parsed_timestamps.append(time_pair)

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

def generate_list(sheet, mode):
	"""Goes through a sheet, bundles values from timestamp columns and descriptions columns into tuples."""
	id_cell = sheet.find(ID_HEADER)
	observation_cell = sheet.find(OBSERVATION_HEADER)
	category_cell = sheet.find(CATEGORY_HEADER)
	timestamps = []

	# Sheet data is a list of lists, which forms a matrix
	# - sheet_data[row][col] where indices start at 0 (real spreadsheet starts at 1)
	sheet_data = sheet.get_all_values()
	debug_print(f'Sheet dumped into memory at {get_current_time()}')

	# Determine the study name.
	study_name = sheet_data[0][0]
	if study_name == '':
		study_name = sheet.spreadsheet.title
	print(f'\nBeginning work on {study_name}.')

	# Normalize study name for filesystem use
	study_name = normalize_study_name(study_name)

	# Get number of participants needed to loop through the worksheet
	num_participants = get_num_participants(sheet.row_values(id_cell.row), id_cell, sheet.col_count)

	# Generate the timestamps, according to the selected mode.
	if mode == 'batch':
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
		timestamps = generate_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name)
	elif mode == 'range':
		while True:
			try:
				start_line = int(input('\nWhich starting line (row number only)?\n>> '))
				end_line = int(input('\nWhich ending line (row number only)?\n>> '))
			except ValueError:
				print('\nInvalid input. Please enter row numbers as integers.')
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
	print(f'Found {num_participants} participants in total, spanning columns {id_cell.col+1} to {num_participants+id_cell.col+1}.')
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

def generate_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name):
	"""Generate videos for one or more line/row numbers (comma-separated)."""
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

	timestamps = []
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
			issue = {
				'cell': cell,
				'desc': sheet_data[line_index][observation_cell.col-1],
				'study': study_name,
				'participant': sheet_data[id_cell.row-1][col_index],
				'category': sheet_data[line_index][observation_cell.col-2]
			}
			debug_print(f"Participant ID at R{id_cell.row},C{col_index} -> '{sheet_data[id_cell.row][col_index]}'")
			debug_print(f"Description at R{line_index},C{observation_cell.col-1} -> '{sheet_data[line_index][observation_cell.col-1]}'")
			debug_print(f"Timestamp at R{cell.row-1},C{cell.col-1} -> '{cell.value}'")
			debug_print(f'Actual cell {cell} at actual address {gspread.utils.rowcol_to_a1(cell.row, cell.col)}')
			timestamps.append(issue)
			print(f"+ Found timestamp: {value.replace(chr(10), ' ')}")

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
	
	# Parse timestamps from cell value
	issue['times'] = parse_timestamps(issue['cell'].value)

	# Clean description: remove bracketed prefix and sanitize
	desc = issue['desc'][issue['desc'].rfind(']')+1:].strip()
	issue['desc'] = sanitize_filename(desc)
	
	# Sanitize category
	issue['category'] = sanitize_filename(issue['category'])

	return issue

# ============================================================================
# Video Operations
# ============================================================================

def run_ffmpeg(input_file, output_file, start_pos, end_pos, reencode):
	"""Calls ffmpeg to cut a video clip. Requires ffmpeg in system PATH.
	
	Returns True if video was generated successfully, False otherwise.
	"""
	duration = get_duration(start_pos, end_pos)
	file_length = get_file_duration(input_file)

	if duration < 0:
		print("Can't work with negative duration for videos. Skipping.")
		return False
	if duration > file_length:
		print('Timestamp duration longer than actual video file. Skipping.')
		return False
	if duration > MAX_CLIP_DURATION_SECONDS:
		yn = input('The generated video will be over 10 minutes long, do you want to still generate it? (y/n)\n>> ')
		if yn == 'n':
			return False

	print(f'Cutting {input_file} from {start_pos} to {end_pos}.')
	if DEBUGGING:
		debug_print(f'Debugging enabled, not calling ffmpeg.\n  input_file: {input_file},\n  output_file: {output_file}')
		return False

	try:
		if not reencode:
			# Use list form to properly handle unicode in filenames
			ffmpeg_command = ['ffmpeg', '-y', '-loglevel', '16', '-ss', start_pos, '-i', input_file, '-t', str(duration), '-c', 'copy', '-avoid_negative_ts', '1', output_file]
			debug_print(f"ffmpeg_command is '{' '.join(ffmpeg_command)}'")
			subprocess.run(ffmpeg_command, encoding='utf-8')
		else:
			subprocess.run(['ffmpeg', '-y', '-loglevel', '16', '-ss', start_pos, '-i', input_file, '-t', str(duration), output_file], encoding='utf-8')
		print(f"+ Generated video '{output_file}' successfully.\n File size: {format_filesize(os.path.getsize(output_file))}\n Expected duration: {duration} s\n")
		return True
	except OSError as e:
		print(f"\n! ERROR ffmpeg could not successfully run.\n  clipgen returned the following error:\n  {e}\n  - Attempted location: '{os.getcwd()}'\n  - Attemped input_file: '{input_file}',\n  - Attempted output_file: '{output_file}'\n")
		return False

def get_file_duration(filepath):
	"""Calls ffprobe, returns duration of video container in seconds."""
	probe_command = ['ffprobe', '-v', 'error', '-show_entries', 'format=duration', '-of', 'default=noprint_wrappers=1:nokey=1', filepath]
	debug_print(f"probe_command is {' '.join(probe_command)}")
	file_length = float(subprocess.check_output(probe_command, encoding='utf-8'))
	return int(file_length)

def get_duration(start_time, end_time):
	"""Returns the duration of a clip as seconds."""
	debug_print(f'start_time is {start_time} with length {len(start_time)}, end_time is {end_time}')
	
	formats = ['%M:%S', '%H:%M:%S'] if len(start_time) <= 5 else ['%H:%M:%S', '%M:%S']
	
	for fmt in formats:
		try:
			start_datetime = datetime.strptime(str(start_time), fmt)
			end_datetime = datetime.strptime(str(end_time), fmt)
			return int((end_datetime - start_datetime).total_seconds())
		except ValueError:
			continue
	
	print('* Timestamp formatting error in get_duration(). Exiting.')
	print('* Formats must match: HH:MM:SS, MM:SS, or M:SS')
	sys.exit(0)

def add_duration(start_time):
	"""Adds one minute to the given timestamp."""
	try:
		if len(start_time) <= 5:
			start_datetime = datetime.strptime(str(start_time), '%M:%S')
			new_time = start_datetime + timedelta(seconds=DEFAULT_DURATION_SECONDS)
			return new_time.strftime('%M:%S')
		else:
			start_datetime = datetime.strptime(start_time, '%H:%M:%S')
			new_time = start_datetime + timedelta(seconds=DEFAULT_DURATION_SECONDS)
			return new_time.strftime('%H:%M:%S')
	except ValueError as e:
		print('* Timestamp formatting error was caught while running add_duration().\n  Returning -1 instead of timestamp')
		print(e)
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
	try:
		credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scopes=scopes)  # pyright: ignore[reportArgumentType]
	except IOError as e:
		print(f'{e}\nCould not find credentials (credentials.json).')
		sys.exit(0)
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
			if input_file_fails <= 1 or input_file_fails >= 3:
				print(e)
				print('\nDid not find spreadsheet. Please try again.')
			else:
				print('\n###')
				print("Did not find spreadsheet. Please try again.\n\nRemember that you need to share the spreadsheet you want to parse.\nShare it with the user listed in the json-file (value of client_email).")
				print("This needs to be done on a per-document basis.\n\nSee available documents by typing 'all' or 'new'")
				print('###\n')

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
		input_mode = input('\nSelect mode: (b)atch, (r)ange, (c)ategory or (l)ine\n>> ')
		try:
			# Check first character or full word
			mode = mode_map.get(input_mode[0]) or mode_map.get(input_mode)
			if mode:
				return generate_list(worksheet, mode)
		except (IndexError, gspread.exceptions.GSpreadException) as e:
			debug_print(f"ERROR Message '{e}', Attempting reconnect")

def process_clips(clips_list):
	"""Process and generate video clips from the clips list. Returns count of videos generated."""
	print('\n* ffmpeg is set to never prompt for input and will always overwrite.\n  Only warns if close to crashing.\n')
	videos_generated = 0

	for clip in clips_list:
		clip = clean_issue(clip)
		
		for vid_in, vid_out in clip['times']:
			try:
				# Ensure all components are strings and handle unicode properly
				vid_name = get_unique_filename(
					f"[{clip['category']}] {clip['study']} {clip['participant']} {clip['desc']}{FILEFORMAT}"
				)
			except (TypeError, UnicodeEncodeError, UnicodeDecodeError) as e:
				print(f'! ERROR Character encoding issue occurred:\n  {e}')
				print(f"  Category: {clip['category']}, Study: {clip['study']}, Participant: {clip['participant']}")
				break

			base_video = f"{clip['study']}_{clip['participant']}{FILEFORMAT}"

			completed = run_ffmpeg(
				input_file=base_video,
				output_file=vid_name,
				start_pos=vid_in,
				end_pos=vid_out,
				reencode=REENCODING
			)
			if completed:
				videos_generated += 1

	return videos_generated

def main():
	# Ensure UTF-8 encoding for stdout/stderr to handle unicode properly
	if sys.stdout.encoding.lower() != 'utf-8':
		import io
		sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
		sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
	
	# Change working directory to place of python script
	os.chdir(os.path.dirname(os.path.abspath(__file__)))
	print('-------------------------------------------------------------------------------')
	print(f'Welcome to clipgen v{VERSIONNUM}\nWorking directory: {os.getcwd()}\nPlace video files and the credentials.json file in this directory.')
	debug_print('Debug mode is ON. Several limitations apply and more things will be printed.')
	
	# Authenticate with Google
	try:
		debug_print('Attempting login...')
		gc = gspread.oauth(credentials_filename='credentials.json')
		debug_print('Login successful!')
	except gspread.exceptions.GSpreadException as e:
		print(f'{e}\n! ERROR Could not authenticate.\n')
		sys.exit(0)

	# Get document list and select spreadsheet
	doc_list = get_all_spreadsheets(gc).split(',')

	# Auto-connect if working directory name matches a spreadsheet
	cwd_name = os.path.basename(os.getcwd())
	auto_match_index = find_spreadsheet_by_name(cwd_name, doc_list)
	if auto_match_index >= 0:
		matched_name = doc_list[auto_match_index].strip()
		print(f'\nAuto-connecting to spreadsheet: {matched_name}')
		worksheet = gc.open(matched_name).worksheet(SHEET_NAME)
	else:
		worksheet = select_spreadsheet(gc, doc_list)
	print('\nConnected to Google Drive!')

	# Main processing loop
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
