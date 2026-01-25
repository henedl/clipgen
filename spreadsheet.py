# -*- coding: utf-8 -*-
"""Spreadsheet data processing for clipgen."""

import webbrowser

import gspread
from icecream import ic

import config
import google_api
import utils


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
	ic(mode, line_numbers, range_start, range_end)
	# Find required headers
	id_cell = sheet.find(config.ID_HEADER)
	observation_cell = sheet.find(config.OBSERVATION_HEADER)
	category_cell = sheet.find(config.CATEGORY_HEADER)
	ic(id_cell, observation_cell, category_cell)
	timestamps = []
	
	# Validate required headers exist
	missing_headers = []
	if id_cell is None:
		missing_headers.append(f"'{config.ID_HEADER}'")
	if observation_cell is None:
		missing_headers.append(f"'{config.OBSERVATION_HEADER}'")
	if category_cell is None:
		missing_headers.append(f"'{config.CATEGORY_HEADER}'")
	
	if missing_headers:
		print(f"! ERROR Required header(s) not found in spreadsheet: {', '.join(missing_headers)}")
		print(f"  The spreadsheet must contain columns with these exact headers: {config.ID_HEADER}, {config.OBSERVATION_HEADER}, {config.CATEGORY_HEADER}")
		print(f"  Please check your spreadsheet structure.")
		return []

	# Sheet data is a list of lists, which forms a matrix
	# - sheet_data[row][col] where indices start at 0 (real spreadsheet starts at 1)
	sheet_data = sheet.get_all_values()
	utils.debug_print(f'Sheet dumped into memory at {utils.get_current_time()}')
	
	# Check if sheet is empty or has only headers
	if len(sheet_data) <= 1:
		print("! ERROR Spreadsheet appears to be empty (no data rows found).")
		print(f"  The spreadsheet only has {len(sheet_data)} row(s).")
		return []

	# Determine the study name.
	study_name = sheet_data[0][0]
	if study_name == '':
		study_name = sheet.spreadsheet.title
	utils.verbose_print(f'\nBeginning work on {study_name}.')

	# Normalize study name for filesystem use
	study_name = utils.normalize_study_name(study_name)

	# Get number of participants needed to loop through the worksheet
	num_participants = get_num_participants(sheet.row_values(id_cell.row), id_cell, sheet.col_count)
	
	# Warn if no participants found
	if num_participants == 0:
		print(f"! WARNING No participant columns found in the spreadsheet.")
		print(f"  Looking for columns starting with: {', '.join(config.PARTICIPANT_PREFIXES)}")
		print(f"  Check that participant column headers start with 'P' or 'G' (e.g., P01, P02, G01).")
		return []

	# Generate the timestamps, according to the selected mode.
	if mode == 'batch':
		if skip_prompts:
			utils.verbose_print('Batch mode: generating all possible clips...')
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
			utils.verbose_print(f'Range mode: lines {range_start} to {range_end}')
			utils.verbose_print(f'Lines selected: {sheet_data[range_start-1][observation_cell.col-1]} to {sheet_data[range_end-1][observation_cell.col-1]}')
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
			if header_row[j][0] in config.PARTICIPANT_PREFIXES:
				num_participants += 1
			elif header_row[j] == config.NOTES_COLUMN:
				break
	utils.verbose_print(f'Found {num_participants} participants in total, spanning columns {id_cell.col+1} to {num_participants+id_cell.col+1}.')
	return num_participants

def generate_batch_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name):
	utils.debug_print('Running method generate_batch_timestamps()')
	timestamps = []
	for i in range(id_cell.row+1, len(sheet_data)):
		utils.debug_print(f'Batching on line {i} (real sheet line {i+1})')
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
	utils.debug_print('Starting method generate_category_timestamps()')
	timestamps = []
	category_col = category_cell.col - 1  # Convert from 1-indexed to 0-indexed
	
	# Start from the row after the category header
	for i in range(category_cell.row, len(sheet_data)):
		row_category = sheet_data[i][category_col].strip()
		if row_category in selected_categories:
			utils.debug_print(f"Row {i+1} matches category '{row_category}'")
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
		utils.verbose_print(f'\nLine mode: processing lines {", ".join(str(n) for n in cli_line_numbers)}')
		utils.verbose_print('\nSelected issues:')
		for line_num in cli_line_numbers:
			if line_num < 1 or line_num > len(sheet_data):
				utils.verbose_print(f'  Line {line_num}: [INVALID - out of range]')
			else:
				desc = sheet_data[line_num-1][observation_cell.col-1]
				utils.verbose_print(f'  Line {line_num}: {desc}')
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
		utils.debug_print(f'Calling get_line_timestamps() from generate_line_timestamps() for line {line_num}')
		line_timestamps = get_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, line_num-1, study_name)
		timestamps.extend(line_timestamps)
	
	utils.debug_print(f'Printing return of get_line_timestamps() in generate_line_timestamps(): {len(timestamps)} total timestamps')
	utils.debug_print(str(timestamps))

	return timestamps

def get_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, line_index, study_name):
	ic(line_index, num_participants, study_name)
	utils.debug_print(f'Running method get_line_timestamps, starting line index {line_index} (real sheet line {line_index+1})')
	
	# Bounds checking
	if line_index < 0 or line_index >= len(sheet_data):
		ic(line_index, len(sheet_data))
		print(f"! ERROR Line index {line_index} (row {line_index+1}) is out of bounds.")
		print(f"  Spreadsheet has {len(sheet_data)} rows.")
		return []

	timestamps = []
	try:
		for col_index, value in enumerate(sheet_data[line_index]):
			utils.debug_print(f"Item {col_index} with value '{value}' being processed.")
			if col_index < id_cell.col:
				utils.debug_print(f"Skipping item {col_index} with value '{value}'")
			elif col_index == id_cell.col + num_participants:
				utils.debug_print(f'Exit for-loop, reached final column {col_index} (real sheet column {col_index+1}).')
				break
			elif value is None or value == '':
				pass
			else:
				cell = gspread.cell.Cell(line_index+1, col_index+1, value)
				utils.debug_print(f'Found something at step {col_index}')
				utils.debug_print(f'study_name is {study_name}')
				
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
				
				ic(participant, desc, category)
				issue = {
					'cell': cell,
					'desc': desc,
					'study': study_name,
					'participant': participant,
					'category': category
				}
				ic(issue)
				utils.debug_print(f"Participant ID at R{id_cell.row},C{col_index} -> '{participant}'")
				utils.debug_print(f"Description at R{line_index},C{observation_cell.col-1} -> '{desc}'")
				utils.debug_print(f"Timestamp at R{cell.row-1},C{cell.col-1} -> '{cell.value}'")
				utils.debug_print(f'Actual cell {cell} at actual address {gspread.utils.rowcol_to_a1(cell.row, cell.col)}')
				timestamps.append(issue)
				utils.verbose_print(f"+ Found timestamp: {value.replace(chr(10), ' ')}")
	except IndexError as e:
		ic(e, line_index)
		print(f"! ERROR Index error while reading row {line_index+1}: {e}")
		print("  The spreadsheet structure may be malformed.")

	utils.debug_print(f'Line completed, returning list of {len(timestamps)} potential timestamps.')
	ic(timestamps)
	return timestamps

def generate_range_timestamps(sheet_data, id_cell, observation_cell, num_participants, study_name, start_line, end_line):
	timestamps = []
	for i in range(start_line-1, end_line):
		utils.debug_print(f'Batching on line {i}')
		timestamps.extend(get_line_timestamps(sheet_data, id_cell, observation_cell, num_participants, i, study_name))
	return timestamps

def browse_spreadsheet(sheet):
	"""Interactive browse mode for viewing spreadsheet rows line by line.
	
	Allows users to navigate through the spreadsheet to inspect issues
	before generating clips. Shows row number, category, description,
	and participant/group timestamps for each row.
	"""
	# Find required headers
	id_cell = sheet.find(config.ID_HEADER)
	observation_cell = sheet.find(config.OBSERVATION_HEADER)
	category_cell = sheet.find(config.CATEGORY_HEADER)
	
	# Validate required headers exist
	missing_headers = []
	if id_cell is None:
		missing_headers.append(f"'{config.ID_HEADER}'")
	if observation_cell is None:
		missing_headers.append(f"'{config.OBSERVATION_HEADER}'")
	if category_cell is None:
		missing_headers.append(f"'{config.CATEGORY_HEADER}'")
	
	if missing_headers:
		print(f"! ERROR Required header(s) not found in spreadsheet: {', '.join(missing_headers)}")
		print(f"  The spreadsheet must contain columns with these exact headers: {config.ID_HEADER}, {config.OBSERVATION_HEADER}, {config.CATEGORY_HEADER}")
		return
	
	# Load sheet data
	sheet_data = sheet.get_all_values()
	utils.debug_print(f'Sheet dumped into memory at {utils.get_current_time()}')
	
	# Check if sheet is empty or has only headers
	if len(sheet_data) <= 1:
		print("! ERROR Spreadsheet appears to be empty (no data rows found).")
		return
	
	# Get participant info
	header_row = sheet.row_values(id_cell.row)
	num_participants = get_num_participants(header_row, id_cell, sheet.col_count)
	
	if num_participants == 0:
		print(f"! WARNING No participant columns found in the spreadsheet.")
		print(f"  Looking for columns starting with: {', '.join(config.PARTICIPANT_PREFIXES)}")
		return
	
	# Calculate bounds for data rows (after header)
	first_data_row = id_cell.row  # 0-indexed, this is the first row after the header
	last_data_row = len(sheet_data) - 1  # 0-indexed
	total_data_rows = last_data_row - first_data_row + 1
	
	# Current position (0-indexed into sheet_data)
	current_row = first_data_row
	
	# Get participant column headers for display
	participant_headers = []
	for col_idx in range(id_cell.col, id_cell.col + num_participants):
		if col_idx < len(header_row):
			participant_headers.append(header_row[col_idx])
	
	print(f'\n=== Browse Mode ===')
	print(f'Total data rows: {total_data_rows} (rows {first_data_row + 1} to {last_data_row + 1})')
	print(f'Participants: {", ".join(participant_headers)}')
	print(f'\nCommands: up/u, down/d, pageup/pu, pagedown/pd, jump/j <row>, open/o, quit/q')
	print(f'Press Enter to move down one row.\n')
	
	def display_rows(start_row, num_rows):
		"""Display num_rows starting from start_row (0-indexed)."""
		print('-' * 60)
		for i in range(num_rows):
			row_idx = start_row + i
			if row_idx > last_data_row:
				break
			
			row_data = sheet_data[row_idx]
			row_num = row_idx + 1  # 1-indexed for display
			
			# Get category
			category_col = category_cell.col - 1  # 0-indexed
			category = row_data[category_col] if category_col < len(row_data) else ''
			
			# Get description (observation)
			desc_col = observation_cell.col - 1  # 0-indexed
			description = row_data[desc_col] if desc_col < len(row_data) else ''
			
			print(f'Row {row_num}')
			print(f'  Category: {category if category else "(empty)"}')
			print(f'  Description: {description if description else "(empty)"}')
			
			# Get participant timestamps
			has_timestamps = False
			participant_data = []
			for j, participant_id in enumerate(participant_headers):
				col_idx = id_cell.col + j  # 0-indexed
				if col_idx < len(row_data):
					timestamp_value = row_data[col_idx]
					if timestamp_value and timestamp_value.strip():
						# Replace newlines with commas for display
						timestamp_display = timestamp_value.replace('\n', ', ').replace('\r', '')
						participant_data.append(f'    {participant_id}: {timestamp_display}')
						has_timestamps = True
			
			if has_timestamps:
				print('  Participants:')
				for p_data in participant_data:
					print(p_data)
			else:
				print('  Participants: (no timestamps)')
			
			print('  ---')
		
		# Show position info
		displayed_end = min(start_row + num_rows, last_data_row + 1)
		print(f'\nShowing rows {start_row + 1}-{displayed_end} of {last_data_row + 1}')
	
	# Initial display
	display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
	
	# Navigation loop
	while True:
		user_input = input('\n>> ').strip().lower()
		
		if user_input in ('quit', 'q'):
			print('Exiting browse mode.')
			break
		elif user_input in ('up', 'u'):
			if current_row > first_data_row:
				current_row -= 1
				display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
			else:
				print('Already at the first row.')
		elif user_input in ('down', 'd', ''):
			if current_row < last_data_row:
				current_row += 1
				display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
			else:
				print('Already at the last row.')
		elif user_input in ('pageup', 'pu'):
			new_row = max(first_data_row, current_row - config.BROWSE_LINES_TO_DISPLAY)
			if new_row != current_row:
				current_row = new_row
				display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
			else:
				print('Already at the first row.')
		elif user_input in ('pagedown', 'pd'):
			new_row = min(last_data_row, current_row + config.BROWSE_LINES_TO_DISPLAY)
			if new_row != current_row:
				current_row = new_row
				display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
			else:
				print('Already at the last row.')
		elif user_input.startswith('jump ') or user_input.startswith('j '):
			try:
				parts = user_input.split()
				if len(parts) >= 2:
					target_row = int(parts[1]) - 1  # Convert to 0-indexed
					if target_row < first_data_row:
						print(f'Row number must be at least {first_data_row + 1}.')
					elif target_row > last_data_row:
						print(f'Row number must be at most {last_data_row + 1}.')
					else:
						current_row = target_row
						display_rows(current_row, config.BROWSE_LINES_TO_DISPLAY)
				else:
					print('Usage: jump <row_number> or j <row_number>')
			except ValueError:
				print('Invalid row number. Usage: jump <row_number> or j <row_number>')
		elif user_input in ('open', 'o'):
			try:
				spreadsheet_url = sheet.spreadsheet.url
				print(f'Opening spreadsheet in browser: {sheet.spreadsheet.title}')
				webbrowser.open(spreadsheet_url)
				print('Spreadsheet opened in your default browser.')
			except AttributeError as e:
				print('! ERROR Could not retrieve spreadsheet URL.')
				print(f'  Error: {e}')
			except Exception as e:
				print('! ERROR Could not open browser.')
				print(f'  Error: {e}')
		else:
			print('Unknown command. Available: up/u, down/d, pageup/pu, pagedown/pd, jump/j <row>, open/o, quit/q')
			print('Press Enter to move down one row.')
