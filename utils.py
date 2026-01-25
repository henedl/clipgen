# -*- coding: utf-8 -*-
"""Utility functions for clipgen."""

import argparse
from datetime import datetime, timedelta

from icecream import ic

import config


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
	if config.DEBUGGING:
		print(f'! DEBUG {message}')

def verbose_print(message):
	"""Print informational messages when VERBOSE is enabled.
	
	In interactive mode, VERBOSE is always True.
	In CLI mode, VERBOSE is False unless -v flag is used.
	"""
	if config.VERBOSE:
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

def add_duration(start_time):
	"""Adds one minute to the given timestamp.
	
	Returns the new timestamp string, or -1 if the timestamp format is invalid.
	"""
	try:
		if len(start_time) <= 5:
			start_datetime = datetime.strptime(str(start_time), '%M:%S')
			new_time = start_datetime + timedelta(seconds=config.DEFAULT_DURATION_SECONDS)
			return new_time.strftime('%M:%S')
		else:
			start_datetime = datetime.strptime(start_time, '%H:%M:%S')
			new_time = start_datetime + timedelta(seconds=config.DEFAULT_DURATION_SECONDS)
			return new_time.strftime('%H:%M:%S')
	except ValueError:
		print(f"! WARNING Could not parse single timestamp '{start_time}' to add default duration.")
		print(f"  Expected format: MM:SS or HH:MM:SS (e.g., 12:34 or 1:23:45)")
		print("  This timestamp will be skipped.")
		return -1

def parse_timestamps(cell_value, cell_ref=None):
	"""Parse timestamp pairs from a cell value string.
	
	Args:
		cell_value: The raw cell value containing timestamps
		cell_ref: Optional cell reference (e.g., 'B5') for error messages
	
	Returns a list of (start_time, end_time) tuples.
	"""
	ic(cell_value, cell_ref)
	parsed_timestamps = []
	skipped_timestamps = []
	raw_times = cell_value.lower().replace('+', ' ').replace(';', ' ').replace(',', ' ').split()
	ic(raw_times)
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
				ic(raw_times[i], time_pair)
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
					ic(time_pair)
					parsed_timestamps.append(time_pair)
				# If -1, warning already printed by add_duration()
			else:
				skipped_timestamps.append(raw_times[i])
		elif raw_times[i]:
			# Non-empty but doesn't match expected patterns
			skipped_timestamps.append(raw_times[i])
	
	# Report skipped timestamps if any
	if skipped_timestamps:
		ic(skipped_timestamps)
		cell_info = f" in cell {cell_ref}" if cell_ref else ""
		print(f"! WARNING Skipped {len(skipped_timestamps)} unparseable timestamp(s){cell_info}:")
		for ts in skipped_timestamps[:3]:  # Show first 3
			print(f"    '{ts}'")
		if len(skipped_timestamps) > 3:
			print(f"    ... and {len(skipped_timestamps) - 3} more")
		print("  Expected formats: MM:SS-MM:SS, HH:MM:SS-HH:MM:SS, or single timestamps like MM:SS")

	ic(parsed_timestamps)
	return parsed_timestamps

def set_program_settings():
	SETTINGSLIST = ['REENCODING', 'FILEFORMAT', 'DEBUGGING']

	print('\nWhich setting? Available:\n')
	print(', '.join(SETTINGSLIST))
	setting_to_change = input('\n>> ')

	print(f"* Current value for '{setting_to_change}' is '{getattr(config, setting_to_change)}'")

	new_value = input('\nWhich new value?\n>> ')

	print(f"* '{setting_to_change}' SET TO '{new_value}'")

	if setting_to_change != '':
		setattr(config, setting_to_change, new_value)
		return True
	return False

def get_current_time():
	return datetime.now().strftime('%Y-%m-%d %H:%M:%S')
