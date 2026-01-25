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
import io
import os
import sys

import gspread
from icecream import ic

import config
import files
import google_api
import spreadsheet
import utils
import video


def select_spreadsheet(gc, doc_list):
	"""Interactive spreadsheet selection. Returns the selected worksheet."""
	input_file_fails = 0
	
	while True:
		input_name = input("\nPlease enter the index, name, URL or key of the spreadsheet ('all' for list, 'new' for list of newest, 'last' to immediately open latest, 'settings' to change settings):\n>> ")
		try:
			if input_name[:4] == 'http':
				return google_api.get_worksheet(gc.open_by_url(input_name))
			elif input_name[:3] == 'all':
				print('\nAvailable documents:')
				for i, doc in enumerate(doc_list):
					print(f'{i+1}. {doc.strip()}')
			elif input_name[:3] == 'new':
				print('\nNewest documents: (modified or opened most recently)')
				for i in range(min(3, len(doc_list))):
					print(f'{i+1}. {doc_list[i].strip()}')
			elif input_name[:4] == 'last':
				latest = google_api.get_all_spreadsheets(gc).split(',')[0]
				return google_api.get_worksheet(gc.open(latest))
			elif input_name[0].isdigit():
				chosen_index = int(input_name) - 1
				print(f'Opening document: {doc_list[chosen_index].strip()}')
				return google_api.get_worksheet(gc.open(doc_list[chosen_index].strip()))
			elif input_name[:8] == 'settings':
				utils.set_program_settings()
			else:
				chosen_index = google_api.find_spreadsheet_by_name(input_name, doc_list)
				if chosen_index >= 0:
					return google_api.get_worksheet(gc.open(google_api.get_all_spreadsheets(gc).split(',')[chosen_index].strip().lstrip()))
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
				print("  - The spreadsheet doesn't contain any worksheets")
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
		'br': 'browse', 'browse': 'browse',
		'test': 'test'
	}
	
	while True:
		input_mode = input('\nSelect mode: (b)atch, (r)ange, (c)ategory, (l)ine, or (br)owse\n>> ').strip().lower()
		
		if not input_mode:
			print("  Please enter a mode (b, r, c, l, or br).")
			continue
		
		try:
			# Check for two-character modes first (like 'br' for browse)
			mode = mode_map.get(input_mode[:2]) or mode_map.get(input_mode[0]) or mode_map.get(input_mode)
			if mode == 'browse':
				# Browse mode doesn't generate clips, just displays data
				spreadsheet.browse_spreadsheet(worksheet)
				return []
			elif mode:
				return spreadsheet.generate_list(worksheet, mode)
			else:
				print(f"  Unknown mode '{input_mode}'. Available modes:")
				print("    b or batch   - Generate all clips in the spreadsheet")
				print("    r or range   - Generate clips from a range of rows")
				print("    c or category - Generate clips by category")
				print("    l or line    - Generate clips from specific line(s)")
				print("    br or browse - Browse spreadsheet rows interactively")
		except gspread.exceptions.GSpreadException as e:
			print(f"! ERROR Google Sheets API error: {e}")
			utils.debug_print(f"ERROR Message '{e}', Attempting reconnect")

def process_clips(clips_list):
	"""Process and generate video clips from the clips list. Returns count of videos generated."""
	ic(len(clips_list))
	# Check if clips_list is empty
	if not clips_list:
		print("! WARNING No clips to process. No timestamps were found or selected.")
		return 0
	
	utils.verbose_print('\n* ffmpeg is set to never prompt for input and will always overwrite.\n  Only warns if close to crashing.\n')
	videos_generated = 0
	videos_skipped = 0
	missing_videos = set()  # Track unique missing video files

	for clip in clips_list:
		ic(clip)
		clip = files.clean_issue(clip)
		
		# Skip if no valid timestamps were parsed
		if not clip['times']:
			videos_skipped += 1
			continue
		
		base_video = f"{clip['study']}_{clip['participant']}{config.FILEFORMAT}"
		
		# Check if base video exists (only warn once per unique file)
		if not os.path.isfile(base_video) and base_video not in missing_videos:
			missing_videos.add(base_video)
			print(f"! ERROR Source video file not found: '{base_video}'")
			print(f"  Expected location: {os.path.join(os.getcwd(), base_video)}")
			print(f"  Clips for participant '{clip['participant']}' in study '{clip['study']}' will be skipped.")
		
		for vid_in, vid_out in clip['times']:
			try:
				# Ensure all components are strings and handle unicode properly
				vid_name = files.get_unique_filename(
					f"[{clip['category']}] {clip['study']} {clip['participant']} {clip['desc']}{config.FILEFORMAT}"
				)
				ic(vid_name)
			except (TypeError, UnicodeEncodeError, UnicodeDecodeError) as e:
				ic(e, clip)
				print(f'! ERROR Character encoding issue occurred:\n  {e}')
				print(f"  Category: '{clip['category']}', Study: '{clip['study']}', Participant: '{clip['participant']}'")
				print("  Try simplifying the description or category names to use only ASCII characters.")
				videos_skipped += 1
				break

			completed = video.run_ffmpeg(
				input_file=base_video,
				output_file=vid_name,
				start_pos=vid_in,
				end_pos=vid_out,
				reencode=config.REENCODING
			)
			if completed:
				videos_generated += 1
			else:
				videos_skipped += 1
	
	# Report summary if any videos were skipped
	if videos_skipped > 0:
		utils.verbose_print(f"\n* Summary: {videos_generated} video(s) generated, {videos_skipped} skipped due to errors.")
	if missing_videos:
		utils.verbose_print(f"* Missing source video files: {len(missing_videos)}")

	return videos_generated

def main():
	# Ensure UTF-8 encoding for stdout/stderr to handle unicode properly
	if sys.stdout.encoding.lower() != 'utf-8':
		sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
		sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
	
	# Parse command-line arguments
	args = utils.parse_arguments()
	ic(args)
	
	# Determine if running in CLI mode (any mode argument provided)
	cli_mode = args.batch or args.lines or args.range
	
	# Set verbose mode: silent by default in CLI mode, verbose in interactive mode
	config.VERBOSE = not cli_mode or args.verbose
	
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
	utils.verbose_print('-------------------------------------------------------------------------------')
	utils.verbose_print(f'Welcome to clipgen v{config.VERSIONNUM}\nWorking directory: {os.getcwd()}\nPlace video files and the credentials.json file in this directory.')
	utils.debug_print('Debug mode is ON. Several limitations apply and more things will be printed.')
	
	# Authenticate with Google
	try:
		utils.debug_print('Attempting login...')
		gc = gspread.oauth(credentials_filename='credentials.json')
		utils.debug_print('Login successful!')
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
	doc_list = google_api.get_all_spreadsheets(gc).split(',')

	# Spreadsheet selection
	worksheet = None
	if args.spreadsheet:
		# CLI-specified spreadsheet
		try:
			if args.spreadsheet.startswith('http'):
				worksheet = google_api.get_worksheet(gc.open_by_url(args.spreadsheet))
			elif args.spreadsheet.isdigit():
				chosen_index = int(args.spreadsheet) - 1
				utils.verbose_print(f'Opening document: {doc_list[chosen_index].strip()}')
				worksheet = google_api.get_worksheet(gc.open(doc_list[chosen_index].strip()))
			else:
				chosen_index = google_api.find_spreadsheet_by_name(args.spreadsheet, doc_list)
				if chosen_index >= 0:
					matched_name = doc_list[chosen_index].strip()
					utils.verbose_print(f'Opening document: {matched_name}')
					worksheet = google_api.get_worksheet(gc.open(matched_name))
				else:
					print(f'Error: Could not find spreadsheet "{args.spreadsheet}"')
					sys.exit(1)
		except (gspread.SpreadsheetNotFound, gspread.exceptions.APIError, gspread.exceptions.GSpreadException) as e:
			print(f'Error: Could not open spreadsheet "{args.spreadsheet}": {e}')
			sys.exit(1)
	else:
		# Auto-connect if working directory name matches a spreadsheet
		cwd_name = os.path.basename(os.getcwd())
		auto_match_index = google_api.find_spreadsheet_by_name(cwd_name, doc_list)
		if auto_match_index >= 0:
			matched_name = doc_list[auto_match_index].strip()
			utils.verbose_print(f'\nAuto-connecting to spreadsheet: {matched_name}')
			worksheet = google_api.get_worksheet(gc.open(matched_name))
		elif cli_mode:
			# CLI mode requires a spreadsheet - can't prompt interactively
			print('Error: No spreadsheet found matching working directory name.')
			print('Use -s to specify a spreadsheet name, URL, or index.')
			sys.exit(1)
		else:
			worksheet = select_spreadsheet(gc, doc_list)
	
	if worksheet:
		ic(worksheet.title)
	utils.verbose_print('\nConnected to Google Drive!')

	if cli_mode:
		# CLI mode - run once and exit
		skip_prompts = args.yes
		
		if args.batch:
			clips_list = spreadsheet.generate_list(worksheet, 'batch', skip_prompts=skip_prompts)
		elif args.lines:
			clips_list = spreadsheet.generate_list(worksheet, 'line', line_numbers=cli_line_numbers, skip_prompts=skip_prompts)
		elif args.range:
			clips_list = spreadsheet.generate_list(worksheet, 'range', range_start=cli_range_start, range_end=cli_range_end, skip_prompts=skip_prompts)
		
		videos_generated = process_clips(clips_list)
		
		if not config.REENCODING:
			utils.verbose_print('* No re-encoding done, expect:\n- inaccurate start and end timings\n- lossy frames until first keyframe\n- bad timecodes at the end\n')
		print(f'All done, created {videos_generated} videos!\nFiles are in {os.getcwd()}\n')
	else:
		# Interactive mode - main processing loop
		while True:
			clips_list = select_mode_and_generate(worksheet)
			videos_generated = process_clips(clips_list)

			if not config.REENCODING:
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
