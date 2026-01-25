# -*- coding: utf-8 -*-
"""Video processing operations for clipgen."""

import os
import subprocess
from datetime import datetime

from icecream import ic

import config
import files
import utils


def run_ffmpeg(input_file, output_file, start_pos, end_pos, reencode):
	"""Calls ffmpeg to cut a video clip. Requires ffmpeg in system PATH.
	
	Returns True if video was generated successfully, False otherwise.
	"""
	ic(input_file, output_file, start_pos, end_pos)
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
	ic(duration, file_length)
	if duration > config.MAX_CLIP_DURATION_SECONDS:
		yn = input(f'The generated video will be {duration}s ({duration//60}m {duration%60}s), over 10 minutes long. Generate anyway? (y/n)\n>> ')
		if yn != 'y':
			return False

	utils.verbose_print(f'Cutting {input_file} from {start_pos} to {end_pos}.')
	if config.DEBUGGING:
		utils.debug_print(f'Debugging enabled, not calling ffmpeg.\n  input_file: {input_file},\n  output_file: {output_file}')
		return False

	try:
		if not reencode:
			# Use list form to properly handle unicode in filenames
			ffmpeg_command = ['ffmpeg', '-y', '-loglevel', '16', '-ss', start_pos, '-i', input_file, '-t', str(duration), '-c', 'copy', '-avoid_negative_ts', '1', output_file]
			utils.debug_print(f"ffmpeg_command is '{' '.join(ffmpeg_command)}'")
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
		
		utils.verbose_print(f"+ Generated video '{output_file}' successfully.\n File size: {files.format_filesize(os.path.getsize(output_file))}\n Expected duration: {duration} s\n")
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
	utils.debug_print(f"probe_command is {' '.join(probe_command)}")
	
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
	ic(start_time, end_time)
	utils.debug_print(f'start_time is {start_time} with length {len(start_time)}, end_time is {end_time}')
	
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
			duration = int((end_datetime - start_datetime).total_seconds())
			ic(duration)
			return duration
		except ValueError:
			continue
	
	print(f"! ERROR Timestamp formatting error in get_duration().")
	print(f"  Start time: '{start_time}', End time: '{end_time}'")
	print("  Accepted formats: HH:MM:SS, MM:SS, or M:SS (e.g., 1:23:45, 12:34, 1:23)")
	return None
