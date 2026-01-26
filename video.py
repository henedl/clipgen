# -*- coding: utf-8 -*-
"""Video processing operations for clipgen."""

import os
import subprocess
from datetime import datetime
from typing import Optional

from icecream import ic

import config
import files
import utils


def run_ffmpeg(input_file: str, output_file: str, start_pos: str, end_pos: str, reencode: bool) -> bool:
    """Calls ffmpeg to cut a video clip. Requires ffmpeg in system PATH.
    
    Args:
        input_file: Path to input video file
        output_file: Path for output video file
        start_pos: Start timestamp (format: HH:MM:SS or MM:SS)
        end_pos: End timestamp (format: HH:MM:SS or MM:SS)
        reencode: If True, re-encode video; if False, use stream copy
        
    Returns:
        True if video was generated successfully, False otherwise.
    """
    ic(input_file, output_file, start_pos, end_pos)
    # Check if input file exists before processing
    if not os.path.isfile(input_file):
        utils.error_print(f"Input video file not found: '{input_file}'",
            [f"Expected location: {os.path.join(os.getcwd(), input_file)}",
             "Skipping this clip."])
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
        utils.error_print("Negative duration calculated for video clip. Skipping.",
            [f"Start: {start_pos}, End: {end_pos}, Duration: {duration}s",
             "The end timestamp must be after the start timestamp."])
        return False
    if duration > file_length:
        utils.error_print(f"Timestamp duration ({duration}s) exceeds video file length ({file_length}s). Skipping.",
            [f"Start: {start_pos}, End: {end_pos}",
             f"Video file: '{input_file}'"])
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
            if config.AUDIO_NORMALIZE:
                # Copy video stream, re-encode audio with normalization
                ffmpeg_command = ['ffmpeg', '-y', '-loglevel', '16', '-ss', start_pos, '-i', input_file, '-t', str(duration), '-c:v', 'copy', '-c:a', 'aac', '-af', 'loudnorm=I=-16:TP=-1.5:LRA=11', '-avoid_negative_ts', '1', output_file]
            else:
                # Copy all streams
                ffmpeg_command = ['ffmpeg', '-y', '-loglevel', '16', '-ss', start_pos, '-i', input_file, '-t', str(duration), '-c', 'copy', '-avoid_negative_ts', '1', output_file]
            utils.debug_print(f"ffmpeg_command is '{' '.join(ffmpeg_command)}'")
            result = subprocess.run(ffmpeg_command, encoding='utf-8', capture_output=True)
        else:
            # Re-encode case
            if config.AUDIO_NORMALIZE:
                # Re-encode with audio normalization
                ffmpeg_command = ['ffmpeg', '-y', '-loglevel', '16', '-ss', start_pos, '-i', input_file, '-t', str(duration), '-af', 'loudnorm=I=-16:TP=-1.5:LRA=11', output_file]
            else:
                # Re-encode without normalization
                ffmpeg_command = ['ffmpeg', '-y', '-loglevel', '16', '-ss', start_pos, '-i', input_file, '-t', str(duration), output_file]
            utils.debug_print(f"ffmpeg_command is '{' '.join(ffmpeg_command)}'")
            result = subprocess.run(ffmpeg_command, encoding='utf-8', capture_output=True)
        
        # Check if ffmpeg succeeded
        if result.returncode != 0:
            error_details = [f"Input: '{input_file}', Output: '{output_file}'",
                f"Timestamps: {start_pos} to {end_pos}"]
            if result.stderr:
                error_details.append(f"ffmpeg error: {result.stderr.strip()}")
            utils.error_print(f"ffmpeg failed with exit code {result.returncode}", error_details)
            return False
        
        # Verify output file was created
        if not os.path.isfile(output_file):
            utils.error_print(f"ffmpeg completed but output file was not created: '{output_file}'")
            return False

        # Check filesize limit and compress if needed
        if config.MAX_FILESIZE_MB and config.MAX_FILESIZE_MB > 0:
            if not compress_to_size(output_file, config.MAX_FILESIZE_MB):
                utils.warning_print(f"Could not compress '{output_file}' to target size")
                # Continue anyway - file was generated, just not compressed

        utils.verbose_print(f"+ Generated video '{output_file}' successfully.\n File size: {files.format_filesize(os.path.getsize(output_file))}\n Expected duration: {duration} s\n")
        return True
    except FileNotFoundError:
        utils.error_print("ffmpeg is not installed or not found in system PATH.",
            ["Please install ffmpeg and ensure it's in your PATH.",
             "Download from: https://www.ffmpeg.org/download.html"])
        return False
    except OSError as e:
        utils.error_print("ffmpeg could not successfully run.",
            [f"Error: {e}",
             f"Working directory: '{os.getcwd()}'",
             f"Input file: '{input_file}'",
             f"Output file: '{output_file}'"])
        return False

def get_file_duration(filepath: str) -> Optional[int]:
    """Calls ffprobe to get duration of video container.
    
    Args:
        filepath: Path to video file
        
    Returns:
        The duration in seconds, or None if the file cannot be probed.
    """
    # Check if file exists before attempting to probe
    if not os.path.isfile(filepath):
        utils.error_print(f"Video file not found: '{filepath}'",
            [f"Expected location: {os.path.join(os.getcwd(), filepath)}",
             "Please ensure the video file exists in the working directory."])
        return None
    
    probe_command = ['ffprobe', '-v', 'error', '-show_entries', 'format=duration', '-of', 'default=noprint_wrappers=1:nokey=1', filepath]
    utils.debug_print(f"probe_command is {' '.join(probe_command)}")
    
    try:
        file_length = float(subprocess.check_output(probe_command, encoding='utf-8'))
        return int(file_length)
    except FileNotFoundError:
        utils.error_print("ffprobe is not installed or not found in system PATH.",
            ["Please install ffmpeg (which includes ffprobe) and ensure it's in your PATH.",
             "Download from: https://www.ffmpeg.org/download.html"])
        return None
    except subprocess.CalledProcessError as e:
        utils.error_print(f"ffprobe failed to read video file: '{filepath}'",
            [f"ffprobe exit code: {e.returncode}",
             "The file may be corrupted, not a valid video, or in an unsupported format."])
        return None
    except ValueError as e:
        utils.error_print(f"Could not parse duration from video file: '{filepath}'",
            [f"ffprobe returned unexpected output. Error: {e}"])
        return None

def get_duration(start_time: str, end_time: str) -> Optional[int]:
    """Calculate the duration between two timestamps.
    
    Args:
        start_time: Start timestamp (format: HH:MM:SS or MM:SS)
        end_time: End timestamp (format: HH:MM:SS or MM:SS)
        
    Returns:
        Duration in seconds, or None if timestamps are invalid.
    """
    ic(start_time, end_time)
    utils.debug_print(f'start_time is {start_time} with length {len(start_time)}, end_time is {end_time}')
    
    # Handle case where add_duration() returned -1 (error)
    if end_time == -1:
        utils.error_print(f"Invalid end timestamp (derived from start: '{start_time}')",
            ["Could not calculate end time. Check the timestamp format."])
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
    
    utils.error_print("Timestamp formatting error in get_duration().",
        [f"Start time: '{start_time}', End time: '{end_time}'",
         "Accepted formats: HH:MM:SS, MM:SS, or M:SS (e.g., 1:23:45, 12:34, 1:23)"])
    return None


def calculate_target_bitrate(target_size_mb: float, duration_seconds: int, audio_bitrate_kbps: int = 128) -> int:
    """Calculate video bitrate needed to achieve target filesize.

    Args:
        target_size_mb: Target file size in megabytes
        duration_seconds: Video duration in seconds
        audio_bitrate_kbps: Audio bitrate in kbps (default: 128)

    Returns:
        Target video bitrate in kbps, or 100 if calculated value is too low
    """
    if duration_seconds <= 0:
        return 100

    target_size_bytes = target_size_mb * 1024 * 1024
    total_bitrate_kbps = (target_size_bytes * 8) / duration_seconds / 1000

    # Subtract audio bitrate to get video-only bitrate
    video_bitrate_kbps = int(total_bitrate_kbps - audio_bitrate_kbps)

    # Ensure minimum viable bitrate (100 kbps minimum)
    return max(video_bitrate_kbps, 100)


def compress_to_size(filepath: str, target_size_mb: float) -> bool:
    """Recompress video to fit within target filesize using two-pass encoding.

    Args:
        filepath: Path to the video file to compress
        target_size_mb: Maximum file size in megabytes

    Returns:
        True if compression succeeded or was unnecessary, False on error
    """
    # Get current file size
    current_size_bytes = os.path.getsize(filepath)
    target_size_bytes = target_size_mb * 1024 * 1024

    # Check if compression is needed
    if current_size_bytes <= target_size_bytes:
        utils.debug_print(f"File already within size limit: {files.format_filesize(current_size_bytes)}")
        return True

    # Get video duration for bitrate calculation
    duration = get_file_duration(filepath)
    if duration is None or duration <= 0:
        utils.error_print(f"Cannot compress: unable to determine duration of '{filepath}'")
        return False

    # Calculate target bitrate (with 5% safety margin)
    target_bitrate = calculate_target_bitrate(target_size_mb * 0.95, duration)

    if target_bitrate <= 100:
        utils.warning_print(f"Target bitrate very low ({target_bitrate} kbps) for {duration}s video.",
            [f"Target size: {target_size_mb}MB, Duration: {duration}s",
             "Quality may be significantly reduced."])

    utils.verbose_print(f"Compressing video to fit within {target_size_mb}MB...")
    utils.verbose_print(f"  Current size: {files.format_filesize(current_size_bytes)}")
    utils.verbose_print(f"  Target bitrate: {target_bitrate} kbps (video) + 128 kbps (audio)")

    # Create temporary output file
    temp_output = filepath + '.temp.mp4'
    passlog_base = filepath + '.passlog'

    try:
        # Two-pass encoding for better quality at target bitrate
        # Pass 1: Analysis pass
        null_output = '/dev/null' if os.name != 'nt' else 'NUL'
        pass1_command = [
            'ffmpeg', '-y', '-loglevel', '16',
            '-i', filepath,
            '-c:v', 'libx264', '-b:v', f'{target_bitrate}k',
            '-pass', '1', '-passlogfile', passlog_base,
            '-an',  # No audio in first pass
            '-f', 'null', null_output
        ]

        utils.debug_print(f"Pass 1 command: {' '.join(pass1_command)}")
        result1 = subprocess.run(pass1_command, encoding='utf-8', capture_output=True)

        if result1.returncode != 0:
            utils.error_print("Compression pass 1 failed",
                [result1.stderr.strip() if result1.stderr else "Unknown error"])
            return False

        # Pass 2: Actual encoding
        pass2_command = [
            'ffmpeg', '-y', '-loglevel', '16',
            '-i', filepath,
            '-c:v', 'libx264', '-b:v', f'{target_bitrate}k',
            '-pass', '2', '-passlogfile', passlog_base,
            '-c:a', 'aac', '-b:a', '128k',
            temp_output
        ]

        utils.debug_print(f"Pass 2 command: {' '.join(pass2_command)}")
        result2 = subprocess.run(pass2_command, encoding='utf-8', capture_output=True)

        if result2.returncode != 0:
            utils.error_print("Compression pass 2 failed",
                [result2.stderr.strip() if result2.stderr else "Unknown error"])
            return False

        # Verify output was created
        if not os.path.isfile(temp_output):
            utils.error_print("Compression failed: output file not created")
            return False

        new_size = os.path.getsize(temp_output)

        # Replace original with compressed version
        os.replace(temp_output, filepath)

        utils.verbose_print(f"  Compressed: {files.format_filesize(current_size_bytes)} -> {files.format_filesize(new_size)}")

        # Warn if still over target (can happen with very low bitrate requirements)
        if new_size > target_size_bytes:
            utils.warning_print(f"Compressed file still exceeds target ({files.format_filesize(new_size)} > {target_size_mb}MB)",
                ["The video may need a higher size limit or shorter duration."])

        return True

    except FileNotFoundError:
        utils.error_print("ffmpeg not found during compression")
        return False
    except OSError as e:
        utils.error_print(f"Compression failed: {e}")
        return False
    finally:
        # Cleanup pass log files
        for ext in ['-0.log', '-0.log.mbtree', '']:
            log_file = passlog_base + ext
            if os.path.exists(log_file):
                try:
                    os.remove(log_file)
                except OSError:
                    pass
        # Cleanup temp file if it exists
        if os.path.exists(temp_output):
            try:
                os.remove(temp_output)
            except OSError:
                pass
