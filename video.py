# -*- coding: utf-8 -*-
"""Video processing operations for clipgen."""

import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from icecream import ic

import config
import utils

INVALID_END_TIMESTAMP = None


def _ffmpeg_install_guidance_lines() -> List[str]:
    """Return actionable install guidance based on the current platform."""
    platform_specific = []
    if sys.platform == 'darwin':
        platform_specific = [
            "macOS: install with Homebrew: brew install ffmpeg",
        ]
    elif sys.platform.startswith('linux'):
        platform_specific = [
            "Linux (Debian/Ubuntu): sudo apt update && sudo apt install ffmpeg",
            "Linux (Fedora): sudo dnf install ffmpeg",
        ]
    elif sys.platform.startswith('win'):
        platform_specific = [
            "Windows (winget): winget install Gyan.FFmpeg",
            "Windows (chocolatey): choco install ffmpeg",
        ]
    else:
        platform_specific = [
            "Install from: https://www.ffmpeg.org/download.html",
        ]

    return platform_specific + [
        "Then verify in a new terminal:",
        "  ffmpeg -version",
        "  ffprobe -version",
    ]


def check_ffmpeg_tools_available() -> bool:
    """Verify ffmpeg and ffprobe are available in PATH at startup."""
    missing_tools = [tool for tool in ('ffmpeg', 'ffprobe') if shutil.which(tool) is None]
    if not missing_tools:
        return True

    details = [
        f"Missing command(s): {', '.join(missing_tools)}",
        "clipgen requires both ffmpeg and ffprobe to cut and inspect videos.",
    ]
    details.extend(_ffmpeg_install_guidance_lines())
    utils.error_print("Required video tools are missing from PATH.", details)
    return False


def _handle_ffmpeg_not_found() -> None:
    """Print a user-facing error when the ffmpeg binary cannot be located."""
    utils.error_print(
        "ffmpeg is not installed or not found in system PATH.",
        [
            "Please install ffmpeg and ensure it's in your PATH.",
            "Download from: https://www.ffmpeg.org/download.html",
        ],
    )


def _run_ffmpeg_process(
    ffmpeg_command: List[str],
    *,
    input_file: str,
    output_file: str,
    os_error_message: str,
) -> Optional[subprocess.CompletedProcess[str]]:
    """Run an ffmpeg subprocess and wrap common OS-level failures."""
    try:
        return subprocess.run(ffmpeg_command, encoding='utf-8', capture_output=True)
    except FileNotFoundError:
        _handle_ffmpeg_not_found()
        return None
    except OSError as error:
        utils.error_print(
            os_error_message,
            [
                f"Error: {error}",
                f"Working directory: '{os.getcwd()}'",
                f"Input file: '{input_file}'",
                f"Output file: '{output_file}'",
            ],
        )
        return None


def _add_ffmpeg_stderr(error_details: List[str], ffmpeg_result: subprocess.CompletedProcess[str]) -> List[str]:
    """Append trimmed ffmpeg stderr output to an error details list when available."""
    if ffmpeg_result.stderr:
        error_details.append(f"ffmpeg error: {ffmpeg_result.stderr.strip()}")
    return error_details


def _verify_output_file(output_file: str, operation_label: str) -> bool:
    """Return True when an expected ffmpeg output file exists, otherwise log an error."""
    if Path(output_file).is_file():
        return True
    utils.error_print(f"{operation_label} completed but output file was not created: '{output_file}'")
    return False


def build_ffmpeg_cut_command(
    input_file: str,
    output_file: str,
    start_pos: str,
    duration_seconds: int,
    reencode: bool,
    audio_normalize: bool,
) -> List[str]:
    """Build ffmpeg argv for cutting a clip. Caller runs subprocess.
    
    Args:
        input_file: Input video path
        output_file: Output video path
        start_pos: Start timestamp
        duration_seconds: Clip duration in seconds
        reencode: If True, re-encode; if False, stream copy
        audio_normalize: If True, apply loudnorm
    Returns:
        argv list for subprocess (e.g. ['ffmpeg', '-y', ...])
    """
    base = ['ffmpeg', '-y', '-loglevel', config.FFMPEG_LOGLEVEL, '-ss', start_pos, '-i', input_file, '-t', str(duration_seconds)]
    if not reencode:
        if audio_normalize:
            # loudnorm: I=-16 (target LUFS), TP=-1.5 (true peak dB), LRA=11 (loudness range)
            # -avoid_negative_ts 1: shift timestamps so output starts at 0 (avoids glitches after cut)
            return base + ['-c:v', 'copy', '-c:a', 'aac', '-af', 'loudnorm=I=-16:TP=-1.5:LRA=11', '-avoid_negative_ts', '1', output_file]
        # Stream copy; -avoid_negative_ts 1 fixes timestamp issues when cutting
        return base + ['-c', 'copy', '-avoid_negative_ts', '1', output_file]
    if audio_normalize:
        return base + ['-af', 'loudnorm=I=-16:TP=-1.5:LRA=11', output_file]
    return base + [output_file]


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
    if config.DEBUGGING:
        ic(input_file, output_file, start_pos, end_pos)
    if not Path(input_file).is_file():
        utils.error_print(f"Input video file not found: '{input_file}'",
            [f"Expected location: {Path(input_file).resolve()}",
             "Skipping this clip."])
        return False

    duration = get_duration(start_pos, end_pos)
    if duration is None:
        # Error already printed by get_duration
        return False

    duration_seconds = get_file_duration(input_file)
    if duration_seconds is None:
        # Error already printed by get_file_duration
        return False

    if duration < 0:
        utils.error_print("Negative duration calculated for video clip. Skipping.",
            [f"Start: {start_pos}, End: {end_pos}, Duration: {duration}s",
             "The end timestamp must be after the start timestamp."])
        return False
    if duration > duration_seconds:
        utils.error_print(f"Timestamp duration ({duration}s) exceeds video file length ({duration_seconds}s). Skipping.",
            [f"Start: {start_pos}, End: {end_pos}",
             f"Video file: '{input_file}'"])
        return False
    if config.DEBUGGING:
        ic(duration, duration_seconds)
    if duration > config.MAX_CLIP_DURATION_SECONDS:
        yn = utils.read_user_input(
            f'The generated video will be {duration}s ({duration//60}m {duration%60}s), over 10 minutes long. Generate anyway? (y/n)\n>> '
        )
        if yn != 'y':
            return False

    utils.verbose_print(f'Cutting {input_file} from {start_pos} to {end_pos}.')
    if config.DEBUGGING:
        utils.debug_print(f'Debugging enabled, not calling ffmpeg.\n  input_file: {input_file},\n  output_file: {output_file}')
        return False

    ffmpeg_command = build_ffmpeg_cut_command(
        input_file, output_file, start_pos, duration, reencode, config.AUDIO_NORMALIZE
    )
    utils.debug_print(f"ffmpeg_command is '{' '.join(ffmpeg_command)}'")
    ffmpeg_result = _run_ffmpeg_process(
        ffmpeg_command,
        input_file=input_file,
        output_file=output_file,
        os_error_message="ffmpeg could not successfully run.",
    )
    if ffmpeg_result is None:
        return False

    if ffmpeg_result.returncode != 0:
        error_details = [f"Input: '{input_file}', Output: '{output_file}'", f"Timestamps: {start_pos} to {end_pos}"]
        utils.error_print(
            f"ffmpeg failed with exit code {ffmpeg_result.returncode}",
            _add_ffmpeg_stderr(error_details, ffmpeg_result),
        )
        return False

    if not _verify_output_file(output_file, "ffmpeg"):
        return False

    if config.MAX_FILESIZE_MB and config.MAX_FILESIZE_MB > 0:
        if not compress_to_size(output_file, config.MAX_FILESIZE_MB):
            utils.warning_print(f"Could not compress '{output_file}' to target size")

    utils.verbose_print(f"+ Generated video '{output_file}' successfully.\n File size: {utils.format_filesize(Path(output_file).stat().st_size)}\n Expected duration: {duration} s\n")
    return True


def extract_screenshot(input_file: str, output_file: str, timestamp: str) -> bool:
    """Extract a single screenshot frame at the given timestamp.

    Args:
        input_file: Path to input video file
        output_file: Path for output screenshot file (.png)
        timestamp: Timestamp to capture (format: HH:MM:SS or MM:SS)

    Returns:
        True if screenshot was generated successfully, False otherwise.
    """
    if config.DEBUGGING:
        ic(input_file, output_file, timestamp)
    if not Path(input_file).is_file():
        utils.error_print(f"Input video file not found: '{input_file}'",
            [f"Expected location: {Path(input_file).resolve()}",
             "Skipping this screenshot."])
        return False

    utils.verbose_print(f"Extracting screenshot from {input_file} at {timestamp}.")
    if config.DEBUGGING:
        utils.debug_print(f'Debugging enabled, not calling ffmpeg.\n  input_file: {input_file},\n  output_file: {output_file}')
        return False

    ffmpeg_command = [
        'ffmpeg', '-y', '-loglevel', config.FFMPEG_LOGLEVEL,
        '-ss', timestamp, '-i', input_file,
        '-vframes', '1', '-q:v', config.FFMPEG_SCREENSHOT_QUALITY,
        output_file,
    ]
    utils.debug_print(f"ffmpeg screenshot command: {' '.join(ffmpeg_command)}")

    ffmpeg_result = _run_ffmpeg_process(
        ffmpeg_command,
        input_file=input_file,
        output_file=output_file,
        os_error_message="ffmpeg could not successfully run for screenshot extraction.",
    )
    if ffmpeg_result is None:
        return False

    if ffmpeg_result.returncode != 0:
        error_details = [f"Input: '{input_file}', Output: '{output_file}'", f"Timestamp: {timestamp}"]
        utils.error_print(
            f"ffmpeg screenshot failed with exit code {ffmpeg_result.returncode}",
            _add_ffmpeg_stderr(error_details, ffmpeg_result),
        )
        return False
    if not _verify_output_file(output_file, "ffmpeg screenshot"):
        return False
    utils.verbose_print(f"+ Generated screenshot '{output_file}' successfully.\n File size: {utils.format_filesize(Path(output_file).stat().st_size)}\n")
    return True


def extract_gif(input_file: str, output_file: str, timestamp: str, duration_seconds: int) -> bool:
    """Extract a GIF segment starting at timestamp.

    Args:
        input_file: Path to input video file
        output_file: Path for output GIF file (.gif)
        timestamp: Start timestamp (format: HH:MM:SS or MM:SS)
        duration_seconds: GIF duration in seconds

    Returns:
        True if GIF was generated successfully, False otherwise.
    """
    if config.DEBUGGING:
        ic(input_file, output_file, timestamp, duration_seconds)
    if not Path(input_file).is_file():
        utils.error_print(f"Input video file not found: '{input_file}'",
            [f"Expected location: {Path(input_file).resolve()}",
             "Skipping this GIF."])
        return False
    if duration_seconds <= 0:
        utils.error_print(f"Invalid GIF duration: {duration_seconds}",
            ["Duration must be greater than 0 seconds."])
        return False

    utils.verbose_print(f"Extracting GIF from {input_file} at {timestamp} ({duration_seconds}s).")
    if config.DEBUGGING:
        utils.debug_print(f'Debugging enabled, not calling ffmpeg.\n  input_file: {input_file},\n  output_file: {output_file}')
        return False

    ffmpeg_command = [
        'ffmpeg', '-y', '-loglevel', config.FFMPEG_LOGLEVEL,
        '-ss', timestamp, '-t', str(duration_seconds), '-i', input_file,
        '-vf', f'fps={config.GIF_FPS},scale={config.GIF_SCALE_WIDTH}:-1:flags=lanczos',
        '-loop', '0',
        output_file,
    ]
    utils.debug_print(f"ffmpeg gif command: {' '.join(ffmpeg_command)}")

    ffmpeg_result = _run_ffmpeg_process(
        ffmpeg_command,
        input_file=input_file,
        output_file=output_file,
        os_error_message="ffmpeg could not successfully run for GIF extraction.",
    )
    if ffmpeg_result is None:
        return False

    if ffmpeg_result.returncode != 0:
        error_details = [
            f"Input: '{input_file}', Output: '{output_file}'",
            f"Timestamp: {timestamp}, Duration: {duration_seconds}s",
        ]
        utils.error_print(
            f"ffmpeg GIF extraction failed with exit code {ffmpeg_result.returncode}",
            _add_ffmpeg_stderr(error_details, ffmpeg_result),
        )
        return False
    if not _verify_output_file(output_file, "ffmpeg GIF extraction"):
        return False
    utils.verbose_print(f"+ Generated GIF '{output_file}' successfully.\n File size: {utils.format_filesize(Path(output_file).stat().st_size)}\n")
    return True


def get_file_duration(filepath: str) -> Optional[int]:
    """Calls ffprobe to get duration of video container.
    
    Args:
        filepath: Path to video file
        
    Returns:
        The duration in seconds, or None if the file cannot be probed.
    """
    if not Path(filepath).is_file():
        utils.error_print(f"Video file not found: '{filepath}'",
            [f"Expected location: {Path(filepath).resolve()}",
             "Please ensure the video file exists in the configured input directory or working directory."])
        return None

    probe_command = ['ffprobe', '-v', 'error', '-show_entries', 'format=duration', '-of', 'default=noprint_wrappers=1:nokey=1', filepath]
    utils.debug_print(f"probe_command is {' '.join(probe_command)}")
    
    try:
        duration_seconds = float(subprocess.check_output(probe_command, encoding='utf-8'))
        return int(duration_seconds)
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

def get_duration(start_time: str, end_time: Optional[str]) -> Optional[int]:
    """Calculate the duration between two timestamps.
    
    Args:
        start_time: Start timestamp (format: HH:MM:SS or MM:SS)
        end_time: End timestamp (format: HH:MM:SS or MM:SS)
        
    Returns:
        Duration in seconds, or None if timestamps are invalid.
    """
    if config.DEBUGGING:
        ic(start_time, end_time)
    utils.debug_print(f'start_time is {start_time} with length {len(start_time)}, end_time is {end_time}')

    if end_time is INVALID_END_TIMESTAMP:
        utils.error_print(f"Invalid end timestamp (derived from start: '{start_time}')",
            ["Could not calculate end time. Check the timestamp format."])
        return None

    formats = ['%M:%S', '%H:%M:%S'] if len(str(start_time)) <= config.MAX_MMSS_LENGTH else ['%H:%M:%S', '%M:%S']

    for time_format in formats:
        try:
            start_datetime = datetime.strptime(str(start_time), time_format)
            end_datetime = datetime.strptime(str(end_time), time_format)
            duration = int((end_datetime - start_datetime).total_seconds())
            if config.DEBUGGING:
                ic(duration)
            return duration
        except ValueError:
            continue
    
    utils.error_print("Timestamp formatting error in get_duration().",
        [f"Start time: '{start_time}', End time: '{end_time}'",
         "Accepted formats: HH:MM:SS, MM:SS, or M:SS (e.g., 1:23:45, 12:34, 1:23)"])
    return None


def calculate_target_bitrate(target_size_mb: float, duration_seconds: int, audio_bitrate_kbps: int = config.AUDIO_BITRATE_KBPS) -> int:
    """Calculate video bitrate needed to achieve target filesize.

    Args:
        target_size_mb: Target file size in megabytes
        duration_seconds: Video duration in seconds
        audio_bitrate_kbps: Audio bitrate in kbps

    Returns:
        Target video bitrate in kbps, minimum MIN_VIDEO_BITRATE_KBPS
    """
    if duration_seconds <= 0:
        return config.MIN_VIDEO_BITRATE_KBPS

    target_size_bytes = target_size_mb * 1024 * 1024
    total_bitrate_kbps = (target_size_bytes * 8) / duration_seconds / 1000

    video_bitrate_kbps = int(total_bitrate_kbps - audio_bitrate_kbps)
    return max(video_bitrate_kbps, config.MIN_VIDEO_BITRATE_KBPS)


def compress_to_size(filepath: str, target_size_mb: float) -> bool:
    """Recompress video to fit within target filesize using two-pass encoding.

    Args:
        filepath: Path to the video file to compress
        target_size_mb: Maximum file size in megabytes

    Returns:
        True if compression succeeded or was unnecessary, False on error
    """
    current_size_bytes = Path(filepath).stat().st_size
    target_size_bytes = target_size_mb * 1024 * 1024
    if current_size_bytes <= target_size_bytes:
        utils.debug_print(f"File already within size limit: {utils.format_filesize(current_size_bytes)}")
        return True

    duration = get_file_duration(filepath)
    if duration is None or duration <= 0:
        utils.error_print(f"Cannot compress: unable to determine duration of '{filepath}'")
        return False

    target_bitrate = calculate_target_bitrate(target_size_mb * config.COMPRESSION_SIZE_FACTOR, duration)
    if target_bitrate <= config.MIN_VIDEO_BITRATE_KBPS:
        utils.warning_print(f"Target bitrate very low ({target_bitrate} kbps) for {duration}s video.",
            [f"Target size: {target_size_mb}MB, Duration: {duration}s",
             "Quality may be significantly reduced."])

    utils.verbose_print(f"Compressing video to fit within {target_size_mb}MB...")
    utils.verbose_print(f"  Current size: {utils.format_filesize(current_size_bytes)}")
    utils.verbose_print(f"  Target bitrate: {target_bitrate} kbps (video) + {config.AUDIO_BITRATE_KBPS} kbps (audio)")

    compressed_temp_path = filepath + '.temp.mp4'
    passlog_base = filepath + '.passlog'

    try:
        null_output = '/dev/null' if os.name != 'nt' else 'NUL'
        pass1_command = [
            'ffmpeg', '-y', '-loglevel', config.FFMPEG_LOGLEVEL,
            '-i', filepath,
            '-c:v', 'libx264', '-b:v', f'{target_bitrate}k',
            '-pass', '1', '-passlogfile', passlog_base,
            '-an',
            '-f', 'null', null_output
        ]

        utils.debug_print(f"Pass 1 command: {' '.join(pass1_command)}")
        pass1_result = _run_ffmpeg_process(
            pass1_command,
            input_file=filepath,
            output_file=null_output,
            os_error_message="ffmpeg could not successfully run during compression pass 1.",
        )
        if pass1_result is None:
            return False

        if pass1_result.returncode != 0:
            utils.error_print(
                "Compression pass 1 failed",
                [pass1_result.stderr.strip() if pass1_result.stderr else "Unknown error"],
            )
            return False

        pass2_command = [
            'ffmpeg', '-y', '-loglevel', config.FFMPEG_LOGLEVEL,
            '-i', filepath,
            '-c:v', 'libx264', '-b:v', f'{target_bitrate}k',
            '-pass', '2', '-passlogfile', passlog_base,
            '-c:a', 'aac', '-b:a', f'{config.AUDIO_BITRATE_KBPS}k',
            compressed_temp_path
        ]

        utils.debug_print(f"Pass 2 command: {' '.join(pass2_command)}")
        pass2_result = _run_ffmpeg_process(
            pass2_command,
            input_file=filepath,
            output_file=compressed_temp_path,
            os_error_message="ffmpeg could not successfully run during compression pass 2.",
        )
        if pass2_result is None:
            return False

        if pass2_result.returncode != 0:
            utils.error_print(
                "Compression pass 2 failed",
                [pass2_result.stderr.strip() if pass2_result.stderr else "Unknown error"],
            )
            return False

        if not _verify_output_file(compressed_temp_path, "Compression"):
            return False

        new_size = Path(compressed_temp_path).stat().st_size

        os.replace(compressed_temp_path, filepath)

        utils.verbose_print(f"  Compressed: {utils.format_filesize(current_size_bytes)} -> {utils.format_filesize(new_size)}")

        if new_size > target_size_bytes:
            utils.warning_print(f"Compressed file still exceeds target ({utils.format_filesize(new_size)} > {target_size_mb}MB)",
                ["The video may need a higher size limit or shorter duration."])

        return True

    except OSError as e:
        utils.error_print(f"Compression failed: {e}")
        return False
    finally:
        for ext in ['-0.log', '-0.log.mbtree', '']:
            log_path = Path(passlog_base + ext)
            if log_path.exists():
                try:
                    log_path.unlink()
                except OSError as e:
                    utils.debug_print(f"Could not remove passlog file '{log_path}': {e}")
        temp_path = Path(compressed_temp_path)
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError as e:
                utils.warning_print(f"Could not remove temp file: {compressed_temp_path}", [str(e)])


def concatenate_clips(clip_paths: List[str], output_file: str, reencode_on_fail: bool = True) -> bool:
    """Concatenate multiple video clips into a single file using ffmpeg concat demuxer.

    Writes a temporary file list for ffmpeg, runs concat demuxer with stream copy,
    and optionally falls back to re-encoding if stream copy fails (e.g. codec mismatch).

    Args:
        clip_paths: List of paths to clip files (order preserved)
        output_file: Path for the concatenated output file
        reencode_on_fail: If True, retry with re-encoding when stream copy fails

    Returns:
        True if concatenation succeeded, False otherwise.
    """
    if not clip_paths:
        utils.error_print("No clips to concatenate.", ["clip_paths must not be empty."])
        return False

    for path in clip_paths:
        if not Path(path).is_file():
            utils.error_print(f"Clip file not found: '{path}'",
                ["Ensure all clips were generated successfully before concatenating."])
            return False

    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as file_handle:
        concat_list_file = file_handle.name
        for path in clip_paths:
            abs_path = str(Path(path).resolve())
            escaped_path = abs_path.replace("'", "'\\''")
            file_handle.write(f"file '{escaped_path}'\n")

    try:
        ffmpeg_command = [
            'ffmpeg', '-y', '-loglevel', config.FFMPEG_LOGLEVEL,
            '-f', 'concat', '-safe', '0', '-i', concat_list_file,
            '-c', 'copy',
            output_file,
        ]
        utils.verbose_print(f'Concatenating {len(clip_paths)} clips into {output_file}.')
        utils.debug_print(f"ffmpeg concat command: {' '.join(ffmpeg_command)}")
        if config.DEBUGGING:
            utils.debug_print('Debugging enabled, not calling ffmpeg for concat.')
            return False

        ffmpeg_result = _run_ffmpeg_process(
            ffmpeg_command,
            input_file=concat_list_file,
            output_file=output_file,
            os_error_message="Concatenation failed.",
        )
        if ffmpeg_result is None:
            return False

        if ffmpeg_result.returncode != 0 and reencode_on_fail:
            utils.warning_print("Stream copy concat failed (e.g. codec mismatch), retrying with re-encoding.")
            ffmpeg_command_reencode = [
                'ffmpeg', '-y', '-loglevel', config.FFMPEG_LOGLEVEL,
                '-f', 'concat', '-safe', '0', '-i', concat_list_file,
                '-c:v', 'libx264', '-c:a', 'aac',
                output_file,
            ]
            ffmpeg_result = _run_ffmpeg_process(
                ffmpeg_command_reencode,
                input_file=concat_list_file,
                output_file=output_file,
                os_error_message="Concatenation failed during re-encoding fallback.",
            )
            if ffmpeg_result is None:
                return False

        if ffmpeg_result.returncode != 0:
            error_details = [f"Output: '{output_file}'", f"Clips: {len(clip_paths)} files"]
            utils.error_print("ffmpeg concat failed.", _add_ffmpeg_stderr(error_details, ffmpeg_result))
            return False

        if not _verify_output_file(output_file, "Concat"):
            return False

        utils.verbose_print(f"+ Generated reel '{output_file}' successfully.")
        return True
    except OSError as e:
        utils.error_print(f"Concatenation failed: {e}")
        return False
    finally:
        concat_path = Path(concat_list_file)
        if concat_path.exists():
            try:
                concat_path.unlink()
            except OSError as e:
                utils.debug_print(f"Could not remove concat list file '{concat_list_file}': {e}")
