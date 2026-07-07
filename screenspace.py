# -*- coding: utf-8 -*-
"""Screenspace analysis engine for clipgen.

Eleven analysis tools (passed as 'type' when creating a task):
  multitool   – chain multiple tools; each subsequent step only checks frames that passed previous steps
  color       – frames where a region's average HSV color matches a target within tolerance
  change      – frames where pixel diff ratio exceeds SCREENSPACE_CHANGE_RATIO_THRESHOLD
  similarity  – frames matching a reference capture via SSIM (SCREENSPACE_SSIM_THRESHOLD)
  text        – OCR fuzzy search for a query string (SCREENSPACE_OCR_FUZZY_THRESHOLD); requires EasyOCR
  numbers     – OCR numeric comparison with a relational condition (eq/gt/lt/gte/lte/range)
  timelapse   – sped-up video of a region over a time range
  template    – find a reference image/template anywhere in the full frame via cv2.matchTemplate
  flow        – detect motion in a region via dense optical flow (cv2.calcOpticalFlowFarneback)
  scene       – classify frames by similarity to user-captured reference scenes
  inactivity  – detect spans of near-duplicate frames via perceptual hashing (loading screens, frozen states)

Workflow: user draws regions on a frame → enqueues tasks → ScreenspaceWorker processes in
a background thread → results are timestamps or artifact files → state persisted to
screenspace_manifest.json. Region coordinates are normalized (0–1); source_width/source_height
are stored for denormalization to target video resolution.

This module is a thin re-export facade. The implementation lives in cohesive sibling
modules (imported deepest-first below); ``import screenspace; screenspace.NAME`` keeps
resolving every public name — and the private names the test suite reaches for — from
their new homes:

  screenspace_primitives  – pure cv2/numpy region + image-analysis primitives
  screenspace_ocr         – cached EasyOCR readers, number/text scoring helpers
  screenspace_frames      – ffmpeg-pipe frame extraction + ffprobe metadata
  screenspace_scans       – the eleven per-tool scan workflows
  screenspace_heatmap     – template/flow/change heatmap PNG + cumulative/rolling GIF generation
  screenspace_tools       – AnalysisTool registry + per-frame dispatch
  screenspace_multitool   – multitool chaining + offset joining
  screenspace_manifest    – task/manifest persistence + event generation
  screenspace_worker      – the background task-queue worker
"""

# ruff: noqa: F401
# Deepest-first re-export so ``screenspace.NAME`` resolves from the new modules.

# ``screenspace.utils`` is part of the public surface (a test monkeypatches
# ``screenspace.utils.warning_print``); ``utils`` is a singleton module so the
# patch is visible to every sibling that does ``import utils``.
import utils

from screenspace_primitives import (
    FULL_FRAME_REGION,
    FULL_FRAME_REGION_NAME,
    ScanCallback,
    _ConsecutiveBuffer,
    _merge_timestamp_spans,
    _morph_kernel,
    _prepare_template,
    _template_correlation_map,
    average_color_hsv,
    color_matches,
    color_present,
    compare_scene_fingerprints,
    compute_frame_diff,
    compute_optical_flow,
    compute_phash,
    compute_scene_fingerprint,
    denormalize_region,
    extract_region,
    match_template,
    regions_are_similar,
    resolve_region_request,
)
from screenspace_ocr import (
    _preprocess_for_ocr,
    _score_numbers_readings,
    _score_text_readings,
    run_calibration_ocr,
)

# Re-exporting a name here only rebinds it on the facade — it does NOT propagate
# to siblings that imported it (e.g. ``screenspace_scans._probe_video_meta``). To
# stub a seam in a test, patch the owning module, not the facade.
from screenspace_frames import (
    _ffmpeg_pipe_frames,
    _probe_video_meta,
    _resolve_scan_window,
    _scan_via_ffmpeg_pipe,
    build_timelapse_command,
    scan_video_frames,
    scan_video_full_frames,
)
from screenspace_scans import (
    generate_timelapse,
    scan_boundaries,
    scan_changes,
    scan_color,
    scan_flow,
    scan_inactivity,
    scan_numbers,
    scan_scene,
    scan_similarity,
    scan_template,
    scan_text,
)
from screenspace_heatmap import (
    generate_change_heatmap,
    generate_flow_heatmap,
    generate_heatmap_gif,
    generate_rolling_heatmap_gif,
    generate_template_heatmap,
)
from screenspace_tools import (
    TOOLS,
    AnalysisTool,
    BoundaryTool,
    ChangeTool,
    ColorTool,
    FlowTool,
    InactivityTool,
    MultitoolTool,
    NumbersTool,
    SceneTool,
    SimilarityTool,
    TemplateTool,
    TextTool,
    TimelapseTool,
    _extract_confidence,
    check_frame_for_tool,
    score_frame_for_tool,
)
from screenspace_multitool import scan_multitool, score_multitool_frame
from screenspace_manifest import (
    TASK_STATUS_CANCELLED,
    TASK_STATUS_COMPLETED,
    TASK_STATUS_FAILED,
    TASK_STATUS_PAUSED,
    TASK_STATUS_QUEUED,
    TASK_STATUS_RUNNING,
    _empty_screenspace_manifest,
    _offset_result_times,
    create_event,
    create_task,
    generate_events_from_results,
    load_screenspace_manifest,
    save_screenspace_manifest,
)
from screenspace_worker import ScreenspaceWorker
