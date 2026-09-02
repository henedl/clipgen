"""Screenspace analysis engine for clipgen.

Fourteen analysis tools (passed as 'type' when creating a task):
  multitool   – chain multiple tools; each subsequent step only checks frames that passed previous steps
  color       – frames where a region's average HSV color matches a target within tolerance
  change      – frames where pixel diff ratio exceeds SCREENSPACE_CHANGE_RATIO_THRESHOLD
  similarity  – frames matching a reference capture via SSIM (SCREENSPACE_SSIM_THRESHOLD)
  text        – OCR fuzzy search for a query string (SCREENSPACE_OCR_FUZZY_THRESHOLD); requires RapidOCR
  numbers     – OCR numeric comparison with a relational condition (eq/gt/lt/gte/lte/range)
  timelapse   – sped-up video of a region over a time range
  template    – find a reference image/template in the run region via cv2.matchTemplate
  shape       – find a reference shape's outline in the run region via scale-swept edge matching (color/size tolerant)
  flow        – detect motion in a region via dense optical flow (cv2.calcOpticalFlowFarneback)
  scene       – classify frames by similarity to user-captured reference scenes
  inactivity  – detect spans of near-duplicate frames via perceptual hashing (loading screens, frozen states)
  boundary    – segment the full frame into scene periods via hash/fingerprint shifts
  attention   – predict where visual attention goes via a saliency composite (full-frame heatmaps + shift events)

Workflow: user draws regions on a frame → enqueues tasks → ScreenspaceWorker processes in
a background thread → results are timestamps or artifact files → state persisted to
the ``screenspace`` section of the output-dir manifest. Region coordinates are normalized (0–1); source_width/source_height
are stored for denormalization to target video resolution.

This module is a thin re-export facade: the implementation lives in the cohesive
``screenspace_*`` siblings imported below (deepest-first), so ``screenspace.NAME``
keeps resolving every public name — and the private names tests reach for.
Per-module roles are tabulated in agents/ARCHITECTURE.md.
"""

# ruff: noqa: F401
# Deepest-first re-export so ``screenspace.NAME`` resolves from the new modules.

# Public: a test monkeypatches ``screenspace.utils.warning_print``, reaching
# every sibling.
import utils

from screenspace_primitives import (
    FULL_FRAME_REGION,
    FULL_FRAME_REGION_NAME,
    ScanCallback,
    PHash,
    _ConsecutiveBuffer,
    _frame_edge_map,
    _merge_timestamp_spans,
    _morph_kernel,
    _prepare_shape_reference,
    _prepare_template,
    _template_correlation_map,
    average_color_hsv,
    attach_capture_mask,
    canny_edges,
    blur_gray,
    color_matches,
    color_present,
    compare_scene_fingerprints,
    compute_color_contrast,
    compute_face_saliency,
    compute_frame_diff,
    compute_frame_diff_gray,
    compute_motion_saliency,
    compute_optical_flow,
    compute_phash,
    compute_saliency_map,
    compute_scene_fingerprint,
    compute_spectral_residual,
    denormalize_region,
    extract_region,
    face_detection_available,
    filter_matches_by_region_mask,
    flow_downscale,
    mask_points_key,
    match_shape,
    match_template,
    mean_gray_diff,
    nms_boxes_iou,
    point_in_mask_points,
    region_mask_for,
    region_masker,
    regions_are_similar,
    resolve_region_request,
    saliency_grid_from_map,
    region_search_window,
    sparse_grid_cells,
    saliency_kwargs_from_params,
    saliency_peak,
    ssim_diff_map,
    structural_similarity,
)
from screenspace_ocr import (
    _OCR_LANG_TO_MODEL,
    _preprocess_for_ocr,
    _resolve_ocr_model,
    _score_numbers_readings,
    _score_text_readings,
    run_calibration_ocr,
)

# Re-exports rebind only the facade; tests must patch the owning module.
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
    scan_attention,
    scan_boundaries,
    scan_changes,
    scan_color,
    scan_flow,
    scan_inactivity,
    scan_numbers,
    scan_scene,
    scan_shape,
    scan_similarity,
    scan_template,
    scan_text,
)
from screenspace_heatmap import (
    GIF_FRAMES,
    GridLayers,
    build_gif_sprite_bytes,
    build_grid_layers,
    generate_attention_heatmap,
    generate_change_heatmap,
    generate_flow_heatmap,
    generate_heatmap_gif,
    generate_rolling_heatmap_gif,
    generate_template_heatmap,
    grid_layer_count,
    sprite_grid,
)
from screenspace_tools import (
    TOOLS,
    AnalysisTool,
    AttentionTool,
    BoundaryTool,
    ChangeTool,
    ColorTool,
    FlowTool,
    InactivityTool,
    MultitoolTool,
    NumbersTool,
    SceneTool,
    ShapeTool,
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
    TASK_BINARY_KEYS,
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
    describe_task,
    generate_events_from_results,
    load_screenspace_manifest,
    save_screenspace_manifest,
    strip_task_param_binaries,
)
from screenspace_worker import ScreenspaceWorker
