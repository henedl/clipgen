"""Screenspace CLI modes: --ss-task / --ss-rerun / --ss-list-* parsing and runs.

Carved out of cli.py; the mode table there points at these callables. Heavy
modules (screenspace, pipeline) are imported inside the functions so --help
stays fast.
"""

from __future__ import annotations

import argparse
import copy
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

import config
import files
import utils
import video


_SS_VALID_TASK_TYPES = (
    "color",
    "change",
    "similarity",
    "text",
    "numbers",
    "timelapse",
    "template",
    "shape",
    "flow",
    "inactivity",
    "scene",
    "attention",
)


def _ss_resolve_videos_for_participant(participant_id: str) -> list[str]:
    """Resolve a participant's ordered source video path(s) via filename discovery.

    Honours ``config.FILENAME_OVERRIDES`` the same way the web tools do.
    A multi-video participant (numbered parts) returns all parts in timeline order;
    a normal participant returns a single-element list. Returns [] when unknown.
    """
    for entry in files.resolve_participant_videos():
        if entry["id"] == participant_id and entry.get("has_video"):
            return list(entry["video_paths"])
    return []


def _ss_frame_extractor(video_paths: list[str]) -> Callable[[float], Any | None]:
    """Return a ``frame_at(global_ts)`` closure mapping into the right sub-video.

    For reference-frame extraction (similarity/template/scene) over a participant
    whose recording spans several files: a global reference timestamp resolves to
    the owning sub-video. Single-video participants extract unchanged (no probe).
    """
    timeline = video.timeline_or_none(video_paths)

    def _extract(global_ts: float) -> Any | None:
        if timeline is None:
            return video.extract_frame_at_timestamp(video_paths[0], global_ts)
        mapped = utils.map_global_to_segment(timeline, global_ts)
        if mapped is None:
            return None
        return video.extract_frame_at_timestamp(timeline[mapped[0]][0], mapped[1])

    return _extract


def _ss_hex_to_hsv(hex_str: str) -> dict[str, int]:
    """Convert a #RRGGBB hex string to OpenCV HSV (H 0–180, S 0–255, V 0–255)."""
    import cv2
    import numpy as np

    s = hex_str.strip().lstrip("#")
    if len(s) != 6:
        raise ValueError(f"Invalid hex colour {hex_str!r} (expected #RRGGBB)")
    try:
        r, g, b = int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)
    except ValueError as exc:
        raise ValueError(f"Invalid hex colour {hex_str!r}") from exc
    bgr = np.array([[[b, g, r]]], dtype=np.uint8)
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)[0][0]
    return {"h": int(hsv[0]), "s": int(hsv[1]), "v": int(hsv[2])}


def _ss_parse_tolerance(tol_str: str) -> dict[str, int]:
    """Parse a comma-separated H,S,V tolerance triple."""
    parts = [p.strip() for p in tol_str.split(",")]
    if len(parts) != 3:
        raise ValueError(
            f"Tolerance must be three comma-separated ints (got {tol_str!r})"
        )
    try:
        h, s, v = (int(parts[0]), int(parts[1]), int(parts[2]))
    except ValueError as exc:
        raise ValueError(
            f"Tolerance must be three comma-separated ints (got {tol_str!r})"
        ) from exc
    return {"h": h, "s": s, "v": v}


def _ss_parse_scene_ref(raw: str) -> dict[str, Any]:
    """Parse a NAME:TIMESTAMP[:THRESHOLD] scene reference (TIMESTAMP in seconds).

    Returns ``{"name", "timestamp"}`` plus optional ``"threshold"``. NAME must not
    contain ``:``; TIMESTAMP and THRESHOLD must be numeric.
    """
    parts = raw.split(":")
    if len(parts) not in (2, 3):
        raise ValueError(
            f"Scene reference must be NAME:TIMESTAMP[:THRESHOLD] (got {raw!r})"
        )
    name = parts[0].strip()
    if not name:
        raise ValueError(f"Scene reference NAME is required (got {raw!r})")
    try:
        ref: dict[str, Any] = {"name": name, "timestamp": float(parts[1])}
        if len(parts) == 3:
            ref["threshold"] = float(parts[2])
    except ValueError as exc:
        raise ValueError(
            f"Scene reference TIMESTAMP/THRESHOLD must be numeric (got {raw!r})"
        ) from exc
    if "threshold" in ref and not (0.0 <= ref["threshold"] <= 1.0):
        raise ValueError(
            f"Scene reference THRESHOLD must be between 0 and 1 (got {raw!r})"
        )
    return ref


def _ss_build_params(
    args: argparse.Namespace,
    task_type: str,
    region_coords: dict[str, int],
    frame_at: Callable[[float], Any | None],
) -> dict[str, Any]:
    """Build a `parameters` dict for create_task() from per-tool CLI flags.

    Validates that the required flags for ``task_type`` are present. For
    ``similarity`` and ``template`` extracts the reference frame at
    ``--ss-reference-timestamp`` via *frame_at* (which maps a global timestamp
    into the owning sub-video for multi-video participants; mirrors the
    server-side path in screenspace_server._extract_tool_media).
    """
    import screenspace

    params: dict[str, Any] = {}

    if args.ss_interval is not None:
        params["interval"] = args.ss_interval
    if args.ss_start is not None:
        params["start_seconds"] = args.ss_start
    if args.ss_end is not None:
        params["end_seconds"] = args.ss_end
    if args.ss_event_label:
        params["event_label"] = args.ss_event_label

    if task_type == "color":
        if not args.ss_target_color:
            raise ValueError("color task requires --ss-target-color HEX")
        if not args.ss_tolerance:
            raise ValueError("color task requires --ss-tolerance H,S,V")
        params["target_color"] = _ss_hex_to_hsv(args.ss_target_color)
        params["tolerance"] = _ss_parse_tolerance(args.ss_tolerance)
        if args.ss_color_mode == "presence":
            params["color_mode"] = "presence"
            if args.ss_min_area is not None:
                params["min_coverage"] = args.ss_min_area / 100.0

    elif task_type == "change":
        if args.ss_threshold is None:
            raise ValueError("change task requires --ss-threshold FLOAT")
        params["threshold"] = args.ss_threshold

    elif task_type == "similarity":
        if args.ss_reference_timestamp is None:
            raise ValueError(
                "similarity task requires --ss-reference-timestamp SECONDS"
            )
        if args.ss_threshold is None:
            raise ValueError("similarity task requires --ss-threshold FLOAT")
        params["reference_timestamp"] = args.ss_reference_timestamp
        params["threshold"] = args.ss_threshold
        frame = frame_at(float(args.ss_reference_timestamp))
        if frame is None:
            raise ValueError(
                f"Could not extract reference frame at {args.ss_reference_timestamp}s"
            )
        params["reference_frame"] = screenspace.extract_region(frame, region_coords)

    elif task_type == "text":
        if not args.ss_text:
            raise ValueError("text task requires --ss-text STR")
        params["search_string"] = args.ss_text
        if args.ss_fuzzy_threshold is not None:
            params["fuzzy_threshold"] = args.ss_fuzzy_threshold

    elif task_type == "numbers":
        if not args.ss_operator:
            raise ValueError("numbers task requires --ss-operator OP")
        params["operator"] = args.ss_operator
        if args.ss_operator == "range":
            if args.ss_range_min is None or args.ss_range_max is None:
                raise ValueError(
                    "numbers task with operator=range requires "
                    "--ss-range-min and --ss-range-max"
                )
            params["range_min"] = args.ss_range_min
            params["range_max"] = args.ss_range_max
        else:
            if args.ss_target_value is None:
                raise ValueError(
                    f"numbers task with operator={args.ss_operator} requires --ss-target-value"
                )
            params["target_value"] = args.ss_target_value

    elif task_type == "timelapse":
        if args.ss_speedup is None:
            raise ValueError("timelapse task requires --ss-speedup FACTOR")
        params["speedup_factor"] = args.ss_speedup
        if args.ss_output_format:
            params["output_format"] = args.ss_output_format

    elif task_type == "template":
        if args.ss_reference_timestamp is None:
            raise ValueError("template task requires --ss-reference-timestamp SECONDS")
        if args.ss_threshold is None:
            raise ValueError("template task requires --ss-threshold FLOAT")
        params["reference_timestamp"] = args.ss_reference_timestamp
        params["threshold"] = args.ss_threshold
        frame = frame_at(float(args.ss_reference_timestamp))
        if frame is None:
            raise ValueError(
                f"Could not extract template frame at {args.ss_reference_timestamp}s"
            )
        params["template_image"] = screenspace.extract_region(frame, region_coords)

    elif task_type == "shape":
        if args.ss_reference_timestamp is None:
            raise ValueError("shape task requires --ss-reference-timestamp SECONDS")
        params["reference_timestamp"] = args.ss_reference_timestamp
        if args.ss_threshold is not None:
            params["threshold"] = args.ss_threshold
        if args.ss_scale_min is not None:
            params["scale_min"] = args.ss_scale_min
        if args.ss_scale_max is not None:
            params["scale_max"] = args.ss_scale_max
        if args.ss_scale_steps is not None:
            params["scale_steps"] = args.ss_scale_steps
        if args.ss_scale_y_min is not None:
            params["scale_y_min"] = args.ss_scale_y_min
        if args.ss_scale_y_max is not None:
            params["scale_y_max"] = args.ss_scale_y_max
        if args.ss_scale_y_steps is not None:
            params["scale_y_steps"] = args.ss_scale_y_steps
        frame = frame_at(float(args.ss_reference_timestamp))
        if frame is None:
            raise ValueError(
                f"Could not extract shape frame at {args.ss_reference_timestamp}s"
            )
        params["shape_image"] = screenspace.extract_region(frame, region_coords)

    elif task_type == "scene":
        raw_refs = getattr(args, "ss_scene_ref", None) or []
        if not raw_refs:
            raise ValueError(
                "scene task requires at least one --ss-scene-ref NAME:TIMESTAMP[:THRESHOLD]"
            )
        parsed = [_ss_parse_scene_ref(r) for r in raw_refs]
        reference_scenes = []
        for ref in parsed:
            frame = frame_at(float(ref["timestamp"]))
            if frame is None:
                raise ValueError(
                    f"Could not extract scene frame for {ref['name']!r} "
                    f"at {ref['timestamp']}s"
                )
            entry: dict[str, Any] = {
                "name": ref["name"],
                "frame": screenspace.extract_region(frame, region_coords),
            }
            if "threshold" in ref:
                entry["threshold"] = ref["threshold"]
            reference_scenes.append(entry)
        params["reference_scenes"] = reference_scenes
        # Frames are stripped on manifest save; the input form keeps --ss-run-task
        # working.
        params["scene_references"] = parsed
        if args.ss_threshold is not None:
            params["threshold"] = args.ss_threshold

    elif task_type == "flow":
        if args.ss_threshold is None:
            raise ValueError("flow task requires --ss-threshold FLOAT (magnitude)")
        params["magnitude_threshold"] = args.ss_threshold

    elif task_type == "inactivity":
        if args.ss_threshold is None:
            raise ValueError("inactivity task requires --ss-threshold FLOAT")
        params["threshold"] = args.ss_threshold

    elif task_type == "attention":
        # All knobs have config defaults; --ss-threshold optionally overrides
        # the normalized peak-jump distance for shift events.
        if args.ss_threshold is not None:
            params["shift_threshold"] = args.ss_threshold

    params.setdefault("cv_resolution_scale", config.SCREENSPACE_CV_RESOLUTION_SCALE)
    return params


def _print_ss_table(
    title: str,
    columns: list[tuple[str, dict[str, Any]]],
    rows: list[list[str]],
    fallback_lines: list[str],
) -> None:
    """Render a Rich table when available, else fall back to plain info_print lines.

    columns is a list of (name, kwargs) tuples passed to Table.add_column.
    """
    if utils._use_rich() and utils.console is not None:
        from rich.table import Table

        table = Table(
            title=title,
            show_header=True,
            header_style="bold cyan",
            border_style="dim",
            row_styles=["", "dim"],
            expand=False,
        )
        for name, kwargs in columns:
            table.add_column(name, **kwargs)
        for row in rows:
            table.add_row(*row)
        utils.console.print(table)
    else:
        for line in fallback_lines:
            utils.info_print(line)


def _run_ss_list_regions(args: argparse.Namespace) -> None:
    """List active Screenspace regions from the manifest."""
    import screenspace

    manifest = screenspace.load_screenspace_manifest()
    regions = manifest.get("regions", {})
    if not regions:
        utils.info_print("No active Screenspace regions in manifest.")
        return

    rows: list[list[str]] = []
    fallback: list[str] = [f"Active regions ({len(regions)}):"]
    for name in sorted(regions.keys()):
        rd = regions[name]
        sw = rd.get("source_width", "?")
        sh = rd.get("source_height", "?")
        rows.append(
            [
                name,
                f"{rd.get('x', 0):.3f}",
                f"{rd.get('y', 0):.3f}",
                f"{rd.get('w', 0):.3f}",
                f"{rd.get('h', 0):.3f}",
                f"{sw}x{sh}",
            ]
        )
        fallback.append(
            f"  {name}: x={rd.get('x', 0):.3f} y={rd.get('y', 0):.3f} "
            f"w={rd.get('w', 0):.3f} h={rd.get('h', 0):.3f}  source={sw}x{sh}"
        )

    _print_ss_table(
        f"Active regions ({len(regions)})",
        [
            ("Name", {"style": "bold"}),
            ("x", {"justify": "right"}),
            ("y", {"justify": "right"}),
            ("w", {"justify": "right"}),
            ("h", {"justify": "right"}),
            ("Source", {"justify": "right"}),
        ],
        rows,
        fallback,
    )


def _run_ss_list_stashes(args: argparse.Namespace) -> None:
    """List Screenspace region stashes."""
    import screenspace

    manifest = screenspace.load_screenspace_manifest()
    stashes = manifest.get("stashes", [])
    if not stashes:
        utils.info_print("No Screenspace region stashes in manifest.")
        return

    rows: list[list[str]] = []
    fallback: list[str] = [f"Stashes ({len(stashes)}):"]
    for stash in stashes:
        regions = stash.get("regions", {})
        names = ", ".join(sorted(regions.keys())) or "(empty)"
        rows.append(
            [
                str(stash.get("id", "?")),
                str(stash.get("name", "(unnamed)")),
                str(len(regions)),
                names,
            ]
        )
        fallback.append(
            f"  {stash.get('id', '?')}  {stash.get('name', '(unnamed)')}: "
            f"{len(regions)} region(s) — {names}"
        )

    _print_ss_table(
        f"Stashes ({len(stashes)})",
        [
            ("ID", {"style": "bold"}),
            ("Name", {}),
            ("Regions", {"justify": "right"}),
            ("Names", {"overflow": "fold"}),
        ],
        rows,
        fallback,
    )


def _run_ss_list_tasks(args: argparse.Namespace) -> None:
    """List Screenspace tasks from the manifest, optionally filtered by status."""
    import screenspace

    manifest = screenspace.load_screenspace_manifest()
    tasks = manifest.get("tasks", [])
    status_filter = (args.ss_list_tasks or "").strip().lower() or None
    if status_filter:
        tasks = [t for t in tasks if (t.get("status") or "").lower() == status_filter]
    if not tasks:
        if status_filter:
            utils.info_print(f"No Screenspace tasks with status={status_filter}.")
        else:
            utils.info_print("No Screenspace tasks in manifest.")
        return

    label = f" (status={status_filter})" if status_filter else ""
    rows: list[list[str]] = []
    fallback: list[str] = [f"Tasks{label}: {len(tasks)}"]
    for t in tasks:
        result = t.get("result")
        result_count = len(result) if isinstance(result, list) else 0
        rows.append(
            [
                str(t.get("id", "?")),
                str(t.get("type", "?")),
                str(t.get("participant", "?")),
                str(t.get("region", "?")),
                str(t.get("status", "?")),
                str(result_count),
            ]
        )
        fallback.append(
            f"  {t.get('id', '?')}  {t.get('type', '?'):10s}  "
            f"{t.get('participant', '?'):8s}  region={t.get('region', '?'):16s}  "
            f"status={t.get('status', '?'):10s}  results={result_count}"
        )

    _print_ss_table(
        f"Tasks{label} ({len(tasks)})",
        [
            ("ID", {"style": "bold"}),
            ("Type", {}),
            ("Participant", {}),
            ("Region", {}),
            ("Status", {}),
            ("Results", {"justify": "right"}),
        ],
        rows,
        fallback,
    )


def _run_ss_task(args: argparse.Namespace) -> None:
    """Run a Screenspace analysis task synchronously and persist the result."""
    import screenspace

    # REGION is optional; two args scan the full frame.
    ss_task = list(args.ss_task)
    if len(ss_task) == 2:
        ss_task.append(screenspace.FULL_FRAME_REGION_NAME)
    if len(ss_task) != 3:
        utils.error_print(
            "--ss-task takes TYPE PARTICIPANT [REGION].",
            ["Omit REGION (or pass 'full_frame') to scan the whole frame."],
        )
        sys.exit(1)
    task_type, participant, region_name = ss_task
    if task_type not in _SS_VALID_TASK_TYPES:
        utils.error_print(
            f"Unknown screenspace task type {task_type!r}.",
            [f"Valid types: {', '.join(_SS_VALID_TASK_TYPES)}"],
        )
        sys.exit(1)
    if task_type == "attention" and region_name != screenspace.FULL_FRAME_REGION_NAME:
        # Attention is full-frame only (the server forces this too); a region would
        # mislabel events.
        utils.warning_print(
            f"attention is full-frame only; ignoring region {region_name!r}."
        )
        region_name = screenspace.FULL_FRAME_REGION_NAME

    manifest = screenspace.load_screenspace_manifest()

    # Active regions win over same-named stashed copies; same resolver as
    # --ss-run-task.
    try:
        _, resolved_region = screenspace.resolve_region_request(
            region_name, None, manifest
        )
    except ValueError:
        # The hint lists every known name, active and stashed.
        available = sorted(
            {
                *manifest.get("regions", {}),
                *(
                    name
                    for stash in manifest.get("stashes", [])
                    for name in stash.get("regions", {})
                ),
            }
        )
        hint = (
            f"Available regions: {', '.join(available)}"
            if available
            else "No regions defined. Use --screenspace to define regions in the web UI first."
        )
        utils.error_print(f"Region {region_name!r} not found.", [hint])
        sys.exit(1)

    video_paths = _ss_resolve_videos_for_participant(participant)
    if not video_paths:
        utils.error_print(
            f"No video found for participant {participant!r}.",
            ["Place the source video in the input directory before running --ss-task."],
        )
        sys.exit(1)

    # Parts share resolution; reference frames map global→sub-video via frame_at.
    props = video.probe_video_properties(video_paths[0])
    if props and props.get("width") and props.get("height"):
        region_coords = screenspace.denormalize_region(
            resolved_region, int(props["width"]), int(props["height"])
        )
    else:
        region_coords = {
            k: int(resolved_region[k])
            for k in ("x", "y", "w", "h")
            if k in resolved_region
        }

    try:
        parameters = _ss_build_params(
            args, task_type, region_coords, _ss_frame_extractor(video_paths)
        )
    except ValueError as exc:
        utils.error_print(str(exc))
        sys.exit(1)

    source_video = Path(video_paths[0]).name
    task = screenspace.create_task(
        task_type=task_type,
        participant=participant,
        source_video=source_video,
        video_paths=video_paths,
        region_name=region_name,
        region_coords=region_coords,
        parameters=parameters,
    )

    _ss_run_and_persist_task(task, manifest)


def _ss_run_and_persist_task(task: dict[str, Any], manifest: dict[str, Any]) -> None:
    """Enqueue a task, poll to completion, persist the manifest, and report.

    Shared by --ss-task (flag-built tasks) and --ss-run-task (manifest re-runs).
    Restores the manifest's historical tasks so they survive the save, then enqueues
    only ``task`` for execution.
    """
    import time

    import screenspace

    task_type = task.get("type", "")
    participant = task.get("participant", "")
    region_name = task.get("region", "")

    worker = screenspace.ScreenspaceWorker()
    worker.restore_tasks(manifest.get("tasks", []))
    worker.start()
    task_id = worker.enqueue(task)
    utils.info_print(f"Running {task_type} on {participant} (region: {region_name})...")

    final_task: dict[str, Any] | None = None
    try:
        with utils.progress_scope(f"{task_type} (queued)", 100) as ps:
            last_progress = -1.0
            while True:
                current = worker.get_task(task_id)
                if current is None:
                    break
                status = current.get("status", "")
                progress = float(current.get("progress", 0.0))
                if ps.live:
                    ps.update(
                        completed=int(progress * 100),
                        description=f"{task_type} ({status})",
                    )
                elif progress - last_progress > 0.05:
                    utils.info_print(f"  {status}: {int(progress * 100)}%")
                    last_progress = progress
                if status in ("completed", "failed", "cancelled"):
                    final_task = current
                    ps.update(completed=100)
                    break
                time.sleep(0.25)
    finally:
        new_events = worker.drain_new_events()
        all_tasks = worker.get_all_tasks()
        events = list(manifest.get("events", [])) + new_events
        screenspace.save_screenspace_manifest(
            manifest.get("regions", {}),
            all_tasks,
            events,
            stashes=manifest.get("stashes", []),
            per_participant=manifest.get("per_participant", {}),
            pins=manifest.get("pins") or {},
        )
        worker.stop()

    if final_task is None:
        utils.error_print("Task did not complete (no final state).")
        sys.exit(1)

    status = final_task.get("status", "")
    if status == "completed":
        result = final_task.get("result")
        result_count = len(result) if isinstance(result, list) else 0
        if task_type == "timelapse":
            output = (
                result[0].get("output_path")
                if isinstance(result, list) and result
                else None
            )
            if output:
                utils.info_print(f"Timelapse written to {output}")
            else:
                utils.info_print("Timelapse complete.")
        else:
            utils.info_print(f"Completed: {result_count} result(s).")
    elif status == "failed":
        utils.error_print(f"Task failed: {final_task.get('error', 'unknown error')}")
        sys.exit(1)
    else:
        utils.info_print(f"Task ended with status={status}.")


def _ss_extract_scene_frames(
    scene_refs: list[dict[str, Any]],
    frame_at: Callable[[float], Any | None],
    region_coords: dict[str, int],
    *,
    context: str = "",
) -> list[dict[str, Any]]:
    """Build reference_scenes (with cropped frames) from saved scene_references.

    Mirrors the scene path of screenspace_server._extract_tool_media. *frame_at*
    maps a global timestamp into the owning sub-video for multi-video
    participants. Raises ValueError when a frame cannot be read.
    """
    import screenspace

    reference_scenes: list[dict[str, Any]] = []
    for ref in scene_refs:
        frame = frame_at(float(ref["timestamp"]))
        if frame is None:
            raise ValueError(
                f"{context}could not read frame for scene {ref.get('name')!r} "
                f"at {ref.get('timestamp')}s"
            )
        entry: dict[str, Any] = {
            "name": ref["name"],
            "frame": screenspace.extract_region(frame, region_coords),
        }
        if "threshold" in ref:
            entry["threshold"] = ref["threshold"]
        reference_scenes.append(entry)
    return reference_scenes


def _ss_reference_coords(
    parameters: dict[str, Any],
    region_coords: dict[str, Any],
    manifest: dict[str, Any],
    dims: tuple[int, int] | None,
) -> dict[str, Any]:
    """Pixel rect a template/shape re-run cuts its sample from.

    The persisted capture region (``reference_region``) wins over the run
    region, mirroring screenspace_server._prepare_task_media.
    """
    import screenspace

    ref_name = str(parameters.get("reference_region") or "").strip()
    if not ref_name:
        return region_coords
    _rn, ref_norm = screenspace.resolve_region_request(ref_name, None, manifest)
    if dims is None:
        return region_coords
    return screenspace.denormalize_region(ref_norm, dims[0], dims[1])


def _ss_rehydrate_task_media(
    task_type: str,
    parameters: dict[str, Any],
    frame_at: Callable[[float], Any | None],
    region_coords: dict[str, Any],
    manifest: dict[str, Any],
    dims: tuple[int, int] | None,
) -> None:
    """Re-extract reference frames/templates/scenes into a saved task's parameters.

    Mirrors screenspace_server._extract_tool_media and _prepare_multitool_steps so a
    manifest task (whose binary frame data was stripped on save) can be re-run.
    *frame_at* maps a global reference timestamp into the owning sub-video for
    multi-video participants. Mutates ``parameters`` in place; raises ValueError
    when a reference cannot be recovered (e.g. a multitool step built from an
    uploaded template image, which has no timestamp to re-extract from).
    """
    import screenspace

    def _extract_frame(ref_ts: float, coords: dict[str, Any], label: str) -> Any:
        frame = frame_at(float(ref_ts))
        if frame is None:
            raise ValueError(f"{label}: could not read reference frame")
        return screenspace.extract_region(frame, coords)

    if task_type == "similarity":
        if parameters.get("reference_timestamp") is None:
            raise ValueError("similarity task has no reference_timestamp to re-extract")
        parameters["reference_frame"] = _extract_frame(
            parameters["reference_timestamp"], region_coords, "similarity"
        )

    elif task_type == "template":
        if parameters.get("reference_timestamp") is None:
            raise ValueError(
                "template task built from an uploaded image cannot be re-run from the "
                "manifest (no reference timestamp was saved)"
            )
        coords = _ss_reference_coords(parameters, region_coords, manifest, dims)
        parameters["template_image"] = _extract_frame(
            parameters["reference_timestamp"],
            coords,
            "template",
        )
        screenspace.attach_capture_mask(
            parameters, "template_image", "template_mask", coords
        )

    elif task_type == "shape":
        if parameters.get("reference_timestamp") is None:
            raise ValueError(
                "shape task built from an uploaded image cannot be re-run from the "
                "manifest (no reference timestamp was saved)"
            )
        coords = _ss_reference_coords(parameters, region_coords, manifest, dims)
        parameters["shape_image"] = _extract_frame(
            parameters["reference_timestamp"],
            coords,
            "shape",
        )
        screenspace.attach_capture_mask(parameters, "shape_image", "shape_mask", coords)

    elif task_type == "scene":
        scene_refs = parameters.get("scene_references")
        if not scene_refs:
            raise ValueError("scene task has no scene_references to re-extract")
        parameters["reference_scenes"] = _ss_extract_scene_frames(
            scene_refs, frame_at, region_coords
        )

    elif task_type == "multitool":
        steps: list[dict[str, Any]] = parameters.get("steps", [])
        for i, step in enumerate(steps):
            stype = step.get("type", "")
            step_region_name = (step.get("region") or "").strip()
            step_region_ref = step.get("region_ref")
            if step_region_name or step_region_ref is not None:
                resolved_name, resolved_region = screenspace.resolve_region_request(
                    step_region_name, step_region_ref, manifest
                )
                step["region"] = resolved_name
                if dims is not None:
                    step_coords = screenspace.denormalize_region(
                        resolved_region, dims[0], dims[1]
                    )
                else:
                    step_coords = region_coords
            else:
                step_coords = region_coords
            step["region_coords"] = step_coords

            if stype == "similarity":
                if step.get("reference_timestamp") is None:
                    raise ValueError(f"Step {i}: no reference_timestamp to re-extract")
                step["reference_frame"] = _extract_frame(
                    step["reference_timestamp"], step_coords, f"Step {i}"
                )
            elif stype == "template":
                if step.get("reference_timestamp") is None:
                    raise ValueError(
                        f"Step {i}: template step built from an uploaded image cannot "
                        "be re-run from the manifest (no reference timestamp saved)"
                    )
                step["template_image"] = _extract_frame(
                    step["reference_timestamp"], step_coords, f"Step {i}"
                )
                screenspace.attach_capture_mask(
                    step, "template_image", "template_mask", step_coords
                )
            elif stype == "scene":
                step_refs = step.get("scene_references")
                if not step_refs:
                    raise ValueError(f"Step {i}: no scene_references to re-extract")
                step["reference_scenes"] = _ss_extract_scene_frames(
                    step_refs, frame_at, step_coords, context=f"Step {i}: "
                )


def _run_ss_rerun_task(args: argparse.Namespace) -> None:
    """Re-run a saved Screenspace task from the manifest by id.

    Re-extracts reference media from the source video (stripped on save), creates a
    fresh task run (new id, preserving the original), and persists the result. This is
    the only headless path for multitool tasks.
    """
    import screenspace

    task_id = args.ss_run_task
    manifest = screenspace.load_screenspace_manifest()
    saved = next((t for t in manifest.get("tasks", []) if t.get("id") == task_id), None)
    if saved is None:
        utils.error_print(
            f"Task {task_id!r} not found in the manifest.",
            ["List available tasks with --ss-list-tasks."],
        )
        sys.exit(1)

    task_type = saved.get("type", "")
    participant = saved.get("participant", "")
    region_name = saved.get("region", "")
    parameters = copy.deepcopy(saved.get("parameters", {}))

    video_paths = _ss_resolve_videos_for_participant(participant)
    if not video_paths:
        utils.error_print(
            f"No video found for participant {participant!r}.",
            ["Place the source video in the input directory before re-running."],
        )
        sys.exit(1)

    props = video.probe_video_properties(video_paths[0])
    dims: tuple[int, int] | None = None
    if props and props.get("width") and props.get("height"):
        dims = (int(props["width"]), int(props["height"]))

    # region_ref-aware resolution first; saved region_coords cover older tasks,
    # multitool steps, and missing dims.
    region_coords: dict[str, int] | None = None
    try:
        _, resolved_region = screenspace.resolve_region_request(
            region_name, saved.get("region_ref"), manifest
        )
        if dims is not None:
            region_coords = screenspace.denormalize_region(
                resolved_region, dims[0], dims[1]
            )
    except ValueError:
        pass  # fall through to saved coords below
    if region_coords is None and isinstance(saved.get("region_coords"), dict):
        region_coords = {
            k: int(saved["region_coords"][k])
            for k in ("x", "y", "w", "h")
            if k in saved["region_coords"]
        }
    if region_coords is None:
        utils.error_print(
            f"Region {region_name!r} for task {task_id!r} could not be resolved.",
            [
                "The named region is no longer in the manifest and no saved coords exist."
            ],
        )
        sys.exit(1)

    try:
        _ss_rehydrate_task_media(
            task_type,
            parameters,
            _ss_frame_extractor(video_paths),
            region_coords,
            manifest,
            dims,
        )
    except ValueError as exc:
        utils.error_print(f"Cannot re-run task {task_id!r}: {exc}")
        sys.exit(1)

    task = screenspace.create_task(
        task_type=task_type,
        participant=participant,
        source_video=Path(video_paths[0]).name,
        video_paths=video_paths,
        region_name=region_name,
        region_coords=region_coords,
        parameters=parameters,
    )

    _ss_run_and_persist_task(task, manifest)


# ---- Event-driven clip cutting (--ss-clips, --transcript-clips) ----
