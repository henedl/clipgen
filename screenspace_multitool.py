# -*- coding: utf-8 -*-
"""Screenspace multitool chaining.

Chains multiple tools (each later step only re-checks frames that passed the
earlier ones), with optional time-offset windows joined into combined events.
Imports the single-frame dispatch from screenspace_tools and a full-frame
sweep / ffprobe helper from screenspace_frames.
"""

import bisect
from typing import Any, Callable

import numpy as np

import config
from screenspace_tools import (
    _extract_confidence,
    check_frame_for_tool,
    score_frame_for_tool,
)
from screenspace_frames import _probe_video_meta, scan_video_full_frames


def _multitool_has_offset(steps: list[dict[str, Any]]) -> bool:
    """True when any chained step (idx > 0) declares an ``offset`` window.

    Offset chains run the two-phase scan path and cannot resume incrementally
    (the join needs every frame from the original start), so the worker uses
    this to decide whether a paused task must restart from scratch.
    """
    return any(isinstance(s.get("offset"), dict) for s in steps[1:])


def scan_multitool(
    video_path: str,
    region: dict[str, int],
    steps: list[dict[str, Any]],
    *,
    start_seconds: float = 0.0,
    end_seconds: float | None = None,
    on_progress: Callable[[float], None] | None = None,
    cancel_flag: Callable[[], bool] | None = None,
    on_result: Callable[[dict[str, Any]], None] | None = None,
    fast_opts: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Run a multi-factor scan chaining several tool types.

    Iterates the video once, checking all steps per frame.  A frame
    must pass every step (in order) to be included in the results.
    Results are emitted incrementally via *on_result* as each passing
    frame is found.

    Each entry in *steps* is a dict with ``"type"`` plus the tool's
    own parameters (e.g. ``target_color``, ``tolerance`` for color).

    When any step (idx > 0) carries an ``"offset"`` window
    (``{"min", "max"}`` seconds relative to the previous step's matched
    frame), the scan switches to a two-phase *collect-then-join* path:
    phase 1 decodes the video once and records every step's pass/fail per
    frame; phase 2 joins them temporally (see
    :func:`_join_multitool_offsets`). Offset events are anchored on the
    first step's frame and emitted after the join (not per-frame). When no
    step has an offset the original single-pass short-circuit path runs,
    behaving exactly as before.

    Returns a list of ``{timestamp, tool_types, steps, min_confidence}``
    dicts.
    """
    if len(steps) < 2:
        raise ValueError("Multitool requires at least 2 steps")

    interval = steps[0].get("interval", config.SCREENSPACE_DEFAULT_INTERVAL)

    # Compute iteration range (intersection of all per-step ranges)
    scan_start = start_seconds
    scan_end = end_seconds
    for step in steps:
        s_start = step.get("start_seconds")
        if s_start is not None:
            scan_start = max(scan_start, s_start)
        s_end = step.get("end_seconds")
        if s_end is not None:
            scan_end = min(scan_end, s_end) if scan_end is not None else s_end

    vid_fps, vid_duration = _probe_video_meta(video_path)
    if vid_fps <= 0:
        return []

    if scan_end is None or scan_end > vid_duration:
        scan_end = vid_duration
    total_range = scan_end - scan_start

    tool_types = [s["type"] for s in steps]
    step_regions = [s.get("region_coords", region) for s in steps]
    prev_frame: list[np.ndarray | None] = [None]
    results: list[dict[str, Any]] = []

    def _cancel() -> bool:
        return bool(cancel_flag and cancel_flag())

    if _multitool_has_offset(steps):
        # ---- Offset path: two-phase collect-then-join --------------------
        # Phase 1: one decode pass, evaluate EVERY step on EVERY frame (no
        # short-circuit), recording pass/fail + detail per sampled timestamp.
        # ``prev_frame`` is maintained exactly as the fast path so temporal
        # tools (change/flow/inactivity) still see the prior frame.
        ts_list: list[float] = []
        passed_cols: list[list[bool]] = [[] for _ in steps]
        detail_cols: list[list[dict[str, Any] | None]] = [[] for _ in steps]

        def _collect(ts: float, frame: np.ndarray) -> bool | None:
            if _cancel():
                return False
            ts_list.append(round(ts, 2))
            for i, step in enumerate(steps):
                passed, rd = check_frame_for_tool(
                    frame, prev_frame[0], step_regions[i], step["type"], step
                )
                passed_cols[i].append(bool(passed))
                detail_cols[i].append(rd)
            prev_frame[0] = frame
            if on_progress and total_range > 0:
                # Reserve the last 10% for the join (cheap, but keeps the bar
                # from sitting at 100% while the join runs).
                on_progress(0.9 * (ts - scan_start) / total_range)
            return None

        scan_video_full_frames(
            video_path,
            interval,
            _collect,
            start_seconds=scan_start,
            end_seconds=scan_end,
            fps=vid_fps,
            duration=vid_duration,
            fast_opts=fast_opts,
        )

        # A user cancel/pause stops the decode early; skip the join+emit and
        # let the worker settle the cancelled/paused status. ``detect_first``
        # does NOT trip _cancel here — it only fires on the first on_result.
        # The join needs every frame from ``scan_start``, so this path cannot
        # resume incrementally; ``ScreenspaceWorker.resume`` restarts offset
        # tasks from scratch (it does not advance ``start_seconds`` for them).
        if _cancel():
            return []

        # Phase 2: temporal join.
        joined = _join_multitool_offsets(
            steps, tool_types, ts_list, passed_cols, detail_cols, interval
        )
        emitted: list[dict[str, Any]] = []
        for rd in joined:
            emitted.append(rd)
            if on_result:
                on_result(rd)
            if _cancel():  # detect_first stops after the first hit
                break
        if on_progress:
            on_progress(1.0)
        return emitted

    def _cb(ts: float, frame: np.ndarray) -> bool | None:
        if _cancel():
            return False

        step_results: list[dict[str, Any]] = []
        chain_ok = True
        for i, step in enumerate(steps):
            passed, rd = check_frame_for_tool(
                frame, prev_frame[0], step_regions[i], step["type"], step
            )
            logic = (step.get("logic") or "AND").upper() if i > 0 else "AND"
            if logic == "NOT":
                if passed:
                    chain_ok = False
                    break
                step_results.append({"negated": True, "type": step["type"]})
            else:
                if not passed or rd is None:
                    chain_ok = False
                    break
                step_results.append(rd)

        prev_frame[0] = frame

        if chain_ok and len(step_results) == len(steps):
            confidences = []
            for i, sr in enumerate(step_results):
                logic = (steps[i].get("logic") or "AND").upper() if i > 0 else "AND"
                if logic == "NOT":
                    continue
                confidences.append(_extract_confidence(steps[i]["type"], sr))
            rd = {
                "timestamp": round(ts, 2),
                "tool_types": tool_types,
                "steps": step_results,
                "min_confidence": round(min(confidences), 4) if confidences else 1.0,
            }
            results.append(rd)
            if on_result:
                on_result(rd)

        if on_progress and total_range > 0:
            on_progress((ts - scan_start) / total_range)
        return None

    scan_video_full_frames(
        video_path,
        interval,
        _cb,
        start_seconds=scan_start,
        end_seconds=scan_end,
        fps=vid_fps,
        duration=vid_duration,
        fast_opts=fast_opts,
    )

    if on_progress:
        on_progress(1.0)
    return results


def _join_multitool_offsets(
    steps: list[dict[str, Any]],
    tool_types: list[str],
    ts_list: list[float],
    passed_cols: list[list[bool]],
    detail_cols: list[list[dict[str, Any] | None]],
    interval: float,
) -> list[dict[str, Any]]:
    """Join per-step frame matches into offset-aware multitool events.

    Phase 2 of :func:`scan_multitool`'s offset path. ``ts_list`` is the sorted
    list of sampled timestamps; ``passed_cols[i]`` / ``detail_cols[i]`` are the
    parallel per-step pass/fail and detail columns from phase 1.

    Semantics (all confirmed with the product owner):

    * **Anchor** — every step-0 match seeds a candidate chain; the emitted
      event's ``timestamp`` is the step-0 (trigger) frame.
    * **Cumulative** — each step ``i`` is evaluated relative to ``ref_ts``, the
      *previous* step's matched frame. An AND match advances ``ref_ts`` to its
      earliest in-window frame; NOT and same-frame steps leave it unchanged.
      The advance is greedy (earliest match wins) with no backtracking — for a
      3+ step chain a later in-window match that would have let a downstream
      step resolve is not reconsidered.
    * **Offset window** — ``[ref_ts + min, ref_ts + max]`` (either bound may be
      negative). Absent offset ⇒ the exact ``ref_ts`` frame (legacy behavior).
    * **NOT** — passes iff the condition matches in *no* frame of the window.

    Adjacent anchors that resolve are coalesced (see
    :func:`_coalesce_offset_events`).
    """
    n = len(steps)
    eps = 1e-6
    ts_to_idx = {ts: idx for idx, ts in enumerate(ts_list)}
    logics = [
        (steps[i].get("logic") or "AND").upper() if i > 0 else "AND" for i in range(n)
    ]
    offsets: list[dict[str, Any] | None] = [
        steps[i].get("offset") if i > 0 else None for i in range(n)
    ]

    raw_events: list[dict[str, Any]] = []
    for a, anchor_ts in enumerate(ts_list):
        # Step 0 is always an AND anchor: needs a pass with a usable detail.
        anchor_detail = detail_cols[0][a]
        if not passed_cols[0][a] or anchor_detail is None:
            continue
        ref_ts = anchor_ts
        chain_details: list[dict[str, Any]] = [anchor_detail]
        chain_ok = True
        for i in range(1, n):
            off = offsets[i]
            logic = logics[i]
            if off is None:
                # Same-frame as the previous matched frame.
                idx = ts_to_idx.get(ref_ts)
                hit = idx is not None and passed_cols[i][idx]
                if logic == "NOT":
                    if hit:
                        chain_ok = False
                        break
                    chain_details.append({"negated": True, "type": steps[i]["type"]})
                else:
                    detail = detail_cols[i][idx] if idx is not None else None
                    if not hit or detail is None:
                        chain_ok = False
                        break
                    chain_details.append(detail)
                    # ref_ts unchanged (same frame).
            else:
                lo = ref_ts + off["min"]
                hi = ref_ts + off["max"]
                lo_idx = bisect.bisect_left(ts_list, lo - eps)
                hi_idx = bisect.bisect_right(ts_list, hi + eps)
                if logic == "NOT":
                    if any(passed_cols[i][j] for j in range(lo_idx, hi_idx)):
                        chain_ok = False
                        break
                    chain_details.append({"negated": True, "type": steps[i]["type"]})
                    # ref_ts unchanged (NOT does not advance the cursor).
                else:
                    match_detail: dict[str, Any] | None = None
                    match_ts = ref_ts
                    for j in range(lo_idx, hi_idx):
                        if passed_cols[i][j] and detail_cols[i][j] is not None:
                            match_detail = detail_cols[i][j]
                            match_ts = ts_list[j]
                            break
                    if match_detail is None:
                        chain_ok = False
                        break
                    chain_details.append(match_detail)
                    ref_ts = match_ts  # cumulative advance to earliest match.
        if not chain_ok:
            continue
        confidences = [
            _extract_confidence(steps[i]["type"], chain_details[i])
            for i in range(n)
            if logics[i] != "NOT"
        ]
        raw_events.append(
            {
                "timestamp": round(anchor_ts, 2),
                "tool_types": tool_types,
                "steps": chain_details,
                "min_confidence": round(min(confidences), 4) if confidences else 1.0,
            }
        )

    return _coalesce_offset_events(raw_events, interval)


def _coalesce_offset_events(
    events: list[dict[str, Any]], interval: float
) -> list[dict[str, Any]]:
    """Collapse runs of adjacent anchor events into one representative each.

    ``events`` arrive in ascending-timestamp order (anchors iterate the frame
    grid in order). Consecutive events within ``interval * 1.5`` of one another
    form a run; the highest-``min_confidence`` member represents the run. This
    mirrors :func:`_merge_timestamp_spans`' adjacency rule so an offset chain
    that fires across a burst of neighbouring frames yields a single event,
    matching how the per-frame path's neighbours collapse downstream.
    """
    if not events:
        return []
    gap = interval * 1.5
    coalesced: list[dict[str, Any]] = []
    group = [events[0]]
    for ev in events[1:]:
        if ev["timestamp"] - group[-1]["timestamp"] <= gap:
            group.append(ev)
        else:
            coalesced.append(max(group, key=lambda e: e["min_confidence"]))
            group = [ev]
    coalesced.append(max(group, key=lambda e: e["min_confidence"]))
    return coalesced


def score_multitool_frame(
    frame: np.ndarray,
    prev_frame: np.ndarray | None,
    steps: list[dict[str, Any]],
    *,
    ocr_reader: "Callable[[str, dict[str, int], dict[str, Any]], list[Any]] | None" = None,
) -> dict[str, Any]:
    """Score every multitool step against one frame for pin calibration.

    Unlike :func:`scan_multitool` (which short-circuits the AND chain), this
    evaluates *all* steps unconditionally so every step row gets a score. Each
    step uses its own ``region_coords`` (populated by the server before calling)
    and the step dict as parameters.

    Returns ``{"steps": [...], "passed": bool | None}``. The chain passes when
    all AND steps pass and all NOT steps fail; it is ``False`` when any step
    definitively fails; ``None`` (not evaluable) when no step definitively fails
    but at least one could not be scored.

    Calibration is single-frame, so a step's ``offset`` window is **not**
    evaluated here — each step is scored against this one pinned frame. Offset
    chains are only resolved temporally at scan time (:func:`scan_multitool`).
    """
    step_results: list[dict[str, Any]] = []
    definitive_fail = False
    any_not_evaluable = False
    for i, step in enumerate(steps):
        region = step.get("region_coords", {})
        res = score_frame_for_tool(
            step["type"], frame, prev_frame, region, step, ocr_reader=ocr_reader
        )
        logic = (step.get("logic") or "AND").upper() if i > 0 else "AND"
        entry = dict(res)
        entry["type"] = step["type"]
        entry["logic"] = logic
        step_results.append(entry)
        if res.get("status") != "ok":
            any_not_evaluable = True
            continue
        step_passed = bool(res.get("passed"))
        if logic == "NOT":
            if step_passed:
                definitive_fail = True
        elif not step_passed:
            definitive_fail = True
    if definitive_fail:
        chain_passed: bool | None = False
    elif any_not_evaluable:
        chain_passed = None
    else:
        chain_passed = True
    return {"steps": step_results, "passed": chain_passed}
