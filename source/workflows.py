"""Workflows node-graph engine: executors, import-time wiring, facade.

A free-form 2D node canvas where users wire typed outputs into typed inputs to
chain capabilities across domains (artifact generation, Screenspace analysis,
transcription + thinking agents); ``WorkflowRunner`` executes the resulting DAG.
See ``plans/archive/WORKFLOWS-PLAN.md``.

Owned here: the per-node ``execute`` callables, each a thin adapter over an
existing pure function and keyed by output-port name. Every domain value a node
emits embeds a "source descriptor", which is what keeps ``ADAPTERS`` pure.

The declarative catalog (``NODE_TYPES``, ``BUILTIN_STASHES``, ``ADAPTERS``) lives
in ``workflows_catalog``, the run engine in ``workflows_runner``; this facade
re-exports both, so ``workflows.NAME`` still resolves every public name and the
private ones tests reach for. The wiring below mutates the *shared* ``NODE_TYPES``
dict rather than a copy — tests patch node executors through
``workflows.NODE_TYPES`` and depend on that identity. Re-binding a name here only
rebinds it on the facade; patch the owning sibling to stub a seam.

Manifest shape (the ``workflows`` section of the output-dir manifest)::

    {
        "blueprints": [ {id, name, nodes, edges, viewport, trigger} ],
        "stashes":    [ {id, name, nodes, edges, createdAt} ],
        "runs":       [ {id, blueprintId, status, nodeStates, startedAt, completedAt} ]
    }

``trigger`` is the auto-launch binding: ``null`` (or ``{"type": <t>, "enabled":
false}``) when disarmed, else ``{"type": <t>, "enabled": true}`` for one of
``TRIGGER_TYPES``. At most one blueprint is armed *per trigger type*, so a
new-video pipeline can chain into a transcript-complete one; the watcher daemon
in ``workflows_server`` fires the match once per arriving participant, bound via
``bind_participant``. No feedback loop is possible: the ``transcribe`` node calls
``transcripts.transcribe_video`` directly and never writes the transcripts
manifest, so a transcript-complete graph containing a Transcribe node cannot
re-fire itself.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

import config
import utils

# Catalog names the executors + wiring below use directly (also part of the
# ``workflows.NAME`` facade surface, like everything imported here).
from workflows_catalog import (
    NODE_TYPES,
    NodeContext,
    ParamSpec,
    _SS_DETECTOR_SPECS,
    _SS_REFERENCE_DETECTORS,
    _clip_source_filename,
    _source_descriptor,
)

# Facade re-exports — moved names that workflows_server and the test suite keep
# reaching as ``workflows.NAME``.
from workflows_catalog import (  # noqa: F401
    ADAPTERS,
    BUILTIN_STASHES,
    NodeType,
    Port,
    serialize_adapters,
    serialize_catalog,
)
from workflows_runner import TRIGGER_TYPE_IDS  # used by load_workflows_manifest
from workflows_runner import (  # noqa: F401
    NODE_STATUS_COMPLETED,
    NODE_STATUS_DEGRADED,
    NODE_STATUS_FAILED,
    NODE_STATUS_QUEUED,
    NODE_STATUS_RUNNING,
    NODE_STATUS_SKIPPED,
    NOTE_NODE_TYPE,
    RUN_STATUS_CANCELLED,
    RUN_STATUS_COMPLETED,
    RUN_STATUS_DEGRADED,
    RUN_STATUS_FAILED,
    RUN_STATUS_QUEUED,
    RUN_STATUS_RUNNING,
    TRIGGER_TYPES,
    WorkflowCycleError,
    WorkflowRunner,
    _inspectable_result,
    bind_participant,
    blueprint_participant_nodes,
    compute_resume_plan,
    inspectable_sidecar_view,
    run_results_dir,
    topo_order,
    write_node_sidecar,
)


def empty_workflows_manifest() -> dict[str, list[Any]]:
    """Return a fresh, empty workflows manifest with all top-level keys present."""
    return {"blueprints": [], "stashes": [], "runs": []}


def load_workflows_manifest() -> dict[str, Any]:
    """Load the ``workflows`` manifest section (empty default).

    Missing or corrupt files fall back to :func:`empty_workflows_manifest` so
    callers always get the full key set, never a partial dict.
    """
    data = utils.load_manifest_section("workflows")
    if not isinstance(data, dict):
        return empty_workflows_manifest()
    # Backfill any missing top-level keys so callers can index unconditionally.
    base = empty_workflows_manifest()
    base.update({k: v for k, v in data.items() if k in base})
    # An unknown trigger type reads as disarmed: the watcher only fires known
    # types, so keeping it enabled would render an armed toolbar that never fires.
    for blueprint in base.get("blueprints", []):
        if not isinstance(blueprint, dict):
            continue
        trigger = blueprint.get("trigger")
        if isinstance(trigger, dict) and trigger.get("type") not in TRIGGER_TYPE_IDS:
            blueprint["trigger"] = None
    return base


def _is_empty_workflows_manifest(payload: dict[str, Any]) -> bool:
    """True when nothing worth persisting exists.

    A blueprint counts only if it carries graph content (nodes or edges); a bare
    auto-created "Untitled" with empty nodes/edges is treated as empty even when
    renamed, so a zero-interaction Workflows launch writes no file. Any stash or
    run is user-meaningful and keeps the manifest.
    """
    if payload.get("stashes") or payload.get("runs"):
        return False
    for blueprint in payload.get("blueprints", []):
        if blueprint.get("nodes") or blueprint.get("edges"):
            return False
    return True


def save_workflows_manifest(
    blueprints: list[dict[str, Any]],
    stashes: list[dict[str, Any]] | None = None,
    runs: list[dict[str, Any]] | None = None,
) -> Path | None:
    """Persist the workflows manifest atomically; returns the path or ``None``.

    Skips the write (and removes any stale file) when the manifest is empty, so
    an unused canvas leaves no junk in the output dir.
    """
    payload = {
        "blueprints": blueprints,
        "stashes": stashes or [],
        "runs": runs or [],
    }
    if _is_empty_workflows_manifest(payload):
        utils.save_manifest_section("workflows", None)
        return None
    return utils.save_manifest_section("workflows", payload)


# ---------------------------------------------------------------------------
# Executors — thin adapters over existing pure functions
# ---------------------------------------------------------------------------
#
# Uniform shape: ``execute(ctx, inputs, params) -> {port: value}``, keyed by
# OUTPUT-port name. Backend modules import lazily inside each executor (mirroring
# ``cli._run_ss_clips``) to avoid import cost and cycles — Workflows sits at the
# top of the dependency DAG. Per-wire value shapes are documented in
# ``plans/archive/WORKFLOWS-PLAN.md``.


# ---- Sources ----


def _exec_video_source(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    # The canvas stores a multi-selection as a list (or the "__all__" sentinel),
    # which the run panel routes to the batch endpoint where bind_participant
    # rewrites it to a scalar. A direct run must not stringify the list into a
    # participant id that resolves zero videos and "completes" empty.
    raw = params.get("participant", "")
    if isinstance(raw, list):
        if len(raw) > 1:
            raise RuntimeError(
                "Video Source has several participants selected — use Run to fan out as a batch"
            )
        raw = raw[0] if raw else ""
    participant = str(raw or "")
    if participant == "__all__":
        raise RuntimeError(
            "Video Source is set to all participants — use Run to fan out as a batch"
        )
    video_paths = ctx.resolve_videos(participant) if participant else []
    return {
        "video": _source_descriptor(participant, video_paths),
        "participant": participant,
    }


def _exec_sheet_selection(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import spreadsheet

    selector = str(params.get("selector", "") or "").strip()
    if ctx.sheet_context is None or not selector:
        return {"clips": {"records": [], "study": ""}}
    records = spreadsheet.generate_list(
        ctx.worksheet,
        "reel",
        ctx=ctx.sheet_context,
        reel_input=selector,
        skip_prompts=True,
    )
    study = str(getattr(ctx.sheet_context, "study_name", "") or "")
    return {"clips": {"records": records, "study": study}}


def _exec_region(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import screenspace

    name = str(params.get("name", "") or "").strip()
    coords: dict[str, Any] | None = None
    if name:
        regions = screenspace.load_screenspace_manifest().get("regions") or {}
        entry = regions.get(name)
        if isinstance(entry, dict):
            coords = {k: entry[k] for k in ("x", "y", "w", "h") if k in entry}
    result: dict[str, Any] = {"region": {"name": name, "coords": coords}}
    if name and coords is None:
        # A named-but-missing region degrades to the full frame downstream
        # (_resolve_region_coords); say so instead of silently scanning it all.
        result["__note__"] = f'Region "{name}" not found — scanning the full frame'
    return result


def _exec_time_range(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    """Manual in/out times → timeRange (scan windows for SS nodes, cuts for clips).

    Parses the same ``MM:SS``/``HH:MM:SS`` (range or single) syntax as a sheet
    cell. Source is left empty — SS detectors take their video from the ``video``
    input, and ``make_clips`` falls back to its wired ``video`` for the source.
    """
    raw = str(params.get("ranges", "") or "").strip()
    ranges: list[tuple[float, float]] = []
    for start_str, end_str in utils.parse_timestamps(raw) if raw else []:
        start = utils.timestamp_to_seconds(start_str)
        end = utils.timestamp_to_seconds(end_str)
        if start is not None and end is not None:
            ranges.append((start, max(start, end)))
    return {"timeRange": {"ranges": ranges, "source": {}}}


# ---- Transcript ----


def _exec_transcribe(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import transcripts
    import video

    src = inputs.get("video") or {}
    paths = list(src.get("video_paths") or [])
    language = str(params.get("language", "") or "").strip()
    lang = None if language in ("", "auto") else language
    model_name = str(params.get("model", "") or "").strip() or None
    # No audio-track param on purpose: omitting audio_index gets speech-track
    # auto-detection, whereas a pinned index looks portable and isn't — "track 1"
    # is the mic for P01 and system audio for P07, whose recorder ordered the
    # streams differently. Any future control here should be a name hint
    # ("audio_track_contains"), not an index.

    result: Any = None
    if not paths:
        return {
            "transcript": {
                "segments": [],
                "language": lang or "",
                "source_file": "",
                "model": "",
                "source": src,
            },
            "segments": {"segments": [], "source": src},
            "__note__": "No video wired",
        }
    if len(paths) >= 2:
        timeline = video.build_source_timeline(paths)
        if timeline is not None:
            result = transcripts.transcribe_timeline(
                timeline,
                model_name=model_name,
                language=lang,
                cancel_flag=ctx.cancel_flag,
            )
    else:
        result = transcripts.transcribe_video(
            paths[0],
            model_name=model_name,
            language=lang,
            cancel_flag=ctx.cancel_flag,
        )

    if result is None:
        # Decode/model failure used to become a successful empty transcript, so
        # Find-word and summaries ran on silence with no error on the node.
        raise RuntimeError("Could not transcribe the wired video")
    transcript_val = dict(result)
    transcript_val["source"] = src
    return {
        "transcript": transcript_val,
        "segments": {"segments": result.get("segments", []), "source": src},
    }


def _exec_find_word(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    seg_val = inputs.get("segments") or {}
    segments = seg_val.get("segments") or []
    source = seg_val.get("source") or {}
    word = str(params.get("word", "") or "").strip().lower()
    pad = float(params.get("pad", 0) or 0)

    ranges: list[tuple[float, float]] = []
    times: list[float] = []
    if word:
        for seg in segments:
            if word in str(seg.get("text", "")).lower():
                start = float(seg.get("start", 0.0) or 0.0)
                end = float(seg.get("end", start) or start)
                ranges.append((max(0.0, start - pad), end + pad))
                times.append(start)
    return {
        "timeRange": {"ranges": ranges, "source": source},
        "timestamps": {"times": times, "source": source},
    }


def _exec_transcript_marks(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    """Read the participant's Transcripts-page marks as padded time ranges.

    Marks live in the transcripts manifest as ``{segment_id: "pid:index", …}``
    records; each resolves against that participant's persisted segments (the
    workflow ``transcribe`` node deliberately never writes that manifest, so
    this reads what the Transcripts page produced).
    """
    import transcripts

    src = inputs.get("video") or {}
    participant = str(src.get("participant", "") or "")
    empty = {
        "timeRange": {"ranges": [], "source": src},
        "timestamps": {"times": [], "source": src},
    }
    if not participant:
        return {**empty, "__note__": "No video wired"}
    manifest = transcripts.load_transcripts_manifest()
    entry = (manifest.get("source_transcripts") or {}).get(participant) or {}
    segments = list(entry.get("segments") or [])
    category = str(params.get("category", "") or "").strip().lower()
    pad = max(0.0, float(params.get("pad", 2) or 0))

    spans: list[tuple[float, float, float]] = []  # (start, padded lo, padded hi)
    for mark in manifest.get("marks") or []:
        if not isinstance(mark, dict):
            continue
        pid, sep, idx_str = str(mark.get("segment_id", "") or "").partition(":")
        if not sep or pid != participant:
            continue
        if category and str(mark.get("category", "") or "").lower() != category:
            continue
        try:
            idx = int(idx_str)
        except ValueError:
            continue
        if not 0 <= idx < len(segments):
            continue
        seg = segments[idx]
        start = float(seg.get("start", 0.0) or 0.0)
        end = max(start, float(seg.get("end", 0.0) or 0.0))
        spans.append((start, max(0.0, start - pad), end + pad))
    spans.sort()
    if not spans:
        scope = f' in category "{category}"' if category else ""
        return {**empty, "__note__": f"No marks for this participant{scope}"}
    return {
        "timeRange": {"ranges": [(lo, hi) for _, lo, hi in spans], "source": src},
        "timestamps": {"times": [start for start, _, _ in spans], "source": src},
    }


def _exec_transcript_export(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import transcripts

    transcript_in = inputs.get("transcript") or {}
    seg_in = inputs.get("segments") or {}
    src = transcript_in.get("source") or seg_in.get("source") or {}
    study = str(src.get("study", "") or "")
    fmt = str(params.get("format", "") or "") or config.TRANSCRIBE_FORMAT
    if fmt not in ("md", "srt", "vtt"):
        fmt = "md"

    # Prefer the full transcript (carries language/model); a bare segments wire
    # still exports, just with empty metadata.
    if transcript_in.get("segments"):
        base: dict[str, Any] = transcript_in
    elif seg_in.get("segments"):
        base = {
            "segments": seg_in.get("segments"),
            "language": "",
            "model": "",
            "source_file": str(src.get("source_filename", "") or ""),
        }
    else:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "No transcript or segments wired",
        }
    result = cast(
        transcripts.TranscriptResult,
        {
            "segments": list(base.get("segments") or []),
            "language": str(base.get("language", "") or ""),
            "model": str(base.get("model", "") or ""),
            "source_file": str(base.get("source_file", "") or ""),
        },
    )

    participant = str(src.get("participant", "") or "")
    stem = f"transcript_{participant}" if participant else "transcript"
    ext = transcripts.get_transcript_extension(fmt)
    output_path = files.get_unique_filename(f"{stem}{ext}", file_format=ext)
    if not transcripts.write_transcript(result, output_path, fmt=fmt):
        files.release_reservation(output_path)
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "Transcript couldn't be written",
        }
    # "export" (not "transcript") — the viewer routes it to the Attachments
    # pane's document card; "transcript" is a timeline card type there.
    rec = _attachment_artifact("export", output_path, src, f"Transcript ({fmt})")
    return {"artifacts": {"artifacts": [rec], "study": study, "count": 1}}


# ---- Thinking (local LLM) ----


def _exec_summarize(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import llm_client
    import thinking_agents

    transcript = inputs.get("transcript") or {}
    segments = transcript.get("segments") or []
    if not llm_client.is_available():
        return {"summary": "", "__note__": "AI server not available. Summary skipped"}
    summary = thinking_agents.summarize_transcript(
        segments, model=params.get("model") or None, cancel_event=ctx.cancel_event
    )
    return {"summary": summary or ""}


def _exec_citations(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import llm_client
    import thinking_agents

    summary = str(inputs.get("summary") or "")
    seg_val = inputs.get("segments") or {}
    segments = seg_val.get("segments") or []
    if not llm_client.is_available():
        return {
            "citations": [],
            "__note__": "AI server not available. Citations skipped",
        }
    cites = thinking_agents.find_citations(
        summary,
        segments,
        model=params.get("model") or None,
        cancel_event=ctx.cancel_event,
    )
    return {"citations": cites or []}


def _exec_friction(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import friction
    import llm_client
    import thinking_agents

    seg_val = inputs.get("segments") or {}
    segments = seg_val.get("segments") or []
    summary = str(inputs.get("summary") or "")
    if not llm_client.is_available():
        return {"friction": [], "__note__": "AI server not available. Friction skipped"}
    scored = friction.score_segments(segments)
    candidates = friction.select_candidates(scored, config.FRICTION_CANDIDATE_LIMIT)
    moments = thinking_agents.find_friction_moments(
        summary,
        segments,
        candidates,
        model=params.get("model") or None,
        cancel_event=ctx.cancel_event,
    )
    if moments is None:
        # "Didn't run" (model failure / unrenderable candidates) must not read
        # as "ran and found nothing".
        return {"friction": [], "__note__": "Friction analysis failed. See log"}
    return {"friction": moments}


def _exec_report(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import llm_client
    import thinking_agents

    summary = str(inputs.get("summary") or "")
    src = inputs.get("video") or {}
    participant = str(src.get("participant", "") or "")
    if not summary:
        return {"report": "", "__note__": "No summary wired"}
    if not llm_client.is_available():
        return {"report": "", "__note__": "AI server not available. Report skipped"}
    # Sheet observations + transcript marks come through the same configured
    # seam the Overview Reports tab uses; unwired (no sheet / CLI) both are
    # empty and the report covers the summary alone.
    observation_lines, mark_lines = thinking_agents.report_source_lines(participant)
    text = thinking_agents.build_report(
        summary,
        "\n".join(observation_lines),
        "\n".join(mark_lines),
        participant=participant or "unknown",
        model=params.get("model") or None,
        cancel_event=ctx.cancel_event,
    )
    if not text:
        return {"report": "", "__note__": "Report generation failed"}
    out: dict[str, Any] = {"report": text}
    if not participant:
        out["__note__"] = "No video wired — the report covers the summary only"
    return out


# ---- Screenspace ----


def _build_ss_scan_params(tool_name: str, params: dict[str, Any]) -> dict[str, Any]:
    """Assemble a detector's flat node params into the scan's ``parameters`` dict.

    Flat number/enum/bool params from the canvas are reshaped into the nested
    structure each ``screenspace_tools`` scan reads (e.g. color's ``target_color``
    /``tolerance`` HSV dicts). Reference-frame extraction is handled separately by
    the caller (it needs the video path + region).
    """

    def _num(key: str, default: float = 0.0) -> float:
        val = params.get(key)
        return float(val) if val not in (None, "") else float(default)

    if tool_name == "color":
        return {
            "target_color": {
                "h": _num("color_h"),
                "s": _num("color_s"),
                "v": _num("color_v"),
            },
            "tolerance": {
                "h": _num("tol_h", 10),
                "s": _num("tol_s", 50),
                "v": _num("tol_v", 50),
            },
            "color_mode": str(params.get("color_mode", "average") or "average"),
            "min_coverage": _num("min_coverage"),
            "interval": _num("interval"),
        }
    if tool_name == "change":
        return {
            "threshold": _num("threshold"),
            "noise_threshold": _num("noise_threshold"),
            "require_consecutive": int(_num("require_consecutive", 1)),
            "interval": _num("interval"),
        }
    if tool_name == "flow":
        return {
            "magnitude_threshold": _num("magnitude_threshold"),
            "require_consecutive": int(_num("require_consecutive", 1)),
            "interval": _num("interval"),
        }
    if tool_name == "text":
        return {
            "search_string": str(params.get("search_string", "") or ""),
            "fuzzy_threshold": _num("fuzzy_threshold", 80),
            "interval": _num("interval", 2.0),
        }
    if tool_name == "numbers":
        out: dict[str, Any] = {
            "operator": str(params.get("operator", "gt") or "gt"),
            "target_value": _num("target_value"),
            "integers_only": bool(params.get("integers_only", False)),
            "interval": _num("interval", 2.0),
        }
        if params.get("range_min") not in (None, ""):
            out["range_min"] = _num("range_min")
        if params.get("range_max") not in (None, ""):
            out["range_max"] = _num("range_max")
        return out
    if tool_name == "similarity":
        return {"threshold": _num("threshold"), "interval": _num("interval")}
    if tool_name == "scene":
        return {"threshold": _num("threshold"), "interval": _num("interval")}
    if tool_name == "template":
        return {
            "threshold": _num("threshold"),
            "template_scale": _num("template_scale", 1.0),
            "interval": _num("interval"),
        }
    if tool_name == "inactivity":
        return {
            "threshold": _num("threshold"),
            "min_duration": _num("min_duration"),
            "interval": _num("interval"),
        }
    if tool_name == "boundary":
        return {
            "threshold": _num("threshold"),
            "min_gap": _num("min_gap"),
            "metric": str(
                params.get("metric", "") or config.SCREENSPACE_BOUNDARY_METRIC
            ),
            "interval": _num("interval"),
        }
    if tool_name == "attention":
        return {
            "shift_threshold": _num("shift_threshold"),
            "ema_alpha": _num("ema_alpha"),
            "interval": _num("interval"),
        }
    return {"interval": _num("interval")}


def _attach_ss_reference(
    tool_name: str,
    base_params: dict[str, Any],
    params: dict[str, Any],
    video_path: str,
    region_coords: dict[str, int],
) -> bool:
    """Self-extract a reference frame from ``video_path`` at ``reference_seconds``.

    Similarity/scene/template need reference image data that the canvas can't
    upload; instead we crop the node's region from a frame at ``reference_seconds``.
    Returns False when the frame can't be read so the executor can short-circuit.
    """
    import screenspace
    import video as video_mod

    ref_ts = float(params.get("reference_seconds", 0.0) or 0.0)
    frame = video_mod.extract_frame_at_timestamp(video_path, ref_ts)
    if frame is None:
        return False
    crop = screenspace.extract_region(frame, region_coords)
    if tool_name == "similarity":
        base_params["reference_frame"] = crop
    elif tool_name == "template":
        base_params["template_image"] = crop
    elif tool_name == "scene":
        base_params["reference_scenes"] = [{"name": "ref", "frame": crop}]
    return True


def _run_ss_detector(
    ctx: NodeContext,
    inputs: dict[str, Any],
    params: dict[str, Any],
    tool_name: str,
) -> dict[str, Any]:
    import screenspace
    import screenspace_manifest
    import screenspace_worker

    src = inputs.get("video") or {}
    paths = list(src.get("video_paths") or [])
    tool = screenspace.TOOLS.get(tool_name)
    if not paths or tool is None:
        note = "No video wired" if not paths else f"Unknown detector: {tool_name}"
        return {
            "events": {"events": [], "source": src, "raw_results": []},
            "__note__": note,
        }

    # Unwired region scans the whole frame; zero-size coords would make the scan
    # a silent no-op (see _resolve_region_coords).
    region_name, region_coords = _resolve_region_coords(
        inputs.get("region") or {}, paths[0]
    )

    base_params = _build_ss_scan_params(tool_name, params)
    if tool_name in _SS_REFERENCE_DETECTORS and not _attach_ss_reference(
        tool_name, base_params, params, paths[0], region_coords
    ):
        return {
            "events": {"events": [], "source": src, "raw_results": []},
            "__note__": "Couldn't read the reference frame at the given time",
        }

    task = screenspace_manifest.create_task(
        tool_name,
        str(src.get("participant", "") or ""),
        str(src.get("source_filename", "") or ""),
        paths,
        region_name,
        region_coords,
        parameters=base_params,
    )

    # Scan windows: each timeRange span scanned separately, else the whole video.
    windows = list((inputs.get("timeRange") or {}).get("ranges") or [])
    raw_results: list[dict[str, Any]] = []
    scan_targets = windows or [None]
    # Each underlying scan reports progress on its own 0->1 scale (multi-video
    # scans even force on_progress(1.0) at the end), so map each window into a
    # span-weighted slice to keep job-level progress monotonic.
    spans = [
        max(0.0, float(w[1]) - float(w[0])) if w is not None else 1.0
        for w in scan_targets
    ]
    total_span = sum(spans)
    if total_span <= 0:
        spans = [1.0] * len(scan_targets)
        total_span = float(len(scan_targets))
    done_span = 0.0
    for window, win_span in zip(scan_targets, spans):
        if ctx.cancel_flag():
            break
        scan_params = dict(task["parameters"])
        if window is not None:
            scan_params["start_seconds"] = float(window[0])
            scan_params["end_seconds"] = float(window[1])
        frac_start = done_span / total_span
        frac_end = (done_span + win_span) / total_span

        def window_progress(
            p: float, _a: float = frac_start, _b: float = frac_end
        ) -> None:
            ctx.on_progress(_a + p * (_b - _a))

        returned = screenspace_worker.dispatch_tool_scan(
            tool,
            paths,
            region_coords,
            scan_params,
            task_id=task["id"],
            scan_mode="normal",
            on_progress=window_progress,
            cancel_flag=ctx.cancel_flag,
            on_result=None,
            fast_opts=None,
        )
        raw_results.extend(returned or [])
        done_span += win_span

    events = screenspace_manifest.generate_events_from_results(task, raw_results)
    # raw_results rides along for the heatmap node (template/flow/change); every
    # other consumer reads only ``events`` and ignores it.
    return {"events": {"events": events, "source": src, "raw_results": raw_results}}


def _make_ss_executor(
    tool_name: str,
) -> Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]:
    """Bind ``_run_ss_detector`` to one tool so all ten nodes share one body."""

    def _exec(
        ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
    ) -> dict[str, Any]:
        return _run_ss_detector(ctx, inputs, params, tool_name)

    return _exec


def _exec_detect(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    """Unified detector node: dispatch to the chosen tool's scan body."""
    tool_name = str(params.get("detector", "text") or "text")
    return _run_ss_detector(ctx, inputs, params, tool_name)


def _resolve_region_coords(
    region_in: dict[str, Any], video_path: str
) -> tuple[str, dict[str, int]]:
    """Resolve a region input to pixel coords, defaulting to the **full frame**.

    A region port is optional on every Screenspace node; when it is unwired (or a
    Region node names nothing / a coord-less entry), this returns the whole
    frame's pixel dimensions rather than zero-size coords. That matters because
    ``scan_video_frames`` rejects a zero-size region and skips the scan entirely
    (a silent no-op), and the reference detectors would otherwise crop an empty
    reference frame. Probes ``video_path`` once for the frame size.
    """
    import screenspace
    import video

    region_name = str(region_in.get("name", "") or "")
    props = video.probe_video_properties(video_path) or {}
    width = int(props.get("width", 0) or 0)
    height = int(props.get("height", 0) or 0)
    coords: dict[str, int] = {"x": 0, "y": 0, "w": 0, "h": 0}
    norm = region_in.get("coords")
    if isinstance(norm, dict) and norm and width > 0 and height > 0:
        coords = screenspace.denormalize_region(norm, width, height)
    if coords["w"] <= 0 or coords["h"] <= 0:
        # Full-frame fallback: unwired region, or a region that didn't resolve.
        coords = {"x": 0, "y": 0, "w": width, "h": height}
    return region_name, coords


def _exec_multitool(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import screenspace
    import screenspace_manifest
    import screenspace_worker

    src = inputs.get("video") or {}
    paths = list(src.get("video_paths") or [])
    raw_steps = list(params.get("steps") or [])
    if not paths or len(raw_steps) < 2:
        return {"events": {"events": [], "source": src, "raw_results": []}}

    region_name, region_coords = _resolve_region_coords(
        inputs.get("region") or {}, paths[0]
    )

    # Reshape each flat step into the {type, region_coords, logic, …} shape
    # scan_multitool expects, reusing the per-detector param builder.
    steps: list[dict[str, Any]] = []
    for idx, raw in enumerate(raw_steps):
        step_type = str(raw.get("type", "") or "")
        if step_type not in screenspace.TOOLS:
            continue
        step = _build_ss_scan_params(step_type, raw)
        step["type"] = step_type
        step["region_coords"] = region_coords
        if idx > 0:
            step["logic"] = str(raw.get("logic", "AND") or "AND").upper()
        steps.append(step)
    if len(steps) < 2:
        return {"events": {"events": [], "source": src, "raw_results": []}}

    task = screenspace_manifest.create_task(
        "multitool",
        str(src.get("participant", "") or ""),
        str(src.get("source_filename", "") or ""),
        paths,
        region_name,
        region_coords,
        parameters={"steps": steps},
    )
    raw_results = (
        screenspace_worker.dispatch_tool_scan(
            screenspace.TOOLS["multitool"],
            paths,
            region_coords,
            task["parameters"],
            task_id=task["id"],
            scan_mode="normal",
            on_progress=ctx.on_progress,
            cancel_flag=ctx.cancel_flag,
            on_result=None,
            fast_opts=None,
        )
        or []
    )
    events = screenspace_manifest.generate_events_from_results(task, raw_results)
    return {"events": {"events": events, "source": src, "raw_results": raw_results}}


# ---- Artifact ----


def _exec_highlights(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import spreadsheet

    clips_in = inputs.get("clips") or {}
    records = list(clips_in.get("records") or [])
    study = str(clips_in.get("study", "") or "")
    if not records:
        return {"clips": {"records": [], "study": study}}
    budget = int(params.get("budget", config.HIGHLIGHTS_REEL_DURATION_SECONDS) or 0)
    if budget <= 0:
        budget = config.HIGHLIGHTS_REEL_DURATION_SECONDS
    # Uniqueness is scored against artifacts already in the output dir (mirrors the
    # ``-H`` CLI path: spreadsheet.generate_reel_timestamps -> score_and_truncate_clips).
    existing_filenames = set(files.discover_clips())
    selected = spreadsheet.score_and_truncate_clips(records, existing_filenames, budget)
    return {"clips": {"records": selected, "study": study}}


def _artifact_padding_params(params: dict[str, Any]) -> tuple[float, float, float]:
    """Read the pad-start/pad-end/max-duration node params for the pipeline.

    Returns ``(pad_pre, pad_post, max_duration)``; all default to a no-op (0.0).
    Pads are signed (negative = trim inward). Shared by the Make Clips and Build
    Reel executors.
    """

    def _num(key: str) -> float:
        try:
            return float(params.get(key) or 0)
        except (TypeError, ValueError):
            return 0.0

    return _num("pad_start"), _num("pad_end"), _num("max_duration")


def _exec_make_clips(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import pipeline

    description = str(params.get("description", "") or "").strip() or "workflow"
    output_format = str(params.get("output_format", "clip") or "clip")
    if output_format not in ("clip", "screen", "gif"):
        output_format = "clip"
    titlecards = bool(params.get("titlecards", False))
    raw_card_dur = params.get("titlecard_duration")
    titlecard_duration = int(raw_card_dur) if raw_card_dur not in (None, "") else None

    records: list[Any] = []
    study = ""
    clips_in = inputs.get("clips")
    if isinstance(clips_in, dict) and clips_in.get("records"):
        records = list(clips_in["records"])
        study = str(clips_in.get("study", "") or "")
    else:
        tr = inputs.get("timeRange")
        if isinstance(tr, dict) and tr.get("ranges"):
            source = tr.get("source") or inputs.get("video") or {}
            study = str(source.get("study", "") or "")
            records = files.build_clip_records(
                participant=str(source.get("participant", "") or ""),
                source_filename=_clip_source_filename(source),
                time_ranges=[(float(s), float(e)) for s, e in tr["ranges"]],
                description=description,
                study=study,
            )

    if not records:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "No clips to render. Wire clips, a time range, or a video",
        }
    pad_pre, pad_post, max_duration = _artifact_padding_params(params)
    count, artifacts = pipeline.process_clips(
        records,
        output_format=output_format,
        include_severity=False,
        cancel_flag=ctx.cancel_flag,
        titlecards_enabled=titlecards,
        titlecard_duration_seconds=titlecard_duration,
        pad_pre=pad_pre,
        pad_post=pad_post,
        max_duration=max_duration,
    )
    return {"artifacts": {"artifacts": artifacts, "study": study, "count": count}}


def _exec_interval_captures(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    """Sample a video into screenshots/GIFs at a fixed interval.

    Iterates each wired time range (or the whole video when none is wired) at
    ``interval`` seconds, expands the samples into point/GIF clip records, and
    reuses ``process_clips`` so the artifacts match Make Clips exactly. Fixes the
    one-artifact-per-range limit of Make Clips' screen/gif output.
    """
    import files
    import pipeline
    import video as video_mod

    src = inputs.get("video") or {}
    paths = list(src.get("video_paths") or [])
    study = str(src.get("study", "") or "")
    empty = {"artifacts": {"artifacts": [], "study": study, "count": 0}}
    if not paths:
        return {**empty, "__note__": "No video wired"}

    interval = float(
        params.get("interval", config.GALLERY_INTERVAL_SECONDS)
        or config.GALLERY_INTERVAL_SECONDS
    )
    interval = max(interval, 0.2)
    fmt = (
        "gif" if str(params.get("output_format", "screen") or "") == "gif" else "screen"
    )
    gif_dur = float(
        params.get("gif_duration", config.GALLERY_GIF_DURATION_SECONDS)
        or config.GALLERY_GIF_DURATION_SECONDS
    )

    ranges = [
        (float(s), float(e))
        for s, e in ((inputs.get("timeRange") or {}).get("ranges") or [])
    ]
    if not ranges:
        duration = video_mod.get_file_duration(paths[0]) or 0
        if duration <= 0:
            return {**empty, "__note__": "Couldn't read the video duration"}
        ranges = [(0.0, float(duration))]

    # Expand each window into per-interval sample points (a point for a
    # screenshot, a [t, t+gif_dur] window for a GIF).
    sample_ranges: list[tuple[float, float]] = []
    for start, end in ranges:
        t = start
        while t < end:
            sample_ranges.append((t, t + gif_dur if fmt == "gif" else t))
            t += interval
    if not sample_ranges:
        return {**empty, "__note__": "No sample points in the given interval/range"}

    records = files.build_clip_records(
        participant=str(src.get("participant", "") or ""),
        source_filename=_clip_source_filename(src),
        time_ranges=sample_ranges,
        description="sample",
        study=study,
    )
    count, artifacts = pipeline.process_clips(
        records,
        output_format=fmt,
        include_severity=False,
        cancel_flag=ctx.cancel_flag,
    )
    return {"artifacts": {"artifacts": artifacts, "study": study, "count": count}}


def _attachment_artifact(
    art_type: str, output_path: str, source: dict[str, Any], description: str
) -> dict[str, Any]:
    """Build an artifact record for a single non-timeline output (timelapse/heatmap).

    ``start``/``end`` are 0 so the viewer keeps it out of the timeline track and
    surfaces it in the Attachments panel instead (branch on ``type``).
    """
    name = Path(output_path).name
    return {
        "id": f"{art_type}-{name}",
        "type": art_type,
        "file": name,
        "thumbnail": "",
        "start": 0,
        "end": 0,
        "study": str(source.get("study", "") or ""),
        "participant": str(source.get("participant", "") or ""),
        "category": "",
        "severity": "",
        "description": description,
        "sourceVideo": str(source.get("source_filename", "") or ""),
        "annotations": [],
    }


def _exec_timelapse(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import screenspace_scans

    src = inputs.get("video") or {}
    paths = list(src.get("video_paths") or [])
    study = str(src.get("study", "") or "")
    if not paths:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "No video wired",
        }

    # _resolve_region_coords already falls back to the full frame when no region
    # is wired; a still-zero size means the probe failed (unreadable video).
    _name, region_coords = _resolve_region_coords(inputs.get("region") or {}, paths[0])
    if region_coords["w"] <= 0 or region_coords["h"] <= 0:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "Couldn't read the video",
        }

    out_format = str(params.get("output_format", "mp4") or "mp4")
    if out_format not in ("mp4", "gif"):
        out_format = "mp4"
    output_path = files.get_unique_filename(
        f"timelapse.{out_format}", file_format=f".{out_format}"
    )
    result = screenspace_scans.generate_timelapse(
        paths[0],
        region_coords,
        float(params.get("speedup_factor", 10.0) or 10.0),
        output_path,
        output_format=out_format,
        sample_interval=float(params.get("sample_interval", 0.0) or 0.0),
        on_progress=ctx.on_progress,
        cancel_flag=ctx.cancel_flag,
    )
    if not result:
        files.release_reservation(output_path)
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "Timelapse couldn't be generated",
        }
    rec = _attachment_artifact("timelapse", result, src, "Timelapse")
    return {"artifacts": {"artifacts": [rec], "study": study, "count": 1}}


def _exec_heatmap(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import screenspace_heatmap
    import video

    events_in = inputs.get("events") or {}
    src = events_in.get("source") or {}
    study = str(src.get("study", "") or "")
    results = list(events_in.get("raw_results") or [])
    style = str(params.get("style", "auto") or "auto")
    paths = list(src.get("video_paths") or [])
    if not results or not paths:
        note = (
            "No detector results. Wire a template/flow/change/attention detector"
            if not results
            else "No video for the heatmap"
        )
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": note,
        }
    if style not in ("template", "flow", "change", "attention"):
        # "auto": infer from the detector data actually present. Keying on the
        # per-frame payload (match boxes / grids) rather than the producing
        # node's type keeps the inference correct through merges and edits.
        style = _infer_heatmap_style(results)
        if not style:
            return {
                "artifacts": {"artifacts": [], "study": study, "count": 0},
                "__note__": (
                    "The wired events carry no heatmap data — use a "
                    "template/flow/change/attention detector upstream"
                ),
            }

    props = video.probe_video_properties(paths[0]) or {}
    width = int(props.get("width", 0) or 0) or 1920
    height = int(props.get("height", 0) or 0) or 1080
    output = str(params.get("output", "image") or "image")
    if output not in ("image", "gif", "rolling_gif"):
        output = "image"
    if output == "image":
        output_path = files.get_unique_filename("heatmap.png", file_format=".png")
        if style == "template":
            result = screenspace_heatmap.generate_template_heatmap(
                results, width, height, output_path
            )
        elif style == "flow":
            result = screenspace_heatmap.generate_flow_heatmap(
                results, width, height, output_path
            )
        elif style == "attention":
            result = screenspace_heatmap.generate_attention_heatmap(
                results, width, height, output_path
            )
        else:
            result = screenspace_heatmap.generate_change_heatmap(
                results, width, height, output_path
            )
        failure_note = "Heatmap couldn't be generated"
    else:
        output_path = files.get_unique_filename("heatmap.gif", file_format=".gif")
        num_frames = int(float(params.get("frames", 24) or 24))
        if output == "gif":
            anim = screenspace_heatmap.generate_heatmap_gif(
                results,
                width,
                height,
                output_path,
                heatmap_type=style,
                num_frames=num_frames,
            )
        else:
            anim = screenspace_heatmap.generate_rolling_heatmap_gif(
                results,
                width,
                height,
                output_path,
                heatmap_type=style,
                num_frames=num_frames,
                window_frames=int(float(params.get("window", 6) or 6)),
            )
        # No sprite sheet here: this node's output is a standalone artifact for
        # the manifest/viewer, not a Screenspace results thumbnail to scrub.
        result = anim["path"] if anim else None
        failure_note = "Not enough detector results for an animated heatmap"
    if not result:
        files.release_reservation(output_path)
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": failure_note,
        }
    rec = _attachment_artifact("heatmap", result, src, f"{style.title()} heatmap")
    return {"artifacts": {"artifacts": [rec], "study": study, "count": 1}}


def _infer_heatmap_style(results: list[Any]) -> str:
    """Pick the heatmap style from the detector payload keys in *results*.

    Each style's generator reads a distinctive key (template match boxes,
    flow/change/saliency grids — see ``screenspace_heatmap._GRID_KEYS``), so
    the first result carrying one decides. Empty string when none match.
    """
    for r in results:
        if not isinstance(r, dict):
            continue
        if r.get("saliency_grid") is not None:
            return "attention"
        if r.get("flow_grid") is not None:
            return "flow"
        if r.get("change_grid") is not None:
            return "change"
        if r.get("matches"):
            return "template"
    return ""


def _reel_start_seconds(rec: Any) -> float:
    """Earliest start (seconds) of a clip record, for chronological reels.

    Adapter-built records carry pre-resolved ``times`` (H:MM:SS); sheet records
    resolve lazily from their ``cell`` (matching
    ``spreadsheet.sort_clips_chronologically``). Unparseable records sort last.
    """
    times = rec.get("times")
    if not times:
        cell = rec.get("cell")
        cell_value = str(getattr(cell, "value", "") or "") if cell is not None else ""
        cleaned, _, _ = utils.parse_cell_annotations(cell_value)
        times = utils.parse_timestamps(cleaned)
    if not times:
        return float("inf")
    seconds = utils.timestamp_to_seconds(times[0][0])
    return seconds if seconds is not None else float("inf")


def _exec_build_reel(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import files
    import pipeline

    clips_in = inputs.get("clips") or {}
    records = list(clips_in.get("records") or [])
    study = str(clips_in.get("study", "") or "")
    if params.get("chronological"):
        records = sorted(records, key=_reel_start_seconds)
    if not records:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "manifest": {"path": None, "records": []},
            "__note__": "No clips to build a reel from",
        }
    # Honor the node's reel name: reserve a unique output path (process_reel
    # treats a supplied output_file as a reservation and releases it on failure).
    name = utils.sanitize_filename(str(params.get("name", "") or "").strip()) or "reel"
    output_file = files.get_unique_filename(f"{name}{config.FILEFORMAT}")
    pad_pre, pad_post, max_duration = _artifact_padding_params(params)
    count, reels = pipeline.process_reel(
        records,
        output_file=output_file,
        cancel_flag=ctx.cancel_flag,
        pad_pre=pad_pre,
        pad_post=pad_post,
        max_duration=max_duration,
    )
    return {
        "artifacts": {"artifacts": reels, "study": study, "count": count},
        "manifest": {"path": None, "records": reels},
    }


def _exec_post_process(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    """Post-process the wired source video: subtitles, loudness, remux, or size.

    Two shapes of operation, mirrored from the app's own actions:

    - **In place** (``normalize_audio``, ``remux_faststart``) rewrite the source
      file exactly like the Transcripts page's quick actions — the original is
      kept beside it as ``.orig`` and the ``video`` output passes the input
      descriptor through unchanged.
    - **Copies** (``embed_subtitles``, ``compress``) write a new file into the
      output dir; the ``video`` output points at the copy so downstream nodes
      chain onto it, and ``artifacts`` carries a record for the manifest/viewer.
    """
    import shutil
    import tempfile

    import files
    import transcripts
    import video as video_mod

    src = inputs.get("video") or {}
    paths = list(src.get("video_paths") or [])
    study = str(src.get("study", "") or "")
    op = str(params.get("operation", "embed_subtitles") or "embed_subtitles")

    def _done(
        video_out: dict[str, Any],
        records: list[dict[str, Any]],
        note: str | None = None,
    ) -> dict[str, Any]:
        out: dict[str, Any] = {
            "video": video_out,
            "artifacts": {
                "artifacts": records,
                "study": study,
                "count": len(records),
            },
        }
        if note:
            out["__note__"] = note
        return out

    if not paths:
        return _done(src, [], "No video wired")

    if op == "embed_subtitles":
        transcript_in = inputs.get("transcript") or {}
        segments = list(transcript_in.get("segments") or [])
        if not segments:
            return _done(src, [], "Embedding subtitles needs a wired transcript")
        if len(paths) > 1:
            # The transcript's timing spans the stitched timeline; muxing it
            # into any single part would be wrong (same rule as the app).
            return _done(
                src,
                [],
                "Subtitle embedding isn't supported for multi-video participants",
            )
        result = cast(
            transcripts.TranscriptResult,
            {
                "segments": segments,
                "language": str(transcript_in.get("language", "") or ""),
                "model": str(transcript_in.get("model", "") or ""),
                "source_file": str(src.get("source_filename", "") or ""),
            },
        )
        source_path = Path(paths[0])
        with tempfile.NamedTemporaryFile(suffix=".srt", delete=False) as tmp:
            srt_path = tmp.name
        try:
            if not transcripts.write_transcript(result, srt_path, fmt="srt"):
                return _done(src, [], "Subtitle file couldn't be written")
            out_path = files.get_unique_filename(
                f"{source_path.stem}-subtitled{source_path.suffix}",
                file_format=source_path.suffix,
            )
            if not video_mod.mux_subtitles(
                str(source_path),
                srt_path,
                out_path,
                set_default=bool(params.get("default_track", True)),
            ):
                files.release_reservation(out_path)
                return _done(src, [], "Subtitle mux failed")
        finally:
            Path(srt_path).unlink(missing_ok=True)
        new_src = {
            **src,
            "video_paths": [out_path],
            "source_filename": Path(out_path).name,
        }
        rec = _attachment_artifact("export", out_path, src, "Subtitled video")
        return _done(new_src, [rec])

    if op == "normalize_audio":
        # Per-part in-place rewrite, sharing the .orig slot (and its
        # already-rewritten skip) with remux — mirrors the Transcripts page.
        import transcripts_server

        failures: list[str] = []
        done = 0
        already = 0
        for path in paths:
            if ctx.cancel_flag():
                failures.append(f"{Path(path).name}: cancelled")
                break
            if video_mod.original_backup_path(path).exists():
                already += 1
                continue
            props = video_mod.probe_video_properties(path)
            if props is None:
                failures.append(f"{Path(path).name}: could not probe the file")
                continue
            indices = transcripts_server._resolve_normalize_indices(props, "auto")
            if isinstance(indices, str):
                failures.append(f"{Path(path).name}: {indices}")
                continue
            success, message = video_mod.normalize_audio_inplace(
                path, indices, cancel_flag=ctx.cancel_flag
            )
            if success:
                done += 1
            else:
                failures.append(f"{Path(path).name}: {message}")
        if failures:
            raise RuntimeError("Normalize audio failed: " + " ".join(failures))
        note = None
        if already and not done:
            note = "Already rewritten; the original is still kept beside the source"
        return _done(src, [], note)

    if op == "remux_faststart":
        failures = []
        done = 0
        already = 0
        for path in paths:
            if ctx.cancel_flag():
                failures.append(f"{Path(path).name}: cancelled")
                break
            seek = video_mod.probe_container_seekability(path)
            if seek is not None and seek.get("browser_seekable"):
                already += 1
                continue
            success, message = video_mod.remux_to_faststart(
                path, cancel_flag=ctx.cancel_flag
            )
            if success:
                done += 1
            else:
                failures.append(f"{Path(path).name}: {message}")
        if failures:
            raise RuntimeError("Remux failed: " + " ".join(failures))
        note = "Already browser-seekable" if already and not done else None
        return _done(src, [], note)

    # compress — write a size-capped copy per part into the output dir. The
    # sources are never touched (the app only ever recompresses generated
    # artifacts, so an in-place source rewrite here would have no precedent).
    target_mb = max(1.0, float(params.get("target_mb", 100) or 100))
    records: list[dict[str, Any]] = []
    out_paths: list[str] = []
    for path in paths:
        if ctx.cancel_flag():
            break
        source_path = Path(path)
        out_path = files.get_unique_filename(
            f"{source_path.stem}-compressed{source_path.suffix}",
            file_format=source_path.suffix,
        )
        try:
            shutil.copyfile(path, out_path)
        except OSError as exc:
            files.release_reservation(out_path)
            raise RuntimeError(f"Couldn't copy {source_path.name}: {exc}") from exc
        if not video_mod.compress_to_size(
            out_path, target_mb, cancel_flag=ctx.cancel_flag
        ):
            # False = encode error or cancel (an already-small file is True).
            files.release_reservation(out_path)
            raise RuntimeError(f"Compression failed for {source_path.name}")
        out_paths.append(out_path)
        records.append(
            _attachment_artifact(
                "export", out_path, src, f"Compressed copy (≤{target_mb:g} MB)"
            )
        )
    new_src = {**src, "video_paths": out_paths} if out_paths else src
    return _done(new_src, records)


def _exec_data_export(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import data_export
    import files

    events_in = inputs.get("events") or {}
    seg_in = inputs.get("segments") or {}
    src = events_in.get("source") or seg_in.get("source") or {}
    study = str(src.get("study", "") or "")
    fmt = str(params.get("format", "both") or "both")
    if fmt not in ("both", "json", "csv"):
        fmt = "both"

    participant = str(src.get("participant", "") or "")
    suffix = f"_{participant}" if participant else ""
    # (stem, rows, preferred CSV column order, description) per wired surface.
    surfaces: list[tuple[str, list[dict[str, Any]], tuple[str, ...], str]] = []
    events = list(events_in.get("events") or [])
    if events:
        surfaces.append(
            (
                f"export_events{suffix}",
                data_export.build_screenspace_events(
                    {"events": events}, include_excluded=True
                ),
                data_export.SCREENSPACE_EVENT_COLUMNS,
                "Events export",
            )
        )
    segments = list(seg_in.get("segments") or [])
    if segments:
        # build_transcript_segments reads a transcripts-manifest envelope;
        # synthesize one around the wired segments.
        manifest = {
            "source_transcripts": {
                (participant or "unknown"): {
                    "segments": segments,
                    "source_file": str(src.get("source_filename", "") or ""),
                }
            }
        }
        surfaces.append(
            (
                f"export_segments{suffix}",
                data_export.build_transcript_segments(manifest),
                data_export._TRANSCRIPT_SEGMENT_BASE_COLS,
                "Segments export",
            )
        )
    # The opt-in tables read the app manifests on disk (not wired ports):
    # calibration pins from Screenspace, friction moments + scored segments
    # from the Transcripts thinking agents. Empty tables are skipped silently —
    # the toggles say "include if present", not "must exist".
    if params.get("include_pins"):
        import screenspace

        pins = data_export.build_screenspace_pins(
            screenspace.load_screenspace_manifest()
        )
        if pins:
            surfaces.append(
                (
                    "export_pins",
                    pins,
                    data_export.SCREENSPACE_PIN_COLUMNS,
                    "Calibration pins export",
                )
            )
    if params.get("include_friction"):
        import transcripts

        t_manifest = transcripts.load_transcripts_manifest()
        moments = data_export.build_friction_moments(t_manifest)
        if moments:
            surfaces.append(
                (
                    "export_friction_moments",
                    moments,
                    data_export._FRICTION_MOMENT_COLS,
                    "Friction moments export",
                )
            )
        scored = data_export.build_friction_segments(t_manifest)
        if scored:
            surfaces.append(
                (
                    "export_friction_segments",
                    scored,
                    data_export._FRICTION_SEGMENT_COLS,
                    "Friction segments export",
                )
            )
    if not surfaces:
        return {
            "artifacts": {"artifacts": [], "study": study, "count": 0},
            "__note__": "No events or segments wired (and no opt-in tables had rows)",
        }

    records: list[dict[str, Any]] = []
    written: list[str] = []
    for stem, rows, columns, description in surfaces:
        writes: list[tuple[str, str, str]] = []  # (extension, payload, label)
        if fmt in ("both", "json"):
            writes.append((".json", data_export.to_json(rows), "JSON"))
        if fmt in ("both", "csv"):
            writes.append(
                (
                    ".csv",
                    data_export.to_csv(rows, preferred_column_order=columns),
                    "CSV",
                )
            )
        for ext, payload, label in writes:
            output_path = files.get_unique_filename(f"{stem}{ext}", file_format=ext)
            try:
                Path(output_path).write_text(payload, encoding="utf-8")
            except OSError:
                # All-or-nothing: remove files this node already wrote, so a
                # half-bundle never orphans behind an artifact-less result.
                files.release_reservation(output_path)
                for prior in written:
                    files.release_reservation(prior)
                return {
                    "artifacts": {"artifacts": [], "study": study, "count": 0},
                    "__note__": "Export couldn't be written",
                }
            written.append(output_path)
            records.append(
                _attachment_artifact(
                    "export", output_path, src, f"{description} ({label})"
                )
            )
    return {"artifacts": {"artifacts": records, "study": study, "count": len(records)}}


def _exec_timeline_viewer(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import viewer

    artifacts_in = inputs.get("artifacts") or {}
    incoming = list(artifacts_in.get("artifacts") or [])
    study = str(artifacts_in.get("study", "") or "")
    # build_reel emits reel records (``components``, no start/end) onto the same
    # ``artifacts`` wire as clip/screen/gif artifacts, but the viewer renders the
    # two from separate slots — timeline vs Attachments pane. Split them here, or
    # reels land in the timeline slot, get filtered for lacking start/end, and the
    # viewer comes up empty.
    reels = [
        a for a in incoming if isinstance(a, dict) and a.get("components") is not None
    ]
    artifacts = [
        a
        for a in incoming
        if not (isinstance(a, dict) and a.get("components") is not None)
    ]
    ss_events = (inputs.get("events") or {}).get("events") or None
    data = viewer.finalize_timeline_data(
        artifacts, reels=reels or None, study=study, screenspace_events=ss_events
    )
    path = viewer.generate_timeline_viewer(data, output_basename="workflow_viewer.html")
    return {"viewer": {"path": str(path) if path else None}}


def _exec_gallery_viewer(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    import video as video_mod
    import viewer

    artifacts_in = inputs.get("artifacts") or {}
    incoming = list(artifacts_in.get("artifacts") or [])
    # The gallery renders image/GIF captures; clips and reels belong to the
    # timeline viewer.
    stills = [
        a
        for a in incoming
        if isinstance(a, dict) and a.get("type") in ("screen", "gif")
    ]
    if not stills:
        return {
            "viewer": {"path": None},
            "__note__": "No screenshot/GIF artifacts wired (clips go to Timeline Viewer)",
        }
    src = artifacts_in.get("source") or {}
    fmt = "gif" if any(a.get("type") == "gif" for a in stills) else "screen"
    paths = list(src.get("video_paths") or [])
    duration = int(video_mod.get_file_duration(paths[0]) or 0) if paths else 0
    data = viewer.finalize_gallery_data(
        stills,
        source_video=str(src.get("source_filename", "") or ""),
        video_duration=duration,
        output_format=fmt,
    )
    path = viewer.generate_gallery_viewer(data, output_basename="workflow_gallery.html")
    return {"viewer": {"path": str(path) if path else None}}


# ---- Control ----


_GATE_OPS: dict[str, Callable[[float, float], bool]] = {
    ">=": lambda a, b: a >= b,
    ">": lambda a, b: a > b,
    "<=": lambda a, b: a <= b,
    "<": lambda a, b: a < b,
    "==": lambda a, b: a == b,
    "!=": lambda a, b: a != b,
}


def _reduce_collection(metric: str, inputs: dict[str, Any]) -> float:
    """Reduce a wired collection to one scalar (the measure / gate-collection core).

    Reads whichever of events / clipRecords / segments is wired (events first).
    ``max_confidence`` only applies to events; it falls back to 0 otherwise.
    """
    events = (inputs.get("events") or {}).get("events")
    records = (inputs.get("clips") or {}).get("records")
    segments = (inputs.get("segments") or {}).get("segments")

    if events is not None:
        items = list(events)
        if metric == "count":
            return float(len(items))
        if metric == "max_confidence":
            confs = [float(e.get("confidence", 0.0) or 0.0) for e in items]
            return max(confs) if confs else 0.0
        return float(
            sum(
                max(
                    0.0,
                    float(e.get("time_out", 0.0) or 0.0)
                    - float(e.get("time_in", 0.0) or 0.0),
                )
                for e in items
            )
        )

    if records is not None:
        items = list(records)
        if metric == "count":
            return float(len(items))
        if metric == "max_confidence":
            return 0.0
        total = 0.0
        for rec in items:
            for start_str, end_str in rec.get("times") or []:
                start = utils.timestamp_to_seconds(start_str)
                end = utils.timestamp_to_seconds(end_str)
                if start is not None and end is not None:
                    total += max(0.0, end - start)
        return total

    if segments is not None:
        items = list(segments)
        if metric == "count":
            return float(len(items))
        if metric == "max_confidence":
            return 0.0
        return float(
            sum(
                max(
                    0.0,
                    float(s.get("end", 0.0) or 0.0) - float(s.get("start", 0.0) or 0.0),
                )
                for s in items
            )
        )

    return 0.0


def _apply_gate(value: float, params: dict[str, Any]) -> bool:
    """Compare *value* to a threshold per the node's ``op`` (shared gate logic)."""
    fn = _GATE_OPS.get(str(params.get("op", ">=") or ">="))
    if fn is None:
        return False
    try:
        return bool(fn(value, float(params.get("threshold", 0))))
    except (TypeError, ValueError):
        return False


def _exec_gate(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    raw = inputs.get("value")
    if raw is None:
        return {"pass": False}
    try:
        value = float(raw)
    except (TypeError, ValueError):
        return {"pass": False}
    return {"pass": _apply_gate(value, params)}


def _exec_gate_collection(
    ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
) -> dict[str, Any]:
    """Reduce a wired collection to a scalar then gate it (measure+gate fused —
    the scalar :func:`_exec_gate` remains for the video→scalar adapter path)."""
    metric = str(params.get("metric", "count") or "count")
    value = _reduce_collection(metric, inputs)
    return {"pass": _apply_gate(value, params)}


# ---- Collection-algebra control nodes (filter / merge / limit / dedup) ----
#
# These thin / combine / branch / cap / dedup the collections already flowing
# through the graph (events, clipRecords, segments). The collections *are* the
# iteration, hence no per-item ``foreach``, which would force runtime DAG
# expansion and break the static ``topo_order`` model. All pure single-pass
# nodes; per-type families (mirroring the ss_* split) keep every port exact-typed,
# so neither adapters nor the frontend's ``canConnect`` need changes.

# One ``{field, op, value}`` clause reuses the gate's comparison table; ``contains``
# is the one string-only addition (kept out of ``_GATE_OPS``, which must stay
# numeric for the gate). ``none`` is the limit node's "keep input order" sentinel.
_COLLECTION_OPS: list[str] = list(_GATE_OPS.keys()) + ["contains"]
_SORT_NONE = "none"

# kind -> envelope metadata. ``port`` is the wire type; ``key`` is the inner list
# key in the envelope; ``preserve`` are the envelope keys carried through unchanged
# (source lineage / study / raw_results); ``fields`` drive the predicate enum and
# ``sort_fields`` the limit sort enum (numeric-only, so the sort key stays
# comparable). ``recount`` (artifacts) rewrites the envelope's ``count`` to the
# kept length. dedup is span-based, registered for events + clips + timeRanges.
#
# Both the pre-clip side (``clipRecords``, ``timeRange``) and the post-clip side
# (``artifacts`` from make_clips/timelapse/heatmap/build_reel) get families, so a
# stream can be thinned/capped/combined before *or* after it becomes artifacts.
_COLLECTION_KINDS: dict[str, dict[str, Any]] = {
    "events": {
        "port": "events",
        "key": "events",
        "label": "Events",
        "preserve": ("source", "raw_results"),
        "fields": ["confidence", "duration", "start"],
        "sort_fields": ["confidence", "duration", "start"],
    },
    "clips": {
        "port": "clipRecords",
        # Labelled "Clip Selections" (not "Clips") so the family reads as operating
        # on pre-render clip specs from sheet_selection/highlights — not the
        # rendered ``artifacts`` Make Clips emits (which has its own artifacts
        # family). The node ids stay ``*_clips`` so saved blueprints are unaffected.
        "key": "records",
        "label": "Clip Selections",
        "preserve": ("study",),
        "fields": ["duration", "category", "severity", "desc"],
        "sort_fields": ["duration"],
    },
    "segments": {
        "port": "segments",
        "key": "segments",
        "label": "Segments",
        "preserve": ("source",),
        "fields": ["text", "duration", "start"],
        "sort_fields": ["duration", "start"],
    },
    "timerange": {
        "port": "timeRange",
        "key": "ranges",
        "label": "Time Ranges",
        "preserve": ("source",),
        "fields": ["duration", "start"],
        "sort_fields": ["duration", "start"],
    },
    "artifacts": {
        "port": "artifacts",
        "key": "artifacts",
        "label": "Artifacts",
        "preserve": ("study",),
        "recount": True,
        "fields": ["duration", "start", "type", "category", "severity", "participant"],
        "sort_fields": ["duration", "start"],
    },
}


def _collection_field(kind: str, item: Any, field: str) -> float | str | None:
    """Extract one predicate/sort field from a collection item.

    Numeric fields return ``float``; text fields (``category``/``severity``/
    ``desc``/``text``/``type``/``participant``) return ``str``. Clip ``duration``
    sums the record's ``times`` spans, mirroring ``_exec_measure``'s total-duration
    path. ``timerange`` items are ``(start, end)`` tuples, not dicts.
    """
    if kind == "timerange":
        if not isinstance(item, (list, tuple)) or len(item) < 2:
            return None
        start = float(item[0] or 0.0)
        end = float(item[1] or 0.0)
        if field == "duration":
            return max(0.0, end - start)
        if field == "start":
            return start
        return None
    if not isinstance(item, dict):
        return None
    if kind == "artifacts":
        if field == "duration":
            return max(
                0.0,
                float(item.get("end", 0.0) or 0.0)
                - float(item.get("start", 0.0) or 0.0),
            )
        if field == "start":
            return float(item.get("start", 0.0) or 0.0)
        if field in ("type", "category", "severity", "participant"):
            return str(item.get(field, "") or "")
        return None
    if kind == "events":
        if field == "confidence":
            return float(item.get("confidence", 0.0) or 0.0)
        if field == "duration":
            return max(
                0.0,
                float(item.get("time_out", 0.0) or 0.0)
                - float(item.get("time_in", 0.0) or 0.0),
            )
        if field == "start":
            return float(item.get("time_in", 0.0) or 0.0)
    elif kind == "clips":
        if field == "duration":
            total = 0.0
            for start_str, end_str in item.get("times") or []:
                start = utils.timestamp_to_seconds(start_str)
                end = utils.timestamp_to_seconds(end_str)
                if start is not None and end is not None:
                    total += max(0.0, end - start)
            return total
        if field in ("category", "severity", "desc"):
            return str(item.get(field, "") or "")
    elif kind == "segments":
        if field == "text":
            return str(item.get("text", "") or "")
        if field == "duration":
            return max(
                0.0,
                float(item.get("end", 0.0) or 0.0)
                - float(item.get("start", 0.0) or 0.0),
            )
        if field == "start":
            return float(item.get("start", 0.0) or 0.0)
    return None


def _eval_predicate(field_val: float | str | None, op: str, raw_value: Any) -> bool:
    """Evaluate one ``{field, op, value}`` clause against an extracted field value.

    ``contains`` is a case-insensitive substring test; text fields support only
    ``==``/``!=`` (case-insensitive, trimmed). Numeric fields route through
    ``_GATE_OPS`` after a ``float()`` coerce (mirrors ``_exec_gate``'s guard).
    """
    if field_val is None:
        return False
    if op == "contains":
        return str(raw_value).strip().lower() in str(field_val).lower()
    if isinstance(field_val, str):
        a = field_val.strip().lower()
        b = str(raw_value).strip().lower()
        if op == "==":
            return a == b
        if op == "!=":
            return a != b
        return False  # ordering ops don't apply to text fields
    fn = _GATE_OPS.get(op)
    if fn is None:
        return False
    try:
        return fn(float(field_val), float(raw_value))
    except (TypeError, ValueError):
        return False


def _eval_clauses(kind: str, item: Any, params: dict[str, Any]) -> bool:
    """Evaluate the primary ``{field, op, value}`` clause plus the optional
    second clause, combined with AND/OR when ``combine`` isn't "off"."""
    meta = _COLLECTION_KINDS[kind]
    field = str(params.get("field") or meta["fields"][0])
    op = str(params.get("op") or ">=")
    first = _eval_predicate(
        _collection_field(kind, item, field), op, params.get("value")
    )
    combine = str(params.get("combine") or "off")
    if combine not in ("AND", "OR"):
        return first
    field2 = str(params.get("field2") or meta["fields"][0])
    op2 = str(params.get("op2") or ">=")
    second = _eval_predicate(
        _collection_field(kind, item, field2), op2, params.get("value2")
    )
    return (first and second) if combine == "AND" else (first or second)


def _wrap_collection(
    kind: str, src_envelope: dict[str, Any], items: list[Any]
) -> dict[str, Any]:
    """Re-wrap an item list in the kind's envelope, preserving source lineage."""
    meta = _COLLECTION_KINDS[kind]
    out: dict[str, Any] = {meta["key"]: items}
    for preserve_key in meta["preserve"]:
        if preserve_key in src_envelope:
            out[preserve_key] = src_envelope[preserve_key]
    if meta.get("recount"):
        out["count"] = len(items)  # artifacts carry a count; keep it honest
    return out


def _make_filter_executor(
    kind: str,
) -> Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]:
    """Keep items matching the clause(s) on ``out``; the rest go to ``unmatched``.

    The second output makes filter subsume the old partition family — leave
    ``unmatched`` unwired for a plain filter, wire it for the gate's data-level
    branch. Same type on every port (runner stores the whole result dict;
    consumers read per-port)."""
    meta = _COLLECTION_KINDS[kind]

    def _exec(
        ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
    ) -> dict[str, Any]:
        env = inputs.get("in") or {}
        items = list(env.get(meta["key"]) or [])
        kept: list[Any] = []
        rejected: list[Any] = []
        for it in items:
            target = kept if _eval_clauses(kind, it, params) else rejected
            target.append(it)
        return {
            "out": _wrap_collection(kind, env, kept),
            "unmatched": _wrap_collection(kind, env, rejected),
        }

    return _exec


def _make_merge_executor(
    kind: str,
) -> Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]:
    """Union 2-3 same-type collections into one (fixes the one-wire-per-input
    wall). Preserves the first wired input's lineage; concatenates ``raw_results``
    for events so a downstream heatmap still sees per-frame coverage."""
    meta = _COLLECTION_KINDS[kind]

    def _exec(
        ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
    ) -> dict[str, Any]:
        merged: list[Any] = []
        raw: list[Any] = []
        base_env: dict[str, Any] = {}
        for port in ("in1", "in2", "in3"):
            env = inputs.get(port)
            if not isinstance(env, dict):
                continue
            if not base_env:
                base_env = env
            seq = env.get(meta["key"])
            if isinstance(seq, list):
                merged.extend(seq)
            extra = env.get("raw_results")
            if isinstance(extra, list):
                raw.extend(extra)
        out = _wrap_collection(kind, base_env, merged)
        if kind == "events":
            out["raw_results"] = raw
        return {"out": out}

    return _exec


def _make_limit_executor(
    kind: str,
) -> Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]:
    """Optionally sort by a numeric field, then keep the first N items."""
    meta = _COLLECTION_KINDS[kind]

    def _exec(
        ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
    ) -> dict[str, Any]:
        env = inputs.get("in") or {}
        items = list(env.get(meta["key"]) or [])
        sort_by = str(params.get("sort_by") or _SORT_NONE)
        order = str(params.get("order") or "desc")
        if sort_by != _SORT_NONE:

            def _key(it: dict[str, Any]) -> float:
                val = _collection_field(kind, it, sort_by)
                return float(val) if isinstance(val, (int, float)) else 0.0

            items.sort(key=_key, reverse=(order == "desc"))
        try:
            take = int(params.get("take", 0) or 0)
        except (TypeError, ValueError):
            take = 0
        if take > 0:
            items = items[:take]
        return {"out": _wrap_collection(kind, env, items)}

    return _exec


def _dedup_events(items: list[dict[str, Any]], gap: float) -> list[dict[str, Any]]:
    """Merge events whose spans overlap or sit within ``gap`` seconds; the merged
    event spans the union and keeps the max confidence."""
    ordered = sorted(items, key=lambda e: float(e.get("time_in", 0.0) or 0.0))
    out: list[dict[str, Any]] = []
    for ev in ordered:
        t_in = float(ev.get("time_in", 0.0) or 0.0)
        t_out = float(ev.get("time_out", t_in) or t_in)
        if out:
            prev = out[-1]
            p_in = float(prev.get("time_in", 0.0) or 0.0)
            p_out = float(prev.get("time_out", p_in) or p_in)
            if t_in <= p_out + gap:
                # Keep the higher-confidence member's fields, union the span.
                ev_conf = float(ev.get("confidence", 0.0) or 0.0)
                prev_conf = float(prev.get("confidence", 0.0) or 0.0)
                merged = dict(ev if ev_conf > prev_conf else prev)
                merged["time_in"] = min(p_in, t_in)
                merged["time_out"] = max(p_out, t_out)
                merged["confidence"] = max(prev_conf, ev_conf)
                out[-1] = merged
                continue
        out.append(dict(ev))
    return out


def _dedup_clips(records: list[dict[str, Any]], gap: float) -> list[dict[str, Any]]:
    """Drop clip records whose overall time-span overlaps (within ``gap``) a
    record already kept; keeps the first, extending the covered span."""

    def _span(rec: dict[str, Any]) -> tuple[float, float] | None:
        starts: list[float] = []
        ends: list[float] = []
        for start_str, end_str in rec.get("times") or []:
            start = utils.timestamp_to_seconds(start_str)
            end = utils.timestamp_to_seconds(end_str)
            if start is not None:
                starts.append(start)
            if end is not None:
                ends.append(end)
        if not starts or not ends:
            return None
        return (min(starts), max(ends))

    indexed = sorted(
        ((_span(rec), rec) for rec in records),
        key=lambda pair: pair[0][0] if pair[0] else 0.0,
    )
    out: list[dict[str, Any]] = []
    last_span: tuple[float, float] | None = None
    for span, rec in indexed:
        if span and last_span and span[0] <= last_span[1] + gap:
            last_span = (last_span[0], max(last_span[1], span[1]))
            continue
        out.append(rec)
        # Keep an untimed record but never let its None span clobber the tracker —
        # otherwise the next overlap check short-circuits and later duplicates leak.
        if span is not None:
            last_span = span
    return out


def _dedup_timeranges(items: list[Any], gap: float) -> list[tuple[float, float]]:
    """Merge ``(start, end)`` windows that overlap or sit within ``gap`` seconds."""
    spans: list[tuple[float, float]] = []
    for it in items:
        if isinstance(it, (list, tuple)) and len(it) >= 2:
            spans.append((float(it[0] or 0.0), float(it[1] or 0.0)))
    spans.sort()
    out: list[tuple[float, float]] = []
    for start, end in spans:
        if out and start <= out[-1][1] + gap:
            out[-1] = (out[-1][0], max(out[-1][1], end))
        else:
            out.append((start, end))
    return out


def _make_dedup_executor(
    kind: str,
) -> Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]:
    """Span-merge near-duplicate items (events, clips, and time ranges)."""
    meta = _COLLECTION_KINDS[kind]

    def _exec(
        ctx: NodeContext, inputs: dict[str, Any], params: dict[str, Any]
    ) -> dict[str, Any]:
        env = inputs.get("in") or {}
        items = list(env.get(meta["key"]) or [])
        try:
            gap = max(0.0, float(params.get("gap", 0) or 0))
        except (TypeError, ValueError):
            gap = 0.0
        if kind == "events":
            items = _dedup_events(items, gap)
        elif kind == "clips":
            items = _dedup_clips(items, gap)
        elif kind == "timerange":
            items = _dedup_timeranges(items, gap)
        return {"out": _wrap_collection(kind, env, items)}

    return _exec


# The predicate/sort fields that compare as text in _collection_field; every
# other field coerces to float. Serialized per field-enum as numericChoices so
# the editor and validation know which values must be numbers.
_TEXT_FIELDS = frozenset(
    {"category", "severity", "desc", "text", "type", "participant"}
)


def _predicate_params(kind: str) -> list[ParamSpec]:
    """The shared ``{field, op, value}`` clause for filter / partition nodes."""
    meta = _COLLECTION_KINDS[kind]
    numeric = [f for f in meta["fields"] if f not in _TEXT_FIELDS]
    # An ordering default on a text field (segments' first field is ``text``)
    # would drop every item; default those kinds to ``contains`` instead.
    default_op = ">=" if meta["fields"][0] not in _TEXT_FIELDS else "contains"
    second_clause: dict[str, Any] = {"param": "combine", "not": "off"}
    return [
        {
            "name": "field",
            "type": "enum",
            "default": meta["fields"][0],
            "choices": list(meta["fields"]),
            "label": "Field",
            "numericChoices": numeric,
        },
        {
            "name": "op",
            "type": "enum",
            "default": default_op,
            "choices": list(_COLLECTION_OPS),
            "label": "Comparison",
        },
        {
            "name": "value",
            "type": "string",
            "default": "",
            "label": "Value",
            "required": True,
            "numericFor": "field",
        },
        # Optional second clause. "off" keeps the node single-clause; value2 is
        # deliberately not required so validation stays quiet in that case.
        {
            "name": "combine",
            "type": "enum",
            "default": "off",
            "choices": ["off", "AND", "OR"],
            "label": "Second clause",
        },
        {
            "name": "field2",
            "type": "enum",
            "default": meta["fields"][0],
            "choices": list(meta["fields"]),
            "label": "Field 2",
            "numericChoices": numeric,
            "showIf": second_clause,
        },
        {
            "name": "op2",
            "type": "enum",
            "default": default_op,
            "choices": list(_COLLECTION_OPS),
            "label": "Comparison 2",
            "showIf": second_clause,
        },
        {
            "name": "value2",
            "type": "string",
            "default": "",
            "label": "Value 2",
            "numericFor": "field2",
            "showIf": second_clause,
        },
    ]


def _limit_params(kind: str) -> list[ParamSpec]:
    """Sort-key + order + count for the limit (top-N) node."""
    meta = _COLLECTION_KINDS[kind]
    return [
        {
            "name": "sort_by",
            "type": "enum",
            "default": _SORT_NONE,
            "choices": [_SORT_NONE] + list(meta["sort_fields"]),
            "label": "Sort by",
        },
        {
            "name": "order",
            "type": "enum",
            "default": "desc",
            "choices": ["desc", "asc"],
            "label": "Order",
        },
        {
            "name": "take",
            "type": "number",
            "default": 10,
            "min": 0,
            "label": "Keep first N",
            "required": True,
        },
    ]


# Wire each declarative NodeType to its executor. ``serialize_catalog`` strips
# ``execute`` again for the JSON catalog endpoint, so this stays server-internal.
_EXECUTORS: dict[
    str, Callable[[NodeContext, dict[str, Any], dict[str, Any]], dict[str, Any]]
] = {
    "video_source": _exec_video_source,
    "sheet_selection": _exec_sheet_selection,
    "region": _exec_region,
    "time_range": _exec_time_range,
    "transcribe": _exec_transcribe,
    "find_word": _exec_find_word,
    "transcript_marks": _exec_transcript_marks,
    "transcript_export": _exec_transcript_export,
    "data_export": _exec_data_export,
    "summarize": _exec_summarize,
    "citations": _exec_citations,
    "friction": _exec_friction,
    "report": _exec_report,
    "multitool": _exec_multitool,
    "highlights": _exec_highlights,
    "make_clips": _exec_make_clips,
    "interval_captures": _exec_interval_captures,
    "build_reel": _exec_build_reel,
    "post_process": _exec_post_process,
    "timelapse": _exec_timelapse,
    "heatmap": _exec_heatmap,
    "timeline_viewer": _exec_timeline_viewer,
    "gallery_viewer": _exec_gallery_viewer,
    "gate": _exec_gate,
    "gate_collection": _exec_gate_collection,
}

# The ten per-detector Screenspace nodes share one body via the factory above;
# the unified ``detect`` node dispatches into the same body by ``detector`` param.
for _ss_tool in _SS_DETECTOR_SPECS:
    _EXECUTORS[f"ss_{_ss_tool}"] = _make_ss_executor(_ss_tool)
_EXECUTORS["detect"] = _exec_detect

# Collection-algebra control nodes — per-type families, all factory-generated.
# Registered here (NODE_TYPES + _EXECUTORS together) so the attach loop below
# wires their ``execute`` like any other node. Category "Collection" groups them
# apart from the gates in the palette.
for _kind, _meta in _COLLECTION_KINDS.items():
    _T = _meta["port"]
    _name = _meta["label"]
    _lname = _name.lower()
    NODE_TYPES[f"filter_{_kind}"] = {
        "id": f"filter_{_kind}",
        "label": f"Filter {_name}",
        "domain": "control",
        "category": "Collection",
        "description": (
            f"Keep the {_lname} matching a field/comparison/value test; "
            "the unmatched output carries the rest."
        ),
        "inputs": [{"name": "in", "type": _T}],
        "outputs": [
            {"name": "out", "type": _T},
            {"name": "unmatched", "type": _T},
        ],
        "params": _predicate_params(_kind),
        "requires": [],
    }
    _EXECUTORS[f"filter_{_kind}"] = _make_filter_executor(_kind)
    NODE_TYPES[f"merge_{_kind}"] = {
        "id": f"merge_{_kind}",
        "label": f"Merge {_name}",
        "domain": "control",
        "category": "Collection",
        "description": f"Combine two or three {_lname} streams into one.",
        "inputs": [
            {"name": "in1", "type": _T},
            {"name": "in2", "type": _T, "optional": True},
            {"name": "in3", "type": _T, "optional": True},
        ],
        "outputs": [{"name": "out", "type": _T}],
        "params": [],
        "requires": [],
    }
    _EXECUTORS[f"merge_{_kind}"] = _make_merge_executor(_kind)
    NODE_TYPES[f"limit_{_kind}"] = {
        "id": f"limit_{_kind}",
        "label": f"Limit {_name}",
        "domain": "control",
        "category": "Collection",
        "description": f"Optionally sort {_lname} by a field, then keep the first N.",
        "inputs": [{"name": "in", "type": _T}],
        "outputs": [{"name": "out", "type": _T}],
        "params": _limit_params(_kind),
        "requires": [],
    }
    _EXECUTORS[f"limit_{_kind}"] = _make_limit_executor(_kind)

# dedup is span-based -> events + clips + time ranges.
for _kind in ("events", "clips", "timerange"):
    _T = _COLLECTION_KINDS[_kind]["port"]
    _name = _COLLECTION_KINDS[_kind]["label"]
    _lname = _name.lower()
    NODE_TYPES[f"dedup_{_kind}"] = {
        "id": f"dedup_{_kind}",
        "label": f"Dedup {_name}",
        "domain": "control",
        "category": "Collection",
        "description": f"Merge overlapping or near-duplicate {_lname} into single spans.",
        "inputs": [{"name": "in", "type": _T}],
        "outputs": [{"name": "out", "type": _T}],
        "params": [
            {
                "name": "gap",
                "type": "number",
                "default": 0,
                "min": 0,
                "label": "Merge gap (s)",
            },
        ],
        "requires": [],
    }
    _EXECUTORS[f"dedup_{_kind}"] = _make_dedup_executor(_kind)

for _node_id, _executor in _EXECUTORS.items():
    NODE_TYPES[_node_id]["execute"] = _executor
