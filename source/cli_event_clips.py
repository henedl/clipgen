"""Event-clip CLI modes: --ss-clips / --transcript-clips / --transcript-mark.

Carved out of cli.py; the mode table there points at these callables. Heavy
modules (screenspace, pipeline) are imported inside the functions so --help
stays fast.
"""

from __future__ import annotations

import argparse
import uuid
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import config
import files
import transcripts
import utils
import viewer
from utils import ClipRecord


_SS_CLIPS_CELL_COL = 1  # synthetic cell column for --ss-clips artifacts
_TRANSCRIPT_CLIPS_CELL_COL = 2  # synthetic cell column for --transcript-clips


def _split_csv_set(value: str | None) -> set[str] | None:
    """Parse a comma-separated CLI value into a set of trimmed non-empty tokens.

    Returns None when the option was not supplied (so callers can distinguish
    "no filter" from "filter rejecting everything").
    """
    if value is None:
        return None
    items = {tok.strip() for tok in value.split(",")}
    items.discard("")
    return items if items else None


def _filter_screenspace_events(
    events: list[dict[str, Any]],
    *,
    detectors: set[str] | None,
    regions: set[str] | None,
    participants: set[str] | None,
    min_confidence: float | None,
    event_type_substr: str | None,
) -> list[dict[str, Any]]:
    """Apply CLI filters to Screenspace events; always drops excluded=True."""
    needle = event_type_substr.lower() if event_type_substr else None
    out: list[dict[str, Any]] = []
    for ev in events:
        if ev.get("excluded"):
            continue
        if detectors and ev.get("detector") not in detectors:
            continue
        if regions and ev.get("region") not in regions:
            continue
        if participants and ev.get("participant") not in participants:
            continue
        if min_confidence is not None:
            try:
                conf = float(ev.get("confidence", 0.0))
            except (TypeError, ValueError):
                conf = 0.0
            if conf < min_confidence:
                continue
        if needle is not None:
            label = str(ev.get("event_type", "")).lower()
            if needle not in label:
                continue
        out.append(ev)
    return out


def _filter_transcript_segments(
    manifest: dict[str, Any],
    *,
    participants: set[str] | None,
    mark_categories: set[str] | None,
    text_substr: str | None,
) -> list[tuple[str, dict[str, Any], list[dict[str, Any]]]]:
    """Return (participant_id, segment, attached_marks) tuples after filtering.

    When ``mark_categories`` is non-empty, only segments referenced by at least
    one matching mark are kept; the matching marks come along for clip metadata.
    Otherwise ``attached_marks`` is empty.
    """
    needle = text_substr.lower() if text_substr else None
    marks = manifest.get("marks") or []
    marks_by_segment: dict[str, list[dict[str, Any]]] = {}
    for mark in marks:
        if mark_categories and mark.get("category") not in mark_categories:
            continue
        seg_id = mark.get("segment_id")
        if not seg_id:
            continue
        marks_by_segment.setdefault(seg_id, []).append(mark)

    rows: list[tuple[str, dict[str, Any], list[dict[str, Any]]]] = []
    sources = manifest.get("source_transcripts") or {}
    corrections = manifest.get("corrections") or []
    for pid, entry in sources.items():
        if participants and pid not in participants:
            continue
        raw_segments = entry.get("segments") or []
        # Match and label on corrected text, like the Transcripts UI.
        corrected = transcripts.apply_corrections(raw_segments, corrections)
        for idx, (raw, seg) in enumerate(zip(raw_segments, corrected, strict=True)):
            seg_id = raw.get("id") or f"{pid}:{idx}"
            attached = marks_by_segment.get(seg_id, [])
            if mark_categories and not attached:
                continue
            if needle is not None and needle not in str(seg.get("text", "")).lower():
                continue
            # apply_corrections drops every key but start/end/text; keep the id.
            row = dict(raw)
            row["text"] = seg.get("text", "")
            rows.append((pid, row, attached))
    return rows


def _cluster_groups(
    keyed_spans: list[tuple[Any, tuple[float, float], int]],
    *,
    gap: float,
    pad_pre: float,
    pad_post: float,
    max_duration: float,
) -> list[tuple[Any, tuple[float, float], list[int]]]:
    """Group items by ``key`` first, then run utils.cluster_spans within each group.

    Returns a list of (group_key, (cluster_start, cluster_end), member_indices),
    where ``member_indices`` index into the original ``keyed_spans`` list (not the
    per-group sublist) — so callers can look up the contributing items directly.
    """
    by_key: dict[Any, list[tuple[tuple[float, float], int]]] = {}
    for key, span, orig_idx in keyed_spans:
        by_key.setdefault(key, []).append((span, orig_idx))

    out: list[tuple[Any, tuple[float, float], list[int]]] = []
    for key, items in by_key.items():
        spans = [span for span, _ in items]
        local_to_orig = [orig for _, orig in items]
        clusters = utils.cluster_spans(
            spans,
            gap_seconds=gap,
            pad_pre=pad_pre,
            pad_post=pad_post,
            max_duration=max_duration,
        )
        for cs, ce, local_members in clusters:
            members = [local_to_orig[i] for i in local_members]
            out.append((key, (cs, ce), members))
    return out


def _build_clusters_from_ss_events(
    events: list[dict[str, Any]],
    *,
    gap: float,
    pad_pre: float,
    pad_post: float,
    max_duration: float,
) -> list[dict[str, Any]]:
    """Group SS events by (participant, source_video, detector) then cluster.

    Returns a list of cluster dicts: ``{participant, source_video, detector,
    region, start, end, regions, event_types, member_event_ids}``.
    """
    keyed: list[tuple[Any, tuple[float, float], int]] = []
    for idx, ev in enumerate(events):
        try:
            t_in = float(ev.get("time_in", 0.0))
            t_out = float(ev.get("time_out", t_in))
        except (TypeError, ValueError):
            continue
        t_out = max(t_out, t_in)
        key = (
            ev.get("participant", ""),
            ev.get("source_video", ""),
            ev.get("detector", ""),
        )
        keyed.append((key, (t_in, t_out), idx))

    clusters_raw = _cluster_groups(
        keyed,
        gap=gap,
        pad_pre=pad_pre,
        pad_post=pad_post,
        max_duration=max_duration,
    )

    out: list[dict[str, Any]] = []
    for key, (cs, ce), members in clusters_raw:
        participant, source_video, detector = key
        regions = sorted(
            {events[m].get("region", "") for m in members if events[m].get("region")}
        )
        event_types = sorted(
            {
                str(events[m].get("event_type", ""))
                for m in members
                if events[m].get("event_type")
            }
        )
        out.append(
            {
                "participant": participant,
                "source_video": source_video,
                "detector": detector,
                "region": regions[0] if len(regions) == 1 else "",
                "regions": regions,
                "event_types": event_types,
                "start": cs,
                "end": ce,
                "member_event_ids": [events[m].get("id", "") for m in members],
            }
        )
    out.sort(
        key=lambda c: (c["participant"], c["source_video"], c["detector"], c["start"])
    )
    return out


def _build_clusters_from_transcript_segments(
    rows: list[tuple[str, dict[str, Any], list[dict[str, Any]]]],
    transcripts_manifest: dict[str, Any],
    *,
    gap: float,
    pad_pre: float,
    pad_post: float,
    max_duration: float,
) -> list[dict[str, Any]]:
    """Group filtered (pid, segment, marks) by participant; cluster on segment spans.

    Returns a list of cluster dicts: ``{participant, source_video, start, end,
    text, mark_categories, member_segment_ids}``.
    """
    keyed: list[tuple[Any, tuple[float, float], int]] = []
    for idx, (pid, seg, _marks) in enumerate(rows):
        try:
            s = float(seg.get("start", 0.0))
            e = float(seg.get("end", s))
        except (TypeError, ValueError):
            continue
        e = max(e, s)
        keyed.append((pid, (s, e), idx))

    clusters_raw = _cluster_groups(
        keyed,
        gap=gap,
        pad_pre=pad_pre,
        pad_post=pad_post,
        max_duration=max_duration,
    )

    sources = transcripts_manifest.get("source_transcripts") or {}
    out: list[dict[str, Any]] = []
    for key, (cs, ce), members in clusters_raw:
        participant = key
        text = " ".join(
            str(rows[m][1].get("text", "")).strip() for m in members
        ).strip()
        mark_cats = sorted(
            {
                str(mark.get("category", ""))
                for m in members
                for mark in rows[m][2]
                if mark.get("category")
            }
        )
        seg_ids = [rows[m][1].get("id") for m in members if rows[m][1].get("id")]
        source_file = ""
        entry = sources.get(participant) or {}
        if isinstance(entry, dict):
            source_file = str(entry.get("source_file", "") or "")
        source_video = Path(source_file).name if source_file else ""
        out.append(
            {
                "participant": participant,
                "source_video": source_video,
                "start": cs,
                "end": ce,
                "text": text,
                "mark_categories": mark_cats,
                "member_segment_ids": seg_ids,
            }
        )
    out.sort(key=lambda c: (c["participant"], c["start"]))
    return out


def _truncate_for_filename(text: str, *, limit: int = 60) -> str:
    """Trim a text snippet for use in a clip description (filename-safe upstream)."""
    text = " ".join(text.split())
    if len(text) <= limit:
        return text
    return files.safe_truncate(text, limit).rstrip() + "…"


def _run_ss_clips(args: argparse.Namespace) -> None:
    """Cut clips from existing Screenspace events and append to the manifest."""
    import pipeline
    import screenspace

    manifest = screenspace.load_screenspace_manifest()
    raw_events = list(manifest.get("events") or [])
    if not raw_events:
        utils.warning_print(
            "No Screenspace events found.",
            [
                (
                    "Run --screenspace (UI) or --ss-task to generate events first, "
                    "or check your input/output directory."
                )
            ],
        )
        return

    filtered = _filter_screenspace_events(
        raw_events,
        detectors=_split_csv_set(args.ss_clips_detector),
        regions=_split_csv_set(args.ss_clips_region),
        participants=_split_csv_set(args.ss_clips_participant),
        min_confidence=args.ss_clips_min_confidence,
        event_type_substr=args.ss_clips_event_type,
    )
    if not filtered:
        utils.warning_print("No Screenspace events match the given filters.")
        return

    clusters = _build_clusters_from_ss_events(
        filtered,
        gap=args.cluster_gap,
        pad_pre=args.clip_pre,
        pad_post=args.clip_post,
        max_duration=args.max_clip_duration,
    )
    if not clusters:
        utils.warning_print("No clusters produced from filtered events.")
        return

    clips_list: list[ClipRecord] = []
    last_study = ""
    for idx, cluster in enumerate(clusters):
        source_video = cluster.get("source_video") or ""
        parsed = utils.parse_source_video_name(source_video) if source_video else None
        study = parsed[0] if parsed else ""
        participant = cluster.get("participant") or ""
        detector = cluster.get("detector") or ""
        region = cluster.get("region") or ""
        desc_parts = [detector]
        if region:
            desc_parts.append(region)
        desc = " ".join(p for p in desc_parts if p).strip() or "event"
        category = f"screenspace-{detector}" if detector else "screenspace"
        clips_list.extend(
            files.build_clip_records(
                participant=participant,
                source_filename=source_video,
                time_ranges=[(cluster["start"], cluster["end"])],
                description=desc,
                category=category,
                study=study,
                cell_col=_SS_CLIPS_CELL_COL,
                cell_row_base=idx,
            )
        )
        last_study = study or last_study

    count, artifacts = pipeline.process_clips(
        clips_list, output_format="clip", include_severity=False
    )
    if artifacts:
        viewer.save_manifest(
            artifacts,
            study=last_study,
            participant="",
            worksheet_title="",
            is_excel=False,
            mode="ss-clips",
        )
    utils.info_print(
        f"Generated {count} clip(s) from {len(filtered)} event(s) "
        f"in {len(clusters)} cluster(s)."
    )


def _run_transcript_clips(args: argparse.Namespace) -> None:
    """Cut clips from transcript segments/marks and append to the manifest."""
    import pipeline

    manifest = transcripts.load_transcripts_manifest()
    if not manifest.get("source_transcripts"):
        utils.warning_print(
            "No transcripts found.",
            [
                (
                    "Run --transcribe (with a clip mode) or --pre-transcribe / --transcripts "
                    "to generate transcripts first."
                )
            ],
        )
        return

    rows = _filter_transcript_segments(
        manifest,
        participants=_split_csv_set(args.transcript_clips_participant),
        mark_categories=_split_csv_set(args.transcript_clips_mark),
        text_substr=args.transcript_clips_text,
    )
    if not rows:
        utils.warning_print("No transcript segments match the given filters.")
        return

    clusters = _build_clusters_from_transcript_segments(
        rows,
        manifest,
        gap=args.cluster_gap,
        pad_pre=args.clip_pre,
        pad_post=args.clip_post,
        max_duration=args.max_clip_duration,
    )
    if not clusters:
        utils.warning_print("No clusters produced from filtered segments.")
        return

    mark_filter = _split_csv_set(args.transcript_clips_mark)
    clips_list: list[ClipRecord] = []
    last_study = ""
    for idx, cluster in enumerate(clusters):
        source_video = cluster.get("source_video") or ""
        parsed = utils.parse_source_video_name(source_video) if source_video else None
        study = parsed[0] if parsed else ""
        participant = cluster.get("participant") or ""
        text = cluster.get("text") or ""
        desc = _truncate_for_filename(text) if text else "transcript"
        if mark_filter and cluster.get("mark_categories"):
            primary = cluster["mark_categories"][0]
            category = f"mark-{primary}"
        else:
            category = "transcript"
        clips_list.extend(
            files.build_clip_records(
                participant=participant,
                source_filename=source_video,
                time_ranges=[(cluster["start"], cluster["end"])],
                description=desc,
                category=category,
                study=study,
                cell_col=_TRANSCRIPT_CLIPS_CELL_COL,
                cell_row_base=idx,
            )
        )
        last_study = study or last_study

    count, artifacts = pipeline.process_clips(
        clips_list, output_format="clip", include_severity=False
    )
    if artifacts:
        viewer.save_manifest(
            artifacts,
            study=last_study,
            participant="",
            worksheet_title="",
            is_excel=False,
            mode="transcript-clips",
        )
    utils.info_print(
        f"Generated {count} clip(s) from {len(rows)} segment(s) "
        f"in {len(clusters)} cluster(s)."
    )


def _post_marks_to_running_server(
    segment_ids: list[str],
    category: str,
    label: str | None,
) -> str:
    """POST marks to a Transcripts server running on localhost.

    Returns ``"posted"``, ``"unreachable"`` (no server), or ``"rejected"`` (the
    server answered but refused). Routing through the API keeps the server's
    in-memory manifest in sync; a CLI-only disk write would be overwritten the
    next time the running server persists its stale state.
    """
    import json
    import urllib.error
    import urllib.request

    url = f"http://127.0.0.1:{config.SERVER_PORT}/transcripts/api/marks"
    payload = json.dumps(
        {"segment_ids": segment_ids, "category": category, "label": label}
    ).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=2.0) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError:
        return "rejected"
    except (urllib.error.URLError, OSError, ValueError):
        return "unreachable"
    if isinstance(data, dict) and data.get("ok"):
        return "posted"
    return "rejected"


def _run_transcript_mark(args: argparse.Namespace) -> None:
    """Batch-mark transcript segments whose text contains a search term.

    Mirrors the Transcripts UI's "Mark all results" action: case-insensitive
    substring match against corrected segment text, upsert into the manifest's
    ``marks`` array (no duplicates per segment). When the Transcripts server is
    running locally, POST through its API so its in-memory state stays in sync;
    otherwise write the manifest directly.
    """
    term = (args.transcript_mark or "").strip()
    if not term:
        utils.error_print("--transcript-mark requires a non-empty TERM.")
        return

    category = (args.transcript_mark_category or "").strip()
    if category not in config.MARK_CATEGORIES:
        valid = ", ".join(sorted(config.MARK_CATEGORIES.keys()))
        utils.error_print(
            "--transcript-mark-category is required and must be a known category.",
            [f"Valid categories: {valid}"],
        )
        return

    label = args.transcript_mark_label
    participants_filter = _split_csv_set(args.transcript_mark_participant)

    manifest = transcripts.load_transcripts_manifest()
    source_transcripts = manifest["source_transcripts"]
    corrections = manifest.get("corrections", [])
    marks: list[dict[str, Any]] = list(manifest.get("marks") or [])

    if not source_transcripts:
        utils.warning_print(
            "No transcripts found.",
            [
                (
                    "Run --transcribe (with a clip mode) or --pre-transcribe / --transcripts "
                    "to generate transcripts first."
                )
            ],
        )
        return

    needle = term.lower()
    matching_seg_ids: list[tuple[str, str]] = []  # (participant, segment_id)
    for pid, entry in source_transcripts.items():
        if participants_filter and pid not in participants_filter:
            continue
        raw_segments = entry.get("segments") or []
        if not raw_segments:
            continue
        corrected = transcripts.apply_corrections(raw_segments, corrections)
        for raw, seg in zip(raw_segments, corrected, strict=True):
            if needle in str(seg.get("text", "")).lower():
                seg_id = raw.get("id") or ""
                if seg_id:
                    matching_seg_ids.append((pid, seg_id))

    if not matching_seg_ids:
        utils.warning_print(f"No transcript segments contain {term!r}.")
        return

    seg_id_list = [sid for _, sid in matching_seg_ids]
    touched_participants = {pid for pid, _ in matching_seg_ids}
    total = len(seg_id_list)

    outcome = _post_marks_to_running_server(seg_id_list, category, label)
    if outcome == "posted":
        utils.info_print(
            f"Marked {total} segment(s) across {len(touched_participants)} participant(s) "
            f"via running Transcripts server."
        )
        return
    if outcome == "rejected":
        utils.error_print("The running Transcripts server rejected the marks.")
        return

    existing_by_seg = {m["segment_id"]: m for m in marks if m.get("segment_id")}
    now = datetime.now(UTC).isoformat()
    created = 0
    updated = 0
    for sid in seg_id_list:
        existing = existing_by_seg.get(sid)
        if existing is not None:
            existing["category"] = category
            if label is not None:
                existing["label"] = label
            updated += 1
        else:
            new_mark = {
                "id": f"m_{uuid.uuid4().hex[:8]}",
                "segment_id": sid,
                "category": category,
                "label": label,
                "created": now,
            }
            marks.append(new_mark)
            existing_by_seg[sid] = new_mark
            created += 1

    transcripts.save_transcripts_manifest(source_transcripts, corrections, marks=marks)
    utils.info_print(
        f"Marked {total} segment(s) across {len(touched_participants)} participant(s) "
        f"(created {created}, updated {updated})."
    )
