"""Central generators for benchmark fixtures: sheets, videos, manifests.

Every benchmark (``scan_bench``, ``clip_bench``, ``tests/ui/ui_bench.py``)
and the profile skill's recipes build their inputs here, so the geometry a
page is measured against is written once. The workbook layout mirrors
``tests/ui/_ui_fixtures._make_workbook`` — ``ID`` at F2 with participant
columns to its right *on row 2*, ``Observation``/``Category`` on row 5, data
from row 6 — because ``spreadsheet.build_sheet_context`` scans for exactly
those labels and a drifted layout yields a silently small grid, not an error.
The manifest sections copy the shapes ``_ui_fixtures`` seeds, scaled up.

Pure: no clipgen imports, so the ``tests/ui`` harness and ``tests/`` unit
tests can load it with ``importlib`` without the source tree on ``sys.path``.
"""

from __future__ import annotations

import json
import subprocess
from collections.abc import Callable
from pathlib import Path
from typing import Any

SEVERITIES = ("Critical", "Serious", "Moderate", "Minor")
HEADERS = ("Count", "Reported", "Severity", "Category", "Observation", "Summary")


def write_sheet(
    path: Path,
    *,
    study: str,
    rows: int,
    participants: int,
    cell: Callable[[int, int], str] | None = None,
) -> Path:
    """Write an ``.xlsx`` with *rows* observations across *participants*.

    *cell(row_index, participant_index)* returns the timestamp text for that
    cell; the default fills every third participant column with ``0:01-0:04``
    (the gridbench recipe). Severities cycle so ``renderGrid`` paints its
    ``.sev-*`` classes — an empty Severity column under-measures it.
    """
    import openpyxl

    def default_cell(_row: int, participant: int) -> str:
        return "0:01-0:04" if participant % 3 == 0 else ""

    fill = cell or default_cell
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Observations"
    ws["A1"] = study
    ws["F2"] = "ID"
    for i in range(participants):
        ws.cell(2, 7 + i, f"P{i + 1:02d}")
    for col, header in enumerate(HEADERS, 1):
        ws.cell(5, col, header)
    for r in range(rows):
        ws.cell(6 + r, 3, SEVERITIES[r % len(SEVERITIES)])
        ws.cell(6 + r, 4, "Onboarding")
        ws.cell(6 + r, 5, f"Observation {r}")
        for i in range(participants):
            ws.cell(6 + r, 7 + i, fill(r, i))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()
    return path


def make_testsrc_video(
    path: Path,
    *,
    duration: int,
    size: str = "1280x720",
    rate: int = 30,
    audio: bool = False,
    fragmented: bool = False,
) -> Path:
    """Encode a deterministic testsrc clip (constant motion defeats phash-skip).

    *audio* adds a sine track (Whisper and the clip pipeline need one);
    *fragmented* writes the OBS-style MP4 browsers cannot seek.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        "ffmpeg",
        "-y",
        "-loglevel",
        "error",
        "-f",
        "lavfi",
        "-i",
        f"testsrc=duration={duration}:size={size}:rate={rate}",
    ]
    if audio:
        cmd += ["-f", "lavfi", "-i", f"sine=frequency=220:duration={duration}"]
    cmd += ["-pix_fmt", "yuv420p", "-c:v", "libx264", "-g", str(rate)]
    if audio:
        cmd += ["-c:a", "aac", "-shortest"]
    if fragmented:
        cmd += ["-movflags", "frag_keyframe+empty_moov+default_base_moof"]
    cmd.append(str(path))
    subprocess.run(cmd, check=True)
    return path


def transcripts_section(participants: dict[str, int]) -> dict[str, Any]:
    """The ``transcripts`` manifest section with N synthetic segments per id."""
    source: dict[str, Any] = {}
    for pid, n in participants.items():
        source[pid] = {
            "segments": [
                {
                    "id": f"{pid}:{i}",
                    "start": i * 3.0,
                    "end": i * 3.0 + 2.8,
                    "text": f"Synthetic segment {i} for render benchmarking.",
                }
                for i in range(n)
            ],
            "language": "en",
            "model": "synthetic",
        }
    return {"source_transcripts": source, "corrections": [], "marks": []}


def screenspace_section(
    *,
    study: str,
    participant: str,
    video_path: Path,
    events: int,
    duration: float,
    source_size: tuple[int, int] = (1280, 720),
) -> dict[str, Any]:
    """A completed ``change`` task with *events* results and matching events.

    Completed only: an active task opens an SSE stream, which defeats every
    settle heuristic (see ``_ui_fixtures._seed_screenspace``).
    """
    width, height = source_size
    task_id = "bench-task-1"
    step = duration / max(events, 1)
    results = [
        {"timestamp": round(i * step, 3), "magnitude": 0.3 + (i % 7) / 10}
        for i in range(events)
    ]
    event_rows = [
        {
            "id": f"bench-evt-{i}",
            "source_video": video_path.name,
            "participant": participant,
            "detector": "change",
            "event_type": "change",
            "time_in": round(i * step, 3),
            "time_out": round(min(i * step + step * 0.8, duration), 3),
            "confidence": round(0.3 + (i % 7) / 10, 2),
            "metadata": {},
            "excluded": False,
            "task_id": task_id,
            "region": "bench",
        }
        for i in range(events)
    ]
    return {
        "regions": {
            "bench": {
                "x": 0.05,
                "y": 0.05,
                "w": 0.40,
                "h": 0.20,
                "source_width": width,
                "source_height": height,
            }
        },
        "tasks": [
            {
                "id": task_id,
                "type": "change",
                "name": "Change · bench",
                "participant": participant,
                "source_video": video_path.name,
                "video_paths": [str(video_path)],
                "region": "bench",
                "region_coords": {
                    "x": int(width * 0.05),
                    "y": int(height * 0.05),
                    "w": int(width * 0.40),
                    "h": int(height * 0.20),
                },
                "parameters": {},
                "status": "completed",
                "progress": 1.0,
                "priority": 100,
                "result": results,
                "error": None,
                "created_at": "2026-01-05T12:00:00+00:00",
                "completed_at": "2026-01-05T12:01:00+00:00",
            }
        ],
        "events": event_rows,
        "stashes": [],
        "per_participant": {},
        "pins": {},
    }


def write_manifest(out_dir: Path, sections: dict[str, Any]) -> Path:
    """Merge *sections* into ``out_dir/clipgen.json`` by absolute path."""
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / "clipgen.json"
    doc: dict[str, Any] = {}
    if path.is_file():
        doc = json.loads(path.read_text(encoding="utf-8"))
    doc.update(sections)
    path.write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def manifest_counts(out_dir: Path) -> dict[str, Any]:
    """What a manifest holds: segments per participant, events, task results."""
    path = out_dir / "clipgen.json"
    if not path.is_file():
        return {"segments": {}, "events": 0, "results": 0}
    doc = json.loads(path.read_text(encoding="utf-8"))
    ts = doc.get("transcripts", {}).get("source_transcripts", {})
    ss = doc.get("screenspace", {})
    return {
        "segments": {pid: len(entry.get("segments", [])) for pid, entry in ts.items()},
        "events": len(ss.get("events", [])),
        "results": sum(len(t.get("result") or []) for t in ss.get("tasks", [])),
    }


def sheet_rows(path: Path) -> int:
    """Observation rows in a sheet written by :func:`write_sheet`."""
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb["Observations"]
    n = sum(1 for row in ws.iter_rows(min_row=6, min_col=5, max_col=5) if row[0].value)
    wb.close()
    return n
