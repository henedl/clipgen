"""Shared remux routes for the Transcripts, Composer and Screenspace blueprints.

A recording made with OBS's "fragmented recording" option is a fragmented MP4:
ffmpeg reads it fine, but a browser cannot seek it without downloading the whole
file first (see :func:`video.probe_container_seekability`). The fix is a stream
copy into a normal container, and all three video-facing pages need to offer it,
so the job registry and the routes live here once rather than three times.

The registry is deliberately **module-level, not per-blueprint**: the pages share
one input directory, so a remux started from Composer must show as in-flight when
the user switches to Transcripts. Jobs are keyed by participant id.

This module owns no blueprint of its own — :func:`register_remux_routes` attaches
the four routes to whichever blueprint asks, mirroring
:func:`utils.register_media_route`.
"""

import threading
from pathlib import Path
from typing import Any

import files
import video
from server_utils import ApiError, err, json_endpoint, ok


# pid -> {"state", "progress", "error", "message"}. "state" is one of
# "running" / "done" / "error".
_jobs: dict[str, dict[str, Any]] = {}
_jobs_lock = threading.Lock()


def _participant_paths(sheet_context_getter: Any, pid: str) -> list[str]:
    """Existing source files for ``pid``, or [] when it has none."""
    record = files.find_participant_record(sheet_context_getter(), pid)
    if record is None:
        return []
    return [str(p) for p in record["video_paths"] if Path(p).is_file()]


def _media_state(sheet_context_getter: Any) -> tuple[dict[str, list[str]], list[str]]:
    """Per participant: kept ``.orig`` filenames, and who is currently unseekable.

    Both are read from disk on every poll rather than remembered, because the
    client's ``/api/participants`` snapshot goes stale the moment anything
    rewrites a file — a reload, a restart, or a remux started from one of the
    other two pages against the same input directory. Trusting that snapshot
    made a freshly-remuxed participant warn about a file that was already fixed.

    Cheap despite the per-poll disk work: ``browser_seekable`` comes from
    ``video.probe_container_seekability``, which is mtime-cached and measured at
    5.7 ms cold across a 16-file, 9 GB study.
    """
    kept: dict[str, list[str]] = {}
    unseekable: list[str] = []
    for participant in files.resolve_participant_videos(sheet_context_getter()):
        names = [
            video.original_backup_path(str(p)).name
            for p in participant["video_paths"]
            if video.original_backup_path(str(p)).is_file()
        ]
        if names:
            kept[participant["id"]] = names
        if participant.get("browser_seekable") is False:
            unseekable.append(participant["id"])
    return kept, unseekable


def _already_remuxed(path: str) -> bool:
    """True when an earlier run of this job already converted this part.

    Its original is parked *and* the file itself is now seekable — a state only
    a completed remux produces, since restore and discard both remove the
    backup. Multi-part participants need this: if part 1 succeeds and part 2
    fails, the retry must skip part 1 rather than hit
    ``remux_to_faststart``'s "an earlier original is still kept" guard, which
    would strand the job until the user hand-deleted the backup.
    """
    if not video.original_backup_path(path).is_file():
        return False
    probed = video.probe_container_seekability(path)
    return bool(probed and probed["browser_seekable"])


def _run_remux(pid: str, paths: list[str], token: dict[str, Any]) -> None:
    """Remux every part of one participant, publishing combined progress."""
    total = len(paths)
    failures: list[str] = []
    messages: list[str] = []

    def _publish(**fields: Any) -> None:
        # Only touch this thread's job; a dead run must not clobber its successor.
        with _jobs_lock:
            if _jobs.get(pid) is not token:
                return
            token.update(fields)

    for index, path in enumerate(paths):
        if _already_remuxed(path):
            messages.append(f"{Path(path).name}: already remuxed.")
            _publish(progress=(index + 1) / total)
            continue

        def _progress(fraction: float, index: int = index) -> None:
            _publish(progress=(index + fraction) / total)

        succeeded, message = video.remux_to_faststart(path, on_progress=_progress)
        if succeeded:
            messages.append(message)
        else:
            failures.append(f"{Path(path).name}: {message}")

    if failures:
        _publish(state="error", progress=1.0, error=" ".join(failures))
    else:
        _publish(state="done", progress=1.0, message=" ".join(messages))


def register_remux_routes(bp: Any, sheet_context_getter: Any) -> None:
    """Attach the remux start / status / discard / restore routes to ``bp``.

    ``sheet_context_getter`` is called per request (never snapshotted) so a
    spreadsheet swapped mid-session resolves against the current one.
    """

    @bp.route("/api/remux/<pid>", methods=["POST"])
    @json_endpoint
    def api_remux_start(pid: str):
        paths = _participant_paths(sheet_context_getter, pid)
        if not paths:
            raise ApiError(f"No source video found for {pid}.", 404)
        token: dict[str, Any] = {
            "state": "running",
            "progress": 0.0,
            "error": "",
            "message": "",
        }
        with _jobs_lock:
            # Check-and-set under one lock: two clicks, one ffmpeg run.
            existing = _jobs.get(pid)
            if existing is not None and existing["state"] == "running":
                return err(f"A remux of {pid} is already running.", 409)
            _jobs[pid] = token
        threading.Thread(
            target=_run_remux,
            args=(pid, paths, token),
            daemon=True,
            name=f"remux-{pid}",
        ).start()
        return ok(participant=pid, parts=len(paths))

    @bp.route("/api/remux/status")
    @json_endpoint
    def api_remux_status():
        with _jobs_lock:
            jobs = {pid: dict(job) for pid, job in _jobs.items()}
        kept, unseekable = _media_state(sheet_context_getter)
        return ok(jobs=jobs, kept=kept, unseekable=unseekable)

    @bp.route("/api/remux/<pid>/discard-original", methods=["POST"])
    @json_endpoint
    def api_remux_discard(pid: str):
        return _apply_to_parts(sheet_context_getter, pid, video.discard_remux_original)

    @bp.route("/api/remux/<pid>/restore-original", methods=["POST"])
    @json_endpoint
    def api_remux_restore(pid: str):
        return _apply_to_parts(sheet_context_getter, pid, video.restore_remux_original)


def _apply_to_parts(sheet_context_getter: Any, pid: str, action: Any):
    """Run a discard/restore over every part, failing loudly if any part fails."""
    paths = _participant_paths(sheet_context_getter, pid)
    if not paths:
        raise ApiError(f"No source video found for {pid}.", 404)
    failures = []
    applied = 0
    for path in paths:
        succeeded, message = action(path)
        if succeeded:
            applied += 1
        else:
            failures.append(f"{Path(path).name}: {message}")
    if failures and not applied:
        raise ApiError(" ".join(failures))
    # Report partial results as such; a silent half-reverted set misleads the user.
    with _jobs_lock:
        _jobs.pop(pid, None)
    return ok(applied=applied, warnings=failures)
