"""Read MindNode mind maps into clipgen's observation records.

Some research teams keep their notes in a mind map rather than a spreadsheet.
The data is the same — a study, research questions, participants, and
timestamped observations — so this module flattens the tree into the same
per-participant note records the Studio intake tabs already consume.

**File layout.** A ``.mindnode`` document is a macOS file *package* (a
directory), not a single file. Its ``contents.xml`` is a **binary plist despite
the extension**, and a plain nested one rather than an NSKeyedArchiver graph, so
``plistlib.load`` opens it with no third-party dependency::

    <bundle>/contents.xml            document (this module's only input)
    <bundle>/viewState.plist         window/zoom state, no document data
    <bundle>/QuickLook/Preview.jpg   pre-rendered map thumbnail
    <bundle>/resources/              attachments
    <bundle>/style.mindnodestyle/    theme

The document is ``canvas.mindMaps[*].mainNode``, each node recursing through
``subnodes``. ``mindMaps`` is a *list* — one document may hold several detached
trees. Node keys unused by a map are absent rather than null (a map with no
notes/tags/tasks simply has no such keys), so everything is read with ``.get()``
and unknown keys are ignored. ``crossConnections`` — MindNode's non-tree arrows
between nodes — are deliberately not followed; only ``subnodes`` is parentage.

**Tree → record mapping.** The participant level is auto-detected rather than
fixed at a depth, so maps of different shapes work::

    clipgen-test          → study
    └── Question 1        → category (every level above the participant,
        └── Question 1a           joined with " / ")
            └── P01       → participant (first node matching ^[PG]\\d+$)
                └── Note 3 0:01:00   → desc "Note 3", times 0:01:00-0:02:00

Timestamps are parsed by the shared :mod:`utils` helpers, so ``MM:SS``,
``H:MM:SS``, ranges, multi-pair separators and ``!key`` annotations behave
exactly as they do in a spreadsheet cell. Notes carrying no timestamp are kept
in the output with an empty ``times`` — they cannot be cut, but dropping them
silently would hide work the researcher did.
"""

from __future__ import annotations

import html
import plistlib
import re
from pathlib import Path
from xml.parsers import expat
from typing import Any

import config
import utils


CONTENTS_FILENAME = "contents.xml"
PREVIEW_RELPATH = "QuickLook/Preview.jpg"

# Titles are HTML fragments styled by MindNode, e.g. "<p style='…'>P01</p>".
_BR_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)
_PARA_BREAK_RE = re.compile(r"</p>\s*<p[^>]*>", re.IGNORECASE)
_TAG_RE = re.compile(r"<[^>]+>")


def _participant_re() -> re.Pattern[str]:
    """Match a participant node title, e.g. ``P01`` / ``G02``."""
    prefixes = "".join(re.escape(p) for p in config.PARTICIPANT_PREFIXES)
    return re.compile(rf"^[{prefixes}]\d+$", re.IGNORECASE)


def node_text(node: dict[str, Any]) -> str:
    """Return a node's title as plain text.

    Strips the HTML wrapper MindNode stores titles in, mapping ``<br>`` and
    paragraph breaks to newlines so a multi-line node keeps its shape.
    """
    raw = str((node.get("title") or {}).get("text") or "")
    if not raw:
        return ""
    raw = _BR_RE.sub("\n", raw)
    raw = _PARA_BREAK_RE.sub("\n", raw)
    return html.unescape(_TAG_RE.sub("", raw)).strip()


def _is_timestamp_token(token: str) -> bool:
    """Whether a whitespace-split token is a timestamp rather than description.

    Mirrors the accept rules of ``utils._parse_single_timestamp_token`` so the
    description is exactly the text that did *not* become a time. Kept separate
    only because the public ``utils.parse_timestamps`` warns about every token
    it rejects, and here rejection is the normal case — most tokens are words.
    ``tests/test_mindnode.py`` asserts the two agree token for token.
    """
    cleaned = token.strip().lower().rstrip(",").rstrip("-").replace(".", ":")
    if not cleaned:
        return False
    # An ignored placeholder ("x") is not description text either.
    if cleaned in utils.get_ignored_timestamp_tokens():
        return True
    dash = cleaned.find("-")
    if dash > 0 and cleaned[dash - 1].isdigit():
        head, tail = cleaned[:dash], cleaned[dash + 1 :]
        return (
            utils.timestamp_to_seconds(head) is not None
            and utils.timestamp_to_seconds(tail) is not None
        )
    return utils.timestamp_to_seconds(cleaned) is not None


def _is_timestamp_run(token: str) -> bool:
    """Whether a whitespace-split token is *entirely* timestamps.

    ``utils._split_timestamp_tokens`` also splits on ``+``, ``;`` and ``,``, so
    ``0:01:00+0:05:00`` parses as two pairs — but a whitespace-only split
    leaves it as one token that ``_is_timestamp_token`` rejects, and the raw
    times then survive into the description and the output filename. Splitting
    on the separators *within* a token rather than splitting the whole string
    on them keeps prose punctuation intact: "obs, 1:00" still describes as
    "obs," because "obs," is not a run of timestamps.
    """
    parts = [p for p in re.split(r"[+;,]", token) if p]
    if len(parts) < 2:
        return False
    return all(_is_timestamp_token(p) for p in parts)


def _describe(text: str) -> str:
    """Strip timestamp and annotation tokens, leaving the observation text."""
    known_annotations = set(utils.get_known_annotation_map().keys())
    kept = [
        tok
        for tok in text.split()
        if tok.lower() not in known_annotations
        and not _is_timestamp_token(tok)
        and not _is_timestamp_run(tok)
    ]
    return " ".join(kept).strip()


def load_document(path: str | Path) -> list[dict[str, Any]]:
    """Return the root nodes of a ``.mindnode`` bundle.

    Raises:
        ValueError: the path is not a MindNode bundle, or its ``contents.xml``
            is unreadable or not shaped like a mind map.
    """
    bundle = Path(path)
    contents = bundle / CONTENTS_FILENAME
    if not contents.is_file():
        raise ValueError(
            f"Not a MindNode document: {bundle} (no {CONTENTS_FILENAME} inside)"
        )
    try:
        with contents.open("rb") as fh:
            data = plistlib.load(fh)
    except (
        OSError,
        plistlib.InvalidFileException,
        ValueError,
        # A truncated XML-format .mindnode raises ExpatError, which is not a ValueError.
        expat.ExpatError,
    ) as exc:
        raise ValueError(f"Could not read {contents}: {exc}") from exc

    canvas = data.get("canvas") if isinstance(data, dict) else None
    maps = canvas.get("mindMaps") if isinstance(canvas, dict) else None
    roots: list[dict[str, Any]] = []
    for entry in maps if isinstance(maps, list) else []:
        main = entry.get("mainNode") if isinstance(entry, dict) else None
        if isinstance(main, dict):
            roots.append(main)
    if not roots:
        raise ValueError(f"No mind maps found in {contents}")
    return roots


def _subnodes(node: dict[str, Any]) -> list[dict[str, Any]]:
    """Child nodes, ignoring anything that is not a node dict."""
    raw = node.get("subnodes")
    if not isinstance(raw, list):
        return []
    return [child for child in raw if isinstance(child, dict)]


def _make_note(
    node: dict[str, Any], *, study: str, participant: str, path: list[str], seq: int
) -> dict[str, Any]:
    """Build one note record from a leaf (or timestamped) node."""
    text = node_text(node)
    cleaned, segment_annotations, cell_annotations = utils.parse_cell_annotations(text)
    times = utils.parse_timestamps(cleaned)
    spans: list[tuple[float, float]] = []
    for start, end in times:
        start_s = utils.timestamp_to_seconds(start)
        end_s = utils.timestamp_to_seconds(end)
        if start_s is not None and end_s is not None:
            spans.append((start_s, end_s))
    node_id = str(node.get("nodeID") or "") or f"{participant}-{seq}"
    return {
        "id": node_id,
        "study": study,
        "participant": participant,
        "category": " / ".join(path),
        "category_path": list(path),
        "desc": _describe(text),
        "text": text,
        "times": times,
        "spans": spans,
        "annotations": sorted(cell_annotations),
        "segment_annotations": {k: sorted(v) for k, v in segment_annotations.items()},
    }


def _collect_notes(
    node: dict[str, Any],
    *,
    study: str,
    participant: str,
    path: list[str],
    notes: list[dict[str, Any]],
) -> None:
    """Emit a note for every leaf, and for any branch that carries a timestamp.

    A branch node with no timestamp is treated as grouping, not as an
    observation, so its text is not emitted twice alongside its children's.
    """
    for child in _subnodes(node):
        grandchildren = _subnodes(child)
        note = _make_note(
            child, study=study, participant=participant, path=path, seq=len(notes)
        )
        if not grandchildren or note["times"]:
            notes.append(note)
        if grandchildren:
            _collect_notes(
                child, study=study, participant=participant, path=path, notes=notes
            )


def _walk(
    node: dict[str, Any],
    *,
    study: str,
    path: list[str],
    notes: list[dict[str, Any]],
    participant_re: re.Pattern[str],
) -> None:
    """Descend until a participant node is found, then collect its notes."""
    for child in _subnodes(node):
        text = node_text(child)
        label = utils.normalize_participant_id(text)
        if participant_re.match(label):
            _collect_notes(
                child, study=study, participant=label.upper(), path=path, notes=notes
            )
            continue
        _walk(
            child,
            study=study,
            path=path + [text] if text else path,
            notes=notes,
            participant_re=participant_re,
        )


def parse_document(path: str | Path) -> dict[str, Any]:
    """Flatten a ``.mindnode`` bundle into per-participant note records.

    Returns a dict with the document's ``study`` (from the first root's title,
    normalized for filesystem use), its ``notes``, and counts of how many of
    those notes carry a usable timestamp. Every root in the document is walked;
    a note keeps the study of the root it came from.

    Raises:
        ValueError: propagated from :func:`load_document`.
    """
    bundle = Path(path)
    roots = load_document(bundle)
    participant_re = _participant_re()

    notes: list[dict[str, Any]] = []
    root_titles: list[str] = []
    for root in roots:
        title = node_text(root)
        root_titles.append(title)
        _walk(
            root,
            study=utils.normalize_study_name(title),
            path=[],
            notes=notes,
            participant_re=participant_re,
        )

    with_times = sum(1 for n in notes if n["times"])
    participants = sorted({n["participant"] for n in notes})
    return {
        "path": str(bundle),
        "name": bundle.name,
        "study": utils.normalize_study_name(root_titles[0]) if root_titles else "",
        "roots": root_titles,
        "participants": participants,
        "notes": notes,
        "with_times": with_times,
        "without_times": len(notes) - with_times,
    }


def find_documents(directory: str | Path) -> list[dict[str, Any]]:
    """List ``.mindnode`` bundles in *directory* for the Start overlay picker."""
    root = Path(directory)
    if not root.is_dir():
        return []
    found: list[dict[str, Any]] = []
    for entry in sorted(root.glob("*.mindnode")):
        if not (entry / CONTENTS_FILENAME).is_file():
            continue
        try:
            modified = entry.stat().st_mtime
        except OSError:
            modified = 0.0
        found.append(
            {
                "path": str(entry),
                "name": entry.name,
                "modified": modified,
                "has_preview": (entry / PREVIEW_RELPATH).is_file(),
            }
        )
    return found
