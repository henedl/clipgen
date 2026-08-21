#!/usr/bin/env python3
"""Render a GitHub Release body for a version tag.

Two sections. **Highlights** comes from the curated CHANGELOG.md, grouped by tool
and written for someone using clipgen. **What's Changed** comes from the git commits
in (previous tag, tag], grouped by conventional-commit type — which works here only
because every PR title in this repo follows `type(scope): description` and GitHub
appends `(#NNN)` on squash-merge.

Run locally:

    uv run build/release_notes.py --tag v0.16.0 --prev-tag v0.15.1

Invoked by .github/workflows/release-notes.yml, which checks out the *default
branch*, never the tag, and feeds this script the tag through git plumbing. Two
reasons, both load-bearing:

  * This script does not exist at v0.15.1 or v0.16.0, so a tag checkout would have
    nothing to run and the two existing releases could never be backfilled.
  * CHANGELOG.md is rewritten in place. At v0.15.1 it is still in the obsolete
    three-part heading format (`## <version> — <date> — <tool>`) and stops at
    v0.14.74, so the tag's copy would yield empty highlights *and* mis-parse every
    heading it did find.

Stdlib-only apart from source/changelog.py, whose CHANGELOG parser is shared so the
heading regex — which has silently mis-parsed the format once already — lives in
exactly one place. That import is stdlib-clean (source/utils.py guards `rich` behind
try/except and imports gspread lazily) and tests/test_release_notes.py pins the
property so a future heavy import fails in the suite rather than on a tag push.

Commit subjects are emitted raw, not markdown-escaped. A subject with paired `_` or
`*` would italicise, but no current subject does, GitHub's own generated notes do not
escape either, and escaping would visibly backslash legitimate underscores in module
names.
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import NamedTuple

_ROOT = Path(__file__).resolve().parents[1]

# The product modules live in `source/`, not the repo root — same insert as
# tests/conftest.py. Only `changelog` is imported, and only for its parser.
sys.path.insert(0, str(_ROOT / "source"))

import changelog


# CHANGELOG.md's own preamble order. Fixed rather than frequency-sorted, so two
# releases' Highlights sections are comparable line for line.
TOOL_ORDER = (
    "Core",
    "Studio",
    "Screenspace",
    "Transcripts",
    "Workflows",
    "Composer",
    "Overview",
)

# Only these types reach the release body, in this order.
SECTIONS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("Features", ("feat",)),
    ("Bug Fixes", ("fix",)),
    ("Performance", ("perf",)),
)

# Real work, but nobody opens a release page to read it. Counted in a footnote
# rather than dropped in silence, so the omission is visible.
OMITTED_TYPES = ("refactor", "build", "docs", "test", "ci", "chore", "style")
OMITTED_LABEL = "refactors, packaging, docs, tests, CI"

# A subject that is not a conventional commit at all lands here instead of being
# thrown away — dropping it would be a silent loss with no rule behind it.
OTHER_SECTION = "Other Changes"

_SUBJECT_RE = re.compile(
    r"^(?P<type>[a-z]+)(?:\((?P<scope>[^)]*)\))?(?P<bang>!)?:\s+(?P<desc>.+)$"
)
_PR_RE = re.compile(r"\s*\(#(?P<pr>\d+)\)\s*$")

# Commit bodies contain newlines, so records need their own framing.
_GIT_SEP = "\x1f"
_GIT_END = "\x1e"

MAX_HIGHLIGHT_VERSIONS = 10  # only applied when there is no previous tag
MAX_BODY_CHARS = 120_000  # GitHub's release-body limit is 125,000


class Commit(NamedTuple):
    """One parsed commit subject. ``type`` is empty when it is not conventional."""

    sha: str
    type: str
    scope: str
    breaking: bool
    description: str  # the `(#NNN)` suffix is already stripped
    pr: str  # empty when the commit was pushed directly


def _warn(message: str) -> None:
    """Annotate the Actions run without failing it. stderr — stdout may be the body."""
    print(f"::warning::{message}", file=sys.stderr)


# --------------------------------------------------------------------------- #
# Versions
# --------------------------------------------------------------------------- #


def parse_version(text: str) -> tuple[int, int, int, bool, str] | None:
    """``v0.15.18`` -> ``(0, 15, 18, True, "")``; unparseable -> ``None``.

    The numeric tuple is the entire point. Compared as strings, ``"0.15.18"`` sorts
    *below* ``"0.15.4"``, which would silently drop ten of the fourteen versions in
    the v0.16.0 highlight range. The fourth element is ``not prerelease`` so that
    ``v0.17.0-rc.1`` sorts below ``v0.17.0``; the fifth orders prereleases among
    themselves.
    """
    core = text.strip().lstrip("vV")
    core, _, pre = core.partition("-")
    parts = core.split(".")
    if len(parts) != 3:
        return None
    try:
        major, minor, patch = (int(p) for p in parts)
    except ValueError:
        return None
    return (major, minor, patch, not pre, pre)


# --------------------------------------------------------------------------- #
# Highlights (CHANGELOG.md)
# --------------------------------------------------------------------------- #


def select_highlights(
    entries: list[dict],
    tag: str,
    prev_tag: str,
    max_versions: int = MAX_HIGHLIGHT_VERSIONS,
) -> list[dict]:
    """Releases in ``(prev_tag, tag]``, in file order (newest first).

    File order is preserved deliberately — source/changelog.py does not re-sort by
    version and neither should this. ``max_versions`` only applies when there is no
    previous tag, where the range would otherwise reach the start of history.
    """
    hi = parse_version(tag)
    if hi is None:
        _warn(f"tag {tag} is not a version; skipping Highlights")
        return []
    lo = parse_version(prev_tag) if prev_tag else None
    if prev_tag and lo is None:
        # An explicit --prev-tag of a SHA is legitimate; it just cannot bound the
        # changelog, so fall back to the no-previous-tag cap.
        _warn(f"previous tag {prev_tag} is not a version; capping Highlights instead")

    selected: list[dict] = []
    for entry in entries:
        if not entry["changes"]:
            continue
        version = parse_version(entry["version"])
        if version is None:
            _warn(f"CHANGELOG heading {entry['version']!r} is not a version; skipped")
            continue
        if version > hi:
            continue
        if lo is not None and version <= lo:
            continue
        selected.append(entry)

    if lo is None:
        return selected[:max_versions]
    return selected


def render_highlights(entries: list[dict]) -> str:
    """Group every change across *entries* by tool. Empty input renders as ``""``.

    Grouping by tool rather than by version is deliberate: the versions between two
    tags are internal build/VERSION bumps that were never released as downloads, so a
    `### v0.15.9` heading would invite the reader to look for a release that does not
    exist. Tool answers the question a release page is actually read for.
    """
    by_tool: dict[str, dict[str, list[str]]] = {}
    for entry in entries:
        for change in entry["changes"]:
            kinds = by_tool.setdefault(change["tool"], {"Feat": [], "Fix": []})
            kinds.setdefault(change["kind"] or "Feat", []).append(change["text"])

    # A tool the CHANGELOG grows later must not vanish just because it is missing
    # from TOOL_ORDER.
    ordered = [t for t in TOOL_ORDER if t in by_tool]
    ordered += sorted(t for t in by_tool if t not in TOOL_ORDER)

    lines: list[str] = []
    for tool in ordered:
        kinds = by_tool[tool]
        bullets = [f"- New: {t}" for t in kinds.get("Feat", [])]
        bullets += [f"- Fixed: {t}" for t in kinds.get("Fix", [])]
        bullets += [
            f"- {t}"
            for kind, texts in kinds.items()
            if kind not in ("Feat", "Fix")
            for t in texts
        ]
        if not bullets:
            continue
        lines.append(f"### {tool}")
        lines.extend(bullets)
        lines.append("")

    if not lines:
        return ""
    return "\n".join(["## Highlights", ""] + lines).rstrip() + "\n"


# --------------------------------------------------------------------------- #
# What's Changed (git)
# --------------------------------------------------------------------------- #


def parse_commit(record: str) -> Commit:
    """Parse one ``sha\\x1fsubject\\x1fbody`` record."""
    sha, _, rest = record.partition(_GIT_SEP)
    subject, _, body = rest.partition(_GIT_SEP)
    subject = subject.strip()

    # Strip the PR suffix *before* matching, so `(#746)` can never survive into the
    # description and get printed a second time alongside the parsed number.
    pr = ""
    pr_match = _PR_RE.search(subject)
    if pr_match is not None:
        pr = pr_match.group("pr")
        subject = subject[: pr_match.start()].rstrip()

    match = _SUBJECT_RE.match(subject)
    if match is None:
        return Commit(sha.strip(), "", "", "BREAKING CHANGE" in body, subject, pr)
    return Commit(
        sha.strip(),
        match.group("type"),
        match.group("scope") or "",
        bool(match.group("bang")) or "BREAKING CHANGE" in body,
        match.group("desc").strip(),
        pr,
    )


def group_commits(commits: list[Commit]) -> tuple[dict[str, list[Commit]], int]:
    """Return ``({section: commits}, omitted_count)`` with empty sections dropped."""
    grouped: dict[str, list[Commit]] = {}
    omitted = 0
    for commit in commits:
        # A breaking change is never routine, whatever its type — it is listed
        # above, so counting it as omitted here would contradict that.
        if commit.type in OMITTED_TYPES and not commit.breaking:
            omitted += 1
            continue
        title = next(
            (t for t, types in SECTIONS if commit.type in types), OTHER_SECTION
        )
        grouped.setdefault(title, []).append(commit)

    order = [t for t, _ in SECTIONS] + [OTHER_SECTION]
    return {t: grouped[t] for t in order if t in grouped}, omitted


def render_commit(commit: Commit) -> str:
    """``- **scope:** description (#NNN)``, falling back to the short SHA."""
    text = (
        f"**{commit.scope}:** {commit.description}"
        if commit.scope
        else commit.description
    )
    ref = f"(#{commit.pr})" if commit.pr else f"({commit.sha})" if commit.sha else ""
    return f"- {text} {ref}".rstrip()


def render_changes(commits: list[Commit]) -> str:
    grouped, omitted = group_commits(commits)
    if not grouped and not omitted:
        return ""

    lines = ["## What's Changed", ""]

    # Breaking changes lead, and repeat below in their own section: flagged once,
    # in context once.
    breaking = [c for c in commits if c.breaking]
    if breaking:
        lines.append("### ⚠️ Breaking Changes")
        lines.extend(render_commit(c) for c in breaking)
        lines.append("")

    for title, section in grouped.items():
        lines.append(f"### {title}")
        lines.extend(render_commit(c) for c in section)
        lines.append("")

    if omitted:
        noun, verb = ("change", "is") if omitted == 1 else ("changes", "are")
        lines.append(f"_{omitted} routine {noun} ({OMITTED_LABEL}) {verb} omitted._")
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


# --------------------------------------------------------------------------- #
# Body
# --------------------------------------------------------------------------- #


def render_body(
    entries: list[dict],
    commits: list[Commit],
    tag: str,
    prev_tag: str,
    repo: str,
    initial_release: bool = False,
) -> str:
    """Assemble the full release body. Empty sections are omitted entirely."""
    blocks: list[str] = []

    highlights = render_highlights(entries)
    if highlights:
        blocks.append(highlights)
    else:
        _warn(f"no CHANGELOG entries in range for {tag}; Highlights omitted")

    commits_url = f"https://github.com/{repo}/commits/{tag}"
    if initial_release:
        # The first tag has no range to diff against — its history reaches the very
        # first commit, which is a thousand subjects nobody will read.
        blocks.append(
            "## What's Changed\n\n"
            "_First tagged release. The full commit history is linked below._\n"
        )
    else:
        changes = render_changes(commits)
        if changes:
            blocks.append(changes)

    if prev_tag and not initial_release:
        compare = f"https://github.com/{repo}/compare/{prev_tag}...{tag}"
        blocks.append(f"**Full Changelog**: {compare}\n")
    else:
        blocks.append(f"**Full Changelog**: {commits_url}\n")

    return "\n".join(blocks)


# --------------------------------------------------------------------------- #
# I/O edges — only these three touch the world
# --------------------------------------------------------------------------- #


def _git(*args: str) -> str:
    return subprocess.run(
        ["git", *args],
        cwd=_ROOT,
        check=True,
        text=True,
        capture_output=True,
    ).stdout


def previous_tag(tag: str) -> str:
    """The nearest tag before *tag*, or ``""`` when *tag* is the first one."""
    try:
        return _git("describe", "--tags", "--abbrev=0", f"{tag}^").strip()
    except subprocess.CalledProcessError:
        return ""


def read_commits(tag: str, prev_tag: str) -> list[Commit]:
    rev_range = f"{prev_tag}..{tag}" if prev_tag else tag
    # --no-merges costs one flag and guards against `Merge branch` noise if anyone
    # ever stops squash-merging.
    out = _git(
        "log",
        "--no-merges",
        f"--pretty=format:%h{_GIT_SEP}%s{_GIT_SEP}%b{_GIT_END}",
        rev_range,
    )
    return [parse_commit(r) for r in out.split(_GIT_END) if r.strip()]


def _default_repo() -> str:
    repo = os.environ.get("GITHUB_REPOSITORY", "").strip()
    if repo:
        return repo
    try:
        url = _git("remote", "get-url", "origin").strip()
    except subprocess.CalledProcessError:
        return "OWNER/REPO"
    return re.sub(r"^.*[:/]([^/]+/[^/]+?)(?:\.git)?$", r"\1", url)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--tag", required=True, help="version tag, e.g. v0.16.0")
    parser.add_argument(
        "--prev-tag",
        default=None,
        help="range start (any git rev). Default: the nearest earlier tag.",
    )
    parser.add_argument("--repo", default=None, help="OWNER/NAME for the links")
    parser.add_argument("--changelog", default=None, type=Path)
    parser.add_argument("--output", default="-", help="'-' for stdout")
    parser.add_argument(
        "--max-highlight-versions", type=int, default=MAX_HIGHLIGHT_VERSIONS
    )
    parser.add_argument(
        "--initial-release",
        action="store_true",
        help="force the first-release shape (implied when there is no earlier tag)",
    )
    args = parser.parse_args()

    prev = args.prev_tag if args.prev_tag is not None else previous_tag(args.tag)
    repo = args.repo or _default_repo()
    path = args.changelog or (_ROOT / "CHANGELOG.md")

    # No earlier tag means the range reaches the first commit ever, so the first
    # release takes the short shape without the workflow having to know that.
    initial = args.initial_release or not prev

    text = path.read_text(encoding="utf-8") if path.is_file() else ""
    if not text:
        _warn(f"no changelog at {path}")
    entries = select_highlights(
        changelog.parse_entries(text), args.tag, prev, args.max_highlight_versions
    )

    commits = [] if initial else read_commits(args.tag, prev)
    body = render_body(entries, commits, args.tag, prev, repo, initial_release=initial)

    if len(body) > MAX_BODY_CHARS:
        _warn(f"body is {len(body)} chars; truncated to {MAX_BODY_CHARS}")
        body = body[:MAX_BODY_CHARS].rsplit("\n", 1)[0] + "\n\n_Truncated._\n"

    if args.output == "-":
        sys.stdout.write(body)
    else:
        Path(args.output).write_text(body, encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
