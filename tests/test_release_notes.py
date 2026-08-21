"""Guard the GitHub Release body renderer in ``build/release_notes.py``.

Every failure this file catches is silent in production: the release still
publishes, it just says the wrong thing, and nobody reads a release page closely
enough to notice a section that quietly went missing. So the assertions here are
mostly about *absence* — a version that should have been in range, a PR number
printed twice, a commit type dropped without a trace.

All of it runs against pure functions. No network, no ``gh``, no git.
"""

import importlib.util
import subprocess
import sys
from pathlib import Path
from types import ModuleType

_ROOT = Path(__file__).resolve().parent.parent


def _load() -> ModuleType:
    """``build/`` is in ``norecursedirs`` and not on ``sys.path`` — load by path.

    Same trick as ``test_packaging.py::_fetch_binaries_module``.
    """
    spec = importlib.util.spec_from_file_location(
        "release_notes", _ROOT / "build" / "release_notes.py"
    )
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


rn = _load()


def _changelog(*versions: str) -> str:
    """A synthetic changelog with one Feat line per version."""
    return "\n".join(
        f"## {v} — 2026-01-01\n**Core:** Feat: did {v}\n" for v in versions
    )


def _entries(*versions: str) -> list[dict]:
    import changelog

    return changelog.parse_entries(_changelog(*versions))


def _commits(*subjects: str) -> list:
    return [rn.parse_commit(f"abc1234{rn._GIT_SEP}{s}{rn._GIT_SEP}") for s in subjects]


# --------------------------------------------------------------------------- #
# Versions
# --------------------------------------------------------------------------- #


def test_patch_numbers_compare_numerically_not_as_strings() -> None:
    """The reason parse_version exists at all.

    Lexically ``"v0.15.18" < "v0.15.4"``, and comparing that way drops ten of the
    fourteen versions in the real v0.16.0 highlight range.
    """
    naive = sorted(["v0.15.18", "v0.15.4"])
    assert naive[0] == "v0.15.18", "the trap: lexical order puts .18 below .4"
    assert rn.parse_version("v0.15.18") > rn.parse_version("v0.15.4")


def test_prerelease_sorts_below_its_release() -> None:
    assert rn.parse_version("v0.17.0-rc.1") < rn.parse_version("v0.17.0")
    assert rn.parse_version("v0.17.0-rc.1") < rn.parse_version("v0.17.0-rc.2")


def test_unparseable_version_returns_none() -> None:
    for text in ("v1.2", "v1.2.3.4", "unreleased", "", "v1.x.0"):
        assert rn.parse_version(text) is None, text


# --------------------------------------------------------------------------- #
# Highlight selection
# --------------------------------------------------------------------------- #


def test_range_excludes_the_previous_tag_and_includes_the_tag() -> None:
    entries = _entries("v1.0.1", "v1.0.0", "v0.9.2", "v0.9.1", "v0.9.0")
    picked = [e["version"] for e in rn.select_highlights(entries, "v1.0.0", "v0.9.1")]
    assert picked == ["v1.0.0", "v0.9.2"]


def test_range_includes_double_digit_patches() -> None:
    """The end-to-end form of the string-compare trap."""
    entries = _entries("v0.15.18", "v0.15.16", "v0.15.4", "v0.15.1")
    picked = [e["version"] for e in rn.select_highlights(entries, "v0.16.0", "v0.15.1")]
    assert picked == ["v0.15.18", "v0.15.16", "v0.15.4"]


def test_file_order_is_preserved_not_re_sorted() -> None:
    entries = _entries("v1.0.0", "v1.0.2", "v1.0.1")
    picked = [e["version"] for e in rn.select_highlights(entries, "v1.0.2", "v0.9.0")]
    assert picked == ["v1.0.0", "v1.0.2", "v1.0.1"]


def test_first_tag_takes_everything_up_to_the_cap() -> None:
    entries = _entries(*[f"v0.1.{n}" for n in range(20, 0, -1)])
    assert len(rn.select_highlights(entries, "v0.1.20", "", max_versions=3)) == 3


def test_cap_does_not_apply_when_a_previous_tag_bounds_the_range() -> None:
    entries = _entries(*[f"v0.1.{n}" for n in range(20, 0, -1)])
    picked = rn.select_highlights(entries, "v0.1.20", "v0.1.10", max_versions=3)
    assert len(picked) == 10


def test_release_with_no_change_lines_is_skipped() -> None:
    import changelog

    text = "## v1.0.0 — 2026-01-01\n\n## v0.9.0 — 2026-01-01\n**Core:** Feat: x\n"
    entries = changelog.parse_entries(text)
    assert len(entries) == 2, (
        "parse_entries keeps empty releases; the caller drops them"
    )
    picked = rn.select_highlights(entries, "v1.0.0", "v0.8.0")
    assert [e["version"] for e in picked] == ["v0.9.0"]


def test_no_matching_entry_renders_no_highlights_section() -> None:
    body = rn.render_body([], _commits("feat: x"), "v1.0.0", "v0.9.0", "o/r")
    assert "## Highlights" not in body
    assert "### Features" in body


# --------------------------------------------------------------------------- #
# Highlight rendering
# --------------------------------------------------------------------------- #


def _highlight_entry(*changes: tuple[str, str, str]) -> list[dict]:
    return [
        {
            "version": "v1.0.0",
            "date": "2026-01-01",
            "changes": [{"tool": t, "kind": k, "text": x} for t, k, x in changes],
        }
    ]


def test_tools_render_in_canonical_order_regardless_of_file_order() -> None:
    out = rn.render_highlights(
        _highlight_entry(
            ("Composer", "Feat", "c"), ("Core", "Feat", "a"), ("Studio", "Feat", "b")
        )
    )
    assert out.index("### Core") < out.index("### Studio") < out.index("### Composer")


def test_features_precede_fixes_within_a_tool() -> None:
    out = rn.render_highlights(
        _highlight_entry(("Core", "Fix", "broke"), ("Core", "Feat", "shiny"))
    )
    assert out.index("- New: shiny") < out.index("- Fixed: broke")


def test_unknown_tool_is_appended_rather_than_dropped() -> None:
    """A tool the CHANGELOG grows later must still reach the release page."""
    out = rn.render_highlights(
        _highlight_entry(("Timeline", "Feat", "new thing"), ("Core", "Feat", "a"))
    )
    assert "### Timeline" in out and "- New: new thing" in out
    assert out.index("### Core") < out.index("### Timeline")


# --------------------------------------------------------------------------- #
# Commit parsing
# --------------------------------------------------------------------------- #


def test_subject_parses_type_scope_and_pr() -> None:
    (c,) = _commits("feat(server): user-definable filename pattern (#746)")
    assert (c.type, c.scope, c.pr) == ("feat", "server", "746")
    assert c.description == "user-definable filename pattern"
    assert c.breaking is False


def test_pr_number_is_stripped_from_the_description() -> None:
    """It arrives in the subject from squash-merge; re-rendering it must not double it."""
    (c,) = _commits("feat(web): tab the right column (#710)")
    assert "(#710)" not in c.description
    assert rn.render_commit(c).count("(#710)") == 1


def test_scopeless_subject_parses() -> None:
    (c,) = _commits("fix: shorten end-of-recording clips (#735)")
    assert (c.type, c.scope) == ("fix", "")
    assert rn.render_commit(c) == "- shorten end-of-recording clips (#735)"


def test_multi_scope_is_kept_verbatim() -> None:
    (c,) = _commits("refactor(server,video): collapse duplication (#694)")
    assert c.scope == "server,video"
    assert "**server,video:**" in rn.render_commit(c)


def test_commit_without_a_pr_falls_back_to_the_short_sha() -> None:
    (c,) = _commits("docs: tighten comments")
    assert c.pr == ""
    assert rn.render_commit(c) == "- tighten comments (abc1234)"


def test_breaking_from_bang_and_from_body_trailer() -> None:
    (bang,) = _commits("feat(web)!: drop the old layout (#1)")
    assert bang.breaking is True
    trailer = rn.parse_commit(
        f"abc1234{rn._GIT_SEP}feat: rework{rn._GIT_SEP}BREAKING CHANGE: gone"
    )
    assert trailer.breaking is True


def test_body_with_newlines_does_not_break_record_splitting() -> None:
    """Commit bodies contain newlines, hence the \\x1e/\\x1f framing."""
    log = (
        f"aaa1111{rn._GIT_SEP}feat: one (#1){rn._GIT_SEP}body line\nsecond line\n"
        f"{rn._GIT_END}"
        f"bbb2222{rn._GIT_SEP}fix: two (#2){rn._GIT_SEP}\n{rn._GIT_END}"
    )
    parsed = [rn.parse_commit(r) for r in log.split(rn._GIT_END) if r.strip()]
    assert [(c.type, c.description) for c in parsed] == [
        ("feat", "one"),
        ("fix", "two"),
    ]


def test_unparsed_subject_lands_in_other_changes() -> None:
    """Not conventional, but dropping it would be a silent loss with no rule behind it."""
    grouped, omitted = rn.group_commits(_commits("WIP", "feat: real (#1)"))
    assert omitted == 0
    assert [c.description for c in grouped[rn.OTHER_SECTION]] == ["WIP"]


# --------------------------------------------------------------------------- #
# Body assembly
# --------------------------------------------------------------------------- #


def _body(*subjects: str, initial: bool = False) -> str:
    return rn.render_body(
        _highlight_entry(("Core", "Feat", "a thing")),
        _commits(*subjects),
        "v1.0.0",
        "v0.9.0",
        "owner/repo",
        initial_release=initial,
    )


def test_empty_sections_are_omitted() -> None:
    body = _body("feat: one (#1)")
    assert "### Features" in body
    assert "### Bug Fixes" not in body and "### Performance" not in body


def test_omitted_types_are_counted_not_listed() -> None:
    body = _body("feat: one (#1)", "docs: two (#2)", "ci: three (#3)")
    assert "### Documentation" not in body and "### CI" not in body
    assert (
        "_2 routine changes (refactors, packaging, docs, tests, CI) are omitted._"
        in body
    )


def test_omitted_note_is_dropped_when_nothing_is_omitted() -> None:
    assert "routine" not in _body("feat: one (#1)")


def test_omitted_note_is_singular_for_one() -> None:
    body = _body("feat: one (#1)", "chore: two (#2)")
    assert (
        "_1 routine change (refactors, packaging, docs, tests, CI) is omitted._" in body
    )


def test_a_breaking_change_is_listed_not_counted_as_routine() -> None:
    """Whatever its type. Counting it as omitted would contradict listing it above."""
    body = _body("feat: one (#1)", "refactor!: rip it out (#2)")
    assert "- rip it out (#2)" in body
    assert "routine" not in body


def test_breaking_changes_section_leads_whats_changed() -> None:
    body = _body("feat!: gone (#1)", "fix: other (#2)")
    assert body.index("### ⚠️ Breaking Changes") < body.index("### Features")
    # Flagged once, in context once.
    assert body.count("- gone (#1)") == 2


def test_compare_link_uses_the_previous_tag() -> None:
    assert (
        "**Full Changelog**: https://github.com/owner/repo/compare/v0.9.0...v1.0.0"
        in _body("feat: one (#1)")
    )


def test_first_release_links_to_the_commit_list_instead() -> None:
    body = _body(initial=True)
    assert "**Full Changelog**: https://github.com/owner/repo/commits/v1.0.0" in body
    assert "_First tagged release." in body
    assert "### Features" not in body
    assert "## Highlights" in body, "an initial release still carries its highlights"


def test_body_render_is_deterministic() -> None:
    """Guards against dict/set iteration order leaking into the output."""
    args = ("feat(b): two (#2)", "feat(a): one (#1)", "fix: three (#3)")
    assert _body(*args) == _body(*args)


# --------------------------------------------------------------------------- #
# The reuse guarantee
# --------------------------------------------------------------------------- #


def test_changelog_import_needs_no_third_party_packages() -> None:
    """``build/release_notes.py`` imports ``source/changelog.py`` in a bare CI Python.

    That only works while ``utils.py`` keeps ``rich`` behind a try/except and
    ``gspread`` function-local. Without this test, adding a top-level heavy import
    there would fail on a tag push instead of in the suite.
    """
    result = subprocess.run(
        [
            sys.executable,
            "-S",
            "-c",
            f"import sys; sys.path.insert(0, {str(_ROOT / 'source')!r}); import changelog",
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr


def test_the_shipped_changelog_parses_and_yields_highlights() -> None:
    """Pins the real-file contract without pinning any prose."""
    import changelog

    text = (_ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    picked = rn.select_highlights(changelog.parse_entries(text), "v0.16.0", "v0.15.1")
    assert picked, "the v0.16.0 range should not be empty"
    out = rn.render_highlights(picked)
    assert "## Highlights" in out and "### Core" in out
    assert "- New: " in out and "- Fixed: " in out
