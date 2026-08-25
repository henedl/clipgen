"""Deterministic guards for AGENTS.md / CODE-REVIEW.md source rules.

Each rule was prose-only; these scans make it mechanical. Frozen per-file
counts ratchet both ways: growth fails, and a fixed site must be removed
from its baseline so the win is locked in.

One pass: every ``source/*.py`` is read exactly once (test-perf budget).
"""

import ast
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "source"

_SOURCES = {path.name: path.read_text(encoding="utf-8") for path in SOURCE.glob("*.py")}

# Version-shaped numbers that are not segments of a dotted quad (IP, port).
_DOC_VERSION = re.compile(r"(?<![\d.])\d+\.\d+\.\d+(?![\d.])")


def _count_hits(pattern: str | re.Pattern[str]) -> dict[str, int]:
    regex = re.compile(pattern) if isinstance(pattern, str) else pattern
    counts = {}
    for name, text in sorted(_SOURCES.items()):
        hits = len(regex.findall(text))
        if hits:
            counts[name] = hits
    return counts


def _assert_frozen(found: dict[str, int], baseline: dict[str, int], rule: str) -> None:
    grew = {n: c for n, c in found.items() if c > baseline.get(n, 0)}
    assert not grew, f"{rule}: new sites in {grew} — use the sanctioned helper instead"
    shrank = {n: c for n, c in baseline.items() if found.get(n, 0) < c}
    assert not shrank, f"{rule}: nice — ratchet the baseline down for {shrank}"


def test_no_hand_rolled_repo_root_paths() -> None:
    """Repo paths go through utils.get_bundled_assets_root(), never Path(__file__)."""
    baseline = {
        "cli.py": 1,  # get_runtime_working_dir: "next to the app", not assets
        "utils.py": 1,  # get_bundled_assets_root itself
    }
    _assert_frozen(_count_hits(r"Path\(__file__\)"), baseline, "Path(__file__)")


def _count_input_calls() -> dict[str, int]:
    counts = {}
    for name, text in sorted(_SOURCES.items()):
        if "input(" not in text:  # skip the parse; ast is the slow part
            continue
        hits = sum(
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id == "input"
            for node in ast.walk(ast.parse(text))
        )
        if hits:
            counts[name] = hits
    return counts


def test_prompts_go_through_read_user_input() -> None:
    """Interactive prompts use utils.read_user_input(), which handles keywords."""
    baseline = {
        "interactive.py": 1,  # raw-terminal key echo: reads the rest of a line
        "utils.py": 1,  # read_user_input's own input() call
    }
    _assert_frozen(_count_input_calls(), baseline, "bare input()")


def test_no_type_ignore_suppressions() -> None:
    """Narrow with `assert x is not None` (CODE-REVIEW.md), never suppress ty."""
    baseline = {
        "llm_client.py": 1,  # urllib3 private ._sock, no public spelling
        "server_utils.py": 1,  # untyped nested Flask generator
        "transcripts_server.py": 1,  # untyped nested Flask generator
        "utils.py": 1,  # rich fallback rebinds Progress to None
    }
    _assert_frozen(_count_hits(r"#\s*type:\s*ignore"), baseline, "type: ignore")


def test_no_numeric_route_converters() -> None:
    """Routes take string params + manual int()/float() (CODE-REVIEW.md)."""
    baseline = {
        "composer_server.py": 1,  # audio-track idx; predates the rule
        "screenspace_server.py": 1,  # audio-track idx; predates the rule
        "transcripts_server.py": 1,  # audio-track idx; predates the rule
    }
    _assert_frozen(_count_hits(r"<(?:int|float):"), baseline, "route converter")


def test_evergreen_docs_carry_no_version_numbers() -> None:
    """AGENTS.md / README.md reference build/VERSION, never a literal version."""
    for name in ("AGENTS.md", "README.md"):
        text = (ROOT / name).read_text(encoding="utf-8")
        hits = _DOC_VERSION.findall(text)
        assert not hits, f"{name} hardcodes {hits}; reference build/VERSION instead"


def test_the_scans_see_planted_violations() -> None:
    """Guard the detectors: each must catch its target spelling."""
    planted = ast.walk(ast.parse("x = input('> ')\ny = utils.read_user_input('> ')"))
    calls = [
        n.func.id
        for n in planted
        if isinstance(n, ast.Call) and isinstance(n.func, ast.Name)
    ]
    assert calls == ["input"]
    assert _DOC_VERSION.search("as of 0.16.11 the")
    assert not _DOC_VERSION.search("bind to 127.0.0.1:8089")
