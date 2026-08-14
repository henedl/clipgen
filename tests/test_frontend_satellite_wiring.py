"""Guard that hub/satellite JS groups have no undefined cross-file calls.

Each hub+satellite page is a set of separate ``(function(){ ... })()`` scopes, so
a function defined in one file is **not** visible in a sibling: cross-file calls
go through a same-named guarded delegator or a late-bound ``SS.f(...)``.

When a carve forgets a delegator, the bare call is syntactically valid — so
``node --check`` and the linter both pass it — and throws ``ReferenceError`` at
runtime, often aborting page init. That shipped at least 3x (``e4f67b2``
screenspace, ``8c7f347`` transcripts, plus the ``_calibrationGen`` bug). This
static check catches it, as ``test_packaging.py`` does for the Python
``py-modules`` analogue.

Detection is a heuristic source scan, no node runtime or browser: in each IIFE
file, every *bare* call must resolve to a local definition/import, an ambient
global (a top-level def in a non-IIFE script like ``utils.js``), or a builtin. One
resolving to none of those **but defined in a sibling of the same group** is the
carve-bug signature.

This covers the function-call class only. The rarer bare-*variable* class (a moved
``var`` read across files) is a manual checklist item in
agents/skills/carve-satellite/SKILL.md.
"""

import re

import pytest

from _frontend_source import WEB as _WEB

# Each group is a hub + satellites discovered by glob; the hub file has no "-"
# and so sorts last. Only IIFE-wrapped files are part of the carve (shared
# non-IIFE files like screenspace-utils.js define globals and are ambient).
_GROUPS = {
    "overview": "overview*.js",
    "screenspace": "screenspace*.js",
    "transcripts": "transcripts*.js",
    "studio": "studio*.js",
    "workflows": "workflows*.js",
    "composer": "composer*.js",
}

_KEYWORDS = {
    "if",
    "for",
    "while",
    "switch",
    "catch",
    "return",
    "function",
    "typeof",
    "do",
    "else",
    "new",
    "delete",
    "void",
    "in",
    "of",
    "instanceof",
    "throw",
    "case",
    "var",
    "let",
    "const",
    "await",
    "yield",
    "with",
}

# A bare call to one of these resolves to a JS/DOM/browser builtin, never a
# project function — so it is never a cross-file ReferenceError.
_BUILTINS = {
    "Array",
    "Object",
    "String",
    "Number",
    "Boolean",
    "Math",
    "JSON",
    "Date",
    "Map",
    "Set",
    "WeakMap",
    "WeakSet",
    "Promise",
    "RegExp",
    "Error",
    "Symbol",
    "parseInt",
    "parseFloat",
    "isNaN",
    "isFinite",
    "encodeURIComponent",
    "decodeURIComponent",
    "setTimeout",
    "clearTimeout",
    "setInterval",
    "clearInterval",
    "requestAnimationFrame",
    "cancelAnimationFrame",
    "fetch",
    "alert",
    "confirm",
    "prompt",
    "atob",
    "btoa",
    "structuredClone",
    "queueMicrotask",
    "URL",
    "URLSearchParams",
    "Blob",
    "FormData",
    "Image",
    "Audio",
    "Event",
    "CustomEvent",
    "AbortController",
    "FileReader",
    "IntersectionObserver",
    "ResizeObserver",
    "MutationObserver",
    "DOMParser",
    "TextEncoder",
    "TextDecoder",
    "Intl",
    "BigInt",
    "Proxy",
    "Reflect",
    "getComputedStyle",
    "matchMedia",
    "WebSocket",
    "EventSource",
    "Function",
}


def _strip(src: str) -> str:
    """Remove comments and single-line string literals from JS source.

    String literals are stripped so ``(`` inside text (e.g. ``"url("``) is not
    misread as a call. The ``\\n`` exclusion bounds any mis-parse to one line —
    a DOTALL match would let one stray quote/regex literal eat real definitions
    across the whole file.

    Line comments go first, and that order is load-bearing: prose inside a ``//``
    comment can accidentally contain ``/*`` (``render*Status/*Generating`` did),
    and the DOTALL block-comment pass would then swallow everything up to the
    next ``*/`` anywhere in the file — silently dropping ~1100 lines of
    definitions from the scan and turning this guard into a no-op for the rest
    of that file. Stripping ``//`` first makes such prose unreachable.
    """
    src = re.sub(r"(?<!:)//[^\n]*", " ", src)  # keep http:// inside strings intact
    src = re.sub(r"/\*.*?\*/", " ", src, flags=re.DOTALL)
    src = re.sub(r'"(?:\\.|[^"\\\n])*"', '""', src)
    src = re.sub(r"'(?:\\.|[^'\\\n])*'", "''", src)
    return src


def _is_iife(src: str) -> bool:
    """True if the file body opens with an IIFE wrapper (isolated scope)."""
    body = re.sub(r'^\s*"use strict";?', "", _strip(src).lstrip()).lstrip()
    return body.startswith(("(function", "(()"))


def _param_names(params: str) -> set[str]:
    out: set[str] = set()
    for piece in params.split(","):
        m = re.match(r"\s*\.{0,3}\s*([A-Za-z_$][\w$]*)", piece)
        if m:
            out.add(m.group(1))
    return out


def _local_defs(src: str) -> set[str]:
    """Names bound in this file: declarations, var-chain imports, params."""
    names: set[str] = set()
    names |= set(re.findall(r"\bfunction\s+([A-Za-z_$][\w$]*)", src))
    names |= set(re.findall(r"\b(?:var|let|const)\s+([A-Za-z_$][\w$]*)", src))
    # var-chain continuation: `, foo = SS.foo`
    names |= set(re.findall(r",\s*([A-Za-z_$][\w$]*)\s*=", src))
    # namespace imports: `foo = SS.foo` / TS / STUDIO / window
    names |= set(re.findall(r"([A-Za-z_$][\w$]*)\s*=\s*(?:SS|TS|STUDIO|window)\.", src))
    for blk in re.findall(r"\b(?:var|let|const)\s*\{([^}]*)\}", src):  # destructuring
        for piece in blk.split(","):
            m = re.match(r"\s*([A-Za-z_$][\w$]*)", piece.split(":")[-1])
            if m:
                names.add(m.group(1))
    for params in re.findall(r"function[^(]*\(([^)]*)\)", src):  # fn params
        names |= _param_names(params)
    for params in re.findall(r"\(([^)]*)\)\s*=>", src):  # arrow params
        names |= _param_names(params)
    names |= set(re.findall(r"(?:^|[(,])\s*([A-Za-z_$][\w$]*)\s*=>", src))  # x => ...
    return names


def _top_level_defs(src: str) -> set[str]:
    """File-scope (column-0) defs — the globals a non-IIFE script exposes."""
    return set(
        re.findall(
            r"^(?:var|let|const|function)\s+([A-Za-z_$][\w$]*)", src, re.MULTILINE
        )
    )


def _bare_calls(src: str) -> set[str]:
    """Identifiers called as ``name(`` and not preceded by ``.`` (not a method)."""
    return set(re.findall(r"(?<![.\w$])([A-Za-z_$][\w$]*)\s*\(", src)) - _KEYWORDS


def _ambient_globals() -> set[str]:
    """Top-level defs of every non-IIFE script — globals visible to all groups."""
    names: set[str] = set()
    for path in _WEB.glob("*.js"):
        text = path.read_text(encoding="utf-8")
        if not _is_iife(text):
            names |= _top_level_defs(_strip(text))
    return names


@pytest.mark.parametrize("group, pattern", sorted(_GROUPS.items()))
def test_no_undefined_cross_file_calls(group: str, pattern: str) -> None:
    files = [
        p for p in sorted(_WEB.glob(pattern)) if _is_iife(p.read_text(encoding="utf-8"))
    ]
    assert len(files) >= 2, f"{group}: expected a hub + satellite(s), found {files}"

    ambient = _ambient_globals()
    stripped = {p.name: _strip(p.read_text(encoding="utf-8")) for p in files}
    defs = {name: _local_defs(s) for name, s in stripped.items()}
    calls = {name: _bare_calls(s) for name, s in stripped.items()}

    offenders: dict[str, list[str]] = {}
    for fname, file_calls in calls.items():
        unresolved = file_calls - defs[fname] - ambient - _BUILTINS
        # The carve-bug signature: a bare call that is owned by a *sibling* file
        # in this group (so it was meant to cross files) but has no local
        # delegator/import here — it will throw ReferenceError at runtime.
        cross_file = sorted(
            name
            for name in unresolved
            if any(name in defs[other] for other in stripped if other != fname)
        )
        if cross_file:
            offenders[fname] = cross_file

    assert not offenders, (
        f"{group}: bare cross-file call(s) with no hub delegator / namespace "
        f"import — these throw ReferenceError at runtime (node --check misses "
        f"them). Add a same-named guarded delegator or late-bind via the "
        f"namespace. See agents/skills/carve-satellite/SKILL.md. Offenders: "
        f"{offenders}"
    )
