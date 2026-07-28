"""Guard resource-lifecycle contracts that prose alone has not held.

Two classes, both re-fixed across several audit sweeps:

* Blob URLs created and never revoked, so a long-lived tab grows unbounded
  (`7c8e751b`, `c895bd2d` — `_thumbCache`, `_ssThumbCache`, `_cvFrameCache`,
  `_preloadedFrames` all leaked).
* Module-level caches with no bound, so a long scrub session accumulates
  decoded frames forever (`186ff45a`, `1b444d0d` both converted a plain dict to
  a bounded LRU after the fact).
"""

import ast
import re
from pathlib import Path

from _frontend_source import WEB

SOURCE = Path(__file__).resolve().parent.parent / "source"

# ---------------------------------------------------------------- blob URLs


def test_every_createobjecturl_file_also_revokes():
    """A file that mints blob URLs must also release them.

    This is the cheap half of the rule: it catches a *new* cache of blob URLs
    landing with no revoke path at all. The other half -- revoking on
    ``pagehide`` as well as on replacement -- stays prose in CODE-REVIEW.md.

    Grepping for ``createObjectURL`` without ``pagehide`` over-reports: the
    download-anchor idiom (``createObjectURL`` -> ``a.click()`` ->
    ``revokeObjectURL`` on the next statement, as in ``overview-metadata.js``
    and ``workflows.js``) caches nothing, so there is nothing for a teardown to
    free. The shape that actually leaks is a blob URL stored in a *module-scoped
    cache*; those live in ``studio.js``, ``overview-convergence.js``,
    ``viewer.js`` and ``screenspace.js``, and all four now revoke on pagehide.
    """
    offenders = [
        p.name
        for p in sorted(WEB.glob("*.js"))
        if "createObjectURL" in (src := p.read_text(encoding="utf-8"))
        and "revokeObjectURL" not in src
    ]
    assert not offenders, (
        "file creates blob URLs but never revokes them (leaks for the tab's "
        "lifetime):\n" + "\n".join(offenders)
    )


# ------------------------------------------------------------ cache bounding

_CACHE_MODULES = (
    "server.py",
    "screenspace_server.py",
    "transcripts_server.py",
    "viewer.py",
)

# Caches whose size is bounded by the domain rather than by an explicit cap.
# Each entry names what bounds it -- add here only with a real justification.
_NATURALLY_BOUNDED = {
    "_sheet_payload_cache": "single entry (the one open worksheet)",
    "_participant_timeline_cache": "one entry per participant in the cohort",
    "_video_metadata_cache": "one entry per participant video",
    "_corrected_cache": "one entry per participant transcript",
    "_ss_events_cache": "single entry, mtime-keyed",
    "_manifest_cache": "single entry, mtime-keyed",
}


def _cache_definitions():
    """Yield (module, name, rhs_source, module_source) for module-level caches.

    Parsed with ``ast`` rather than a regex: several caches carry a multi-line
    type annotation, which a line-oriented pattern silently skips -- a false
    negative in exactly the direction that matters here.
    """
    for mod in _CACHE_MODULES:
        src = (SOURCE / mod).read_text(encoding="utf-8")
        tree = ast.parse(src)
        for node in tree.body:  # module level only
            if isinstance(node, ast.AnnAssign):
                targets, value = [node.target], node.value
            elif isinstance(node, ast.Assign):
                targets, value = node.targets, node.value
            else:
                continue
            if value is None:
                continue
            for target in targets:
                if not isinstance(target, ast.Name):
                    continue
                name = target.id
                if not name.startswith("_") or "cache" not in name.lower():
                    continue
                # ALL-CAPS names are the cap constants / lock handles themselves
                # (_FRAME_CACHE_MAX, _MANIFEST_CACHE_LOCK), not containers.
                if name.isupper() or name.endswith("_lock"):
                    continue
                # Plain aliases (`_MediaCache = MediaCache`) hold no entries.
                if isinstance(value, ast.Name):
                    continue
                yield mod, name, ast.unparse(value), src


def test_module_level_caches_are_bounded_or_justified():
    """Every server-side cache states its bound.

    Accepted forms: an ``OrderedDict`` with a sibling ``_*_CACHE_MAX`` constant,
    a ``MediaCache(max)``, or an entry in ``_NATURALLY_BOUNDED`` naming the key
    that bounds it. A new unbounded dict fails here rather than surviving until
    someone notices the memory growth in a long session.
    """
    offenders = []
    for mod, name, rhs, src in _cache_definitions():
        if name in _NATURALLY_BOUNDED:
            continue
        if "MediaCache(" in rhs:
            continue
        if "OrderedDict" in rhs:
            # Require a companion cap constant, e.g. _FRAME_CACHE_MAX.
            cap = name.upper() + "_MAX"
            if re.search(rf"^{cap}\s*=", src, re.MULTILINE):
                continue
            offenders.append(f"{mod}:{name} is an OrderedDict with no {cap}")
            continue
        offenders.append(
            f"{mod}:{name} = {rhs} — unbounded; use OrderedDict + "
            f"{name.upper()}_MAX, a MediaCache, or justify it in _NATURALLY_BOUNDED"
        )
    assert not offenders, "\n".join(offenders)


def test_naturally_bounded_allowlist_has_no_dead_entries():
    """Keep the allowlist honest: every entry must still name a real cache."""
    live = {name for _, name, _, _ in _cache_definitions()}
    dead = sorted(set(_NATURALLY_BOUNDED) - live)
    assert not dead, f"allowlisted caches no longer exist: {dead}"
