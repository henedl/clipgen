"""Source-level guards for the shared hotkey registry (assets/web/hotkeys.js).

Keyboard handling is centralized: pages register actions against the catalog
in hotkeys.js instead of adding their own document-level keydown listeners.
These tests keep that contract honest:

  1. Only an allowlisted set of files may attach document-level keydown
     listeners (transient popover owners + the dispatcher itself).
  2. The catalog literal stays parseable, ids/combos stay valid against the
     server's structural regexes (imported, not duplicated), and Escape/Tab
     stay reserved.
  3. Every action id registered by page code exists in the catalog, and every
     catalog action is registered by some page (no dead entries).
  4. hotkeys.js stays ES5 (exported viewers inline it verbatim).
  5. The hand-written per-page cheatsheets stay deleted and every template
     loads hotkeys.js directly after utils.js.
"""

import json
import re
from pathlib import Path

import pytest

import server

WEB = Path(__file__).resolve().parent.parent / "assets" / "web"
HOTKEYS_SRC = (WEB / "hotkeys.js").read_text(encoding="utf-8")

# Files allowed to attach document-level keydown listeners: the dispatcher
# itself plus transient popover/modal owners that must capture keys while
# open (each detaches on close, or gates on its own open state).
KEYDOWN_ALLOWLIST = {
    "hotkeys.js",
    "utils.js",  # openBlockingModal focus trap + frontend switcher
    "settings-modal.js",  # hotkey recorder (capture-phase, recording only)
    "color-picker.js",
    "studio-trim.js",
    "topnav.js",
    "workflows.js",  # bindMenuToggle (run split-button menu)
    "workflows-wires.js",  # armed-wire Escape (attached only while armed)
    "start-overlay.js",
}

_DOC_KEYDOWN_RE = re.compile(
    r'document\.addEventListener\(\s*"keydown"|on\(document,\s*"keydown"'
)


def _js_files():
    return sorted(p for p in WEB.glob("*.js") if "vendor" not in str(p))


def test_document_keydown_allowlist():
    offenders = []
    for path in _js_files():
        if _DOC_KEYDOWN_RE.search(path.read_text(encoding="utf-8")):
            if path.name not in KEYDOWN_ALLOWLIST:
                offenders.append(path.name)
    assert not offenders, (
        f"Files attach document-level keydown listeners outside the allowlist: "
        f"{offenders}. Register actions via ClipgenHotkeys.register / "
        f"registerEscape in assets/web/hotkeys.js instead."
    )


def test_keyup_only_in_dispatcher():
    offenders = []
    for path in _js_files():
        if path.name == "hotkeys.js":
            continue
        if '"keyup"' in path.read_text(encoding="utf-8"):
            offenders.append(path.name)
    assert not offenders, (
        f"keyup listeners outside hotkeys.js: {offenders}. Hold-to-act "
        f"shortcuts use the onRelease option of ClipgenHotkeys.register."
    )


def _parse_literal(name: str):
    match = re.search(
        r"var\s+" + re.escape(name) + r"\s*=\s*(\[.+?\n  \]);",
        HOTKEYS_SRC,
        re.DOTALL,
    )
    assert match, f"{name} literal not found in hotkeys.js"
    raw = match.group(1)
    raw = re.sub(r"(\b\w+)\s*:", r'"\1":', raw)
    raw = re.sub(r",\s*([}\]])", r"\1", raw)
    return json.loads(raw)


@pytest.fixture(scope="module")
def catalog():
    return _parse_literal("HOTKEY_CATALOG")


@pytest.fixture(scope="module")
def sections():
    return _parse_literal("HOTKEY_SECTIONS")


def test_catalog_ids_unique_and_valid(catalog, sections):
    ids = [a["id"] for a in catalog]
    assert len(ids) == len(set(ids)), "Duplicate action ids in HOTKEY_CATALOG"
    section_ids = {s["id"] for s in sections}
    for action in catalog:
        assert server._HOTKEY_ID_RE.match(action["id"]), (
            f"Action id {action['id']!r} does not match server._HOTKEY_ID_RE"
        )
        assert action["section"] in section_ids, (
            f"Action {action['id']} references unknown section {action['section']!r}"
        )
        assert action.get("label"), f"Action {action['id']} has no label"


def test_catalog_combos_valid(catalog):
    for action in catalog:
        combos = action.get("combos")
        if action.get("note"):
            assert combos is None, f"Note entry {action['id']} must not define combos"
            continue
        assert combos, f"Action {action['id']} has no default combos"
        for combo in combos:
            assert server._HOTKEY_COMBO_RE.match(combo), (
                f"Combo {combo!r} of {action['id']} does not match "
                f"server._HOTKEY_COMBO_RE"
            )
            assert combo.split("+")[-1] not in server._HOTKEY_RESERVED_KEYS, (
                f"Combo {combo!r} of {action['id']} uses a reserved key"
            )


_REGISTERED_ID_RE = re.compile(r'id:\s*"([a-z][a-zA-Z0-9]*(?:\.[a-zA-Z0-9]+)+)"')


def _registered_ids_by_file():
    found = {}
    for path in _js_files():
        if path.name == "hotkeys.js":
            continue
        ids = _REGISTERED_ID_RE.findall(path.read_text(encoding="utf-8"))
        if ids:
            found[path.name] = ids
    return found


def test_registered_ids_exist_in_catalog(catalog):
    known = {a["id"] for a in catalog}
    unknown = []
    for fname, ids in _registered_ids_by_file().items():
        for action_id in ids:
            if action_id not in known:
                unknown.append(f"{fname}: {action_id}")
    assert not unknown, (
        f"Page code registers action ids missing from HOTKEY_CATALOG: {unknown}"
    )


def test_catalog_actions_all_registered(catalog):
    registered = set()
    for ids in _registered_ids_by_file().values():
        registered.update(ids)
    # global.cheatsheet is registered by hotkeys.js itself.
    registered.add("global.cheatsheet")
    dead = [a["id"] for a in catalog if a["id"] not in registered]
    assert not dead, (
        f"Catalog actions never registered by any page: {dead}. "
        f"Remove the entry or add the ClipgenHotkeys.register call."
    )


def test_hotkeys_js_is_es5():
    assert "=>" not in HOTKEYS_SRC, "hotkeys.js must not use arrow functions"
    assert not re.search(r"\b(let|const)\s", HOTKEYS_SRC), (
        "hotkeys.js must not use let/const"
    )
    # Template literals: a bare backtick check would trip on a backtick combo
    # string in the catalog, so look for interpolation specifically.
    assert not re.search(r"`[^`\n]*\$\{", HOTKEYS_SRC), (
        "hotkeys.js must not use template literals"
    )


ALL_TEMPLATES = [
    "studio.html",
    "screenspace.html",
    "transcripts.html",
    "composer.html",
    "overview.html",
    "workflows.html",
    "viewer.html",
    "gallery.html",
    "timeline-viewer.html",
]


def test_templates_load_hotkeys_after_utils():
    for name in ALL_TEMPLATES:
        html = (WEB / name).read_text(encoding="utf-8")
        utils_idx = html.find('<script src="utils.js" defer></script>')
        hotkeys_idx = html.find('<script src="hotkeys.js" defer></script>')
        assert utils_idx != -1, f"{name} missing utils.js script tag"
        assert hotkeys_idx != -1, f"{name} missing hotkeys.js script tag"
        assert hotkeys_idx > utils_idx, f"{name}: hotkeys.js must load after utils.js"
        assert '<link rel="stylesheet" href="hotkeys.css">' in html, (
            f"{name} missing hotkeys.css link"
        )


def test_handwritten_cheatsheets_removed():
    removed = {
        "composer.html": "coShortcutsMenu",
        "transcripts.html": "trShortcuts",
        "workflows.html": "wfShortcutsMenu",
    }
    for name, marker in removed.items():
        html = (WEB / name).read_text(encoding="utf-8")
        assert marker not in html, (
            f"{name} still contains the hand-written cheatsheet ({marker}); "
            f"the shared hotkeys.js overlay replaces it"
        )


def test_coerce_hotkey_overrides():
    ok = server._coerce_hotkey_overrides(
        {"transport.playPause": "Space", "edit.redo": "Mod+Shift+Z Mod+Y"}
    )
    assert ok == {"transport.playPause": "Space", "edit.redo": "Mod+Shift+Z Mod+Y"}
    # Empty string disables a shortcut.
    assert server._coerce_hotkey_overrides({"nav.next": ""}) == {"nav.next": ""}
    # Structural rejections.
    assert server._coerce_hotkey_overrides("nope") is None
    assert server._coerce_hotkey_overrides({"BadId": "X"}) is None
    assert server._coerce_hotkey_overrides({"nav.next": 5}) is None
    assert server._coerce_hotkey_overrides({"nav.next": "Bogus+X"}) is None
    # Escape/Tab are reserved.
    assert server._coerce_hotkey_overrides({"nav.next": "Escape"}) is None
    assert server._coerce_hotkey_overrides({"nav.next": "Mod+Tab"}) is None


def test_hotkey_overrides_flow_through_frontend_config():
    import config
    import utils

    original = config.HOTKEY_OVERRIDES
    try:
        config.HOTKEY_OVERRIDES = {"nav.next": "N"}
        assert utils.get_frontend_config()["hotkeyOverrides"] == {"nav.next": "N"}
    finally:
        config.HOTKEY_OVERRIDES = original
