"""Static checks on the wording of the export affordances.

The shared Export quick action writes files server-side into the output folder
and only covers Screenspace/Transcripts data — neither fact is guessable from a
bare "Export" label, which is also what the command palette shows. These pin the
copy so it does not regress to a naked verb. Mirrors the Workflows copy checks in
tests/test_workflows_frontend_source.py.
"""

from _frontend_source import read

EXPORT_ACTIONS_JS = "export-actions.js"
SCREENSPACE_HTML = "screenspace.html"
STUDIO_JS = "studio.js"


def test_shared_export_label_names_object_and_format():
    src = read(EXPORT_ACTIONS_JS)
    assert 'label: "Export",' not in src, "bare 'Export' label says nothing"
    assert 'label: "Export analysis data (JSON+CSV)"' in src


def test_shared_export_tooltip_names_data_format_and_destination():
    src = read(EXPORT_ACTIONS_JS)
    # Enabled tooltip: what data, what files, where they land, what they are not.
    assert "output folder" in src
    assert "clipgen_export_*.json" in src
    assert "Does not export clips." in src
    # Disabled tooltip stays an imperative unblocking step.
    assert "Run a Screenspace scan or transcribe a video first" in src


def test_shared_export_toast_names_the_surfaces():
    src = read(EXPORT_ACTIONS_JS)
    assert "function surfaceNames(" in src
    assert "clipgen_export_" in src


def test_studio_quick_actions_all_have_tooltips():
    src = read(STUDIO_JS)
    start = src.index('label: "Build Viewer"')
    end = src.index("exportQuickAction()", start)
    block = src[start:end]
    # Three items, each carrying a title (the palette uses it as the subtitle).
    for label in ('"Build Viewer"', '"Open Timeline"', '"Open Gallery"'):
        assert label in block
    assert block.count("title:") == 3


def test_screenspace_export_menu_states_its_scope():
    html = read(SCREENSPACE_HTML)
    start = html.index('id="exportEventsMenu"')
    end = html.index("</div>", html.index('data-format="json"'))
    menu = html[start:end]
    assert 'class="rp-export-note"' in menu
    # The one thing the UI otherwise hides: the file is not the filtered list.
    assert "certainty slider filters this list only, not the file" in menu
    # Self-describing item labels, unchanged format values.
    assert ">Save as CSV<" in menu
    assert ">Save as JSON<" in menu
    assert 'data-format="csv"' in menu
    assert 'data-format="json"' in menu


def test_screenspace_export_note_is_styled_and_inert():
    css = read("screenspace.css")
    start = css.index(".rp-export-note {")
    block = css[start : css.index("}", start)]
    # Not a menu item: clicks must fall through to the wrapper's outside-click
    # handler rather than looking clickable.
    assert "pointer-events: none;" in block
