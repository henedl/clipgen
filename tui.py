# -*- coding: utf-8 -*-
"""Textual TUI screens for clipgen interactive features."""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Optional

import config

if TYPE_CHECKING:
    from utils import BrowseRow

try:
    from textual.app import App, ComposeResult
    from textual.binding import Binding
    from textual.containers import Horizontal, Vertical, VerticalScroll
    from textual.widgets import (
        Button,
        DataTable,
        Footer,
        Header,
        Input,
        Label,
        Select,
        SelectionList,
        Static,
        Switch,
    )
    TEXTUAL_AVAILABLE = True
except ImportError:
    TEXTUAL_AVAILABLE = False


def use_textual() -> bool:
    """Check if Textual TUI screens should be used."""
    return TEXTUAL_AVAILABLE and getattr(config, 'TEXTUAL_TUI', True)


# ---------------------------------------------------------------------------
# Settings definitions: (config_attr, display_label, widget_type, extra)
#   widget_type: 'switch' | 'select' | 'input'
#   extra:       None for switch, list of choices for select, 'integer' for input
# ---------------------------------------------------------------------------

_SETTINGS_DEFS = [
    ('REENCODING',                'Re-encode clips',          'switch', None),
    ('AUDIO_NORMALIZE',           'Audio normalization',      'switch', None),
    ('FILEFORMAT',                'File format',              'select', ['.mp4', '.mkv', '.mov', '.webm']),
    ('MAX_FILESIZE_MB',           'Max file size (MB, 0=off)', 'input', 'integer'),
    ('DEFAULT_DURATION_SECONDS',  'Default duration (sec)',   'input',  'integer'),
    ('MAX_CLIP_DURATION_SECONDS', 'Max clip duration (sec)',  'input',  'integer'),
    ('DEBUGGING',                 'Debug mode',               'switch', None),
    ('RICH_COLORS',               'Rich colors',              'switch', None),
    ('RICH_PANELS',               'Rich panels',              'switch', None),
    ('RICH_PROGRESS',             'Rich progress bars',       'switch', None),
    ('TEXTUAL_TUI',               'Textual TUI',              'switch', None),
]


# ---------------------------------------------------------------------------
# Textual App classes  (only defined when the library is installed)
# ---------------------------------------------------------------------------

if TEXTUAL_AVAILABLE:

    class SettingsApp(App):
        """Interactive settings editor with grid-based card layout."""

        TITLE = "clipgen Settings"

        CSS = """
        Screen {
            align: center middle;
        }
        #settings-scroll {
            width: 1fr;
            max-width: 96;
            max-height: 90%;
            border: solid $accent;
            padding: 1 2;
        }
        #settings-title {
            text-style: bold;
            width: 100%;
            content-align: center middle;
            margin-bottom: 1;
        }
        #settings-grid {
            layout: grid;
            grid-size: 2;
            grid-gutter: 1 2;
            grid-rows: auto;
            height: auto;
        }
        .setting-card {
            height: auto;
            padding: 1 2;
            border: round $surface-lighten-2;
        }
        .setting-name {
            text-style: bold;
        }
        .setting-desc {
            color: $text 60%;
            margin-top: 1;
        }
        .setting-card Switch,
        .setting-card Select,
        .setting-card Input {
            margin-top: 1;
        }
        #button-row {
            width: 1fr;
            max-width: 96;
            height: auto;
            align: center middle;
            padding: 1 0 0 0;
        }
        #button-row Button {
            margin: 0 2;
        }
        """

        BINDINGS = [
            Binding("escape", "cancel", "Cancel", show=True),
        ]

        def compose(self) -> ComposeResult:
            descriptions = getattr(config, 'SETTINGS_DESCRIPTIONS', {})
            yield Header()
            with VerticalScroll(id="settings-scroll"):
                yield Static("Settings", id="settings-title")
                with Vertical(id="settings-grid"):
                    for attr, label, wtype, extra in _SETTINGS_DEFS:
                        current = getattr(config, attr)
                        desc = descriptions.get(attr, '')
                        with Vertical(classes="setting-card"):
                            yield Label(label, classes="setting-name")
                            if wtype == 'switch':
                                yield Switch(value=bool(current), id=attr)
                            elif wtype == 'select':
                                options = [(o, o) for o in (extra or [])]
                                yield Select(
                                    options,
                                    value=str(current),
                                    id=attr,
                                    allow_blank=False,
                                )
                            elif wtype == 'input':
                                yield Input(
                                    value=str(current),
                                    id=attr,
                                    type="integer" if extra == 'integer' else "text",
                                )
                            if desc:
                                yield Static(desc, classes="setting-desc")
            with Horizontal(id="button-row"):
                yield Button("Save", variant="primary", id="save")
                yield Button("Cancel", variant="default", id="cancel")
            yield Footer()

        def on_button_pressed(self, event: Button.Pressed) -> None:
            if event.button.id == "save":
                self._apply_settings()
                self.exit(True)
            else:
                self.exit(False)

        def action_cancel(self) -> None:
            self.exit(False)

        def _apply_settings(self) -> None:
            for attr, _, wtype, extra in _SETTINGS_DEFS:
                if wtype == 'switch':
                    switch = self.query_one(f"#{attr}", Switch)
                    setattr(config, attr, switch.value)
                elif wtype == 'select':
                    select = self.query_one(f"#{attr}", Select)
                    if select.value != Select.BLANK:
                        setattr(config, attr, select.value)
                elif wtype == 'input':
                    try:
                        input_widget = self.query_one(f"#{attr}", Input)
                        setattr(config, attr, int(input_widget.value))
                    except (ValueError, TypeError):
                        self.notify(f"Invalid value for {attr}, keeping current.", severity="warning")

    # -------------------------------------------------------------------

    class CategorySelectApp(App):
        """Checkbox list for interactive category selection."""

        TITLE = "Select Categories"

        CSS = """
        Screen {
            align: center middle;
        }
        #category-scroll {
            width: 60;
            max-height: 80%;
            border: solid $accent;
            padding: 1 2;
        }
        #category-title {
            text-style: bold;
            margin-bottom: 1;
        }
        #button-row {
            width: 60;
            height: auto;
            align: center middle;
            padding: 1 0 0 0;
        }
        #button-row Button {
            margin: 0 1;
        }
        """

        BINDINGS = [
            Binding("escape", "cancel", "Cancel", show=True),
            Binding("a", "toggle_all", "Toggle All", show=True),
        ]

        def __init__(self, categories: List[str]):
            super().__init__()
            self.categories = categories

        def compose(self) -> ComposeResult:
            yield Header()
            with VerticalScroll(id="category-scroll"):
                yield Static("Choose one or more categories:", id="category-title")
                yield SelectionList(
                    *[(cat, cat) for cat in self.categories],
                    id="category-list",
                )
            with Horizontal(id="button-row"):
                yield Button("Confirm", variant="primary", id="confirm")
                yield Button("Toggle All", variant="default", id="toggle-all")
                yield Button("Cancel", variant="error", id="cancel")
            yield Footer()

        def on_button_pressed(self, event: Button.Pressed) -> None:
            if event.button.id == "confirm":
                self._confirm()
            elif event.button.id == "toggle-all":
                self.action_toggle_all()
            else:
                self.exit(None)

        def action_cancel(self) -> None:
            self.exit(None)

        def action_toggle_all(self) -> None:
            sl = self.query_one("#category-list", SelectionList)
            if sl.selected:
                sl.deselect_all()
            else:
                sl.select_all()

        def _confirm(self) -> None:
            sl = self.query_one("#category-list", SelectionList)
            selected = list(sl.selected)
            if selected:
                self.exit(selected)
            else:
                self.notify("Select at least one category", severity="warning")

    # -------------------------------------------------------------------

    class BrowseApp(App):
        """Scrollable spreadsheet browser using DataTable."""

        TITLE = "Browse Spreadsheet"

        CSS = """
        DataTable {
            height: 1fr;
        }
        """

        BINDINGS = [
            Binding("q", "quit", "Quit", show=True),
            Binding("escape", "quit", "Quit", show=False),
        ]

        def __init__(
            self,
            rows_data: List[BrowseRow],
            participant_headers: List[str],
            title: str = "",
        ):
            super().__init__()
            self._rows_data = rows_data
            self._participant_headers = participant_headers
            self._browse_title = title or "Browse Spreadsheet"

        def compose(self) -> ComposeResult:
            yield Header()
            yield DataTable(id="browse-table", zebra_stripes=True)
            yield Footer()

    # -------------------------------------------------------------------

    class SpreadsheetSelectApp(App):
        """Spreadsheet selector for recent documents."""

        TITLE = "Select Spreadsheet"

        CSS = """
        Screen {
            align: center middle;
        }
        #sheet-panel {
            width: 72;
            max-height: 80%;
            border: solid $accent;
            padding: 1 2;
        }
        #sheet-title {
            text-style: bold;
            width: 100%;
            content-align: center middle;
            margin-bottom: 1;
        }
        #sheet-list {
            height: 1fr;
        }
        #button-row {
            width: 72;
            height: auto;
            align: center middle;
            padding: 1 0 0 0;
        }
        #button-row Button {
            margin: 0 2;
        }
        """

        BINDINGS = [
            Binding("escape", "cancel", "Cancel", show=True),
            Binding("enter", "confirm", "Open", show=False),
        ]

        def __init__(self, doc_list: List[str]):
            super().__init__()
            self._doc_list = doc_list

        def compose(self) -> ComposeResult:
            yield Header()
            with Vertical(id="sheet-panel"):
                yield Static("Choose a spreadsheet:", id="sheet-title")
                options = [
                    (f"{idx+1}. {name.strip()}", name.strip())
                    for idx, name in enumerate(self._doc_list)
                ]
                yield SelectionList(*options, id="sheet-list")
            with Horizontal(id="button-row"):
                yield Button("Open", variant="primary", id="open")
                yield Button("Cancel", variant="default", id="cancel")
            yield Footer()

        def on_mount(self) -> None:
            self.query_one("#sheet-list", SelectionList).focus()

        def action_cancel(self) -> None:
            self.exit(None)

        def action_confirm(self) -> None:
            sl = self.query_one("#sheet-list", SelectionList)
            selected = list(sl.selected)
            if not selected:
                self.notify("Select a spreadsheet to open.", severity="warning")
                return
            # SelectionList.selected is a set of option IDs (values we passed in)
            self.exit(selected[0])

        def on_button_pressed(self, event: Button.Pressed) -> None:
            if event.button.id == "open":
                self.action_confirm()
            else:
                self.action_cancel()

    # -------------------------------------------------------------------

    class ModeSelectApp(App):
        """Main mode selection menu for interactive runs."""

        TITLE = "Select Mode"

        CSS = """
        Screen {
            align: center middle;
        }
        #mode-panel {
            width: 72;
            max-height: 80%;
            border: solid $accent;
            padding: 1 2;
        }
        #mode-title {
            text-style: bold;
            width: 100%;
            content-align: center middle;
            margin-bottom: 1;
        }
        #mode-list {
            height: 1fr;
        }
        #button-row {
            width: 72;
            height: auto;
            align: center middle;
            padding: 1 0 0 0;
        }
        #button-row Button {
            margin: 0 2;
        }
        """

        BINDINGS = [
            Binding("escape", "cancel", "Cancel", show=True),
            Binding("enter", "confirm", "Select", show=False),
        ]

        def compose(self) -> ComposeResult:
            yield Header()
            with Vertical(id="mode-panel"):
                yield Static("Choose what you want to generate:", id="mode-title")
                options = [
                    ("Batch – all clips", "batch"),
                    ("Range – rows between two lines", "range"),
                    ("Category – by category label", "category"),
                    ("Line – specific row numbers", "line"),
                    ("Cell – specific cells (P01.11)", "cell"),
                    ("Participant – all clips for participant(s)", "participant"),
                    ("Filter – key-marked clips only", "filter"),
                    ("Screen – screenshots (.png)", "screen"),
                    ("GIF – GIFs (.gif)", "gif"),
                    ("Reel – combined highlight reel", "reel"),
                    ("Reel-late – combine existing clips", "reellate"),
                    ("Browse – inspect spreadsheet rows", "browse"),
                    ("Custom selectors / type manually", "custom"),
                ]
                yield SelectionList(*options, id="mode-list")
            with Horizontal(id="button-row"):
                yield Button("Select", variant="primary", id="select")
                yield Button("Cancel", variant="default", id="cancel")
            yield Footer()

        def on_mount(self) -> None:
            self.query_one("#mode-list", SelectionList).focus()

        def action_cancel(self) -> None:
            self.exit(None)

        def action_confirm(self) -> None:
            sl = self.query_one("#mode-list", SelectionList)
            selected = list(sl.selected)
            if not selected:
                self.notify("Select a mode to continue.", severity="warning")
                return
            self.exit(selected[0])

        def on_button_pressed(self, event: Button.Pressed) -> None:
            if event.button.id == "select":
                self.action_confirm()
            else:
                self.action_cancel()

    # -------------------------------------------------------------------

    class ReelBuilderApp(App):
        """Compound reel selector: participants, rows, and categories."""

        TITLE = "Build Reel"

        CSS = """
        Screen {
            align: center middle;
        }
        #reel-root {
            width: 96;
            max-height: 90%;
            border: solid $accent;
            padding: 1 2;
        }
        #reel-title {
            text-style: bold;
            width: 100%;
            content-align: center middle;
            margin-bottom: 1;
        }
        #selectors {
            height: 1fr;
        }
        .selector-column {
            width: 1fr;
        }
        .selector-column Static {
            text-style: bold;
            margin-bottom: 1;
        }
        #button-row {
            width: 100%;
            height: auto;
            align: center middle;
            padding: 1 0 0 0;
        }
        #button-row Button {
            margin: 0 2;
        }
        """

        BINDINGS = [
            Binding("escape", "cancel", "Cancel", show=True),
            Binding("enter", "confirm", "Build reel", show=False),
        ]

        def __init__(
            self,
            rows_data: List[BrowseRow],
            participant_headers: List[str],
            categories: List[str],
        ):
            super().__init__()
            self._rows_data = rows_data
            self._participant_headers = participant_headers
            self._categories = categories

        def compose(self) -> ComposeResult:
            yield Header()
            with Vertical(id="reel-root"):
                yield Static("Build your reel visually:", id="reel-title")
                with Horizontal(id="selectors"):
                    with VerticalScroll(classes="selector-column"):
                        yield Static("Participants")
                        yield SelectionList(
                            *[(pid, pid) for pid in self._participant_headers],
                            id="participants",
                        )
                    with VerticalScroll(classes="selector-column"):
                        yield Static("Rows")
                        row_items = [
                            (
                                f"{row['row_num']}: [{row['category']}] {row['description']}",
                                str(row['row_num']),
                            )
                            for row in self._rows_data
                        ]
                        yield SelectionList(*row_items, id="rows")
                    with VerticalScroll(classes="selector-column"):
                        yield Static("Categories")
                        yield SelectionList(
                            *[(cat, cat) for cat in self._categories],
                            id="categories",
                        )
                with Horizontal(id="button-row"):
                    yield Switch(value=False, id="batch-switch")
                    yield Label("Include all clips (batch)")
                    yield Switch(value=False, id="filter-switch")
                    yield Label("Key-marked only (filter)")
                    yield Switch(value=False, id="timeline-switch")
                    yield Label("Timeline (exactly one participant)")
                    yield Button("Build reel", variant="primary", id="build")
                    yield Button("Cancel", variant="default", id="cancel")
            yield Footer()

        def action_cancel(self) -> None:
            self.exit(None)

        def action_confirm(self) -> None:
            self._finish()

        def on_button_pressed(self, event: Button.Pressed) -> None:
            if event.button.id == "build":
                self._finish()
            else:
                self.action_cancel()

        def _finish(self) -> None:
            participants_sl = self.query_one("#participants", SelectionList)
            rows_sl = self.query_one("#rows", SelectionList)
            categories_sl = self.query_one("#categories", SelectionList)
            batch_sw = self.query_one("#batch-switch", Switch)
            filter_sw = self.query_one("#filter-switch", Switch)
            timeline_sw = self.query_one("#timeline-switch", Switch)

            participants = list(participants_sl.selected)
            rows = [int(r) for r in rows_sl.selected]
            categories = list(categories_sl.selected)

            if not (batch_sw.value or filter_sw.value or timeline_sw.value or participants or rows or categories):
                self.notify("Select at least one option for the reel.", severity="warning")
                return

            if timeline_sw.value:
                if len(participants) != 1:
                    self.notify("Timeline requires exactly one participant.", severity="warning")
                    return

            tokens: List[str] = []
            if batch_sw.value:
                tokens.append("batch")
            if filter_sw.value:
                tokens.append("filter")
            if timeline_sw.value:
                tokens.append("timeline")
            tokens.extend(str(r) for r in sorted(set(rows)))
            tokens.extend(f'"{c}"' for c in categories)
            tokens.extend(participants)

            reel_input = ", ".join(tokens)
            self.exit(reel_input)

    # -------------------------------------------------------------------

    class LongClipConfirmApp(App):
        """Styled confirmation dialog for very long clips."""

        TITLE = "Long Clip Warning"

        CSS = """
        Screen {
            align: center middle;
        }
        #dialog {
            width: 72;
            border: solid $warning;
            padding: 2 3;
        }
        #dialog-title {
            text-style: bold;
            color: $warning;
            margin-bottom: 1;
        }
        #dialog-body {
            margin-bottom: 1;
        }
        #button-row {
            width: 100%;
            align: center middle;
        }
        #button-row Button {
            margin: 0 2;
        }
        """

        BINDINGS = [
            Binding("y", "confirm", "Generate", show=True),
            Binding("n", "cancel", "Skip", show=True),
            Binding("escape", "cancel", "Skip", show=False),
        ]

        def __init__(self, duration_seconds: int):
            super().__init__()
            self._duration_seconds = duration_seconds

        def compose(self) -> ComposeResult:
            minutes = self._duration_seconds // 60
            seconds = self._duration_seconds % 60
            pretty_duration = f"{self._duration_seconds}s ({minutes}m {seconds}s)"

            yield Header()
            with Vertical(id="dialog"):
                yield Static("Generate very long clip?", id="dialog-title")
                yield Static(
                    f"The selected segment is {pretty_duration}, which is longer than the recommended maximum.",
                    id="dialog-body",
                )
                yield Static(
                    "Long clips can be slow to process and large on disk.\nDo you still want to generate this clip?",
                )
                with Horizontal(id="button-row"):
                    yield Button("Generate", variant="warning", id="yes")
                    yield Button("Skip clip", variant="default", id="no")
            yield Footer()

        def action_confirm(self) -> None:
            self.exit(True)

        def action_cancel(self) -> None:
            self.exit(False)

        def on_button_pressed(self, event: Button.Pressed) -> None:
            if event.button.id == "yes":
                self.action_confirm()
            else:
                self.action_cancel()


# ---------------------------------------------------------------------------
# Public entry-point functions (safe to call even when Textual is unavailable)
# ---------------------------------------------------------------------------

def run_settings() -> bool:
    """Launch interactive settings editor.

    Returns True if settings were saved, False otherwise.
    Falls back (returns False) when Textual is unavailable.
    """
    if not use_textual():
        return False
    app = SettingsApp()
    return app.run() or False


def run_category_select(categories: List[str]) -> Optional[List[str]]:
    """Launch interactive category picker.

    Returns list of selected category names, or None on cancel / unavailable.
    """
    if not use_textual():
        return None
    app = CategorySelectApp(categories)
    return app.run()


def run_browse(rows_data: List[BrowseRow], participant_headers: List[str], title: str = "") -> None:
    """Launch interactive spreadsheet browser.

    No-op when Textual is unavailable.
    """
    if not use_textual():
        return
    app = BrowseApp(rows_data, participant_headers, title)
    app.run()


def run_spreadsheet_select(doc_list: List[str]) -> Optional[str]:
    """Launch spreadsheet selector. Returns chosen name or None on cancel/unavailable."""
    if not use_textual():
        return None
    app = SpreadsheetSelectApp(doc_list)
    return app.run()


def run_mode_select() -> Optional[str]:
    """Launch mode selection menu. Returns canonical mode or 'custom'."""
    if not use_textual():
        return None
    app = ModeSelectApp()
    return app.run()


def run_reel_builder(
    rows_data: List[BrowseRow],
    participant_headers: List[str],
    categories: List[str],
) -> Optional[str]:
    """Launch visual reel builder. Returns reel selector string or None."""
    if not use_textual():
        return None
    app = ReelBuilderApp(rows_data, participant_headers, categories)
    return app.run()


def confirm_long_clip(duration_seconds: int) -> bool:
    """Ask user to confirm generation of a very long clip.

    Returns True when user confirms, False otherwise. Falls back to False when
    Textual is unavailable; caller should handle CLI confirmation separately.
    """
    if not use_textual():
        return False
    app = LongClipConfirmApp(duration_seconds)
    return bool(app.run())
