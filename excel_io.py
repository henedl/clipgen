# -*- coding: utf-8 -*-
"""Local Excel (.xlsx) support for clipgen.

Provides a sheet adapter that matches the gspread Worksheet interface
so spreadsheet.py and clipgen can use local Excel files the same way as Google Sheets.
"""
import os
from typing import Any, List, Optional

import openpyxl

import config
import utils


def _cell_value_to_str(value: Any) -> str:
    """Convert a cell value to string for consistency with gspread."""
    if value is None:
        return ''
    return str(value).strip() if isinstance(value, str) else str(value)


class _CellLike:
    """Minimal cell-like object with .row and .col (1-based) for header lookup."""
    __slots__ = ('row', 'col')

    def __init__(self, row: int, col: int) -> None:
        self.row = row
        self.col = col


class _SpreadsheetLike:
    """Minimal spreadsheet-like object with .title and .url = None for Excel."""
    __slots__ = ('title', 'url')

    def __init__(self, title: str) -> None:
        self.title = title
        self.url = None


class ExcelSheetAdapter:
    """Adapter that makes an openpyxl worksheet look like a gspread Worksheet."""

    def __init__(self, ws: Any, workbook_path: str) -> None:
        self._ws = ws
        self.title = getattr(ws, 'title', os.path.splitext(os.path.basename(workbook_path))[0])
        self._workbook_path = workbook_path
        self._data: List[List[str]] = []
        self._load_data()
        self.spreadsheet = _SpreadsheetLike(
            title=os.path.splitext(os.path.basename(workbook_path))[0]
        )

    def _load_data(self) -> None:
        """Load all cell values into _data as List[List[str]], padded to max column."""
        rows: List[List[str]] = []
        max_col = 0
        for row in self._ws.iter_rows(values_only=True):
            str_row = [_cell_value_to_str(v) for v in row]
            rows.append(str_row)
            max_col = max(max_col, len(str_row))
        # Pad rows to same length
        for row in rows:
            while len(row) < max_col:
                row.append('')
        self._data = rows

    def find(self, text: str) -> Optional[_CellLike]:
        """Find first cell with exact match. Returns cell-like with .row, .col (1-based)."""
        for row_idx, row in enumerate(self._data):
            for col_idx, cell_value in enumerate(row):
                if cell_value == text:
                    return _CellLike(row=row_idx + 1, col=col_idx + 1)
        return None

    def get_all_values(self) -> List[List[str]]:
        """Return all sheet data as list of rows (list of strings)."""
        return self._data

    def row_values(self, row_1based: int) -> List[str]:
        """Return one row as list of strings. row_1based is 1-based."""
        idx = row_1based - 1
        if 0 <= idx < len(self._data):
            return self._data[idx]
        return []

    @property
    def col_count(self) -> int:
        """Number of columns (max length of any row)."""
        if not self._data:
            return 0
        return max(len(row) for row in self._data)


def _get_worksheet_from_workbook(wb: Any) -> Any:
    """Pick worksheet by config.WORKSHEET_PRIORITY, else active/first."""
    sheet_names = wb.sheetnames
    utils.debug_print(f'Available worksheets: {sheet_names}')
    for priority_name in config.WORKSHEET_PRIORITY:
        if priority_name in sheet_names:
            utils.verbose_print(f'Using worksheet: {priority_name}')
            return wb[priority_name]
    if sheet_names:
        first = wb[sheet_names[0]]
        utils.verbose_print(f'No matching worksheet found. Using first worksheet: {first.title}')
        return first
    raise ValueError('Workbook contains no worksheets')


def open_excel_workbook(path: str) -> Optional[ExcelSheetAdapter]:
    """Load an .xlsx file and return an ExcelSheetAdapter for the chosen worksheet.

    Args:
        path: Path to the .xlsx file.

    Returns:
        ExcelSheetAdapter, or None on error (prints error).
    """
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        utils.error_print(f"Excel file not found: {path}")
        return None
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
        ws = _get_worksheet_from_workbook(wb)
        adapter = ExcelSheetAdapter(ws, path)
        wb.close()
        return adapter
    except Exception as e:
        utils.error_print(f"Could not open Excel file {path}: {e}")
        return None


def list_excel_in_cwd() -> List[str]:
    """Return list of .xlsx file paths in the current working directory (case-insensitive extension)."""
    cwd = os.getcwd()
    result = []
    for name in os.listdir(cwd):
        if name.lower().endswith('.xlsx'):
            result.append(os.path.join(cwd, name))
    return sorted(result)


def select_excel_file() -> Optional[ExcelSheetAdapter]:
    """Discover .xlsx in cwd: 0 -> error; 1 -> open and return; 2+ -> list and prompt.

    Returns:
        ExcelSheetAdapter on success, None otherwise.
    """
    paths = list_excel_in_cwd()
    if not paths:
        utils.error_print('No .xlsx files found in the current directory.',
            ['Place one or more Excel files (.xlsx) in the working directory.'])
        return None
    if len(paths) == 1:
        utils.verbose_print(f'Opening Excel file: {os.path.basename(paths[0])}')
        return open_excel_workbook(paths[0])
    # Multiple files: list and prompt
    utils.info_print('\nExcel files in current directory:')
    for i, p in enumerate(paths, 1):
        utils.info_print(f'  {i}. {os.path.basename(p)}')
    while True:
        choice = input('\nEnter index (1-based) or filename to open (or Enter to cancel):\n>> ').strip()
        if not choice:
            return None
        # Try as index
        if choice.isdigit():
            idx = int(choice)
            if 1 <= idx <= len(paths):
                return open_excel_workbook(paths[idx - 1])
            utils.info_print(f'Invalid index. Enter a number between 1 and {len(paths)}.')
            continue
        # Try as filename
        for p in paths:
            if os.path.basename(p) == choice:
                return open_excel_workbook(p)
        utils.info_print(f'No file named "{choice}". Enter an index or exact filename.')
