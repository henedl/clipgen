# -*- coding: utf-8 -*-
"""Google Sheets API integration for clipgen."""

import time

from collections.abc import Callable
from typing import Any, TypeVar

import gspread
from icecream import ic

import config
import utils

_T = TypeVar("_T")


def _is_transient_api_error(exc: gspread.exceptions.APIError) -> bool:
    """Return True if the APIError is worth retrying (5xx or rate limit)."""
    try:
        code = exc.response.status_code
    except AttributeError:
        return False
    return code is not None and (code >= 500 or code == 429)


def _call_with_api_retry(fn: Callable[[], _T], operation: str) -> _T:
    """Call *fn*, retrying on transient Google API errors with exponential backoff."""
    max_retries = config.GOOGLE_API_MAX_RETRIES
    for attempt in range(max_retries + 1):
        try:
            return fn()
        except gspread.exceptions.APIError as e:
            if not _is_transient_api_error(e) or attempt == max_retries:
                raise
            delay = 2 ** (attempt + 1)
            utils.warning_print(
                f"Google API error during {operation} "
                f"(attempt {attempt + 1}/{max_retries + 1}): {e}. "
                f"Retrying in {delay}s..."
            )
            time.sleep(delay)
    raise RuntimeError(f"Google API {operation} failed after retries")


def get_worksheet(spreadsheet: gspread.Spreadsheet) -> gspread.Worksheet:
    """Get a worksheet from a spreadsheet using priority-based name matching.

    Tries to find a worksheet matching names in WORKSHEET_PRIORITY order.
    If no match is found, returns the first worksheet (index 0).

    Args:
        spreadsheet: A gspread Spreadsheet object

    Returns:
        A gspread Worksheet object
    """
    # Get all worksheet titles from the spreadsheet
    worksheets = spreadsheet.worksheets()
    worksheet_titles = [ws.title for ws in worksheets]

    utils.debug_print(f"Available worksheets: {worksheet_titles}")

    # Try each name in priority order
    for priority_name in config.WORKSHEET_PRIORITY:
        if priority_name in worksheet_titles:
            utils.standard_print(f"Using worksheet: {priority_name}")
            return spreadsheet.worksheet(priority_name)

    # No match found - use first worksheet
    if worksheets:
        first_sheet = worksheets[0]
        utils.standard_print(
            f"No matching worksheet found. Using first worksheet: {first_sheet.title}"
        )
        return first_sheet

    # This shouldn't happen, but handle empty spreadsheet case
    raise gspread.WorksheetNotFound("Spreadsheet contains no worksheets")


def get_all_spreadsheets(connection: gspread.Client) -> list[str]:
    """Returns list of all accessible Google Spreadsheet names.

    Retries on transient Google API errors (429, 5xx) with exponential backoff.

    Args:
        connection: Google client connection object

    Returns:
        List of spreadsheet name strings
    """

    def fetch() -> list[str]:
        spreadsheet_files = list(connection.list_spreadsheet_files())
        for doc in spreadsheet_files:
            utils.debug_print(str(doc))
        return [doc["name"] for doc in spreadsheet_files]

    return _call_with_api_retry(fetch, "list_spreadsheet_files")


def get_sheet_values(sheet: Any) -> list[list[str]]:
    """Return all worksheet cell values, retrying on transient Google API errors."""
    return _call_with_api_retry(sheet.get_all_values, "get_all_values")


def find_spreadsheet_by_name(search_name: str, doc_list: list[str]) -> int:
    """Find a matching Google Sheet name from doc_list.

    Args:
        search_name: Name to search for
        doc_list: List of spreadsheet names

    Returns:
        The index of matching sheet, or -1 if not found.
    """
    if config.DEBUGGING:
        ic(search_name)
    utils.debug_print("Running method find_spreadsheet_by_name()")
    search_name = search_name.strip().lower()
    search_name_guess = search_name + " data set"
    utils.debug_print(
        f"Using search_name '{search_name}', search_name_guess '{search_name_guess}'"
    )

    for i, doc in enumerate(doc_list):
        doc_name = doc.strip().lower()
        if config.DEBUGGING:
            ic(doc_name, search_name)
        utils.debug_print(
            f"Attempting exact match with '{doc}', formatted as '{doc_name}'"
        )
        if doc_name == search_name:
            utils.debug_print(f"Matched sheet '{doc_name}' with input '{search_name}'")
            if config.DEBUGGING:
                ic(i)
            return i
        utils.debug_print(f"Found no exact match at step {i}")

    for i, doc in enumerate(doc_list):
        doc_name = doc.strip().lower()
        utils.debug_print(
            f"Attempting guess match with '{doc}', formatted as '{doc_name}'"
        )
        if doc_name == search_name_guess:
            utils.debug_print(
                f"Matched sheet '{doc_name}' with guess '{search_name_guess}'"
            )
            if config.DEBUGGING:
                ic(i)
            return i
        utils.debug_print(f"Found no guess match at step {i}")

    if config.DEBUGGING:
        ic(-1)
    return -1
