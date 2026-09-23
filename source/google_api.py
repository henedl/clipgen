"""Google Sheets API integration for clipgen."""

from __future__ import annotations

import time

from collections.abc import Callable
from typing import TYPE_CHECKING, Any, TypeVar

import config
import profiling
import utils

if TYPE_CHECKING:
    import gspread

_T = TypeVar("_T")


def _is_transient_api_error(exc: gspread.exceptions.APIError) -> bool:
    """Return True if the APIError is worth retrying (5xx or rate limit)."""
    try:
        code = exc.response.status_code
    except AttributeError:
        return False
    return code is not None and (code >= 500 or code == 429)


def _call_with_api_retry(fn: Callable[[], _T], operation: str) -> _T:
    """Call *fn*, retrying on transient Google API errors with exponential backoff.

    Profiled as ``sheets.<operation>``. PERFORMANCE.md's first rule is that these
    calls are precious and rate-limited, and AGENTS.md warns that hitting the
    limit "can appear as bugs (e.g. silently skipping timestamps)" — the *count*
    is the invariant worth watching, not the duration. It is what makes
    ``build_sheet_context``'s "makes exactly one API call" docstring claim
    self-reporting rather than prose.

    Backoff sleep is recorded under its own label so waiting is never confused
    with working: a slow ``sheets.*`` total plus a large ``sheets.backoff_sleep``
    means throttling, not a slow sheet.
    """
    import gspread

    max_retries = config.GOOGLE_API_MAX_RETRIES
    for attempt in range(max_retries + 1):
        try:
            with profiling.span(f"sheets.{operation}"):
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
            profiling.count("sheets.retry")
            profiling.add("sheets.backoff_sleep", float(delay))
            time.sleep(delay)
    raise RuntimeError(f"Google API {operation} failed after retries")


def get_worksheet(
    spreadsheet: gspread.Spreadsheet, preferred_name: str | None = None
) -> gspread.Worksheet:
    """Get a worksheet from a spreadsheet.

    Honors an explicit *preferred_name* when it exists in the spreadsheet,
    otherwise falls back to ``config.WORKSHEET_PRIORITY`` order, then the first
    worksheet (index 0). Selection logic is shared with the Excel path via
    ``utils.pick_worksheet_title``.

    Args:
        spreadsheet: A gspread Spreadsheet object
        preferred_name: A worksheet title chosen by the user, if any

    Returns:
        A gspread Worksheet object
    """
    import gspread

    # Both round-trips bypass _call_with_api_retry, so profile them here; the Sheets counter must be complete.
    with profiling.span("sheets.worksheets"):
        worksheet_titles = [ws.title for ws in spreadsheet.worksheets()]

    utils.debug_print(f"Available worksheets: {worksheet_titles}")

    chosen = utils.pick_worksheet_title(worksheet_titles, preferred_name)
    if chosen is None:
        # Empty spreadsheet - shouldn't happen but handle it.
        raise gspread.WorksheetNotFound("Spreadsheet contains no worksheets")
    utils.standard_print(f"Using worksheet: {chosen}")
    with profiling.span("sheets.worksheet"):
        return spreadsheet.worksheet(chosen)


def get_all_spreadsheet_meta(connection: gspread.Client) -> list[dict[str, str]]:
    """Return metadata for all accessible Google Spreadsheets.

    ``list_spreadsheet_files()`` already carries ``id``/``name``/``createdTime``/
    ``modifiedTime`` for every file in a single (paged) Drive call, so exposing
    the last-edit time costs no extra API round-trips. Retries on transient
    Google API errors (429, 5xx) with exponential backoff.

    Args:
        connection: Google client connection object

    Returns:
        List of ``{"name", "id", "modifiedTime"}`` dicts (newest fields empty
        when Drive omits them).
    """

    def fetch() -> list[dict[str, str]]:
        spreadsheet_files = list(connection.list_spreadsheet_files())
        metas: list[dict[str, str]] = []
        for doc in spreadsheet_files:
            utils.debug_print(str(doc))
            metas.append(
                {
                    "name": doc.get("name", ""),
                    "id": doc.get("id", ""),
                    "modifiedTime": doc.get("modifiedTime", ""),
                }
            )
        return metas

    return _call_with_api_retry(fetch, "list_spreadsheet_files")


def get_all_spreadsheets(connection: gspread.Client) -> list[str]:
    """Returns list of all accessible Google Spreadsheet names.

    Thin wrapper over :func:`get_all_spreadsheet_meta` for callers (the CLI
    selectors) that only need bare names.

    Args:
        connection: Google client connection object

    Returns:
        List of spreadsheet name strings
    """
    return [m["name"] for m in get_all_spreadsheet_meta(connection)]


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
        config.debug_ic(search_name)
    utils.debug_print("Running method find_spreadsheet_by_name()")
    search_name = search_name.strip().lower()
    search_name_guess = search_name + " data set"
    utils.debug_print(
        f"Using search_name '{search_name}', search_name_guess '{search_name_guess}'"
    )

    for i, doc in enumerate(doc_list):
        doc_name = doc.strip().lower()
        if config.DEBUGGING:
            config.debug_ic(doc_name, search_name)
        utils.debug_print(
            f"Attempting exact match with '{doc}', formatted as '{doc_name}'"
        )
        if doc_name == search_name:
            utils.debug_print(f"Matched sheet '{doc_name}' with input '{search_name}'")
            if config.DEBUGGING:
                config.debug_ic(i)
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
                config.debug_ic(i)
            return i
        utils.debug_print(f"Found no guess match at step {i}")

    if config.DEBUGGING:
        config.debug_ic(-1)
    return -1
