# clipgen – Project context for AI assistants

## Project overview

clipgen is a Python CLI tool that generates video clips from timestamps stored in a Google Sheet or a local Excel file. It uses **gspread** for the Google Sheets API, **openpyxl** for Excel, and **ffmpeg** for video cutting and concatenation. The target audience is UX Researchers and professionals who manage playtest videos locally.

**Data flow:** Timestamps in spreadsheet → clipgen reads records (descriptions, study name, participant IDs, categories) → ffmpeg → individual clips or a single highlight reel.

## Architecture

| File | Role |
| ------ | ------ |
| [clipgen.py](clipgen.py) | Main entry point, CLI parsing, spreadsheet selection, mode dispatch, clip/reel processing |
| [spreadsheet.py](spreadsheet.py) | Spreadsheet parsing, header validation, timestamp generation for all modes (~1266 lines) |
| [video.py](video.py) | ffmpeg operations: cut clips, concatenate, compress to target size, duration/bitrate helpers |
| [files.py](files.py) | Filename handling (unique names, truncation), `prepare_clip()` (parse timestamps, sanitize desc/category) |
| [utils.py](utils.py) | Timestamp parsing, argument parsing, `error_print` / `warning_print` / `verbose_print` / `info_print` |
| [config.py](config.py) | Global constants and settings (version, headers, limits, commands) |
| [google_api.py](google_api.py) | Google Sheets auth, worksheet selection by priority, spreadsheet listing/search |
| [excel_io.py](excel_io.py) | Excel adapter: `ExcelSheetAdapter` mimics gspread Worksheet interface for local .xlsx |

## Key data structures

**Clip record** (built in spreadsheet layer, enriched in files):

```python
{
    'cell': gspread.Cell,      # Cell with timestamp value (1-based row/col)
    'desc': str,               # Observation text from Observation column
    'study': str,              # Normalized study name (filesystem-safe)
    'participant': str,        # e.g. 'P01', 'G02' from header
    'category': str,           # Row category (sanitized; empty → 'uncategorized')
    'times': [(start, end)]    # Added by files.prepare_clip() – list of (start_time, end_time) strings
}
```

Source video filenames follow `{study}_{participant}.mp4` (e.g. `mystudy_P01.mp4`).

## Conventions and patterns

- **Coordinates:** gspread uses **1-based** row/col. `sheet.get_all_values()` is a list of lists with **0-based** indices: `sheet_data[row_idx][col_idx]`. Conversions: sheet row = `row_idx + 1`, sheet col = `col_idx + 1`.
- **Timestamps:** Formats `MM:SS` or `HH:MM:SS`. Ranges with `-` (e.g. `1:23-1:45`). Multiple pairs separated by `,`, `;`, `+`, or space. Single time gets end = start + `DEFAULT_DURATION_SECONDS`.
- **Participant IDs:** Headers must start with `P` (individual) or `G` (group); see `config.PARTICIPANT_PREFIXES`.
- **User feedback:** Use `utils.error_print()`, `utils.warning_print()`, `utils.verbose_print()`, `utils.info_print()`. Do not `print()` directly for user-facing messages.
- **Debug:** Set `config.DEBUGGING = True` to enable icecream output and to skip actual ffmpeg calls in [video.py](video.py).

## Modes

| Mode | Description |
| ------ | ------------- |
| **batch** | All non-empty timestamp cells in the sheet |
| **line** | One or more specific row numbers (e.g. 5, 7, 12) |
| **range** | Contiguous rows from start to end (inclusive) |
| **category** | Rows matching selected category names |
| **cell** | Specific cells as `participant.row` (e.g. P01.11, P03.11) |
| **participant** | All clips for one or more participants (e.g. P01, P03) |
| **reel** | Mixed selectors (lines, ranges, categories, cells, participants) combined into one video; deduped by cell, sorted by row then col |
| **browse** | Interactive view of spreadsheet rows (no clip generation) |

Reel selectors: `batch`, line numbers, ranges like `13-16`, quoted categories, cells like `P01.11`, participant IDs like `P01`.

## Important configuration ([config.py](config.py))

- `WORKSHEET_PRIORITY` – Worksheet names tried first (e.g. 'Sheet1', 'Data', 'Observations')
- `ID_HEADER`, `OBSERVATION_HEADER`, `CATEGORY_HEADER` – Required column headers
- `PARTICIPANT_PREFIXES` – `('P', 'G')`
- `FILEFORMAT` – `.mp4`
- `MAX_CLIP_DURATION_SECONDS` – 600 (10 min); prompts before generating longer clips
- `DEFAULT_DURATION_SECONDS` – 60 (used when only start time is given)
- `MAX_FILENAME_LENGTH` – 255
- `COMMAND_LIST_ALL`, `COMMAND_LIST_NEW`, `COMMAND_OPEN_LAST`, `COMMAND_EXCEL`, `COMMAND_HTTP_PREFIX`, `COMMAND_SETTINGS` – Interactive spreadsheet selection commands

## Spreadsheet layout

- Row 0 (A1): Study name (optional; falls back to spreadsheet title).
- Header row: Must include columns **ID**, **Observation**, **Category** (exact names from config).
- Participant columns: Immediately after ID; headers start with P or G (e.g. P01, P02, G01). Each holds timestamp strings; non-empty cells become clip candidates.
- Observation column: Human-readable description per row.
- Category column: Label per row for category/reel selection.

Reference spreadsheet layout is described in [README.md](README.md).

## Version

- The version is stored as `VERSIONNUM` in [config.py](config.py) (e.g. `'0.7.4'`).
- **When making substantive code changes** (bug fixes or features), increment the **last segment only** (patch) in `config.py`, e.g. `0.7.4` → `0.7.5`. Do not bump for docs-only, comment-only, or refactor-only changes unless they affect user-visible behavior.

## Testing notes

- There is no test suite in the repo.
- With `config.DEBUGGING = True`, icecream is enabled and [video.py](video.py) does not invoke ffmpeg (returns without writing files where applicable).
