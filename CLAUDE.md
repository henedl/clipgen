# clipgen

Video clip generator for UX researchers. Extracts clips from video recordings based on timestamps stored in Google Sheets or Excel files.

## Key Files

- `clipgen.py` - Main entry point, orchestrates modes and processing
- `spreadsheet.py` - Core data parsing for timestamps and clip specifications
- `video.py` - FFmpeg integration for video cutting
- `config.py` - Configuration constants (headers, formats, limits)
- `utils.py` - Argument parsing, logging, timestamp utilities
- `google_api.py` - Google Sheets API integration
- `excel_io.py` - Local Excel support (adapter pattern for gspread compatibility)
- `files.py` - Filename generation and file operations

## Running

```bash
python clipgen.py           # Interactive mode
python clipgen.py -b        # Batch: all clips
python clipgen.py -l 5      # Line: specific row
python clipgen.py -r 1-10   # Range: rows 1-10
python clipgen.py -p P01    # Participant: all clips for P01
python clipgen.py -R "1-5"  # Reel: combined highlight video
python clipgen.py -v        # Verbose output
python clipgen.py -y        # Skip confirmations
```

## Dependencies

**Python:** gspread, openpyxl, icecream (see requirements.txt)

**External:** FFmpeg and FFprobe must be installed and in PATH

**Auth:** Google Sheets requires credentials.json (OAuth2) in working directory or ~/.gspread

## Key Patterns

- **Mode-based processing:** batch, line, range, category, cell, participant, reel
- **Clip dict structure:** `{cell, desc, study, participant, category, times: [(start, end), ...]}`
- **Spreadsheet headers:** ID, Observation, Category, Time (optional for timestamp conversion)
- **Participant columns:** Prefixed with P (individual) or G (group), e.g., P01, G01
