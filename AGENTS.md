# Project learnings for agents

## Learned User Preferences

- When the user attaches an implementation plan that already has created todos, do not edit the plan file; mark those todos in_progress as you work and complete them without recreating the list.
- Prefer inlining functions rather than creating yet another helper.
- Don't extract helpers unless it is called more than once.
- Prefer placing generic index/letter conversion utilities (e.g. index_to_letter, letter_to_index) in utils.py rather than in domain-specific modules like files.py.
- Prefer minimal, focused edits over broad rewrites.
- Prefer naming new helpers to match existing method naming patterns in the same module.
- Never write a class when a function will do.
- No comments on obvious code.
- Treat spreadsheet layout and timestamp semantics as domain rules; if tests conflict with these, reconsider or adjust the tests rather than changing core semantics to satisfy them.

## Learned Workspace Facts

- Baseline time row placement in the spreadsheet layer is tied to header/`id_cell` row math (e.g. offsets from `id_cell.row`); changing that offset without aligning tests and sheet layout has broken baseline timestamp handling before.
- When making substantive code changes (fixes or features), increment the patch (last number) of VERSIONNUM in config.py.
- Interactive prompts use a keyword-aware helper: `quit`/`exit` exit the program, `top` returns to spreadsheet selection, and `back` returns to mode selection (or spreadsheet selection when already at mode selection).
- Textual-based TUI support (tui.py, TEXTUAL_TUI) has been removed; prefer CLI prompts and the HTML timeline viewer for interactive features.
- Browse mode scrolling is controlled via `BROWSE_LINES_TO_SCROLL` in `config.py`, with a default of 5 rows per up/down step.
- The program runs everything in sequence, no multithreading. Implementing multithreading was too much a headache, though the performance upside was notable; shaving up to 30% of runtimes with 4 threads.
