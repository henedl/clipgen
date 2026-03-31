# Project learnings for agents

## Learned User Preferences

- When the user attaches an implementation plan that already has created todos, do not edit the plan file; mark those todos in_progress as you work and complete them without recreating the list.
- Don't extract helpers unless it is called more than once.
- Prefer placing generic index/letter conversion utilities (e.g. index_to_letter, letter_to_index) in utils.py rather than in domain-specific modules like files.py.
- Prefer naming new helpers to match existing method naming patterns in the same module.
- Never write a class when a function will do.
- Treat spreadsheet layout and timestamp semantics as domain rules; if tests conflict with these, reconsider or adjust the tests rather than changing core semantics to satisfy them.
- All web UIs use vanilla JavaScript (ES5-style `.then()` chaining, not async/await), hand-written CSS with CSS variables for theming, and plain HTML. No React, TypeScript, CSS frameworks, or build tools.
- **CSS design tokens**: Use tokens from `assets/web/tokens.css` for spacing (`--space-N`), font sizes (`--text-N`), border radius (`--radius-N`), shadows (`--shadow-N`), transitions (`--duration-N`), and z-index (`--z-N`). Never write raw `rem`/`px` values for these properties in new code. When editing existing CSS, convert touched values to tokens.
- Thin server, thick client: keep the Flask server focused on data/media endpoints; UI logic, state management, and rendering happen client-side.
- Plan-driven development: detailed implementation plans with specific files, line numbers, code structure, and verification steps are written before coding begins. Follow attached plans closely.
- Features are often built incrementally across multiple sessions. Check for existing groundwork before starting from scratch.
- Manifest-driven state persistence: JSON manifest files (clipgen, insights, screenspace, stashes, settings) follow the pattern of load-on-startup, save-after-mutations.
- No hardcoded version numbers in evergreen docs (CLAUDE.md, README.md). Reference `VERSIONNUM` in `config.py` instead.
- **Icons**: Prefer SVG icons from `assets/icons/` (316 Heroicons outline, kebab-case names like `pencil-square.svg`) over crafting new inline SVG paths or using text/emoji glyphs in web UIs.
- **Linting/formatting**: Run `uv run ruff check --fix && uv run ruff format` after editing Python files. Run `uv run ty check` for type checking.
- Commit early and commit often, so we can roll back changes more easily.
- If a problem is reoccurring and survives fix attempts, check git logs for clues.
- Never edit .gitignore automatically, always confirm changes to this file with the user.

## Learned Workspace Facts

- Baseline time row placement in the spreadsheet layer is tied to header/`id_cell` row math (e.g. offsets from `id_cell.row`); changing that offset without aligning tests and sheet layout has broken baseline timestamp handling before.
- When making substantive code changes (fixes or features), increment the patch (last number) of VERSIONNUM in config.py.
- Interactive prompts use a keyword-aware helper: `quit`/`exit` exit the program, `top` returns to spreadsheet selection, and `back` returns to mode selection (or spreadsheet selection when already at mode selection).
- Textual-based TUI support (tui.py, TEXTUAL_TUI) has been removed; prefer CLI prompts and the HTML timeline viewer for interactive features.
- Browse mode scrolling is controlled via `BROWSE_LINES_TO_SCROLL` in `config.py`, with a default of 5 rows per up/down step.
- The program runs everything in sequence, no multithreading. Implementing multithreading was too much a headache, though the performance upside was notable; shaving up to 30% of runtimes with 4 threads.
- Always use `uv run` to execute Python commands (e.g. `uv run pytest`, `uv run clipgen.py`). This ensures the correct venv is used, even in worktrees where no `.venv` exists yet.
- Be careful about using the `generate_list()`, `sheet.find()`, `sheet.get_all_values()` methods as they are API calls to Google Sheets and are heavily rate-limited. Repeatedly calling the Google API will lead to rate-limiting without warnings, which can appear as bugs (e.g. silently skipping timestamps) and make development difficult.
