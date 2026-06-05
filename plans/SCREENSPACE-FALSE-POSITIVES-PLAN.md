# Screenspace: Reduce False Positives

## Context

The Screenspace analysis subsystem over-triggers on compressed video — most painfully in the **text** and **numbers** OCR tools (low resolution + naive fuzzy matching). Compute is constrained (the tool runs on laptops) and every source/game looks different, so the strategy is to give users **levers** (confidence gating, preprocessing, allowlists, normalization, consecutive-frame requirements) and **visibility** (real confidence on events, a distribution histogram) rather than chase a universal detector.

This is a multi-phase effort ordered by impact-per-LOC. Each phase is independently shippable.

## Status

| Phase | Summary | State |
| --- | --- | --- |
| 1 | Gate OCR text/numbers by detection confidence | **Done** (`37f00a4`) |
| 2 | Opt-in OCR ROI preprocessing (upscale + CLAHE) | **Done** (`df35431`) |
| 3a | Numbers tool: EasyOCR digit allowlist | Planned |
| 3b | Text tool: opt-in character normalization | Planned |
| 5 | Extract hardcoded static-frame-skip threshold | Planned |
| 4 | Temporal coherence (`require_consecutive`) | Planned |
| 6 | Confidence histogram in results UI | Planned |

Recommended ship order for the remainder: **3a → 3b → 5 → 4 → 6** (trivial/low-risk first; temporal coherence and the histogram last).

## Conventions established in Phases 1–2 (reuse these)

- **Two OCR call sites per tool**: the standalone `scan_text()` / `scan_numbers()` *and* the `TextTool.check_frame` / `NumbersTool.check_frame` paths (used by multitool). Every per-task OCR param must touch **both**, plus forward from `TextTool.scan` / `NumbersTool.scan`.
- **Frontend plumbing points** for any new per-task param (5 spots): `renderTextParams`/`renderNumbersParams` (standalone editor), `_mtRenderText`/`_mtRenderNumbers` (multitool steps), `gatherWorkflowParams`, `restoreTaskToWorkflow` (edit-rehydrate), `snapshotMultitoolStepValues`. Add a `PARAM_DESCRIPTIONS` entry for the tooltip (both the standalone label and the shorter multitool label).
- **Multitool row helpers**: `_mtAddNumberRow`, `_mtAddCheckboxRow` (added in Phase 2).
- **Mirrored constants** (values JS also needs) flow through `utils.get_frontend_config()` → `CLIPGEN_CONFIG` + `clipgenApplyConfig` in `assets/web/utils.js`; the contract is asserted in `tests/test_shared_constants.py` (both the `CLIPGEN_CONFIG` match test and `test_get_frontend_config_shape`). Constants used **only** in Python (e.g. `SCREENSPACE_OCR_MIN_HEIGHT`) do not need mirroring.
- **"0/absent → config default"** idiom: `params.get("x") or config.DEFAULT` in `check_frame`; `if x <= 0: x = config.DEFAULT` in standalone scans. Keeps both paths gating identically.
- **No backwards-compat layers** (AGENTS.md hard rule): new task-param keys and event fields just change the shape; users re-run. Update tests to the new shape.
- All UI is vanilla JS (ES5 `.then()` chaining); use `assets/web/tokens.css` design tokens for any new spacing/color.

Critical files throughout: `screenspace.py`, `assets/web/screenspace.{html,js,css}`, `config.py`, `utils.py`, `assets/web/utils.js`, `tests/test_screenspace.py`, `tests/test_shared_constants.py`.

---

## Phase 3a — Numbers tool: EasyOCR digit allowlist

Numbers mode is digits-only by definition, so constraining EasyOCR's character set kills a large class of misreads (O↔0, S↔5, l↔1) at the source.

- **Backend**: in `scan_numbers._cb` and `NumbersTool.check_frame`, pass `allowlist="0123456789.,-"` to `reader.readtext(...)`. Gate behind `if languages == ["en"]` initially — some EasyOCR language combos reject `allowlist`. No UI; always-on for numbers in English.
- **Manifest**: none.
- **Tests**: `TestScanNumbers::test_allowlist_passed` — stub reader records kwargs; assert `allowlist` is forwarded when `languages == ["en"]` and omitted otherwise.
- **Risk**: an allowlist tighter than the parsing regex could occasionally drop a character that helped EasyOCR localize a digit (e.g. `%` in "30%"). Worth a quick real-footage check before widening the set.
- **Effort**: S.

## Phase 3b — Text tool: opt-in character normalization

For the text tool, let users opt into collapsing common OCR confusions before the fuzzy compare.

- **Backend**: module constant `_OCR_NORMALIZATION_TABLE = str.maketrans({"o":"0","l":"1","i":"1","|":"1","s":"5","b":"8"})` and helper `_normalize_ocr_text(s)` (lowercase + translate). In `scan_text` and `TextTool.check_frame`, when `params.ocr_normalize` is true, compute the fuzzy ratio on normalized search + OCR text. New `ocr_normalize: bool = False` kwarg on `scan_text`.
- **Frontend**: `Normalize characters` checkbox (`paramTextOcrNormalize`) in `renderTextParams` + multitool, via the established plumbing; tooltip "Collapse O→0, l→1, S→5 before matching (helps fonts with weak letter/digit distinction)". Numbers tool: no UI (handled by 3a's allowlist).
- **Manifest**: text task gains `ocr_normalize: bool`.
- **Tests**: `TestScanText::test_normalize_o_for_zero` — stub yields `"l00"`, search `"100"`; matches only when `ocr_normalize=True`.
- **Risk**: over-fires on words containing O/I/l when the search contains digits. Default off; user's tradeoff.
- **Effort**: S.

## Phase 5 — Extract the static-frame-skip threshold

Four scan loops hardcode a `< 2.0` mean-abs-diff "skip near-identical frame" check (similarity, text, numbers, scene). Make it tunable.

- **Backend**: `config.SCREENSPACE_STATIC_FRAME_SKIP_THRESHOLD: float = 2.0`; replace the four literals. Add to `SETTINGS_DESCRIPTIONS` and `STUDIO_SETTINGS` (range 0.5–10.0) so noisy footage can skip more aggressively and subtle-change footage can skip less.
- **Explicitly NOT changing**: `SCREENSPACE_PHASH_THRESHOLD = 15` — every existing task is calibrated against it; defer until data shows it's the bottleneck. Note this in the PR.
- **Tests**: `test_static_skip_uses_config` — source-scan `screenspace.py` to assert the four sites reference the constant, not a literal.
- **Risk**: a Studio setting needs a crisp description to avoid confusion.
- **Effort**: S.

## Phase 4 — Temporal coherence (`require_consecutive`)

Single-frame matches at multi-second sampling are inherently flaky on compressed video. Let users require N consecutive sampled matches before an event fires.

- **Backend**: a small `_ConsecutiveBuffer` helper (push → emit only after N consecutive matches; emit the **median** timestamp of the window so the event centers on the run; any miss clears the buffer). Wire into `scan_text`, `scan_numbers`, `scan_changes`, `scan_flow` via `require_consecutive: int = 1` (1 = today's behavior). Multitool's per-frame `check_frame` can't buffer cheaply — document that `require_consecutive` is honored at the standalone-scan level only; multitool's AND-chaining is the cross-tool noise filter there.
- **Frontend**: an "Advanced" numeric `paramXxxConsecutive` (1–10, default 1) in the four tools' editors via the standard plumbing.
- **Manifest**: task gains `require_consecutive: int` (omit when 1).
- **Tests**: `TestConsecutiveBuffer` — emits only after N; median timestamp; miss clears; N=1 emits immediately.
- **Risk**: delays detection by `n × interval`; can miss brief flashes. Default 1 → opt-in.
- **Effort**: M.

## Phase 6 — Confidence histogram in results UI

Give users a sightline on the confidence distribution before they move the existing certainty-cutoff slider.

- **Frontend**: `<div id="confHistogram">` next to `#certaintyCutoff` in `screenspace.html`; `.conf-hist*` rules in `screenspace.css` using design tokens. `renderConfidenceHistogram(events, taskType)` buckets `event.confidence` into 10 bins, renders bars (fill via `--color-task-{type}`), overlays a marker line at the current cutoff and updates it on the slider's input handler. Hook into `renderResults`. Hide for tools with degenerate confidence (e.g. timelapse).
- **Backend/manifest**: none (events already carry `confidence`, including numbers as of Phase 1).
- **Tests**: none required (pure frontend); manual verification — bars match per-bucket counts, marker tracks the slider.
- **Risk**: pure UX polish; ship only if Phases 1–5 leave a real gap.
- **Effort**: M (mostly CSS/render).

---

## Verification (per phase)

1. `uv run --extra dev pytest -c tests/pytest.ini` green; new tests cover new behavior.
2. `/check` (ruff format → ruff check → ty).
3. Manual: launch `uv run clipgen.py --screenspace -i DIR -o DIR`, open `http://127.0.0.1:8089/screenspace/`, exercise the new control on real compressed footage, and confirm default-off settings reproduce prior output (regression check).
