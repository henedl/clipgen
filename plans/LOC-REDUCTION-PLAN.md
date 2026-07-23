# LOC-reduction opportunities — plan

Status: **pass 1 complete; pass 2 in progress** (2026-07-23). Items **1 (1a + 1b)**, **2**, **3**,
**4**, and **5** are closed (see check-marks below); a second sweep (three parallel exploration
agents over Python / JS / CSS+tests) found the remaining items tracked in **Pass 2** at the bottom.
Investigation targets for genuinely *reducing* total lines of code (not relocating them). Each
item is sized as one or more focused `refactor:` commits. Check items off and add a "Done" note
as they land (per AGENTS.md plan-maintenance rule).

## Framing

Hub→satellite carves (see [FRONTEND-REFACTORS-PLAN.md](FRONTEND-REFACTORS-PLAN.md)) **relocate**
lines — the hub shrinks but the total is unchanged. Real reduction comes from **deduplication
and deleting redundancy**. Everything below targets repeated shapes, ordered by leverage.

Baseline (2026-06-30): Python ~40.2k lines, JS ~38.8k, CSS ~17.9k. The four Flask server
files alone are ~9.8k lines.

## Do NOT chase

- **The test suite** (~26k lines) — it's guard rails, not fat. Don't compress it.
- **Comments / docstrings** — the house style is deliberately heavy inline documentation
  (AGENTS.md). Not a target.
- **More carves** — they don't change the total; they live in the other plan.

---

## 1. Flask route boilerplate — highest leverage

The four server files (`server.py`, `screenspace_server.py`, `transcripts_server.py`,
`workflows_server.py`) span ~9.8k lines across **149 routes**, with **416 `jsonify(` calls**,
**281 `{"ok": False, "error": …}` returns**, and **113 try/excepts**. Sampling
(`screenspace_server.py` pins routes ~461–540) confirms near-identical per-route scaffolding:

```python
return jsonify({"ok": False, "error": "timestamp must be a number"}), 400
...
return jsonify({"ok": True, "pin": pin})
```

- [x] **1a. Response helpers.** **Done** — `ok(**fields)` / `err(msg, code=400)` live in the new
  Flask-only `server_utils.py` (not `utils.py`, which is deliberately Flask-free). Applied across all
  four blueprints: workflows, transcripts, screenspace, server. Non-standard envelopes left raw on
  purpose (`{"ok": False}` with no `error`; `{"ok": False, "generating": True, ...}` poll/cancelled
  payloads; `{"errors": [...]}` plural; the 202 google-auth envelope; bare `jsonify(var)`; the two
  `api_sheet` payloads `test_shared_constants` guards by source text; binary/SSE/VTT routes).
- [x] **1b. `@json_endpoint` decorator** — **Done** (added to `server_utils.py`). Catches **only**
  `ApiError` (not bare `Exception`) so it never swallows real 500s or fights routes with their own
  cleanup/`finally`. Applied conservatively: most server try/excepts catch *specific* exceptions or do
  resource cleanup (`_release_busy`), so they were left intact; the decorator currently backs the one
  `parse_number_arg` site (screenspace `api_pins_create`). It's in place for future raise-based routes.
- **Risk:** low. **Approach:** one blueprint at a time; every route already has test coverage, so
  convert + run that blueprint's tests before the next. Watch for routes with non-standard
  envelopes (not all use `{"ok": …}`) — leave those or special-case, don't force-fit.
- **Payoff (realized):** net **~−310 lines** across the four files (workflows −15, transcripts −38,
  screenspace −120, server −138) + far more scannable routes. Full suite green (1745 passed).

## 2. Request-arg parsing/validation (rides on #1)

`CODE-REVIEW.md` mandates manual `float()/int()` parsing of string route params (Flask's `<float:>`
converter 404s when JS sends ints). The "parse → on failure return error dict" block is copy-pasted
dozens of times.

- [x] **2. `parse_number_arg(raw, name, *, int_only=False, min_=None, max_=None, finite=False)`** —
  **Done** (in `server_utils.py`). Parses one numeric value and raises `ApiError(400)` (caught by 1b's
  decorator) on bad value / out-of-bounds / non-finite. Applied where the clean *error-on-failure* shape
  existed (screenspace `api_pins_create`: an isinstance + `float()` + finite/bounds block → one line).
  **Note:** most numeric parsing in these files is *silent-fallback* (`except: x = default`) or
  *returns `None`*, which deliberately must NOT raise — those sites were left as-is, so #2's reach is
  smaller than the "dozens" first estimated. The helper is in place for future error-on-failure sites.

## 3. Cross-cutting JS dedup (these *delete* lines)

Shared with [FRONTEND-REFACTORS-PLAN.md](FRONTEND-REFACTORS-PLAN.md) Theme B — listed here because,
unlike the carves, they remove duplication rather than move it.

- [x] **3a. `createSSEStream(url, {onMessage, onOpen, onError, onUnsupported})`** in `utils.js` —
  **Done.** Collapsed **3** (not 4 — `workflows-runs.js:1019` is a plain `createPoller` call, no
  `EventSource`) near-identical `new EventSource → JSON-parse onmessage → onerror-fallback` blocks
  (`screenspace-tasks.js startSSE`, `workflows-runs.js subscribeRun`/`subscribeBatch`). Each caller
  keeps its own stream/poller state and drop handling; only the EventSource setup + parse wrapper are
  shared. `tests/test_workflows_frontend_source.py` now asserts `createSSEStream` (the raw
  `EventSource` literal moved to `utils.js`).
- [x] **3b. Generalize the modal focus-trap** — **Done.** Promoted `openBlockingModal` /
  `closeBlockingModal` to `utils.js` (focus-trap + focus-restore opt-in flags, optional
  `onBackdropClick`). `studio.js openModalTrap`/`closeModalTrap` became thin delegators
  (`trapFocus`+`restoreFocus` on) — the ~50-line implementation + `_activeTrap`/`_TRAP_FOCUSABLE`
  moved to `utils.js`; Studio has 4 trap call sites (gallery/status/confirm/log), so wrappers beat
  inlining. `transcripts.js _confirmModelInstallNow` dropped its hand-rolled Escape/backdrop
  listeners for `openBlockingModal` (no trap, matching prior behavior). Singleton pickers
  (`color-picker.js`, `settings-modal.js`) left owning their own lifecycle.
- [x] **3c. Form-input factories + color-conversion** — **Descoped (investigated).** The form-input
  factories (`rangeInput`/`numberInput`/`textInput`) are *already* shared: they live once in
  `screenspace-utils.js` and `screenspace-multitool-params.js` just calls them — no duplication to
  remove. The only real dup is `hexToRgb`/`rgbToHex` between `color-picker.js` and
  `screenspace-utils.js` (~8 lines), and the two `rgbToHsv`/`hsvToRgb` pairs use **incompatible
  ranges** (standard 0–360/[0,1] vs OpenCV 0–180/0–255) and must stay separate. Not worth churning
  ~8 lines across two pages' load graphs; left as-is.
- **Payoff (realized):** 3a + 3b deduped the SSE-setup and modal-lifecycle shapes onto `utils.js`
  globals (one implementation each instead of per-page copies). Frontend-source + satellite-wiring
  tests green.

## 4. CSS shared-component promotion

Continues today's pass (radius-full, `--color-backdrop`, theme-toggle dedup landed in `0f1012b`).
The CSS audit flagged more repeated blocks across the page stylesheets:

- [x] **4. Promote duplicated component blocks** — **Done (popover + scrollbar; icon/badge
  descoped).** Two new primitives live in `tokens.css` (loaded on every page; `primitives.css` is
  Studio-only so it can't host cross-page shared classes):
  - **`.cg-menu`** — the floating-menu/dropdown *visual* shell (bg/border/radius/shadow/padding/
    z-index). Deliberately sets **no `display`**, so it never fights a page's per-component
    `.hidden`/`display:none` toggle. Adopted by adding `cg-menu` alongside the page class (like
    `.btn`) and deleting the duplicated shell props from page CSS — `.wf-run-menu`,
    `.wf-shortcuts-menu` (workflows), `.topnav-qa-panel` (topnav), `.mark-popover` (transcripts),
    and the two start-overlay dropdowns. The two byte-identical start-overlay blocks
    (`.recent-pop` / `.sheet-picker__menu`) were also merged into one grouped body (their
    hardcoded `z-index: 10` now resolves to `var(--z-dropdown)`). Left as-is on purpose:
    `.trim-popover` (bespoke dark blur), `.cgcp-popover` (singleton picker, `z-toast`), and
    Screenspace's `.rp-switcher-panel`/`.rp-export-menu` (legacy `--color-*` + `--shadow-lg`, not
    `--shadow-pop` — adopting would change the shadow).
  - **`.cg-scroll-thin`** — theme-aware thin scrollbar (thumb `var(--border)`, hover
    `var(--border-strong)`). Applied to `#studioSidebar` and the start-overlay Changelog/About
    panels; the two always-dark launcher panels shifted from a hardcoded `rgba(255,255,255,0.08)`
    thumb to the theme border color (accepted). `#regionChips` (hidden scrollbar) left untouched.
  - **Descoped:** the **icon-mask base** (dedup needs either grouped selectors coupling the shared
    file to page-local names, or markup+JS churn across pages; `.xref-badge-icon` already is the
    base — not worth the churn) and the **badge/pill** micro-label (good `.filter-chip`/
    `.participant-pill` primitives already exist in `primitives.css`; page-local badges are
    context-specific — the item-5 over-abstraction risk).
- **Payoff (realized):** net **~−55 lines** across `tokens.css`/`studio.css`/`start-overlay.css`/
  `workflows.css`/`topnav.css`/`transcripts.css`, plus consistent menu + scrollbar styling.
  **Risk:** medium (visual) — needs a browser check (menus + the two launcher scrollbars).

## 5. Manifest load/save patterns (investigation first)

`PERFORMANCE.md` notes the "load-on-startup, save-after-mutation" JSON-manifest shape recurs for
clipgen / screenspace / transcripts / workflows / stashes / settings.

- [x] **5. Audit whether the manifest implementations are truly parallel.** — **Investigated &
  closed (2026-07-23): not duplication.** All manifest loaders already delegate to
  `utils.load_json_manifest`/`save_json_manifest`; the per-domain wrappers each add genuinely
  distinct logic (transcripts mtime-cache, workflows key-backfill + trigger sanitize, screenspace
  binary-strip). No shared abstraction to extract.

## Suggested order

1. ~~**1a** (response helpers — biggest, safest win).~~ ✓ 2. ~~**1b** + **2** (decorator + arg
   parsing, building on 1a).~~ ✓ 3. ~~**3a/3b** (JS dedup; 3c descoped — factories already
   shared).~~ ✓ 4. ~~**4** (CSS `.cg-menu` + `.cg-scroll-thin`; icon/badge descoped).~~ ✓
   5. ~~**5** only after the investigation confirms it's worthwhile.~~ ✓ (closed, not worthwhile)

Quantify before committing to a tier: a duplication scan (jscpd for JS/CSS, a `pylint`-style
duplicate-code pass for Python) would put hard block counts behind each item.

---

# Pass 2 (2026-07-23)

Second sweep after items 1–5 closed. Scope decision: mechanical **test-suite** dedup that
preserves every case/assertion is now in scope (supersedes the "Do NOT chase the test suite" note
above — parametrize/shared-helper refactors keep the guard rails, they only remove boilerplate).

## A. Mechanical dedup (safe)

- [x] **A1. `readNDJSONStream` → `utils.js`.** — **Done.** Byte-identical copies deleted from
  `studio.js` and `composer.js`; single ambient definition in `utils.js` (non-IIFE, inlined into
  exports). The `STUDIO.readNDJSONStream` export/import pair is gone;
  `test_studio_frontend_source.py` now asserts the utils.js definition + no raw `.getReader()`
  in studio sources. (~−50)
- [x] **A2. Delete dead `server._resolve_source_video`.** — **Done.** Zero callers. (~−5)
- [x] **A3. `viewer._sanitize_event_metadata` → `utils.sanitize_floats`.** — **Done.** The viewer
  helper was a strict subset (sanitize_floats also normalizes numpy scalars — more correct for
  screenspace-derived events); `import math` dropped too. (~−12)
- [x] **A4. Screenspace task-binary-strip dedup.** — **Done.** One `TASK_BINARY_KEYS` +
  `strip_task_param_binaries()` in `screenspace_manifest.py` (re-exported via the facade), used
  by both `save_screenspace_manifest` and `screenspace_server._clean_task`. (~−25)

## B. Test-suite boilerplate dedup (all cases/assertions preserved)

- [x] **B1. Parametrize `test_cli_screenspace_args.py`** — **Done (−37).** Argv-parse cluster
  (7 cases), flag-conflict exits (2), mode-conflict exits (3), scene-ref/conversion valid+raises
  (8) collapsed into 5 parametrized tests; all 49 original cases preserved (same test count).
  **Learning:** ruff-format's one-element-per-line list expansion eats most parametrize savings —
  the win only materialized after expressing each argv case as a single space-split string.
  Estimates for parametrize refactors here must be made against *formatted* code.
- [x] **B2. Shared `tests/_frontend_source.py`** — **Done (~−45).** `WEB`, `concat_js(prefix)`,
  `read(name)`, `strip_comments`, `assert_es5` adopted across the 11 frontend source-test files.
- [x] **B3. `SheetContext` builder in `conftest.py`** — **Done (~−55).** Plain
  `make_sheet_context()` helper (not a fixture; imported `from conftest import ...`); the
  `test_selectors` / `test_files_and_artifacts` copies deleted outright (all call sites pass by
  keyword), `test_spreadsheet_generation` keeps its thin sheet/cells-defaults wrapper.
- [x] **B4 (stretch). Blueprint-client fixture envelope** — **Skipped (investigated).** The
  shareable envelope is only ~3 lines/file (Flask app + register_blueprint + test_client); the
  fixture bulk is genuinely per-module seed state interleaved between construction and yield.
  ~10 real lines against a new abstraction with teardown subtleties — not worth it.

## C. Medium-risk dedup (needs browser check)

- [x] **C1. `.cg-modal-overlay` in `tokens.css`** — **Done (~−15), 4 adopters not 5.** The
  fixed/inset-0/flex-center/z-modal shell adopted by screenspace `.modal-overlay`, transcripts
  `.modal` (×2 modals), workflows `.wf-dialog-overlay`, and `.settings-overlay`; backdrop stays
  per-page. **`.hk-overlay` excluded** — hotkeys.css is inlined into exported viewers where
  tokens.css is stripped, so it must stay self-contained (its header says so). Unlike `.cg-menu`
  this primitive sets `display: flex`; each adopter's hide mechanism out-cascades it
  (screenspace `.hidden !important`; the others use compound `.x.hidden` rules).
  ⚠️ Needs a browser check: open/close all four modals in light+dark.
- [x] **C2. `createSeekCoalescer` in `utils.js`** — **Done (~−55).** The pending-seek /
  loadedmetadata-deferral / RAF-coalesce scaffolding moved to a utils.js factory
  (`getVideo`/`onDeferred`/`applySeek` hooks keep the page differences: transcripts re-dispatches
  through `seekVideo` and auto-plays after a seek write; composer stays paused).
  `cancelPendingSeek`/`seekLocal`/`_seekLocal` remain as thin wrappers so all call sites and the
  `TS.cancelPendingSeek` publication are unchanged. viewer.js's partial third instance left
  as-is (structurally divergent). ⚠️ Needs a browser check: rapid scrub + seek-before-metadata +
  part switching on Composer and Transcripts.

## D. Owner-decision deletions

- [x] **D1. `viewer.load_manifest_reels` / `friction.smooth_scores`** — **Both deleted (~−40).**
  `load_manifest_reels` was a 3-line convenience over `load_manifest_both` used only by
  `test_manifest.py` (call sites switched to `load_manifest_both()[1]`; the empty-file test now
  asserts on `load_manifest_both`). `smooth_scores` was never wired to a route — the transcripts
  timeline band does its own smoothing client-side in `transcripts-video.js` (EMA over the shared
  friction state), so the Python rolling mean was duplicate intent; its `TestSmoothScores` class
  went with it and the friction.py / ARCHITECTURE.md docs now point at the client-side smoothing.

## Pass-2 dead ends (verified, don't re-chase)

Delegator loop-generation (blocked: `test_frontend_satellite_wiring.py` requires literal
`function NAME` text); theme-toggle CSS dedup (gallery/viewer copies must survive export
stripping); `confirm()` y/n helper (case-sensitivity semantics diverge between interactive.py
and clipgen.py); config.py dead settings (none — all referenced); unused imports (ruff-clean);
facade compression (re-export contract); ffmpeg builder sharing (flags genuinely diverge);
CSS dead-selector hunt (dynamic class names defeat grep); `dev-token-tweak.js` (dev-only,
export-stripped); big-hub copy-paste blocks (none found).
