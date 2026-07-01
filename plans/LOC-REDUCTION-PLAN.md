# LOC-reduction opportunities — plan

Status: **in progress** (2026-07-01). Items **1 (1a + 1b)**, **2**, **3**, and **4** landed (see
check-marks below); item **5** still open. Investigation targets for genuinely *reducing* total
lines of code (not relocating them). Each item is sized as one or more focused `refactor:`
commits. Check items off and add a "Done" note as they land (per AGENTS.md plan-maintenance rule).

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

- [ ] **5. Audit whether the manifest implementations are truly parallel.** If so, a small
  load/save helper could dedup them. **Confirm they're not subtly different first** — manifests have
  different schemas and mutation semantics; only the I/O envelope is a candidate, and forcing a
  shared abstraction over diverging shapes would be net-negative.

## Suggested order

1. ~~**1a** (response helpers — biggest, safest win).~~ ✓ 2. ~~**1b** + **2** (decorator + arg
   parsing, building on 1a).~~ ✓ 3. ~~**3a/3b** (JS dedup; 3c descoped — factories already
   shared).~~ ✓ 4. ~~**4** (CSS `.cg-menu` + `.cg-scroll-thin`; icon/badge descoped).~~ ✓
   5. **5** only after the investigation confirms it's worthwhile.

Quantify before committing to a tier: a duplication scan (jscpd for JS/CSS, a `pylint`-style
duplicate-code pass for Python) would put hard block counts behind each item.
