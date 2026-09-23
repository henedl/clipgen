# Test suite evolution plan

> **Status: Done** (phases 1–3 shipped; phase 4 remains optional and unstarted).
> Journey #2 shipped adapted — see the note on its checkbox.

**Goal:** Buy confidence that researcher paths still work, without slowing `/check`.

**Architecture:** Keep the existing pytest suite (mocked API/domain + JS architecture
ratchets + opt-in `/ui-check` boot smoke). Add two thin "real" layers on top: a
handful of ffmpeg I/O tests in the default suite, and a capped set of Playwright
journeys in `tests/ui/`. Freeze the source-scan style as a ratchet, not a default
frontend test.

## Diagnosis (do not relitigate)

The suite is ~2,500 tests / ~50k lines, ~20s serial, ~6s at `-n 4`. It is strong at
fast isolated Python (Flask `test_client`, monkeypatched `run_ffmpeg`, argv
asserts) and at JS *architecture* lints (satellite wiring, packaging, ES5, shared
constants, dead-function ratchet). It is weak at JS *behavior* and at proving the
product still produces media.

`/ui-check` already builds a real fixture project (ffmpeg `testsrc` videos, xlsx,
seeded manifests), boots the real Flask app, and loads six pages in Chromium. It
clicks almost nothing. The gap is "curtain up" vs "do the job."

`test_container_seekability.py` already runs real ffmpeg for fragmented-MP4
detect/remux. `test_clip_pipeline.py` already locks the silent-wrong reel class
under mocks. Do not duplicate either.

Two missing "real"s, different prices:

| Kind | Proves | Where it lives |
|---|---|---|
| I/O integration | ffmpeg actually writes a usable file | default suite, `skipif` no ffmpeg, ~1s budget each |
| UI journey | click → API → DOM still hangs together | `tests/ui/` only, `CLIPGEN_UI_CHECK=1` |

## Constraints (every task inherits these)

- Do **not** start a new suite, add Jest/Vitest, or change the IIFE/ES5 frontend
  to make it unit-testable. Playwright on the existing harness is the JS runner.
- Do **not** put Playwright on `/check` or CI. `tests/pytest.ini` `norecursedirs`
  already excludes `ui`.
- Do **not** "cover the UI." Cap journeys at five until one of them has earned a
  sixth.
- Do **not** add `pytest-xdist` to `pyproject.toml` (CI-only, same as today).
- Do **not** chase coverage %. `fail_under = 60` is not a selection policy.
- New frontend tests: pin a **rule** or a **shipped bug**, not a spelling.
  `assert "function foo" in src` is frozen — allowed only when the bug is "this
  symbol vanished," preferably as a shrink-only allowlist (`KNOWN_DEAD` in
  `tests/test_js_dead_functions.py` is the pattern).
- I/O tests that land in `--durations=20` go through
  [agents/skills/test-perf/SKILL.md](../agents/skills/test-perf/SKILL.md) before
  they ship.
- Prefer existing DOM ids/classes over new `data-testid` attributes. Add a test
  id only when the live selector would rot on a rename the user never sees.

## Taxonomy (write the right kind)

When adding a test, pick one. Do not mix kinds in one file.

| Kind | Example | Default `/check`? |
|---|---|---|
| **Ratchet** | `test_frontend_satellite_wiring.py`, `test_packaging.py`, `test_shared_constants.py` | yes, must stay fast |
| **Domain** | `test_utils_timestamps.py` | yes |
| **API** | `test_studio_api.py` route + JSON under mocks | yes |
| **I/O** | real `process_clips` cut of a 2s `testsrc` | yes, max three, `skipif` no ffmpeg |
| **Journey** | Studio generate → artifact appears | no, `tests/ui/` only |

Teaching copies: domain → `tests/test_utils_timestamps.py`; API → a *small* test
in `tests/test_studio_api.py` (not the whole 4.5k-line file); journey →
`tests/ui/test_ui_smoke.py` + `tests/ui/_ui_states.py`.

---

## Phase 1 — Policy in the skill (no new tests yet)

Update [agents/skills/test/SKILL.md](../agents/skills/test/SKILL.md) so agents
stop defaulting to "another mocked route test" and "another `in src` pin."

- [x] Add the taxonomy table above (kinds, where they live, `/check` or not).
- [x] Add the source-scan freeze: no new `assert "<code>" in src` unless the
      bug is a vanished symbol; prefer `KNOWN_DEAD`-style shrink-only lists.
- [x] Add: every new CLI mode / flag / selector still needs a smoke test
      (existing rule). That is the *only* additive-by-default rule.
- [x] Add: if the bug was "page booted but the click did nothing," add a
      journey — only if the journey list is still ≤ 5.
- [x] Add: if the bug was "ffmpeg produced garbage / success-with-partial,"
      add an I/O test, not another mocked argv assert. Check
      `test_clip_pipeline.py` first; the silent-wrong reel class is already
      locked under mocks.
- [x] Add optional local speed tip (not a CI change):

```bash
uv run --extra dev --with pytest-xdist pytest -c tests/pytest.ini -n auto
```

Do not add pytest markers (`unit` / `integration`) in this phase. The `ui/`
directory split is already the slow-path marker. Markers are YAGNI until
someone is actually running a subset inner loop.

**Done when:** `/check` still green; skill is the selection policy; no test
files changed.

---

## Phase 2 — Three I/O tests (default suite)

New file: `tests/test_pipeline_io.py`. Follow `tests/test_container_seekability.py`
for the ffmpeg skip and 2s `testsrc` encode helper (160×120, 15 fps is enough).
Do not import `tests/ui/_ui_fixtures.py` — that module is opt-in UI and must
stay out of default collection.

Reuse `make_clip` from `tests/conftest.py`. Point the clip at the encoded file
via the same `{study}_{participant}.mp4` contract `files` expects. Set
`config.OUTPUT_DIR` to `tmp_path`, `CLIP_PARALLEL_WORKERS = 1`.

- [x] **Cut:** `pipeline.process_clips([clip], output_format="clip")` returns
      count 1, the output `.mp4` exists and is non-empty, and
      `video.get_file_duration` is in range for the requested window (a 2s
      source, ~1s cut).
- [x] **Screenshot:** same setup, `output_format="screen"`, output `.png`
      exists and `PIL.Image.open` can read it (dimensions > 0).
- [x] **Reel of two cuts:** two adjacent windows on the same 2s source,
      `pipeline.process_reel(...)`. Assert the reel path exists, duration is
      roughly the sum of the windows (not one window), and a mocked-False
      second cut does **not** produce a reel file (locks the silent-wrong
      class on the real concat path; the mock-only version already lives in
      `test_clip_pipeline.py`).

Skip the whole module when ffmpeg/ffprobe are missing, same
`requires_ffmpeg` marker pattern as `test_container_seekability.py`.

**Out of scope here:** titlecards (Homebrew `drawtext` is machine-dependent),
EasyOCR, Whisper, remux (already covered), more combinatorics (padding ×
format × workers stay mocked in `test_clip_pipeline.py`).

**Done when:**

```bash
uv run --extra dev pytest -c tests/pytest.ini tests/test_pipeline_io.py -p no:randomly --durations=20
```

All three pass when ffmpeg is present; none appear as duration outliers
relative to the ~1s EasyOCR ceiling. Full `/check` still green.

---

## Phase 3 — Five journeys (opt-in `/ui-check` only)

Keep `tests/ui/test_ui_smoke.py` as boot + overlay open. Add
`tests/ui/test_ui_journeys.py`. Reuse `live_server`, `browser_context`,
`_ui_pages.open_and_settle`, `_ui_pages.wire_listeners`, `PageLog`. Same
`CLIPGEN_UI_CHECK=1` module-level skip. Same `-p no:randomly`. Fail on
`log.fatal` *and* on the journey's own DOM contract.

Reach states the way `_ui_states.py` does: public `window` openers and real
clicks, not internal render functions. Discover live selectors; put the
fragile ones in a small table at the top of the journeys file (same honesty
as `PAGES` in `_ui_pages.py`).

Cap: these five, no more, until a real shipped bug is a sixth.

- [x] **Studio generate.** Open Studio (`#sheetGrid tbody tr` already ready).
      Select one valid timestamp cell (`.valid-ts`), get it into the artifact
      work area (`#artifactsList` / `#artifactsCount`), click `#generateBtn`,
      wait until the spinner (`#artifactsSpinner`) is gone and the list has
      at least one generated item (not just a queued cell). Budget: the
      fixture videos are 20s; cut a short cell. If generate is flaky because
      of job/stream timing, drive the same click path but assert on a
      completion signal the page already exposes (progress snapshot / queue
      item class) rather than adding a `data-testid`.
- [x] **Screenspace seek-to-event.** *Shipped adapted:* the seeded events
      render nowhere as written (the timeline is a canvas that only draws
      tasks, and the fixture seeded none), and Screenspace seeks never touch
      `video.currentTime` (`loadFrame` pauses the video and moves only the
      playhead state). The fixture now seeds one **completed** task (safe:
      SSE/polling only start for queued/running/paused) with `ui-evt-1`
      attached; the journey clicks the task, then a `.result-row`, and
      asserts the playhead (`#timestampInput`) landed on 2.0s. Still no
      detector scan.
- [x] **Transcripts seek-to-segment.** Open Transcripts `#P01`. Click
      `#segmentList .segment-row` (already the readiness selector). Assert
      player time matches the segment's start. (Shipped against the row's
      `.segment-timestamp` child — row-level clicks are deliberately a no-op.)
- [x] **Settings persist.** Open settings via `window.openSettingsModal`
      (already in `GLOBAL_OVERLAYS`). Change one durable setting that the
      fixture does not depend on (e.g. a numeric pad or a theme-adjacent
      toggle that round-trips through `/api/settings`). Reload Studio. Re-open
      settings. Assert the value stuck. Do not use `localStorage` as the
      assertion if the product persists through the settings file.
- [x] **Start overlay → Studio rows.** Needs a context **without**
      `clipgen.startOverlayDismissed` (today `tests/ui/_ui_session.py`
      `build_init_script` always sets it). Add a parameter or a second
      factory so this one test can see the overlay. Pick the fixture workbook
      through the overlay's own UI (not by stuffing sessionStorage). Assert
      Studio shows `#sheetGrid tbody tr` after confirm.

**Done when:**

```bash
CLIPGEN_UI_CHECK=1 uv run --extra dev --extra ui pytest -c tests/pytest.ini \
  tests/ui -p no:randomly -q
```

Boot smoke + five journeys pass; warm run stays in the same ballpark as
today's ~20s smoke (journeys will add time; if the total exceeds ~90s warm,
cut generate's wait/fixture length, do not drop the cap). Screenshots for
failed journeys still land under `.context/ui-check/screenshots/`. `/check`
is unchanged (does not collect `tests/ui/`).

Update [agents/skills/ui-check/SKILL.md](../agents/skills/ui-check/SKILL.md)
to mention journeys exist and that `/ui-check` now covers them.

---

## Phase 4 — Optional, only if a later session feels the pain

Not required to close this plan. Do not start these in the same change as
phases 1–3.

- [ ] Split `tests/test_studio_api.py` / `test_screenspace_api.py` /
      `test_transcripts_api.py` by resource (sheet, generate, reel, …) if
      navigation is the actual drag. Same tests, new files; no behavior
      change.
- [ ] Delete source-scan tests that only pin a spelling and name no shipped
      bug. Keep ratchets. Do not bulk-delete in a cleanup pass without a
      named replacement.
- [ ] A sixth journey, only for a bug that boot smoke + the five missed.

## Explicitly out of scope

- Replacing pytest, adding a JS unit runner, rewriting IIFEs into modules.
- Putting browser tests in the default pytest path.
- Growing `/check` with scans, OCR, Whisper, or titlecards.
- A coverage-driven rewrite, or lowering/raising `fail_under` as a proxy for
  this work.
- Matching test count to JS line count with more source scans.

## How a later agent should pick work

1. Phase 1 first (policy). Then 2 (I/O). Then 3 (journeys). Stop.
2. If `/check` got slower, run `--durations=20` and fix per `test-perf` —
   do not "optimize" by deleting API tests (they are already fast).
3. When this plan's phases 1–3 are done, mark the header **Done** and move
   the file to `plans/archive/`.
