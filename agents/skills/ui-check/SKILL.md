# clipgen-ui-check — Headless page smoke, screenshots, and in-page probing

Boots the real combined server against a generated fixture project, loads all six
pages in headless Chromium, and fails on any uncaught error. It also runs five
click-through journeys (`tests/ui/test_ui_journeys.py`): Studio generate,
Screenspace result-row seek, Transcripts segment seek, settings persistence, and
opening the fixture workbook through the Start overlay. Use it after touching
anything in `assets/web/` — it is the runtime confirmation that used to be "ask the
human to open the page".

It catches the class the static guards cannot: a bare cross-file reference in a
hub/satellite carve throws a `ReferenceError` on boot, which `node --check` and
`tests/test_frontend_satellite_wiring.py`'s regex scan both miss by design.

## 1. Preflight — never install without asking

```bash
uv run --no-sync python -c "import playwright" 2>/dev/null || echo "extra missing"
```

If missing, **ask the user first**, then:

```bash
uv sync --extra dev --extra ui                   # playwright wheel + driver, ~40 MB
uv run --extra ui playwright install chromium    # ONLY if the run reports no browser found
```

The browser download is ~150 MB into `~/Library/Caches/ms-playwright`, never the
repo. The harness globs that cache and passes the newest Chromium as
`executable_path`, so it works with whatever build is already there — the pip
package's own expected build number is ignored on purpose.

`ffmpeg` + `ffprobe` must be on PATH; the run skips with a clear reason otherwise.
**A skipped run is not a passing run — read the skip reason.**

## 2. Run

```bash
CLIPGEN_UI_CHECK=1 uv run --extra dev --extra ui pytest -c tests/pytest.ini \
  tests/ui -p no:randomly -q
```

`-p no:randomly` because shuffling independent page loads buys nothing and
makes failure output harder to read. Cold run ~45 s (ffmpeg encodes the fixture
videos once), warm ~30 s — the boot smoke plus the five journeys.

## 3. Inspect — actually look at the screenshots

```
.context/ui-check/screenshots/{studio,screenspace,transcripts,workflows,composer,overview}.png
.context/ui-check/screenshots/journey-{generate,screenspace,transcripts,settings,start-overlay}.png
.context/ui-check/ui-report.json
```

`Read` the PNG for each page you touched. The Read tool renders images, so **look
at the UI** rather than trusting the exit code. What you are checking for:

- Does the page show real content, or an empty state? The fixture project has 6
  observations, 2 participants with video, 4 transcript segments, 2 Screenspace
  events plus 1 completed Screenspace task, 1 Composer cut and 1 Workflows
  blueprint. A zero-state screenshot means the seeding in
  `tests/ui/_ui_fixtures.py` drifted, even though the test passed.
- Does your change look right — spacing, alignment, colors, truncation?

`ui-report.json` holds the non-fatal detail that never reaches stdout: XHR 404s
(`/api/thumbnail`, `/api/sprite`, `/api/preview` legitimately 404 with no
extracted frames), request failures, and the absolute screenshot paths.

## 4. Iterate with `shot.py`, not the full suite

One page in ~4 s, plus the thing that replaces "paste this DevTools snippet":

```bash
uv run --extra ui python tests/ui/shot.py studio
uv run --extra ui python tests/ui/shot.py studio --selector "#sheetGrid"
uv run --extra ui python tests/ui/shot.py transcripts \
  --eval "return document.querySelectorAll('.segment-row').length"
uv run --extra ui python tests/ui/shot.py composer --viewport 1280x800 --wait 1500
```

`--eval` runs arbitrary JS in the live page and prints the return value as JSON —
read computed styles, count rendered nodes, dump page state, call
`el.getAnimations()` to bisect an animation that was created but never painted.
Use `--eval-file` for anything longer than a shell-quotable line. Exit is non-zero
on a `pageerror` or a throwing snippet, so it composes in a shell chain.

Do not loop `shot.py` over all six pages — it boots its own server per invocation.
That is what the suite is for.

For performance questions ("is this render slow?", "does that poller pause when
hidden?") don't eyeball — `shot.py --perf` captures CDP metrics, timing entries and
the page's profiling spans as grep-able `perf | ` lines, and `--trace` writes a
Perfetto-compatible Chrome trace. Workflow: [/profile](../profile/SKILL.md).

### States and themes

The six-page smoke renders each page's boot state, in dark, and clicks nothing.
Most of the frontend is not in that: every modal, every tab past the default, and
the whole Start overlay (both boot paths pre-dismiss it). Reach the rest with:

```bash
uv run --extra ui python tests/ui/shot.py studio --theme light
uv run --extra ui python tests/ui/shot.py studio --state settings
uv run --extra ui python tests/ui/shot.py screenspace --state tool:template
uv run --extra ui python tests/ui/shot.py screenspace --all-states   # 20 states, one boot
```

Every page has `settings`, `settings-hotkeys`, `cheatsheet`, `palette` and
`start`. Tab states are **discovered from the live DOM**, not hard-coded, so a tab
added to the HTML becomes a state for free — pass an unknown `--state` name and
the error lists what that page actually has.

`--all-states` writes `<page>-<state>.png` and drives them all from one boot,
which is the thing looping `shot.py` cannot do. It reports every state it tried,
reached or not: an unreachable state prints as `MISS`, never silently skipped.
`--state` exits non-zero when the named state can't be reached; `--all-states`
does not, because some states legitimately don't exist in the fixture.

Two honest limits, worth knowing before you read the output:

- Screenspace's 13 tool tabs are hidden whenever the grouped category nav is on.
  They are activated with a DOM `.click()` — the same delegation the grouped nav
  itself uses (`screenspace.css:1401-1410`) — and are labelled
  `hidden; activated via DOM click` so you can tell that from a real gesture.
- The fixture is small on purpose, so an empty panel in a state screenshot may be
  "no data" rather than "broken". Check `tests/ui/_ui_fixtures.py` before filing it.

## Diagnosing a failure

| Symptom | Where to look |
| --- | --- |
| `pageerror: X is not defined` | The carve bug. A satellite calls a hub function with no delegator, or reads a moved `var` bare. See [carve-satellite](../carve-satellite/SKILL.md). |
| Timeout on `nav.topnav` | The bundle threw before boot — a syntax or load-order problem, not a data problem. |
| Timeout on the per-page selector | Either a fetch never resolved, or the selector in `tests/ui/_ui_pages.py` rotted because an id/class was renamed. Check the page's render function. |
| The whole run hangs with no output | An SSE stream open at teardown. See the `block_on_close` note in `tests/ui/_ui_server.py`. |

## Non-goals — do not wire this in

- **Never** add `tests/ui` to `/check` or to `.github/workflows/`. It needs a
  browser, a fixture encode and ~20 s; `/check` must stay fast and hermetic.
  `tests/test_packaging.py::test_ui_suite_stays_opt_in` guards the locks.
- **Never** remove `ui` from `norecursedirs` in `tests/pytest.ini`, or the
  `CLIPGEN_UI_CHECK` gates in `test_ui_smoke.py` / `test_ui_journeys.py`. They
  are two independent locks and both are deliberate.
- `tests/test_frontend_syntax.py` (the `node --check` gate) is **not** part of this
  suite and *does* belong in `/check`. Don't move it here.
- This is not an interaction crawl. The journey list is capped at five until a
  real shipped bug earns a sixth (`agents/skills/test/SKILL.md` has the policy).
  Feel, motion, drag behaviour and real-media playback still need a human.

## Housekeeping

The fixture project lives in `.context/ui-check/` — `.context/` is the agent
scratch dir Conductor already uses per worktree, so the harness adds no top-level
directory. `.gitignore` ignores the whole subtree, so nothing it writes can be
committed. Delete the directory to force a full rebuild (~2 s of ffmpeg).

That `.gitignore` line is load-bearing and easy to mistake for redundant:
Conductor also excludes `.context/` via `.git/info/exclude`, but that file is
local-only and absent from a plain clone.

Note the uv extras churn: `/check` (`--extra dev`) after `/ui-check`
(`--extra dev --extra ui`) uninstalls playwright, and the next `/ui-check`
reinstalls it from cache in about a second. Surprising, not broken.
