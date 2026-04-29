# JavaScript type-checking — deferred plan

## Context

`assets/web/` is hand-written ES5-style vanilla JavaScript. There is no
build step, no transpiler, no TypeScript. Recent ESLint cleanup
(branch `henedl/js-lint-fixes`) caught spurious `no-undef` issues from a
speculatively-written config and a small handful of real bugs (dead
helpers, shadow redeclarations, unused locals).

ESLint catches structural problems but doesn't know types. The next class
of bugs we'd catch with relatively little extra scaffolding:

- Wrong arity on a `utils.js` helper call.
- Typos in DOM property names (`.classList.tggle` etc.).
- Accessing fields on `window.CLIPGEN_DATA` / API responses that aren't in
  the data contract.
- Returning the wrong shape from a function consumed cross-file.

These are exactly the kind of cross-file API-surface bugs that a
speculative-globals list invited. Worth catching automatically.

## Approach: JSDoc + `tsc --checkJs --noEmit`

Use TypeScript's compiler **as a checker only** — no transpilation, no
emission. Annotate types via JSDoc comments. Source stays pure JavaScript;
no file extensions change; no build step; no runtime change. To roll back,
delete `tsconfig.json` and the JSDoc comments.

### Minimal scaffolding

1. Add a `tsconfig.json` at repo root:
   ```json
   {
     "compilerOptions": {
       "allowJs": true,
       "checkJs": true,
       "noEmit": true,
       "target": "ES2020",
       "module": "none",
       "lib": ["ES2020", "DOM"],
       "strict": false,
       "noImplicitAny": false
     },
     "include": ["assets/web/**/*.js"]
   }
   ```
2. Add a `lint:ts` script to `package.json`:
   `"lint:ts": "tsc --noEmit -p tsconfig.json"`
3. Add `typescript` to `devDependencies`.
4. Add a `js-typecheck` job to `.github/workflows/tests.yml`, parallel to
   the existing `js-lint` job.

That's the floor. ~15 lines of config, one new dev dependency, no source
changes yet.

### Where to annotate

**Tier 1 — annotate first:**
- `assets/web/utils.js` exports. These are the cross-file API surface and
  the source of the problem we just cleaned up. Annotating each top-level
  helper's signature catches every page-script caller that misuses it.
- `window.CLIPGEN_DATA` shape (used by viewer, gallery, insights-viewer).
  One typedef in `utils.js` (or a new `types.js`) covers all consumers.
- API response shapes for the most-touched endpoints: `/api/sheet-data`,
  `/api/artifacts`, `/api/tasks`. One typedef per endpoint.

**Tier 2 — only if a real bug surfaces in tier 1:**
- Page-script internal helpers (per-file).
- Less-touched API endpoints.

**Don't annotate:**
- Tight loops, ad-hoc DOM construction, single-call-site helpers. Cost
  exceeds benefit for a codebase this size.

### Cost estimate

- Setup (config + CI job + first checkJs run with current code): half a day,
  mostly just resolving the initial wave of "implicit any" warnings by
  setting `noImplicitAny: false` (already in the config above) or by
  selectively annotating where the checker can't infer.
- Tier-1 annotation pass: ~1 day, concentrated in `utils.js` and one
  typedef file for the data contract.
- Maintenance: low — JSDoc on new helpers as they're added, same way
  Python type hints work today.

## Why not TypeScript-the-language

- Adds a build step (`.ts` → `.js`) — the project's frontend explicitly
  has no build pipeline (`AGENTS.md`: "Vanilla JS (ES5 .then()),
  hand-written CSS, no frameworks or build tools").
- File extensions change, every consumer has to follow.
- Onboarding cost for a project that's intentionally low-ceremony.

JSDoc + `checkJs` gives us most of TypeScript's bug-finding value without
crossing the no-build-step line.

## Why not all-or-nothing

`checkJs` runs file-by-file. We can opt files in incrementally with
`// @ts-check` comments, or opt the whole project in with `checkJs: true`
in tsconfig and tolerate weak inference where we don't annotate. Either
works; the strictness setting (`noImplicitAny: false`) controls how loud
the unannotated parts are.

## Triggers — when to revisit

This is deferred work. Pick it up when one of these happens:

1. A bug ships that a type-checker would have caught (e.g. a renamed
   utils.js export breaks a page script silently — same bug class as the
   speculative-globals issue).
2. The cross-file API surface grows enough that ESLint's globals list
   becomes painful to maintain manually.
3. A new contributor joins and the lack of inline type info becomes a
   review-time tax.

## Future work (also deferred)

### Prettier

We considered Prettier alongside this. Decision: defer. The codebase has
its own consistent hand-rolled style; introducing Prettier would either
churn every file (style change) or require careful config to match. The
lint pass already enforces structural quality, and ESLint's `--fix`
catches the formatting bits we actually care about. Revisit only if:

- Multi-contributor style drift becomes a real problem.
- We're already touching every JS file for another reason (e.g. a
  framework migration) and bundling Prettier into that pass is cheap.

### Other ESLint plugins

`eslint-plugin-import`, `eslint-plugin-promise`, `eslint-plugin-unicorn`,
etc. — each catches a few edge cases but collectively buries the config.
Not worth it for a codebase this size unless a specific pain point
surfaces.
