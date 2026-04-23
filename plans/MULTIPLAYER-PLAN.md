# Multiplayer / live collaboration plan

Planning doc. No code is written yet. Scope is clipgen's web frontends
(Studio, Transcripts, Timeline viewer, Screenspace, Insights).

## Context

clipgen is a single-user desktop tool today. A researcher runs
`clipgen --studio` (or `--transcripts`, `--screenspace`, …) and a
local Flask server on port 8089 serves the UI to their own browser.
State lives in JSON manifests in the output directory, written
optimistically after each mutation.

The goal is to make Studio (first) feel like a shared workspace when
**colleagues on the same LAN point their browsers at the running
researcher's machine**. Think "Figma-lite over LAN": you see their
cursor, the cell they're hovering, and — eventually — their comments
and marks as they appear.

We explicitly do **not** want a cloud deployment, a user database,
OAuth, or CRDT-style hard co-editing. This is an awareness layer with
a couple of collaborative seams, and nothing more.

Inspiration: PartyKit's room model (stateful server process
addressable by URL, connected clients broadcast through it). We don't
adopt PartyKit itself — we mimic the pattern inside Flask.

## Framing decisions (locked via Q&A)

| Decision | Choice |
| --- | --- |
| Deployment | LAN — one host runs clipgen, others connect to `http://<host>:8089` |
| Collaboration strength | Mixed: awareness everywhere, co-edit on comments/marks later |
| v1 hero surface | Studio |
| v1 scope | Presence + cursors only. Comments and marks are phase 2. |
| Identity | Name prompted on first visit, stored in `localStorage`, random color auto-assigned |
| Transport | WebSockets via `flask-sock` |
| Comments storage (phase 2) | New `comments_manifest.json`, keyed by target ID |
| Conflicts | Last-write-wins, matches current behavior |
| Ghost shape in Studio | Cursor dot + name label when outside grid, cell highlight when hovering a cell, active-action indicator while typing/marking |
| Room model | One room per surface (`studio`, `transcripts`, `timeline`, `screenspace`) per running instance |
| Presence persistence | Transient in memory. Session *history* appended to disk for later analytics. |
| Feature gate | `config.MULTIPLAYER_ENABLED`, off by default |

## Architecture

### Room model

One in-memory `Room` per surface, keyed by surface name. All clients
that open `/studio/` join the `studio` room; all clients that open
`/transcripts/` join the `transcripts` room; etc.

A room holds:

- `members: dict[client_id, Presence]` — live connected clients
- `broadcast(event)` — fan out to all sockets
- `on_join` / `on_leave` — lifecycle hooks

Because clipgen runs as a single Flask process on one host, a room is
just a Python object. No Redis, no pub/sub, no external platform.

### Transport

`flask-sock` adds a `/ws/<surface>` endpoint alongside existing HTTP
routes. Each connected client gets one long-lived WebSocket; the
server pushes cursor updates from other clients and any server-side
broadcast (e.g. "a new artifact manifest write happened, reload it").

Reasons over SSE:

- Bidirectional — clients send cursor positions often (throttled to
  ~20 Hz). With SSE we'd need a parallel POST firehose.
- Screenspace SSE stays as-is; the two patterns can coexist.

Reasons over plain POST polling: obvious, presence would feel laggy.

### Message shape

```json
{ "type": "presence:update",
  "client_id": "c-ab12",
  "name": "Henrik",
  "color": "#6aa9ff",
  "surface": "studio",
  "cursor": { "x": 420, "y": 188 },
  "focus":  { "kind": "cell", "row": 14, "col": 3 },
  "action": "marking" }
```

Three message families:

- `presence:join` / `presence:leave` / `presence:update` — awareness layer
- `manifest:changed` — server tells all clients "reload artifacts"
  after any HTTP POST that writes a manifest
- `comments:*` — phase 2

### Identity

On first visit the frontend prompts for a name (reusing the existing
settings modal pattern). Name + auto-generated color + random
`client_id` are stored in `localStorage` under a single key
`clipgen.identity`. No server-side identity store; the server trusts
whatever the client sends and never persists a "user row".

### Session history on disk

Every `presence:join` and `presence:leave` is appended to
`sessions_log.ndjson` in the output directory. This is the *only*
piece that touches disk for multiplayer. It's for later analytics
("who joined when"), not for presence replay.

## Files to add or touch

New:

- `multiplayer.py` — `Room`, `Presence`, `broadcast_manifest_change`,
  session-log append. Thin server-side module.
- `assets/web/multiplayer.js` — shared client: WebSocket connect,
  identity bootstrap, cursor-throttle, ghost renderer. Loaded before
  each page's main script, like `utils.js`.
- `assets/web/multiplayer.css` — cursor/ghost styles using existing
  tokens (`--color-accent`, `--shadow-1`, etc.). No new design tokens.

Touched:

- `config.py` — `MULTIPLAYER_ENABLED = False`, `MULTIPLAYER_WS_PATH`
- `server.py` — register `/ws/<surface>` via `flask-sock`; call
  `broadcast_manifest_change()` after every manifest-writing route
- `studio.js` / `studio.html` — mount ghost overlay; emit
  `presence:update` on mousemove (throttled) and on cell focus
- `pyproject.toml` — add `flask-sock` to deps (not dev)

Phase 2 touches (not in v1):

- `transcripts.js` — segment-scoped cursors, comments panel
- `viewer.js` — timeline comments
- new `comments_manifest.json` + `/api/comments` CRUD

## v1 scope — what actually ships

1. Flag `MULTIPLAYER_ENABLED` off by default. When off, no
   `multiplayer.js` is injected and no `/ws/*` route is registered.
2. When on:
   - Studio page prompts for name on first load (settings-modal style).
   - Studio page opens one WebSocket to `/ws/studio`.
   - Mouse movement inside the Studio page posts `presence:update`
     throttled to ~50 ms.
   - Hovering a sheet cell adds `focus: { kind: "cell", row, col }`.
   - While the mark popover is open or the user is typing a label,
     `action` flips to `marking` / `typing` respectively.
   - Other clients' cursors render as a 10px dot + name label that
     follows them; cells under someone else's focus get an outline in
     that client's color.
   - Server broadcasts `manifest:changed` after `/api/manifest` POST
     so peers refresh their artifacts list without an F5.
3. Sessions log appended to `sessions_log.ndjson` on join/leave.

Not in v1: comments, transcript cursors, timeline cursors,
Screenspace cursors, mark ownership, conflict UI.

## Technical fundamentals

These are the parts to commit to now so phase 2 doesn't fight the
foundation.

### Throttling and bandwidth

- Cursor updates: client-side `requestAnimationFrame` coalesce, then
  rate-limit to 20 Hz max. With 5 clients that's ≤100 msgs/sec total
  — trivial for a local Flask + one WebSocket.
- Server never stores cursor state durably. Rooms are memory-only.

### Reconnection

- Client reconnects with exponential backoff (1 s → 30 s cap).
- On reconnect the client re-sends its last `presence:update` so
  peers rediscover it. No server-side replay needed.

### Ordering and conflicts

- Presence is eventually-consistent and frame-dropped by design. No
  ordering guarantees.
- Manifest writes keep the existing HTTP POST path. WebSocket is
  *notification only* for writes ("something changed, reload"); it
  never carries the write payload. This keeps current optimistic-
  update code paths untouched and keeps last-write-wins semantics
  identical to today.

### Failure modes

- WebSocket server disabled or crashed → page still works, cursors
  just don't appear. Detect with a 10 s join-ack timeout and fall
  back silently.
- User without a name (hit cancel on prompt) → treated as local-only,
  no WebSocket opened.
- LAN peer loses route → server sees WS close, removes from room,
  broadcasts `presence:leave`.

### Security boundary

- clipgen already listens on `0.0.0.0` in the existing LAN flow
  (implicit — user binds to a routable IP when colleagues connect).
  We document this explicitly now, and flag that multiplayer mode
  exposes the same surface. No auth is added. This is intentional —
  same-LAN researcher tool, not an internet service.
- Name field is length-capped (32 chars) and HTML-escaped before
  rendering. Color must match `/^#[0-9a-f]{6}$/i`. Messages with
  unknown `type` are ignored.

### Testing approach

- Unit: `Room.broadcast` fan-out, `Presence` validation, session-log
  append, `broadcast_manifest_change` triggers on each manifest route.
- Integration: spin up Flask test client, open two WebSocket clients,
  verify join/leave/update round-trip.
- Manual: open Studio in two browsers pointed at the same host.
  Confirm cursor, cell highlight, and `manifest:changed` reload.

## Verification

End-to-end manual check (blocking before merge):

1. Set `config.MULTIPLAYER_ENABLED = True`.
2. `uv run clipgen.py --studio` on host machine.
3. Open `http://<host>:8089/studio/` in two browsers (or one browser
   + one phone on the same Wi-Fi). Enter different names.
4. Move mouse in browser A → cursor appears in B within ~100 ms.
5. Hover a cell in A → cell gets coloured outline in B.
6. Open the mark popover in A → B sees "Alice is marking" badge.
7. Save a manifest change in A → B's artifact list refreshes
   without reload.
8. Close tab A → B sees the ghost disappear within ~2 s.
9. Check `sessions_log.ndjson` — contains join and leave entries.

Automated:

- `uv run --extra dev pytest -c tests/pytest.ini` (new tests for
  Room/Presence/session-log).
- `uv run ruff check --fix && uv run ruff format`.
- `uv run ty check`.

## Open questions for phase 2 (not blocking v1)

- Do comments live in their own manifest or ride inside the existing
  ones? Current lean: separate `comments_manifest.json` keyed by
  `target_kind + target_id`, mirroring how insights is separate.
- Should Screenspace's existing SSE be folded into the WebSocket
  transport, or left alone? Default: leave alone — SSE works and
  rewriting buys nothing.
- Per-user mark attribution: keep last-write-wins visually but stamp
  `author` + `author_color` on each mark so the UI can show who did
  what without blocking concurrent edits.

---

# Phase 3 — collaborative writing in Insights

## Why Insights is different

Every earlier phase is about *awareness*: who is here, who is hovering
what, who is about to save. Insights is the first surface where the
primary interaction is **writing prose into free-text fields**. A
single insight has six editable text fields — `title`, `summary`,
three bucket `narrative`s (causes/behaviors/impacts), and
`timelineContext` — plus structured pickers and a drag-and-drop
artifact list. Today the full insight object is PUT on save, and
concurrent edits silently clobber: if Alice edits `title` and Bob
edits `causes.narrative` and both save within ~100 ms, whoever lands
second wins *for the whole record*, including fields they didn't
touch. See `insights.py` `update_insight()` and `_save_insights()`
in `insights_server.py` for the merge path.

That's the bug phase 3 fixes. The feature on top is making
multi-person writing *feel* live.

## Framing decisions (locked via Q&A)

| Decision | Choice |
| --- | --- |
| What "collaborative writing" means | Parallel editing across fields — two people on different fields of the same insight, no collisions |
| Next-best-feel if full co-edit is too much | Field-level LWW + live preview of others' typing |
| Expected frequency of simultaneous same-insight edits | Rare. Conflicts must not silently destroy work. |
| Complexity ceiling | Vanilla JS + flask-sock only. No CRDT, no Yjs, no build step. (Recommendation — see below.) |
| Conflict resolution | Auto-merge if edits don't overlap (per-field). Prompt on true overlap. |
| Live preview granularity | Figma-style full-content preview as the other user types |
| Persistence trigger | Debounced 500 ms after last keystroke (replaces current explicit Save) |
| Attribution | None. Insights are team artifacts. |

## Why no CRDT / no Yjs

CRDTs solve three problems: (a) concurrent character-level merges,
(b) cursor re-positioning after remote edits, (c) offline sync. You
said simultaneous same-insight editing is *rare* and the real concern
is *not silently destroying work*. Field-level granularity plus
versioning plus a clear prompt on overlap hits that bar for ~1% of
CRDT's complexity. Live-preview-of-typing is a separate ephemeral
broadcast (same class as cursors), not a merge problem.

We revisit this if pair-writing the same paragraph ever becomes a
core workflow.

## Architecture

Phase 3 adds three things on top of the phase 1 WebSocket transport:

1. **Field-scoped presence** — which user is focused on which field,
   broadcast like cursors.
2. **Live typing preview** — ephemeral field content broadcast at ~10
   Hz (debounced, never persisted).
3. **Versioned per-field writes** — the PUT payload changes from
   "whole insight" to "this field, at version N".

### Data model change

Each insight gains a `fieldVersions` map — one integer per editable
field. The shape, added to the record in `insights.py`:

```json
{ "id": "ins_abc123",
  "title": "…",
  "summary": "…",
  "causes":     { "narrative": "…", "artifacts": […] },
  "behaviors":  { "narrative": "…", "artifacts": […] },
  "impacts":    { "narrative": "…", "artifacts": […] },
  "timelineContext": "…",
  "fieldVersions": {
    "title": 7,
    "summary": 3,
    "causes.narrative": 12,
    "behaviors.narrative": 5,
    "impacts.narrative": 9,
    "timelineContext": 1
  },
  "updatedAt": "…" }
```

`fieldVersions` is the conflict detector. Every accepted write to
field `f` increments `fieldVersions[f]` by one. Clients always send
the version they based their edit on. Old clients without
`fieldVersions` are migrated server-side on first read (all fields
initialised to 0).

### API change

Replace the single PUT-the-whole-insight endpoint with a narrower one:

- `PATCH /api/insights/<id>/fields` — body:
  ```json
  { "edits": [
      { "field": "causes.narrative",
        "baseVersion": 11,
        "value": "New text…" }
  ] } ```
- Server behaviour per edit:
  - If `fieldVersions[field] == baseVersion` → accept, bump version,
    broadcast `insight:field-changed` on the WebSocket.
  - If `fieldVersions[field] > baseVersion` → **conflict**. Return
    the current server value + version. Client decides (see below).
- Whole-insight PUT is kept for backward compatibility with
  non-multiplayer clients but marked legacy; when multiplayer is
  enabled the frontend always uses PATCH.

### Conflict resolution policy

Server is dumb about overlap — it just rejects stale writes. The
*client* decides auto-merge vs. prompt:

- **Auto-merge window**: if the user's local edit and the server's
  newer value differ only by changes that don't touch the same
  character range, the client silently re-bases (accepts server,
  re-applies own change on top, re-submits). "Don't touch the same
  range" is computed with a simple diff on the two strings against
  the last-known-common base. No OT, no CRDT — it's one three-way
  diff at conflict time, which happens rarely by assumption.
- **True overlap**: modal — "Henrik edited this field while you were
  writing. Keep yours, take theirs, or see diff?" Three-button prompt.
  This is the "must not silently destroy work" safeguard.

### Live typing preview (Figma-style)

Separate from persistence. Every keystroke:

1. Local textarea updates immediately (no change).
2. A debounced (~100 ms) `insight:typing` WebSocket message carries
   `{ insightId, field, value, clientId }`. Ephemeral — server
   broadcasts, nobody writes to disk.
3. Other clients, if currently viewing that insight and *not focused
   on that field*, render the incoming value in a subtly-styled
   overlay or diff-highlight. On field focus switch, the overlay
   is dismissed and the field shows whatever the persisted value is.

The 500 ms debounced PATCH and the ~100 ms typing broadcast are
independent: typing preview is best-effort; the PATCH is the
authoritative write.

### Field-focus presence

Extends phase 1 presence:

```json
{ "type": "presence:update",
  "surface": "insights",
  "focus": { "kind": "insight-field",
             "insightId": "ins_abc123",
             "field": "causes.narrative" } }
```

UI effect: the focused field gets a 2 px outline in the focusing
user's color, plus their name label at the top-right of the textarea.
If two users focus the same field, both outlines/labels stack. This
is the *visual warning* that makes conflicts rare in practice — you
see Alice is typing in that field before you start.

### Debounced auto-save replaces explicit Save

Current behaviour: `markDirty` → UI shows "unsaved" → user presses
Save → whole-object PUT. With multiplayer enabled:

- Keystroke → local optimistic update → 500 ms debounce → PATCH with
  baseVersion.
- Save button and Cmd+S become a "flush pending edits now" shortcut
  rather than the only save path. Dirty indicator still shown while
  debounce is pending or a PATCH is in flight.
- `markDirty` / `saveAll` stay as fallbacks when multiplayer is off.

## Files to touch

New:
- `assets/web/multiplayer-insights.js` — extends the shared
  `multiplayer.js` with insight-specific handlers (field-focus,
  typing preview, conflict modal, three-way diff).
- `assets/web/multiplayer-insights.css` — field-focus rings, typing
  overlay style, conflict modal.

Touched:
- `insights.py` — add `fieldVersions` to the Insight shape. New
  helper `apply_field_edit(insight, field, value, baseVersion)`
  returning `("ok", new_version)` or `("conflict", current_value,
  current_version)`. Migration on load: initialise `fieldVersions`
  for records that lack it.
- `insights_server.py` — new `PATCH /api/insights/<id>/fields`
  route. On accepted edits, call `multiplayer.broadcast` with
  `insight:field-changed` payload.
- `assets/web/insights-builder.js` — swap explicit Save for
  debounced PATCH, wire up field-focus presence, handle incoming
  `insight:field-changed` and `insight:typing` WebSocket messages.
- `assets/web/insights-builder.html` — small DOM additions:
  typing-preview overlay slot per textarea, focus-label element.
- `config.py` — nothing new (already gated by `MULTIPLAYER_ENABLED`).

## Technical fundamentals

### Write amplification

500 ms debounce + 6 fields per insight + N users typing = bounded.
Worst realistic case: 2 users typing simultaneously in 2 different
fields = ~4 PATCH/s + ~20 typing broadcasts/s. Same order of magnitude
as phase 1 cursor traffic. No concern.

### Ordering and versioning

Field versions are monotonic integers scoped to `(insightId, field)`.
Server is the sole incrementer; clients never set the version. The
server processes PATCH edits serially (Flask single-process) so
version comparisons are race-free at the edit boundary.

### Offline / disconnected behaviour

- If WebSocket is down, clients keep editing locally. Dirty state
  accumulates.
- PATCH still works over plain HTTP — writes land, just without live
  broadcast. On reconnect, clients re-fetch `/api/insights/<id>` for
  any record they had open and reconcile versions.
- Conflict prompts on reconnect are expected in this mode.

### Failure modes specific to phase 3

- PATCH on a deleted insight → 404, client removes it from view.
- PATCH on an unknown field → 400, client logs and does nothing (this
  catches schema drift bugs).
- Typing broadcast for a field the recipient already has focus on →
  recipient discards the message. Never overwrite what a focused user
  is typing.
- User closes tab with pending debounce → flush synchronously on
  `beforeunload` via `navigator.sendBeacon()` to the PATCH endpoint.

### Security boundary

Same as phase 1 — LAN trust. But: PATCH values are now arbitrary user
text that ends up rendered in other users' DOM. Render through the
existing escape helpers (textarea `.value =` is safe; any span/badge
rendering goes through `utils.js`'s escape helper).

### Testing approach

- Unit: `apply_field_edit` accept, reject-stale, version increment,
  migration from missing `fieldVersions`.
- Integration: two Flask test clients; A PATCHes then B PATCHes same
  field at same base version — B gets conflict. A PATCHes, B PATCHes
  *different* field — both succeed.
- Manual (blocking before merge):
  1. Open same insight in two browsers.
  2. Alice focuses `title`, Bob focuses `causes.narrative` — both see
     colored outlines + name labels.
  3. Both type simultaneously — no collisions, both persisted, each
     sees the other's value update on blur.
  4. Both focus the *same* field and both type → conflict modal
     appears on whoever PATCHes second, with three-button choice.
  5. Disable WebSocket, keep editing, re-enable — verify reconcile.
  6. Close tab mid-typing → verify beacon flush landed.

## Open questions for phase 3

- Do we want a "history" view per insight (field-level change log)
  once every edit has a version? Low cost to add, defer unless users
  ask.
- Should the conflict modal's three-way diff be character-level
  (`diff-match-patch`) or line-level (trivial hand-rolled)? Default:
  line-level for v1; upgrade if reviewers complain.
- Typing preview is always-on when two users view the same insight —
  is there a surface where it'd be distracting and want an opt-out?
  Revisit after dogfooding.
