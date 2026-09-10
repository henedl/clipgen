# clipgen-new-workflow-node — Add a Workflows node type

A node is one `NODE_TYPES` entry (declarative shape) plus one executor callable. The canvas fetches the catalog from `GET /workflows/api/catalog`, so **no server route or frontend registry edits** are needed. `tests/test_workflows_api.py::test_every_node_type_has_a_callable_executor` enforces that every catalog entry has an executor.

## Checklist

1. **Catalog entry** (`workflows_catalog.py`, the `NODE_TYPES` dict)
   - Add a `NodeType` literal: `id`, `label`, `description` (one line; the palette tooltip), `domain` (`artifact | screenspace | transcript | thinking | control`), `category` (palette group), `inputs`, `outputs`, `params`, `requires` (subset of `{"sheet", "videoDir"}`).
   - Ports are `{"name", "type", "optional"?}`. Reuse an existing port type (`segments`, `timeRange`, `timestamps`, `events`, `clipRecords`, `transcript`, …) so wires stay compatible; a new type needs `ADAPTERS` entries for every conversion the canvas should accept.
   - Params are `ParamSpec` dicts; the frontend renders the editor from `type`/`default`/`choices`/`min`/`max`. `required: True` blocks Run on empty; `showIf` hides a row behind a sibling param.
   - `hidden: True` keeps a node out of the palette (the per-detector `ss_<tool>` nodes); `multitoolStep: True` offers a detector as a Multitool step.
   - Keep `workflows_catalog.py` import-light: only `config`/`utils` at module level (`tests/test_import_layering.py` enforces it).

2. **Executor** (`workflows.py`)
   - Write `_exec_<id>(ctx: NodeContext, inputs, params) -> dict[str, Any]` returning one key per output port name.
   - Late-import heavy modules inside the body (`import video`, `import screenspace`, …), matching the neighbours; the module is imported by the server at startup.
   - Reuse `pipeline` / `screenspace` / `transcripts` / `viewer`; never reimplement them.
   - Report progress through `ctx.on_progress(fraction)` and honour `ctx.cancel_event` in loops.
   - Add the id to `_EXECUTORS`. The import-time loop at the bottom of the module attaches it to `NODE_TYPES[id]["execute"]`.

3. **Recipe (optional)** (`workflows_catalog.py`, `BUILTIN_STASHES`)
   - Ship a read-only example blueprint if the node needs context to be discoverable.

4. **Tests**
   - `tests/test_workflows_executors.py`: call `_run("<id>", ctx, inputs, params)` with a synthetic input and assert each output port.
   - `tests/test_workflows_api.py`: the catalog parity test runs unchanged; add a catalog assertion only for a new `ParamSpec` feature.
   - `tests/test_workflows_runner.py`: only when the node changes adapter or control-flow behaviour.

5. **Version bump** — a new node is a `feat:`; bump `build/VERSION` (see [bump](../bump/SKILL.md)).

## No edits needed

- `workflows_server.py` (routes serve `serialize_catalog()`, run nodes by id)
- `workflows_runner.py` (reads `NODE_TYPES[...]["execute"]` at call time)
- `assets/web/workflows-*.js` (palette, cards, and param editors derive from the fetched catalog)
