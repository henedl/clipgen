# clipgen-sync-constants — Audit Python ↔ JS constant mirroring

Any value shared between Python and the frontend must flow through `utils.get_frontend_config()`. Hardcoding in JS causes silent drift when the Python value changes.

## Audit checklist

1. **Python source** (`utils.py`)
   - Open `get_frontend_config()` — list every key it returns

2. **JS defaults** (`assets/web/utils.js`)
   - Check `CLIPGEN_CONFIG` has a default for every key
   - Check `clipgenApplyConfig()` copies every key from the payload onto `CLIPGEN_CONFIG`

3. **Server routes** — each must embed `"config": utils.get_frontend_config()` in its JSON response:
   - `server.py`: `/api/sheet` endpoint
   - `insights_server.py`: `/api/artifacts` endpoint
   - `viewer.py`: `finalize_timeline_data()`, gallery finalize, insights viewer finalize

4. **Test assertions** (`tests/test_shared_constants.py`)
   - Every mirrored value must have an assertion that the JS default matches the Python value

## Adding a new mirrored constant

1. Add the value to `get_frontend_config()` in `utils.py`
2. Add the JS default to `CLIPGEN_CONFIG` in `utils.js`
3. Add a copy line in `clipgenApplyConfig()` in `utils.js`
4. Add an assertion in `tests/test_shared_constants.py`
5. Confirm all server routes already embed the config (they should, but verify)

## Known previously-drifted constants

Severity labels, `!key` annotation keyphrase, `x` ignored token, and `DEFAULT_DURATION_SECONDS` (60s) all previously drifted across 3–5 JS files before being centralized. These are now guarded by `test_shared_constants.py`.
