# clipgen-debug — Diagnostic checklist for common issues

## Clips not generating / missing output

1. **Wrong video filename** — source videos must be named `{study}_{participant}.mp4` exactly (e.g. `mystudy_P01.mp4`). The study name is lowercased and filesystem-safe. Run `ls INPUT_DIR` and compare against what clipgen expects.

2. **Participant ID mismatch** — participant column headers must start with `P` (individual) or `G` (group). Check `config.PARTICIPANT_PREFIXES`. IDs are case-sensitive.

3. **Timestamp silently skipped** — Google Sheets rate limiting. Repeated calls to `get_all_values()` / `sheet.find()` / `generate_list()` can be silently throttled, causing rows to appear empty. Wait a minute and retry. In development, avoid calling the API repeatedly in a loop.

## Wrong clip timing

4. **Clip too short** — a single timestamp (e.g. `1:23`) gets `end = start + DEFAULT_DURATION_SECONDS` (60 seconds). To specify an explicit end time, use a range: `1:23-2:30`.

5. **Baseline time offset wrong** — the `Baseline time` marker row offset is calculated relative to the header/`id_cell` row. If the spreadsheet layout changed, the offset math in `spreadsheet.py` may be off. Check `id_cell.row` and the baseline row placement.

6. **Annotation keyphrases stripped** — `!key` is stripped before timestamp parsing (configured in `ANNOTATION_KEYPHRASES`). Tokens like `x` are ignored entirely (configured in `IGNORED_TIMESTAMP_TOKENS`). If a timestamp contains one of these, it will be silently skipped.

## Web UI issues

7. **UI not loading** — the combined Flask server always starts on port `8089` (`config.SERVER_PORT`). Check that nothing else is bound to that port: `lsof -i :8089`.

8. **DBus errors in logs** — harmless. They come from Chrome attempting DBus connections in headless environments. Ignore them.

9. **`objc[...] Class AVFFrameReceiver/AVFAudioReceiver is implemented in both ...` on macOS** — harmless. Both `opencv-python-headless` (cv2, Screenspace) and `av` (PyAV, pulled in by `faster-whisper` for Transcripts) bundle their own FFmpeg `libavdevice`, an AVFoundation capture-device library clipgen never uses; the ObjC runtime warns when both load in one process. `start_combined_server()` pre-loads both via `utils.preload_av_libs_quietly()` with native stderr silenced, so the combined web server stays quiet. A rare pure-CLI run that loads both libs may still print it — it's benign.

## Development / debugging

10. **Enable debug mode** — set `config.DEBUGGING = True` to:
   - Enable icecream (`ic()`) output
   - Skip ffmpeg execution (video.py returns stubs)
   - Return stub transcript results without loading a Whisper model

11. **Type errors in ty** — see [agents/skills/check/SKILL.md](../check/SKILL.md) for common ty failure patterns.

12. **Tests unexpectedly passing despite broken behavior** — check whether the test is mocking at too high a level. The project avoids mocking the database/sheets layer; if a test uses a mock that doesn't reflect reality, consider replacing it with a real fixture.
