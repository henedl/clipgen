import types

import clipgen
import tui
import video


class DummyWorksheet:
    """Minimal worksheet stub for reel/selection tests."""

    def __init__(self, sheet_data, col_count):
        self._sheet_data = sheet_data
        self.col_count = col_count
        self.spreadsheet = types.SimpleNamespace(title="TestSheet", url="http://example.com")

    def get_all_values(self):
        return self._sheet_data

    def row_values(self, row_index):
        # gspread uses 1-based row indices
        return self._sheet_data[row_index - 1]

    def find(self, value):
        # Very small helper to satisfy validate_spreadsheet_headers()
        for r_idx, row in enumerate(self._sheet_data, start=1):
            for c_idx, cell in enumerate(row, start=1):
                if cell == value:
                    # Return an object with .row and .col attributes
                    return types.SimpleNamespace(row=r_idx, col=c_idx)
        return None


def test_select_spreadsheet_uses_textual_first_choice(monkeypatch):
    """select_spreadsheet should prefer the Textual selector when available."""

    chosen_name = "Second Sheet"
    called_with = {}

    monkeypatch.setattr(tui, "use_textual", lambda: True)
    monkeypatch.setattr(tui, "run_spreadsheet_select", lambda doc_list: chosen_name)

    def fake_handle(client, doc_list, input_name):
        called_with["doc_list"] = doc_list
        called_with["input_name"] = input_name
        return "DUMMY_WORKSHEET"

    monkeypatch.setattr(clipgen, "_handle_spreadsheet_command", fake_handle)

    fake_client = object()
    doc_list = ["First Sheet", "Second Sheet", "Third Sheet"]

    worksheet = clipgen.select_spreadsheet(fake_client, doc_list)

    assert worksheet == "DUMMY_WORKSHEET"
    assert called_with["doc_list"] == doc_list
    assert called_with["input_name"] == chosen_name


def test_select_mode_and_generate_textual_mode_batch(monkeypatch):
    """Mode selection should respect Textual menu choice and dispatch to generate_list."""

    monkeypatch.setattr(tui, "use_textual", lambda: True)
    monkeypatch.setattr(tui, "run_mode_select", lambda: "batch")

    called = {}

    def fake_generate_list(ws, mode, **kwargs):
        called["worksheet"] = ws
        called["mode"] = mode
        called["kwargs"] = kwargs
        return ["clip1", "clip2"]

    monkeypatch.setattr(clipgen, "spreadsheet", types.SimpleNamespace(generate_list=fake_generate_list))

    worksheet = "WS"
    clips_list, is_reel, reel_output_file = clipgen.select_mode_and_generate(worksheet)

    assert clips_list == ["clip1", "clip2"]
    assert is_reel is False
    assert reel_output_file is None
    assert called["worksheet"] == worksheet
    assert called["mode"] == "batch"


def test_run_ffmpeg_uses_textual_long_clip_dialog(monkeypatch, tmp_path):
    """run_ffmpeg should call tui.confirm_long_clip when duration exceeds limit."""

    # Create a dummy input file so os.path.isfile returns True for it.
    input_file = tmp_path / "input.mp4"
    input_file.write_bytes(b"dummy")

    output_file = tmp_path / "output.mp4"

    # Force long duration
    monkeypatch.setattr(video, "get_duration", lambda start, end: video.config.MAX_CLIP_DURATION_SECONDS + 10)
    # Pretend the container duration is long enough
    monkeypatch.setattr(video, "get_file_duration", lambda path: video.config.MAX_CLIP_DURATION_SECONDS + 20)

    # Avoid actually running ffmpeg
    monkeypatch.setattr(video, "_run_ffmpeg_process", lambda *args, **kwargs: None)

    called = {}

    monkeypatch.setattr(tui, "use_textual", lambda: True)

    def fake_confirm_long_clip(duration_seconds):
        called["duration"] = duration_seconds
        # Simulate user choosing not to generate the long clip
        return False

    monkeypatch.setattr(tui, "confirm_long_clip", fake_confirm_long_clip)

    ok = video.run_ffmpeg(
        input_file=str(input_file),
        output_file=str(output_file),
        start_pos="00:00",
        end_pos="20:00",
        reencode=True,
    )

    assert ok is False
    assert called["duration"] == video.config.MAX_CLIP_DURATION_SECONDS + 10

