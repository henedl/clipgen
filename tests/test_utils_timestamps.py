import utils


def test_parse_timestamps_parses_single_and_range():
    single = utils.parse_timestamps("01:23")
    assert single == [("01:23", utils.add_duration("01:23"))]

    range_only = utils.parse_timestamps("1:23-1:45")
    assert range_only == [("1:23", "1:45")]


def test_parse_timestamps_parses_multiple_separators_and_ignored_tokens(monkeypatch):
    monkeypatch.setattr(
        utils.config,
        "IGNORED_TIMESTAMP_TOKENS",
        {"x"},
        raising=False,
    )
    value = "00:10-00:20, 00:30; 00:40 + x"
    parsed = utils.parse_timestamps(value)
    assert len(parsed) == 3
    assert parsed[0] == ("00:10", "00:20")
    assert parsed[1][0] == "00:30"
    assert parsed[2][0] == "00:40"


def test_timestamp_to_seconds_valid_and_invalid():
    assert utils.timestamp_to_seconds("00:10") == 10.0
    assert utils.timestamp_to_seconds("01:02:03") == 3723.0
    assert utils.timestamp_to_seconds("  00:05  ") == 5.0
    assert utils.timestamp_to_seconds("") is None
    assert utils.timestamp_to_seconds("not-a-time") is None


def test_parse_cell_annotations_splits_segment_and_cell_annotations():
    # !key should annotate the preceding timestamp and also appear as a cell-level annotation.
    cleaned, segment_annotations, cell_annotations = utils.parse_cell_annotations(
        "00:10-00:20 !key 00:30-00:40"
    )
    assert cleaned == "00:10-00:20 00:30-00:40"
    assert "key" in cell_annotations
    assert segment_annotations["key"] == {0}

