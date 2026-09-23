import utils


def test_parse_timestamps_parses_single_and_range():
    single = utils.parse_timestamps("01:23")
    assert single == [("01:23", utils.add_duration("01:23"))]

    range_only = utils.parse_timestamps("1:23-1:45")
    assert range_only == [("1:23", "1:45")]


def test_parse_timestamps_rejects_half_garbage_dash_ranges():
    # A dash range must have a valid timestamp on BOTH sides; half-garbage
    # ranges are skipped like any other unparseable token instead of flowing
    # to ffmpeg as bogus (start, end) pairs.
    assert utils.parse_timestamps("1:23-abc") == []
    assert utils.parse_timestamps("5-10") == []
    assert utils.parse_timestamps("1:23-1:45-2:00") == []
    # Valid ranges still parse, including hours and the MM:SS >59-minute form.
    assert utils.parse_timestamps("1:23-1:45") == [("1:23", "1:45")]
    assert utils.parse_timestamps("1:02:03-1:04:05") == [("1:02:03", "1:04:05")]
    assert utils.parse_timestamps("75:00-80:00") == [("75:00", "80:00")]


def test_parse_timestamps_parses_multiple_separators_and_ignored_tokens(monkeypatch):
    monkeypatch.setattr(
        utils.config,
        "IGNORED_TIMESTAMP_TOKENS",
        {"x"},
        raising=False,
    )
    # get_ignored_timestamp_tokens() is @functools.cache; clear it so the
    # monkeypatched config value is read, and clear again at the end so the
    # cached set can't leak to later tests.
    utils.get_ignored_timestamp_tokens.cache_clear()
    try:
        value = "00:10-00:20, 00:30; 00:40 + x"
        parsed = utils.parse_timestamps(value)
        assert len(parsed) == 3
        assert parsed[0] == ("00:10", "00:20")
        assert parsed[1][0] == "00:30"
        assert parsed[2][0] == "00:40"
    finally:
        utils.get_ignored_timestamp_tokens.cache_clear()


def test_timestamp_to_seconds_valid_and_invalid():
    assert utils.timestamp_to_seconds("00:10") == 10.0
    assert utils.timestamp_to_seconds("01:02:03") == 3723.0
    assert utils.timestamp_to_seconds("  00:05  ") == 5.0
    assert utils.timestamp_to_seconds("") is None
    assert utils.timestamp_to_seconds("not-a-time") is None


def test_timestamp_to_seconds_allows_minutes_over_59_in_mmss():
    # MM:SS with minutes >= 60 is a valid way to write a long-session offset
    # without an hours component; it must parse rather than be dropped.
    assert utils.timestamp_to_seconds("75:00") == 4500.0
    assert utils.timestamp_to_seconds("90:30") == 5430.0
    # Single MM:SS over 59 minutes also flows through parse_timestamps.
    assert utils.parse_timestamps("75:00") == [("75:00", utils.add_duration("75:00"))]
    # Seconds field still bounded to < 60, and minutes in HH:MM:SS to < 60.
    assert utils.timestamp_to_seconds("10:60") is None
    assert utils.timestamp_to_seconds("01:60:00") is None
    assert utils.timestamp_to_seconds("1:2:3:4") is None


def test_split_selector_tokens_splits_on_comma_and_plus():
    # Shared helper for CLI/spreadsheet selectors: split on ',' or '+',
    # strip whitespace, drop empty tokens.
    assert utils.split_selector_tokens("P01,P02") == ["P01", "P02"]
    assert utils.split_selector_tokens("P01+P02") == ["P01", "P02"]
    assert utils.split_selector_tokens("P01 + P02, P03") == ["P01", "P02", "P03"]
    assert utils.split_selector_tokens("1++2") == ["1", "2"]
    assert utils.split_selector_tokens("  ,+ ") == []
    assert utils.split_selector_tokens("") == []


def test_seconds_to_timestamp_handles_float_input():
    # Shared helper: a float arg must not crash the int format specs.
    assert utils.seconds_to_timestamp(75.9) == "1:15"
    assert utils.seconds_to_timestamp(3723.0, force_hours=True) == "1:02:03"
    assert utils.seconds_to_timestamp(-5) == "0:00"


def test_format_filesize_promotes_at_exact_power_of_1024():
    # >= boundary: an exact KiB/MiB must promote to the next unit.
    assert utils.format_filesize(1024) == "1.00KB"
    assert utils.format_filesize(1024 * 1024) == "1.00MB"
    assert utils.format_filesize(512) == "512.00B"
    assert utils.format_filesize(1536) == "1.50KB"


def test_convert_clock_pairs_to_relative_uses_baseline_and_skips_invalid():
    pairs = [("22:02:12", "22:02:45"), ("22:01:00", "22:01:30")]
    baseline = "22:00"  # HH:MM = 22:00:00
    converted = utils.convert_clock_pairs_to_relative(pairs, baseline, cell_ref="B5")
    # Both pairs are after baseline; converted to relative (2:12-2:45 and 1:00-1:30).
    assert converted == [("0:02:12", "0:02:45"), ("0:01:00", "0:01:30")]


def test_convert_clock_pairs_to_relative_handles_post_midnight_wraparound():
    # Baseline at 22:00, segment at 01:30 the next morning → +3:30:00 from start.
    pairs = [("01:30:00", "01:32:00")]
    baseline = "22:00"
    converted = utils.convert_clock_pairs_to_relative(pairs, baseline)
    assert converted == [("3:30:00", "3:32:00")]


def test_convert_clock_pairs_to_relative_handles_midnight_crossing_span():
    # Baseline 23:30, span 23:55-00:05 should yield 0:25:00-0:35:00.
    pairs = [("23:55:00", "00:05:00")]
    baseline = "23:30"
    converted = utils.convert_clock_pairs_to_relative(pairs, baseline)
    assert converted == [("0:25:00", "0:35:00")]


def test_convert_clock_pairs_to_relative_rejects_unreasonable_pre_baseline():
    # Baseline 08:00, segment at 07:00 is more likely a typo than a 23-hour
    # overnight recording; should be skipped, not wrapped.
    pairs = [("07:00:00", "07:05:00")]
    baseline = "08:00"
    converted = utils.convert_clock_pairs_to_relative(pairs, baseline)
    assert converted == []


def test_cluster_spans_empty_returns_empty():
    assert utils.cluster_spans([], gap_seconds=5.0) == []


def test_cluster_spans_disabled_yields_one_per_input():
    spans = [(0.0, 1.0), (10.0, 11.0), (5.0, 6.0)]
    out = utils.cluster_spans(spans, gap_seconds=0.0)
    # Sorted by start; one cluster per span, padding zero so identical bounds.
    assert out == [
        (0.0, 1.0, [0]),
        (5.0, 6.0, [2]),
        (10.0, 11.0, [1]),
    ]


def test_cluster_spans_merges_within_gap():
    # gap=2.0 means spans with <= 2.0s separation merge.
    spans = [(0.0, 1.0), (2.5, 3.0), (10.0, 11.0)]
    out = utils.cluster_spans(spans, gap_seconds=2.0)
    assert out == [(0.0, 3.0, [0, 1]), (10.0, 11.0, [2])]


def test_cluster_spans_pads_and_clamps_low_to_zero():
    spans = [(2.0, 3.0)]
    out = utils.cluster_spans(spans, gap_seconds=0.0, pad_pre=10.0, pad_post=4.0)
    assert out == [(0.0, 7.0, [0])]


def test_cluster_spans_splits_on_max_duration():
    spans = [(0.0, 30.0)]
    out = utils.cluster_spans(spans, gap_seconds=0.0, max_duration=10.0)
    assert out == [
        (0.0, 10.0, [0]),
        (10.0, 20.0, [0]),
        (20.0, 30.0, [0]),
    ]


def test_cluster_spans_preserves_member_indices_for_unsorted_input():
    # Index 1 is the earliest by start; should appear first in the cluster.
    spans = [(5.0, 6.0), (0.0, 1.0)]
    out = utils.cluster_spans(spans, gap_seconds=10.0)
    assert out == [(0.0, 6.0, [1, 0])]


def test_apply_span_padding_noop_by_default():
    assert utils.apply_span_padding(10.0, 20.0) == (10.0, 20.0)


def test_apply_span_padding_extends_both_ends():
    assert utils.apply_span_padding(10.0, 20.0, pad_pre=3.0, pad_post=2.0) == (
        7.0,
        22.0,
    )


def test_apply_span_padding_negative_pads_trim_inward():
    assert utils.apply_span_padding(10.0, 20.0, pad_pre=-2.0, pad_post=-3.0) == (
        12.0,
        17.0,
    )


def test_apply_span_padding_floors_start_at_zero():
    assert utils.apply_span_padding(1.0, 5.0, pad_pre=5.0) == (0.0, 5.0)


def test_apply_span_padding_enforces_minimum_one_second_span():
    # Trimming the end below start+1 clamps the end back up.
    assert utils.apply_span_padding(10.0, 10.4, pad_post=-1.0) == (10.0, 11.0)


def test_apply_span_padding_caps_at_max_duration():
    assert utils.apply_span_padding(10.0, 40.0, max_duration=10.0) == (10.0, 20.0)


def test_apply_span_padding_max_duration_never_below_one_second():
    # A sub-1s max_duration cap can't shrink the span below the 1-second floor.
    assert utils.apply_span_padding(10.0, 20.0, max_duration=0.5) == (10.0, 11.0)


def test_apply_span_padding_clamps_end_to_limit():
    # A large pad_post past EOF is clamped to the limit, not left to be skipped.
    assert utils.apply_span_padding(90.0, 100.0, pad_post=20.0, limit=105.0) == (
        90.0,
        105.0,
    )


def test_apply_span_padding_keeps_span_inside_tiny_limit():
    # start is pulled back so at least a 1s span survives within the limit.
    assert utils.apply_span_padding(50.0, 60.0, pad_post=10.0, limit=3.0) == (2.0, 3.0)


def test_parse_cell_annotations_splits_segment_and_cell_annotations():
    # !key should annotate the preceding timestamp and also appear as a cell-level annotation.
    cleaned, segment_annotations, cell_annotations = utils.parse_cell_annotations(
        "00:10-00:20 !key 00:30-00:40"
    )
    assert cleaned == "00:10-00:20 00:30-00:40"
    assert "key" in cell_annotations
    assert segment_annotations["key"] == {0}


def test_normalize_track_language_maps_two_letter_to_iso639_2():
    """Whisper reports ISO 639-1; media containers store only ISO 639-2."""
    assert utils.normalize_track_language("en") == "eng"
    assert utils.normalize_track_language("zh") == "zho"
    assert utils.normalize_track_language("de") == "deu"
    # Case and stray whitespace are the user's, not the muxer's, problem.
    assert utils.normalize_track_language("  EN  ") == "eng"


def test_normalize_track_language_prefers_the_terminological_variant():
    """ISO 639-2 has two codes for ~20 languages and ffmpeg stores whichever it
    is handed, so the choice is ours: /T is what Matroska specifies and what
    current players expect."""
    assert utils.normalize_track_language("de") == "deu"  # not "ger"
    assert utils.normalize_track_language("fr") == "fra"  # not "fre"
    assert utils.normalize_track_language("nl") == "nld"  # not "dut"


def test_normalize_track_language_falls_back_to_und_not_junk():
    """transcripts.py falls back to the literal "unknown" when detection fails.
    Passed through, the mp4 muxer truncates it to the nonsense tag "unk" — so
    anything unrecognized has to become the standard undetermined tag instead."""
    for junk in ("", "   ", None, "unknown", "english", "x", "12"):
        assert utils.normalize_track_language(junk) == "und", junk


def test_normalize_track_language_passes_through_three_letter_codes():
    # Already 639-2/639-3 (including Whisper's own "haw"/"yue"), or a code the
    # table does not carry — either way it is already container-shaped.
    for code in ("eng", "haw", "yue", "und"):
        assert utils.normalize_track_language(code) == code


def test_normalize_track_language_keeps_only_the_primary_subtag():
    # TRANSCRIBE_LANGUAGE is user-editable, so a BCP 47 tag is plausible input.
    assert utils.normalize_track_language("en-US") == "eng"
    assert utils.normalize_track_language("pt-BR") == "por"
    assert utils.normalize_track_language("zh_Hans") == "zho"


def test_normalize_track_language_covers_every_whisper_language():
    """The table exists to serve transcription output, so a Whisper code that
    silently degrades to "und" is the exact bug this function was added to fix.
    Skips rather than fails where faster-whisper is not installed."""
    import transcripts

    # av is deliberately not installed (pyproject override); without the stub
    # the package import fails and importorskip would silently drop this test.
    transcripts._ensure_av_stub()
    pytest = __import__("pytest")
    tokenizer = pytest.importorskip("faster_whisper.tokenizer")
    unmapped = [
        code
        for code in sorted(tokenizer._LANGUAGE_CODES)
        if utils.normalize_track_language(code) == "und"
    ]
    assert not unmapped, f"Whisper languages missing from ISO639_1_TO_2: {unmapped}"
