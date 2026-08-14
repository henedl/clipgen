"""Tests for the MindNode mind-map reader and its Studio intake plumbing.

Two kinds of fixture, deliberately:

* ``sample_bundle`` is a real document written by MindNode, checked in under
  ``tests/fixtures/``. It is the only guard against the parser drifting from
  what the app actually writes — a hand-built plist can agree with our
  assumptions and still not match reality.
* ``make_bundle`` synthesizes bundles for the shapes one sample cannot cover
  (deeper nesting, several roots, entity-escaped titles, malformed files).
"""

import json
import plistlib
from pathlib import Path

import pytest

import config
import mindnode
import utils


FIXTURES = Path(__file__).parent / "fixtures"
SAMPLE = FIXTURES / "mindnode-sample.mindnode"


# ---- Fixture construction ----------------------------------------------------


def _node(title, children=(), node_id=None):
    """One MindNode node, titled with the HTML fragment the app really writes."""
    return {
        "title": {
            "text": (
                "<p style='color: rgba(109, 109, 109, 1.000000); font: 20px "
                f'"Helvetica"; text-align: left; \'>{title}</p>'
            ),
            "maxWidth": 300.0,
            "allowToShrinkWidth": True,
        },
        "subnodes": list(children),
        "nodeID": node_id or f"id-{title}",
        "location": "{0, 0}",
        "hasFoldedSubnodes": False,
    }


@pytest.fixture
def make_bundle(tmp_path):
    """Write a ``.mindnode`` package holding the given root nodes."""

    def _make(*roots, name="test.mindnode", contents=None):
        bundle = tmp_path / name
        bundle.mkdir()
        if contents is None:
            contents = {
                "version": 9,
                "canvas": {"mindMaps": [{"mainNode": r} for r in roots]},
            }
        with (bundle / mindnode.CONTENTS_FILENAME).open("wb") as fh:
            plistlib.dump(contents, fh, fmt=plistlib.FMT_BINARY)
        return bundle

    return _make


# ---- The real document -------------------------------------------------------


@pytest.mark.skipif(not SAMPLE.is_dir(), reason="sample bundle not checked out")
def test_real_bundle_parses_to_expected_notes():
    """The checked-in MindNode document yields exactly the tree it draws."""
    doc = mindnode.parse_document(SAMPLE)

    assert doc["study"] == "clipgen-test"
    assert doc["participants"] == ["P01", "P02"]
    assert doc["with_times"] == 4
    assert doc["without_times"] == 4

    got = [
        (n["category"], n["participant"], n["desc"], n["times"]) for n in doc["notes"]
    ]
    assert got == [
        ("Question 1", "P01", "Note 1", []),
        ("Question 1", "P01", "Note 2", []),
        ("Question 1", "P01", "Note 3", [("0:01:00", "0:02:00")]),
        ("Question 1", "P02", "Note 1", []),
        # A bare-timestamp node keeps an empty description on purpose: the
        # filename already carries study, participant and category.
        ("Question 1", "P02", "", [("0:04:00", "0:07:00")]),
        ("Question 1", "P02", "Note 2", [("0:03:00", "0:04:00")]),
        ("Question 2", "P01", "Note 4", []),
        ("Question 2", "P01", "Note 5", [("0:02:00", "0:03:00")]),
    ]


@pytest.mark.skipif(not SAMPLE.is_dir(), reason="sample bundle not checked out")
def test_real_bundle_spans_are_seconds():
    """`spans` mirrors `times` in seconds, which is what the frontend cuts on."""
    doc = mindnode.parse_document(SAMPLE)
    spans = {n["desc"]: n["spans"] for n in doc["notes"] if n["spans"]}
    assert spans["Note 3"] == [(60.0, 120.0)]
    assert spans["Note 5"] == [(120.0, 180.0)]


# ---- Title text --------------------------------------------------------------


def test_node_text_strips_html_and_unescapes(make_bundle):
    bundle = make_bundle(
        _node("Tom &amp; Jerry &lt;3", [_node("P01", [_node("x 0:01:00")])])
    )
    roots = mindnode.load_document(bundle)
    assert mindnode.node_text(roots[0]) == "Tom & Jerry <3"


def test_node_text_maps_breaks_to_newlines():
    node = {"title": {"text": "<p style='x'>one<br>two</p><p style='y'>three</p>"}}
    assert mindnode.node_text(node) == "one\ntwo\nthree"


def test_node_text_of_untitled_node_is_empty():
    assert mindnode.node_text({}) == ""
    assert mindnode.node_text({"title": {}}) == ""


# ---- Participant auto-detection and category paths ---------------------------


def test_participant_level_is_detected_at_any_depth(make_bundle):
    """Depth is not fixed: the ^[PG]\\d+$ node is found wherever it sits."""
    bundle = make_bundle(
        _node(
            "study",
            [
                _node("Question 1", [_node("P01", [_node("shallow 0:01:00")])]),
                _node(
                    "Question 2",
                    [_node("Sub A", [_node("G02", [_node("deep 0:02:00")])])],
                ),
            ],
        )
    )
    doc = mindnode.parse_document(bundle)
    by_participant = {n["participant"]: n for n in doc["notes"]}
    assert by_participant["P01"]["category"] == "Question 1"
    assert by_participant["G02"]["category"] == "Question 2 / Sub A"


def test_group_prefix_is_accepted(make_bundle):
    bundle = make_bundle(
        _node("study", [_node("Q", [_node("G07", [_node("n 0:01:00")])])])
    )
    doc = mindnode.parse_document(bundle)
    assert doc["participants"] == ["G07"]


def test_participant_id_is_uppercased(make_bundle):
    bundle = make_bundle(
        _node("study", [_node("Q", [_node("p03", [_node("n 0:01:00")])])])
    )
    doc = mindnode.parse_document(bundle)
    assert doc["participants"] == ["P03"]


def test_non_participant_leaf_is_not_mistaken_for_one(make_bundle):
    """A node has to match the id shape exactly, not merely start with P."""
    bundle = make_bundle(
        _node("study", [_node("Pricing", [_node("P01", [_node("n 0:01:00")])])])
    )
    doc = mindnode.parse_document(bundle)
    assert doc["notes"][0]["category"] == "Pricing"
    assert doc["notes"][0]["participant"] == "P01"


def test_participant_directly_under_root_has_empty_category(make_bundle):
    bundle = make_bundle(_node("study", [_node("P01", [_node("n 0:01:00")])]))
    doc = mindnode.parse_document(bundle)
    assert doc["notes"][0]["category"] == ""


# ---- Note collection ---------------------------------------------------------


def test_untimestamped_notes_are_kept_not_dropped(make_bundle):
    """They cannot be cut, but losing them silently would read as data loss."""
    bundle = make_bundle(
        _node(
            "study",
            [_node("Q", [_node("P01", [_node("no time"), _node("has 0:01:00")])])],
        )
    )
    doc = mindnode.parse_document(bundle)
    assert [n["desc"] for n in doc["notes"]] == ["no time", "has"]
    assert doc["with_times"] == 1
    assert doc["without_times"] == 1


def test_note_with_several_pairs_keeps_them_all(make_bundle):
    bundle = make_bundle(
        _node(
            "study",
            [_node("Q", [_node("P01", [_node("many 0:01:00, 0:05:00-0:06:00")])])],
        )
    )
    (note,) = mindnode.parse_document(bundle)["notes"]
    assert note["times"] == [("0:01:00", "0:02:00"), ("0:05:00", "0:06:00")]
    assert note["spans"] == [(60.0, 120.0), (300.0, 360.0)]


def test_nested_notes_are_flattened_without_double_counting(make_bundle):
    """A timestamp-free branch groups its children; it is not an observation."""
    bundle = make_bundle(
        _node(
            "study",
            [
                _node(
                    "Q",
                    [
                        _node(
                            "P01",
                            [
                                _node("grouping", [_node("child 0:01:00")]),
                                _node("observation 0:02:00", [_node("detail 0:03:00")]),
                            ],
                        )
                    ],
                )
            ],
        )
    )
    doc = mindnode.parse_document(bundle)
    # "grouping" carries no timestamp and has children → skipped as a heading.
    # "observation" carries one → kept, alongside its own child.
    assert [n["desc"] for n in doc["notes"]] == ["child", "observation", "detail"]


def test_annotations_survive(make_bundle):
    """`!key` behaves exactly as it does in a spreadsheet cell."""
    bundle = make_bundle(
        _node("study", [_node("Q", [_node("P01", [_node("marked 0:01:00 !key")])])])
    )
    (note,) = mindnode.parse_document(bundle)["notes"]
    assert note["annotations"] == ["key"]
    assert note["desc"] == "marked"


def test_ignored_token_is_not_description_text(make_bundle):
    bundle = make_bundle(
        _node("study", [_node("Q", [_node("P01", [_node("skip x")])])])
    )
    (note,) = mindnode.parse_document(bundle)["notes"]
    assert note["desc"] == "skip"
    assert note["times"] == []


@pytest.mark.parametrize(
    "title",
    [
        "Note 3 0:01:00+0:05:00",
        "Note 3 0:01:00,0:05:00",
        "Note 3 0:01:00;0:05:00",
        "Note 3 0:01:00, 0:05:00",
    ],
)
def test_separator_joined_timestamps_leave_the_description(make_bundle, title):
    """utils._split_timestamp_tokens splits on + ; , as well as whitespace, so
    all four forms parse to the same two pairs. A whitespace-only split in
    _describe left the raw times in the description — and desc becomes the
    intake event_type, so they ended up in the output filename too."""
    bundle = make_bundle(_node("study", [_node("Q", [_node("P01", [_node(title)])])]))
    (note,) = mindnode.parse_document(bundle)["notes"]
    assert note["desc"] == "Note 3"
    assert note["times"] == [("0:01:00", "0:02:00"), ("0:05:00", "0:06:00")]


def test_prose_punctuation_survives_the_timestamp_strip(make_bundle):
    """Splitting the whole string on , would also eat commas out of real prose,
    so only tokens that are *entirely* timestamps are dropped."""
    bundle = make_bundle(
        _node("study", [_node("Q", [_node("P01", [_node("obs, then this 1:00")])])])
    )
    (note,) = mindnode.parse_document(bundle)["notes"]
    assert note["desc"] == "obs, then this"


def test_malformed_xml_plist_raises_value_error(tmp_path):
    """plistlib raises ExpatError (not ValueError/OSError/InvalidFileException)
    on a truncated *XML* plist. Every caller catches only ValueError, so
    without normalizing it here the Start overlay's preview and Open both
    return a 500 with a stack trace instead of a readable message."""
    bundle = tmp_path / "broken.mindnode"
    bundle.mkdir()
    (bundle / mindnode.CONTENTS_FILENAME).write_bytes(
        b'<?xml version="1.0"?>\n<!DOCTYPE plist><plist version="1.0"><dict>'
    )
    with pytest.raises(ValueError, match="Could not read"):
        mindnode.parse_document(bundle)


def test_node_ids_are_carried_through(make_bundle):
    bundle = make_bundle(
        _node(
            "study",
            [_node("Q", [_node("P01", [_node("n 0:01:00", node_id="UUID-1")])])],
        )
    )
    (note,) = mindnode.parse_document(bundle)["notes"]
    assert note["id"] == "UUID-1"


# ---- Several roots -----------------------------------------------------------


def test_multiple_roots_are_all_walked(make_bundle):
    """`mindMaps` is a list — a document can hold several detached trees."""
    bundle = make_bundle(
        _node("first study", [_node("P01", [_node("a 0:01:00")])], node_id="r1"),
        _node("second study", [_node("P02", [_node("b 0:02:00")])], node_id="r2"),
    )
    doc = mindnode.parse_document(bundle)
    assert doc["roots"] == ["first study", "second study"]
    assert doc["study"] == "first_study"  # doc-level study is the first root
    assert [n["study"] for n in doc["notes"]] == ["first_study", "second_study"]


# ---- Failure paths -----------------------------------------------------------


def test_missing_bundle_raises(tmp_path):
    with pytest.raises(ValueError, match="Not a MindNode document"):
        mindnode.load_document(tmp_path / "nope.mindnode")


def test_directory_without_contents_raises(tmp_path):
    plain = tmp_path / "plain.mindnode"
    plain.mkdir()
    with pytest.raises(ValueError, match="Not a MindNode document"):
        mindnode.load_document(plain)


def test_unparseable_contents_raises(tmp_path):
    bundle = tmp_path / "broken.mindnode"
    bundle.mkdir()
    (bundle / mindnode.CONTENTS_FILENAME).write_bytes(b"this is not a plist")
    with pytest.raises(ValueError, match="Could not read"):
        mindnode.load_document(bundle)


@pytest.mark.parametrize(
    "contents",
    [
        {"version": 9},
        {"canvas": {}},
        {"canvas": {"mindMaps": []}},
        {"canvas": {"mindMaps": [{"notMainNode": {}}]}},
    ],
    ids=["no_canvas", "no_mindmaps", "empty_mindmaps", "no_mainnode"],
)
def test_plist_without_a_mind_map_raises(make_bundle, contents):
    bundle = make_bundle(contents=contents)
    with pytest.raises(ValueError, match="No mind maps found"):
        mindnode.load_document(bundle)


def test_unknown_node_keys_are_ignored(make_bundle):
    """Richer maps carry notes/tags/tasks; the parser must not choke on them."""
    node = _node("n 0:01:00")
    node.update(
        {"note": "a long note", "tags": ["x"], "taskState": 1, "completed": True}
    )
    bundle = make_bundle(_node("study", [_node("Q", [_node("P01", [node])])]))
    (parsed,) = mindnode.parse_document(bundle)["notes"]
    assert parsed["desc"] == "n"


def test_non_dict_subnodes_are_skipped(make_bundle):
    root = _node("study", [_node("Q", [_node("P01", [_node("n 0:01:00")])])])
    root["subnodes"].append("not a node")
    bundle = make_bundle(root)
    assert len(mindnode.parse_document(bundle)["notes"]) == 1


# ---- Timestamp-token agreement with utils ------------------------------------


@pytest.mark.parametrize(
    "token",
    [
        "0:01:00",
        "1:23",
        "1:23-1:45",
        "0:04:00-0:07:00",
        "75:00",
        "0.01.00",
        "note",
        "Note",
        "3",
        "x",
        "",
        "-5",
        "5-10",
        "1:23-abc",
        "P01",
    ],
)
def test_timestamp_token_test_agrees_with_utils(token):
    """`_is_timestamp_token` must accept exactly what `parse_timestamps` parses.

    They are separate only because the public parser warns on every rejected
    token, and here rejection is the normal case. If they drift, description
    text silently gains timestamps or loses words.
    """
    parses = bool(utils.parse_timestamps(token))
    ignored = token.strip().lower() in utils.get_ignored_timestamp_tokens()
    assert mindnode._is_timestamp_token(token) == (parses or ignored)


def test_participant_regex_follows_config(monkeypatch):
    """The prefixes are config, not a literal in the parser."""
    monkeypatch.setattr(config, "PARTICIPANT_PREFIXES", ("Z",))
    pattern = mindnode._participant_re()
    assert pattern.match("Z01")
    assert not pattern.match("P01")


# ---- Document discovery ------------------------------------------------------


def test_find_documents_lists_bundles(tmp_path, make_bundle):
    make_bundle(_node("study", [_node("P01", [_node("n 0:01:00")])]), name="a.mindnode")
    (tmp_path / "not-a-bundle.mindnode").mkdir()  # no contents.xml → skipped
    (tmp_path / "sheet.xlsx").write_text("x")
    found = mindnode.find_documents(tmp_path)
    assert [f["name"] for f in found] == ["a.mindnode"]
    assert found[0]["has_preview"] is False


def test_find_documents_reports_preview(tmp_path, make_bundle):
    bundle = make_bundle(
        _node("s", [_node("P01", [_node("n 0:01:00")])]), name="p.mindnode"
    )
    (bundle / "QuickLook").mkdir()
    (bundle / mindnode.PREVIEW_RELPATH).write_bytes(b"jpegdata")
    assert mindnode.find_documents(tmp_path)[0]["has_preview"] is True


def test_find_documents_of_missing_dir_is_empty(tmp_path):
    assert mindnode.find_documents(tmp_path / "gone") == []


# ---- Server routes -----------------------------------------------------------


@pytest.fixture
def mn_client(monkeypatch, tmp_path, make_bundle):
    """A combined-app client with the input dir pointed at a bundle."""
    pytest.importorskip("flask")
    import server

    bundle = make_bundle(
        _node(
            "My Study",
            [_node("Q1", [_node("P01", [_node("note 0:01:00"), _node("no time")])])],
        )
    )
    monkeypatch.setattr(config, "INPUT_DIR", str(tmp_path))
    monkeypatch.setattr(server, "_worksheet", None)
    monkeypatch.setattr(server, "_sheet_context", None)
    # Both are module-level and survive the request; a test that opens a
    # document would otherwise leak it into the next one.
    monkeypatch.setattr(server, "_mindnode_doc", None)
    monkeypatch.setattr(server, "_active_project_source", None)
    monkeypatch.setattr(server, "_active_sheet_meta", None)
    # The routes write the real per-user settings file otherwise.
    monkeypatch.setattr(
        "start_settings.record_recent_spreadsheet", lambda *a, **k: None
    )
    monkeypatch.setattr("start_settings.record_project_session", lambda *a, **k: None)
    app = server.build_combined_app(default_page="studio")
    with app.test_client() as client:
        yield client, bundle, server
    monkeypatch.setattr(server, "_mindnode_doc", None)
    monkeypatch.setattr(server, "_active_project_source", None)


def test_route_lists_documents(mn_client):
    client, bundle, _ = mn_client
    body = client.get("/api/spreadsheets/mindnode").get_json()
    assert body["ok"] is True
    assert [f["name"] for f in body["files"]] == [bundle.name]


def test_route_previews_without_opening(mn_client):
    client, bundle, server = mn_client
    body = client.get(
        "/api/spreadsheets/mindnode/preview", query_string={"path": str(bundle)}
    ).get_json()
    assert body["study"] == "my_study"
    assert body["participants"] == ["P01"]
    assert body["notes"] == 2
    assert body["with_times"] == 1
    assert body["without_times"] == 1
    assert body["categories"] == ["Q1"]
    # Read-only: previewing must not make it the session's document.
    assert server._mindnode_doc is None


def test_route_preview_rejects_a_bad_path(mn_client):
    client, _, _ = mn_client
    resp = client.get(
        "/api/spreadsheets/mindnode/preview", query_string={"path": "/nope"}
    )
    assert resp.status_code == 400
    assert resp.get_json()["ok"] is False


def test_route_preview_requires_a_path(mn_client):
    client, _, _ = mn_client
    assert client.get("/api/spreadsheets/mindnode/preview").status_code == 400


def test_thumb_404s_without_a_quicklook_render(mn_client):
    client, bundle, _ = mn_client
    resp = client.get(
        "/api/spreadsheets/mindnode/thumb", query_string={"path": str(bundle)}
    )
    assert resp.status_code == 404


def test_thumb_serves_the_bundle_render(mn_client):
    client, bundle, _ = mn_client
    (bundle / "QuickLook").mkdir()
    (bundle / mindnode.PREVIEW_RELPATH).write_bytes(b"\xff\xd8jpeg")
    resp = client.get(
        "/api/spreadsheets/mindnode/thumb", query_string={"path": str(bundle)}
    )
    assert resp.status_code == 200
    assert resp.mimetype == "image/jpeg"


def test_open_close_round_trip(mn_client):
    client, bundle, server = mn_client

    assert client.get("/api/status").get_json()["mindnode_loaded"] is False
    assert client.get("/studio/api/mindnode").get_json()["mindnode_loaded"] is False

    opened = client.post(
        "/api/spreadsheets/open", json={"type": "mindnode", "id_or_path": str(bundle)}
    ).get_json()
    assert opened["ok"] is True
    assert opened["mindnode_loaded"] is True
    # A mind map is not a worksheet; Studio stays in its no-sheet state.
    assert opened["sheet_loaded"] is False
    assert opened["study"] == "my_study"

    status = client.get("/api/status").get_json()
    assert status["mindnode_loaded"] is True
    assert status["mindnode_path"] == str(bundle)
    assert status["sheet_loaded"] is False

    doc = client.get("/studio/api/mindnode").get_json()["document"]
    assert len(doc["notes"]) == 2

    closed = client.post(
        "/api/spreadsheets/close", json={"type": "mindnode"}
    ).get_json()
    assert closed["mindnode_loaded"] is False
    assert server._mindnode_doc is None


def test_status_reports_the_recorded_source(mn_client):
    """`active_source` must mirror what record_project_session stored.

    The Start overlay keys its "current session" highlight (and the restored
    project name) off this. Derived from the spreadsheet fields alone, a
    mind-map session keyed to an empty source and matched no stored project.
    """
    client, bundle, _ = mn_client
    assert client.get("/api/status").get_json()["active_source"] is None

    client.post(
        "/api/spreadsheets/open", json={"type": "mindnode", "id_or_path": str(bundle)}
    )
    source = client.get("/api/status").get_json()["active_source"]
    assert source["type"] == "mindnode"
    assert source["id_or_path"] == str(bundle)
    assert source["label"] == bundle.name


def test_closing_the_map_clears_the_recorded_source(mn_client):
    client, bundle, _ = mn_client
    client.post(
        "/api/spreadsheets/open", json={"type": "mindnode", "id_or_path": str(bundle)}
    )
    client.post("/api/spreadsheets/close", json={"type": "mindnode"})
    assert client.get("/api/status").get_json()["active_source"] is None


def test_sheet_route_carries_the_map_study_when_no_sheet(mn_client):
    """The subheader needs a study and a cohort in a mind-map-only session."""
    client, bundle, _ = mn_client

    empty = client.get("/studio/api/sheet").get_json()
    assert empty["study"] == ""
    assert empty["mindnodeParticipants"] == []

    client.post(
        "/api/spreadsheets/open", json={"type": "mindnode", "id_or_path": str(bundle)}
    )
    loaded = client.get("/studio/api/sheet").get_json()
    assert loaded["sheet_loaded"] is False
    assert loaded["study"] == "my_study"
    assert loaded["mindnodeParticipants"] == ["P01"]
    # `participants`/`rows` mean *sheet columns and rows* to Studio's grid and
    # every Overview tab — a mind map must not invent a cohort there.
    assert loaded["participants"] == []
    assert loaded["rows"] == []


def test_open_rejects_a_bad_document(mn_client):
    client, _, server = mn_client
    resp = client.post(
        "/api/spreadsheets/open",
        json={"type": "mindnode", "id_or_path": "/nope.mindnode"},
    )
    assert resp.status_code == 400
    assert server._mindnode_doc is None


def test_open_still_rejects_unknown_types(mn_client):
    client, _, _ = mn_client
    resp = client.post(
        "/api/spreadsheets/open", json={"type": "wat", "id_or_path": "x"}
    )
    assert resp.status_code == 400


def test_document_route_rereads_from_disk(mn_client):
    """Editing the map in MindNode shows up without reopening the workspace."""
    client, bundle, _ = mn_client
    client.post(
        "/api/spreadsheets/open", json={"type": "mindnode", "id_or_path": str(bundle)}
    )
    assert len(client.get("/studio/api/mindnode").get_json()["document"]["notes"]) == 2

    with (bundle / mindnode.CONTENTS_FILENAME).open("wb") as fh:
        plistlib.dump(
            {
                "canvas": {
                    "mindMaps": [
                        {
                            "mainNode": _node(
                                "My Study",
                                [
                                    _node(
                                        "Q1",
                                        [_node("P01", [_node("only one 0:01:00")])],
                                    )
                                ],
                            )
                        }
                    ]
                }
            },
            fh,
            fmt=plistlib.FMT_BINARY,
        )
    doc = client.get("/studio/api/mindnode").get_json()["document"]
    assert [n["desc"] for n in doc["notes"]] == ["only one"]


def test_document_refresh_does_not_resurrect_a_closed_map(mn_client, monkeypatch):
    """The re-parse runs with _mindnode_lock released (it is slow, and the
    route re-reads on every request). A close landing in that window must win —
    otherwise the refresh writes the map back and the server believes it is
    open while the UI has already shut it."""
    import server as server_mod

    client, bundle, _ = mn_client
    client.post(
        "/api/spreadsheets/open", json={"type": "mindnode", "id_or_path": str(bundle)}
    )
    assert server_mod._mindnode_doc is not None

    real_parse = mindnode.parse_document

    def _parse_then_close(path):
        parsed = real_parse(path)
        # The close landing mid-parse, applied directly: a nested test-client
        # request would unwind Flask's request-context stack out of order.
        # This is exactly the state /api/spreadsheets/close leaves behind.
        server_mod._mindnode_doc = None
        return parsed

    # The route imports mindnode function-locally, so patch the module itself.
    monkeypatch.setattr(mindnode, "parse_document", _parse_then_close)

    body = client.get("/studio/api/mindnode").get_json()
    assert body["mindnode_loaded"] is False
    assert body["document"] is None
    assert server_mod._mindnode_doc is None


def test_document_route_404s_when_the_bundle_disappears(mn_client):
    import shutil

    client, bundle, _ = mn_client
    client.post(
        "/api/spreadsheets/open", json={"type": "mindnode", "id_or_path": str(bundle)}
    )
    shutil.rmtree(bundle)
    assert client.get("/studio/api/mindnode").status_code == 404


# ---- Generation plumbing -----------------------------------------------------


def test_intake_video_paths_follow_the_input_dir(monkeypatch, tmp_path):
    """POST /api/dirs moves INPUT_DIR; intake must not keep a boot-time scan.

    The MindNode-only Start overlay never hits /api/participants on Screenspace
    or Transcripts, so those caches stay empty. Generating from the map would
    then 404 every clip with "No video for P01" even though the file is there.
    """
    import screenspace_server
    import transcripts_server
    import server

    boot = tmp_path / "boot"
    boot.mkdir()
    chosen = tmp_path / "chosen"
    chosen.mkdir()
    video = chosen / "study_P01.mp4"
    video.write_bytes(b"fake")

    monkeypatch.setattr(config, "INPUT_DIR", str(chosen))
    monkeypatch.setattr(server, "_sheet_context", None)
    # Stale caches as after boot, before anyone opens Transcripts/Screenspace.
    monkeypatch.setattr(screenspace_server, "_participants", [])
    monkeypatch.setattr(transcripts_server, "_participants", [])

    paths = server._resolve_intake_video_paths("P01", "mindnode")
    assert paths == [str(video)]


def test_effective_study_falls_back_to_the_mind_map(monkeypatch):
    """With no spreadsheet, artifacts must still get a study name."""
    import server

    monkeypatch.setattr(server, "_sheet_context", None)
    monkeypatch.setattr(server, "_mindnode_doc", {"study": "from_the_map"})
    assert server._effective_study() == "from_the_map"

    monkeypatch.setattr(server, "_mindnode_doc", None)
    assert server._effective_study() == ""


def test_intake_item_carries_category_and_study(monkeypatch, tmp_path):
    """A mind-map item's question becomes the artifact's category."""
    import server

    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("video.run_ffmpeg", lambda *a, **kw: True)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))

    result = server._process_intake_item(
        {
            "participant": "P01",
            "start": 60.0,
            "end": 120.0,
            "event_type": "Note 3",
            "event_ids": ["uuid#0"],
            "source": "mindnode",
            "category": "Question 1",
            "study": "per_item_study",
        },
        "clip",
        "batch_study",
    )
    assert result["_ok"] is True
    assert result["category"] == "Question 1"
    assert result["description"] == "Note 3"
    # The item's own study wins — one document can hold several roots.
    assert result["study"] == "per_item_study"
    assert result["source"] == "mindnode"


def test_bare_timestamp_note_gets_a_generic_description(monkeypatch, tmp_path):
    """An empty desc must not fall through to "Screenspace intake"."""
    import server

    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("video.run_ffmpeg", lambda *a, **kw: True)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))

    result = server._process_intake_item(
        {
            "participant": "P01",
            "start": 240.0,
            "end": 420.0,
            "event_type": "",
            "source": "mindnode",
        },
        "clip",
        "study",
    )
    assert result["description"] == "MindNode intake"


def test_other_sources_keep_their_empty_category(monkeypatch, tmp_path):
    """The new category field is additive — it must not change existing sources."""
    import server

    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("video.run_ffmpeg", lambda *a, **kw: True)
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))

    result = server._process_intake_item(
        {"participant": "P01", "start": 0.0, "end": 5.0, "event_type": "blur"},
        "clip",
        "study",
    )
    assert result["category"] == ""
    assert result["description"] == "blur"


def test_generate_intake_accepts_mindnode_items(monkeypatch, tmp_path, make_bundle):
    """End to end: a mind-map item streams an artifact carrying its category."""
    pytest.importorskip("flask")
    import server

    monkeypatch.setattr(
        server, "_resolve_intake_video_paths", lambda p, s="": ["/fake/video.mp4"]
    )
    monkeypatch.setattr("video.run_ffmpeg", lambda *a, **kw: True)
    monkeypatch.setattr(server, "_save_manifest_quiet", lambda: None)
    monkeypatch.setattr(server, "_sheet_context", None)
    monkeypatch.setattr(server, "_mindnode_doc", {"study": "map_study", "path": "x"})
    monkeypatch.setattr(config, "OUTPUT_DIR", str(tmp_path))

    app = server.build_combined_app(default_page="studio")
    with app.test_client() as client:
        resp = client.post(
            "/studio/api/generate-intake",
            json={
                "items": [
                    {
                        "participant": "P01",
                        "start": 60.0,
                        "end": 120.0,
                        "event_type": "Note 3",
                        "event_ids": ["uuid#0"],
                        "source": "mindnode",
                        "category": "Question 1",
                    }
                ],
                "format": "clip",
            },
        )
        assert resp.status_code == 200
        (line,) = [json.loads(x) for x in resp.data.decode().strip().split("\n")]

    assert line["ok"] is True
    assert line["artifact"]["category"] == "Question 1"
    assert line["artifact"]["study"] == "map_study"
