"""Unit tests for the shared Flask response/parse helpers in server_utils.

These back the LOC-reduction refactor of the four server blueprints: the
helpers must preserve the exact ``{"ok": ...}`` envelope + status codes the
blueprint tests assert on, and ``parse_number_arg`` must raise a uniform
``ApiError`` that ``json_endpoint`` turns into a 400.
"""

import math

import pytest

Flask = pytest.importorskip("flask").Flask

import server_utils


@pytest.fixture
def app():
    return Flask(__name__)


def test_ok_envelope(app):
    with app.app_context():
        resp = server_utils.ok(pin=3, items=["a"])
        body = resp.get_json()
    assert body == {"ok": True, "pin": 3, "items": ["a"]}


def test_ok_empty(app):
    with app.app_context():
        body = server_utils.ok().get_json()
    assert body == {"ok": True}


def test_err_default_code(app):
    with app.app_context():
        resp, code = server_utils.err("bad input")
    assert code == 400
    assert resp.get_json() == {"ok": False, "error": "bad input"}


def test_err_custom_code(app):
    with app.app_context():
        resp, code = server_utils.err("missing", 404)
    assert code == 404
    assert resp.get_json() == {"ok": False, "error": "missing"}


# ---- parse_number_arg ----


def test_parse_float():
    assert server_utils.parse_number_arg("1.5", "x") == 1.5


def test_parse_int_only_from_float_string():
    # int_only parses via int(float(...)) so "3.0" and "3" both work.
    assert server_utils.parse_number_arg("3.0", "n", int_only=True) == 3
    assert server_utils.parse_number_arg(3, "n", int_only=True) == 3
    assert isinstance(server_utils.parse_number_arg("3.0", "n", int_only=True), int)


def test_parse_bad_value_raises():
    with pytest.raises(server_utils.ApiError) as ei:
        server_utils.parse_number_arg("abc", "timestamp")
    assert ei.value.code == 400
    assert "timestamp must be a number" in ei.value.message


def test_parse_none_raises():
    with pytest.raises(server_utils.ApiError):
        server_utils.parse_number_arg(None, "x")


def test_parse_min_bound():
    assert server_utils.parse_number_arg("0", "x", min_=0) == 0.0
    with pytest.raises(server_utils.ApiError) as ei:
        server_utils.parse_number_arg("-1", "x", min_=0)
    assert "must be >= 0" in ei.value.message


def test_parse_max_bound():
    with pytest.raises(server_utils.ApiError) as ei:
        server_utils.parse_number_arg("11", "x", max_=10)
    assert "must be <= 10" in ei.value.message


def test_parse_finite_rejects_inf():
    with pytest.raises(server_utils.ApiError) as ei:
        server_utils.parse_number_arg("inf", "x", finite=True)
    assert "finite" in ei.value.message


def test_parse_int_only_rejects_nan():
    # int_only implies finite (int(nan) would raise anyway).
    with pytest.raises(server_utils.ApiError):
        server_utils.parse_number_arg("nan", "x", int_only=True)


def test_parse_non_finite_allowed_without_flag():
    # Without finite/int_only, inf passes through (matches lenient call sites).
    assert math.isinf(server_utils.parse_number_arg("inf", "x"))


# ---- json_endpoint ----


def test_json_endpoint_catches_apierror(app):
    @server_utils.json_endpoint
    def handler():
        raise server_utils.ApiError("nope", 422)

    with app.app_context():
        resp, code = handler()
    assert code == 422
    assert resp.get_json() == {"ok": False, "error": "nope"}


def test_json_endpoint_passes_through_success(app):
    @server_utils.json_endpoint
    def handler():
        return server_utils.ok(value=1)

    with app.app_context():
        body = handler().get_json()
    assert body == {"ok": True, "value": 1}


def test_json_endpoint_does_not_swallow_other_exceptions(app):
    @server_utils.json_endpoint
    def handler():
        raise ValueError("real bug")

    with app.app_context(), pytest.raises(ValueError):
        handler()


# ---- require_json_body ----


def test_require_json_body_returns_the_object(app):
    with app.test_request_context(json={"a": 1}):
        assert server_utils.require_json_body() == {"a": 1}


@pytest.mark.parametrize("body", [None, {}, [1, 2]])
def test_require_json_body_rejects_missing_empty_or_non_object(app, body):
    with (
        app.test_request_context(json=body),
        pytest.raises(server_utils.ApiError) as ei,
    ):
        server_utils.require_json_body("Missing JSON body")
    assert ei.value.code == 400
    assert ei.value.message == "Missing JSON body"


def test_find_by_id_returns_first_match():
    items = [{"id": "a", "n": 1}, {"id": "b", "n": 2}, {"id": "b", "n": 3}]
    assert server_utils.find_by_id(items, "b") == {"id": "b", "n": 2}


def test_find_by_id_missing_and_idless_entries():
    assert server_utils.find_by_id([], "x") is None
    assert server_utils.find_by_id([{"name": "no id"}], "x") is None


def test_remove_by_id_pops_in_place():
    items = [{"id": "a"}, {"id": "b"}]
    removed = server_utils.remove_by_id(items, "a")
    assert removed == {"id": "a"}
    assert items == [{"id": "b"}]


def test_remove_by_id_missing_leaves_list_untouched():
    items = [{"id": "a"}]
    assert server_utils.remove_by_id(items, "z") is None
    assert items == [{"id": "a"}]


def test_opt_number_missing_returns_default():
    assert server_utils.opt_number({}, "x") is None
    assert server_utils.opt_number({}, "x", 3.0) == 3.0


def test_opt_number_parses_and_falls_back():
    assert server_utils.opt_number({"x": "1.5"}, "x") == 1.5
    assert server_utils.opt_number({"x": "nope"}, "x", 7.0) == 7.0
