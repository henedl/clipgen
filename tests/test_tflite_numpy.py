"""tflite_numpy: flatbuffer parsing on hand-built models, and each op against numpy."""

import math
from typing import Any

import numpy as np
import pytest

import tflite_numpy

flatbuffers = pytest.importorskip("flatbuffers")

_ADD, _FC = 0, 9
_F32, _I8 = 0, 9


def _vector(b: Any, values: list[int], prepend: str) -> int:
    b.StartVector(4, len(values), 4)
    for v in reversed(values):
        getattr(b, prepend)(v)
    return b.EndVector()


def _offsets(b: Any, offsets: list[int]) -> int:
    b.StartVector(4, len(offsets), 4)
    for off in reversed(offsets):
        b.PrependUOffsetTRelative(off)
    return b.EndVector()


def _build(
    tensors: list[dict[str, Any]],
    ops: list[tuple[int, list[int], list[int], int | None]],
    buffers: list[bytes],
    external: dict[int, int] | None = None,
) -> bytes:
    """A one-subgraph model; ``ops`` hold (builtin code, inputs, outputs, fused act)."""
    b = flatbuffers.Builder(1024)
    codes = sorted({op[0] for op in ops})
    code_tables = []
    for code in codes:
        b.StartObject(4)
        b.PrependInt8Slot(0, min(code, 127), 0)
        b.PrependInt32Slot(3, code, 0)
        code_tables.append(b.EndObject())
    buffer_tables = []
    for index, data in enumerate(buffers):
        inline = data and not (external and index in external)
        vec = (
            b.CreateNumpyVector(np.frombuffer(data, dtype=np.uint8)) if inline else None
        )
        b.StartObject(3)
        if vec is not None:
            b.PrependUOffsetTRelativeSlot(0, vec, 0)
        if external and index in external:
            b.PrependUint64Slot(1, external[index], 0)
            b.PrependUint64Slot(2, len(data), 0)
        buffer_tables.append(b.EndObject())
    tensor_tables = []
    for t in tensors:
        name = b.CreateString(t["name"])
        shape = _vector(b, t["shape"], "PrependInt32")
        quant = None
        if "scale" in t:
            scale = b.CreateNumpyVector(np.asarray(t["scale"], dtype=np.float32))
            zero = b.CreateNumpyVector(np.zeros(len(t["scale"]), dtype=np.int64))
            b.StartObject(7)
            b.PrependUOffsetTRelativeSlot(2, scale, 0)
            b.PrependUOffsetTRelativeSlot(3, zero, 0)
            b.PrependInt32Slot(6, t.get("axis", 0), 0)
            quant = b.EndObject()
        b.StartObject(5)
        b.PrependUOffsetTRelativeSlot(0, shape, 0)
        b.PrependInt8Slot(1, t["type"], 0)
        b.PrependUint32Slot(2, t.get("buffer", 0), 0)
        b.PrependUOffsetTRelativeSlot(3, name, 0)
        if quant is not None:
            b.PrependUOffsetTRelativeSlot(4, quant, 0)
        tensor_tables.append(b.EndObject())
    op_tables = []
    for code, ins, outs, act in ops:
        in_vec = _vector(b, ins, "PrependInt32")
        out_vec = _vector(b, outs, "PrependInt32")
        options = None
        if act is not None:
            b.StartObject(1)
            b.PrependInt8Slot(0, act, 0)
            options = b.EndObject()
        b.StartObject(5)
        b.PrependUint32Slot(0, codes.index(code), 0)
        b.PrependUOffsetTRelativeSlot(1, in_vec, 0)
        b.PrependUOffsetTRelativeSlot(2, out_vec, 0)
        if options is not None:
            b.PrependUOffsetTRelativeSlot(4, options, 0)
        op_tables.append(b.EndObject())
    graph_tensors = _offsets(b, tensor_tables)
    graph_inputs = _vector(
        b, [i for i, t in enumerate(tensors) if t.get("io") == "in"], "PrependInt32"
    )
    graph_outputs = _vector(
        b, [i for i, t in enumerate(tensors) if t.get("io") == "out"], "PrependInt32"
    )
    graph_ops = _offsets(b, op_tables)
    b.StartObject(4)
    b.PrependUOffsetTRelativeSlot(0, graph_tensors, 0)
    b.PrependUOffsetTRelativeSlot(1, graph_inputs, 0)
    b.PrependUOffsetTRelativeSlot(2, graph_outputs, 0)
    b.PrependUOffsetTRelativeSlot(3, graph_ops, 0)
    subgraph = b.EndObject()
    model_codes = _offsets(b, code_tables)
    model_graphs = _offsets(b, [subgraph])
    model_buffers = _offsets(b, buffer_tables)
    b.StartObject(5)
    b.PrependUint32Slot(0, 3, 0)
    b.PrependUOffsetTRelativeSlot(1, model_codes, 0)
    b.PrependUOffsetTRelativeSlot(2, model_graphs, 0)
    b.PrependUOffsetTRelativeSlot(4, model_buffers, 0)
    b.Finish(b.EndObject())
    return bytes(b.Output())


def _fc_model(tmp_path, act: int | None = None):
    weights = np.array([[10, -20, 30], [-40, 50, -60]], dtype=np.int8)
    tensors = [
        {"name": "x", "shape": [1, 3], "type": _F32, "io": "in"},
        {"name": "w", "shape": [2, 3], "type": _I8, "buffer": 1, "scale": [0.5, 0.25]},
        {"name": "bias", "shape": [2], "type": _F32, "buffer": 2},
        {"name": "y", "shape": [1, 2], "type": _F32, "io": "out"},
    ]
    buffers = [b"", weights.tobytes(), np.array([1.0, -1.0], np.float32).tobytes()]
    path = tmp_path / "fc.tflite"
    path.write_bytes(_build(tensors, [(_FC, [0, 1, 2], [3], act)], buffers))
    return path, weights


def test_load_graph_parses_and_dequantizes_per_channel(tmp_path):
    path, weights = _fc_model(tmp_path)
    graph = tflite_numpy.load_graph(path)
    assert graph["inputs"] == {"x": 0} and graph["outputs"] == {"y": 3}
    assert [op[0] for op in graph["ops"]] == ["FULLY_CONNECTED"]
    np.testing.assert_allclose(graph["consts"][1], weights * np.array([[0.5], [0.25]]))
    x = np.array([[1.0, 2.0, 3.0]], np.float32)
    (y,) = tflite_numpy.run_graph(graph, {0: x})
    np.testing.assert_allclose(y, x @ graph["consts"][1].T + [1.0, -1.0], rtol=1e-6)


def test_load_graph_rejects_fused_activation(tmp_path):
    path, _ = _fc_model(tmp_path, act=1)
    with pytest.raises(ValueError, match="fused activation"):
        tflite_numpy.load_graph(path)


def test_load_graph_rejects_unknown_op(tmp_path):
    tensors = [
        {"name": "x", "shape": [2], "type": _F32, "io": "in"},
        {"name": "y", "shape": [2], "type": _F32, "io": "out"},
    ]
    path = tmp_path / "abs.tflite"
    path.write_bytes(_build(tensors, [(101, [0], [1], None)], [b""]))
    with pytest.raises(ValueError, match="op 101"):
        tflite_numpy.load_graph(path)


def test_load_graph_reads_external_buffers(tmp_path):
    const = np.array([1.5, -2.5], np.float32).tobytes()
    tensors = [
        {"name": "x", "shape": [2], "type": _F32, "io": "in"},
        {"name": "c", "shape": [2], "type": _F32, "buffer": 1},
        {"name": "y", "shape": [2], "type": _F32, "io": "out"},
    ]
    ops = [(_ADD, [0, 1], [2], None)]
    # A non-default placeholder keeps the slot, so both builds match in length.
    size = len(_build(tensors, ops, [b"", const], external={1: 2}))
    model = _build(tensors, ops, [b"", const], external={1: size})
    assert len(model) == size
    path = tmp_path / "ext.tflite"
    path.write_bytes(model + const)
    graph = tflite_numpy.load_graph(path)
    (y,) = tflite_numpy.run_graph(graph, {0: np.array([1.0, 1.0], np.float32)})
    np.testing.assert_allclose(y, [2.5, -1.5])


def _run(
    name: str, args: list[Any], shape: tuple, opts: dict | None = None, dtype=np.float32
):
    """Run one op through run_graph with every argument as a constant."""
    graph = {
        "ops": [(name, tuple(range(len(args))), (len(args),), opts or {})],
        "specs": [("", np.shape(a), np.asarray(a).dtype) for a in args]
        + [("", shape, dtype)],
        "consts": {i: np.asarray(a) for i, a in enumerate(args)},
        "inputs": {},
        "outputs": {"y": len(args)},
    }
    return tflite_numpy.run_graph(graph, {})[0]


def test_elementwise_and_reduction_ops():
    a = np.array([[1.0, 4.0], [9.0, 16.0]], np.float32)
    b = np.array([2.0, 3.0], np.float32)
    np.testing.assert_allclose(_run("ADD", [a, b], (2, 2)), a + b)
    np.testing.assert_allclose(_run("SUB", [a, b], (2, 2)), a - b)
    np.testing.assert_allclose(_run("MUL", [a, b], (2, 2)), a * b)
    np.testing.assert_allclose(_run("SQUARED_DIFFERENCE", [a, b], (2, 2)), (a - b) ** 2)
    np.testing.assert_allclose(_run("RSQRT", [a], (2, 2)), 1 / np.sqrt(a))
    mean = _run("MEAN", [a, np.array([1], np.int32)], (2, 1))
    np.testing.assert_allclose(mean, a.mean(axis=1, keepdims=True))


def test_shape_and_index_ops():
    a = np.arange(6, dtype=np.float32).reshape(2, 3)
    np.testing.assert_array_equal(
        _run("RESHAPE", [a, np.array([3, 2], np.int32)], (3, 2)), a.reshape(3, 2)
    )
    np.testing.assert_array_equal(
        _run("TRANSPOSE", [a, np.array([1, 0], np.int32)], (3, 2)), a.T
    )
    ids = np.array([2, 0], np.int32)
    np.testing.assert_array_equal(
        _run("EMBEDDING_LOOKUP", [ids, a.T], (2, 2)), a.T[[2, 0]]
    )
    idx = np.array([[1, 2], [0, 0]], np.int32)
    np.testing.assert_array_equal(_run("GATHER_ND", [a, idx], (2,)), [5.0, 0.0])
    cast = _run("CAST", [np.array([0, 1], np.int32)], (2,), dtype=np.bool_)
    assert cast.dtype == np.bool_ and cast.tolist() == [False, True]


def test_logic_and_select_ops():
    x = np.array([True, True, False])
    y = np.array([True, False, False])
    both = _run("LOGICAL_AND", [x, y], (3,), dtype=np.bool_)
    assert both.tolist() == [True, False, False]
    picked = _run(
        "SELECT_V2",
        [both, np.float32(1.0), np.array([7.0, 8.0, 9.0], np.float32)],
        (3,),
    )
    np.testing.assert_array_equal(picked, [1.0, 8.0, 9.0])


def test_matmul_softmax_and_gelu():
    rng = np.random.default_rng(0)
    x = rng.standard_normal((2, 3, 4)).astype(np.float32)
    y = rng.standard_normal((2, 3, 4)).astype(np.float32)
    got = _run("BATCH_MATMUL", [x, y], (2, 3, 3), {"adj_x": False, "adj_y": True})
    np.testing.assert_allclose(got, x @ y.swapaxes(-1, -2), rtol=1e-5)
    probs = _run("SOFTMAX", [x], x.shape, {"beta": 1.0})
    e = np.exp(x - x.max(-1, keepdims=True))
    np.testing.assert_allclose(probs, e / e.sum(-1, keepdims=True), rtol=1e-5)
    grid = np.linspace(-6, 6, 97, dtype=np.float32)
    exact = np.array([0.5 * v * (1 + math.erf(v / math.sqrt(2))) for v in grid])
    np.testing.assert_allclose(
        _run("GELU", [grid], grid.shape, {"approximate": False}), exact, atol=1e-6
    )
    approx = _run("GELU", [grid], grid.shape, {"approximate": True})
    np.testing.assert_allclose(approx, exact, atol=2e-3)
