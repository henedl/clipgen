"""A numpy evaluator for small static TFLite graphs, so Redact needs no LiteRT runtime.

``load_graph`` reads the ``.tflite`` flatbuffer with ``struct`` (field slots
from the LiteRT ``schema.fbs``), dequantizes int8 constants once, and keeps
only the op subset in ``_OPS``; anything else raises ``ValueError`` so a model
bump fails loudly instead of computing garbage. ``run_graph`` walks the
operators in file order and reshapes every result to the tensor's declared
shape, which covers ``keep_dims``/``keep_num_dims`` without reading them.

Hybrid (dynamic-range) FULLY_CONNECTED runs in float32 against the
dequantized weights. LiteRT quantizes the activations per token instead, so
logits drift by ~0.2 at most; argmax tags agreed on every Redact window tried.
"""

from __future__ import annotations

import struct
from pathlib import Path
from typing import Any

import numpy as np

_OPS = {
    0: "ADD",
    7: "EMBEDDING_LOOKUP",
    9: "FULLY_CONNECTED",
    18: "MUL",
    22: "RESHAPE",
    25: "SOFTMAX",
    39: "TRANSPOSE",
    40: "MEAN",
    41: "SUB",
    53: "CAST",
    76: "RSQRT",
    86: "LOGICAL_AND",
    99: "SQUARED_DIFFERENCE",
    107: "GATHER_ND",
    123: "SELECT_V2",
    126: "BATCH_MATMUL",
    150: "GELU",
}
_DTYPES = {0: np.float32, 2: np.int32, 4: np.int64, 6: np.bool_, 9: np.int8}
_FUSED_ACTIVATION = frozenset({"ADD", "FULLY_CONNECTED", "MUL", "SUB"})


# ---------------------------------------------------------------------------
# Flatbuffer reading
# ---------------------------------------------------------------------------


def _u32(buf: bytes, pos: int) -> int:
    return struct.unpack_from("<I", buf, pos)[0]


def _field(buf: bytes, table: int, slot: int) -> int | None:
    """Absolute position of a table field, or None when absent."""
    vtable = table - struct.unpack_from("<i", buf, table)[0]
    entry = 4 + 2 * slot
    if entry >= struct.unpack_from("<H", buf, vtable)[0]:
        return None
    offset = struct.unpack_from("<H", buf, vtable + entry)[0]
    return table + offset if offset else None


def _scalar(buf: bytes, table: int, slot: int, fmt: str, default: Any) -> Any:
    pos = _field(buf, table, slot)
    return default if pos is None else struct.unpack_from(fmt, buf, pos)[0]


def _ref(buf: bytes, table: int, slot: int) -> int | None:
    pos = _field(buf, table, slot)
    return None if pos is None else pos + _u32(buf, pos)


def _vector(buf: bytes, table: int, slot: int, dtype: Any) -> np.ndarray:
    start = _ref(buf, table, slot)
    if start is None:
        return np.zeros(0, dtype=dtype)
    return np.frombuffer(buf, dtype=dtype, count=_u32(buf, start), offset=start + 4)


def _tables(buf: bytes, table: int, slot: int) -> list[int]:
    start = _ref(buf, table, slot)
    if start is None:
        return []
    items = range(start + 4, start + 4 + 4 * _u32(buf, start), 4)
    return [pos + _u32(buf, pos) for pos in items]


def _string(buf: bytes, table: int, slot: int) -> str:
    start = _ref(buf, table, slot)
    if start is None:
        return ""
    return buf[start + 4 : start + 4 + _u32(buf, start)].decode("utf-8", "replace")


def _options(buf: bytes, op: int, name: str) -> dict[str, Any]:
    """The builtin options the evaluator honours; unsupported values raise."""
    table = _ref(buf, op, 4)
    if table is None:
        return {"beta": 1.0} if name == "SOFTMAX" else {}
    if name in _FUSED_ACTIVATION and _scalar(buf, table, 0, "<b", 0):
        raise ValueError(f"{name} with a fused activation is unsupported")
    if name == "SOFTMAX":
        return {"beta": _scalar(buf, table, 0, "<f", 0.0)}
    if name == "GELU":
        return {"approximate": bool(_scalar(buf, table, 0, "<B", 0))}
    if name == "BATCH_MATMUL":
        return {
            "adj_x": bool(_scalar(buf, table, 0, "<B", 0)),
            "adj_y": bool(_scalar(buf, table, 1, "<B", 0)),
        }
    return {}


def _constant(
    buf: bytes, tensor: int, data: np.ndarray, dtype: Any, shape: tuple
) -> np.ndarray:
    """A constant tensor's value, int8 dequantized to float32."""
    array = np.frombuffer(data, dtype=dtype).reshape(shape)
    if dtype is not np.int8:
        return array
    quant = _ref(buf, tensor, 4)
    if quant is None or not _vector(buf, quant, 2, "<f4").size:
        raise ValueError("int8 constant without a scale")
    scale = _vector(buf, quant, 2, "<f4")
    zero = _vector(buf, quant, 3, "<i8")
    bcast = [1] * array.ndim
    if scale.size > 1:
        bcast[_scalar(buf, quant, 6, "<i", 0)] = scale.size
    values = array.astype(np.float32)
    if zero.any():
        values -= zero.astype(np.float32).reshape(bcast if zero.size > 1 else -1)
    return values * scale.reshape(bcast)


def load_graph(path: str | Path) -> dict[str, Any]:
    """Parse subgraph 0 into ops, tensor specs, constants and named inputs/outputs."""
    buf = Path(path).read_bytes()
    root = _u32(buf, 0)
    codes = [
        max(_scalar(buf, oc, 0, "<b", 0), _scalar(buf, oc, 3, "<i", 0))
        for oc in _tables(buf, root, 1)
    ]
    buffers = []
    for entry in _tables(buf, root, 4):
        # Offsets past 1 point at data appended after the flatbuffer.
        offset = _scalar(buf, entry, 1, "<Q", 0)
        if offset > 1:
            size = _scalar(buf, entry, 2, "<Q", 0)
            buffers.append(
                np.frombuffer(buf, dtype=np.uint8, count=size, offset=offset)
            )
        else:
            buffers.append(_vector(buf, entry, 0, np.uint8))
    subgraph = _tables(buf, root, 2)[0]

    specs: list[tuple[str, tuple[int, ...], Any]] = []
    consts: dict[int, np.ndarray] = {}
    for index, tensor in enumerate(_tables(buf, subgraph, 0)):
        type_id = _scalar(buf, tensor, 1, "<b", 0)
        if type_id not in _DTYPES:
            raise ValueError(f"tensor type {type_id} is unsupported")
        dtype = _DTYPES[type_id]
        shape = tuple(int(d) for d in _vector(buf, tensor, 0, "<i4"))
        specs.append((_string(buf, tensor, 3), shape, dtype))
        data = buffers[_scalar(buf, tensor, 2, "<I", 0)]
        if data.size:
            consts[index] = _constant(buf, tensor, data, dtype, shape)

    ops = []
    for op in _tables(buf, subgraph, 3):
        code = codes[_scalar(buf, op, 0, "<I", 0)]
        if code not in _OPS:
            raise ValueError(f"TFLite op {code} is unsupported")
        name = _OPS[code]
        ins = tuple(int(i) for i in _vector(buf, op, 1, "<i4"))
        outs = tuple(int(i) for i in _vector(buf, op, 2, "<i4"))
        ops.append((name, ins, outs, _options(buf, op, name)))

    def named(slot: int) -> dict[str, int]:
        return {specs[int(i)][0]: int(i) for i in _vector(buf, subgraph, slot, "<i4")}

    return {
        "ops": ops,
        "specs": specs,
        "consts": consts,
        "inputs": named(1),
        "outputs": named(2),
    }


# ---------------------------------------------------------------------------
# Evaluation
# ---------------------------------------------------------------------------


def _erf(x: np.ndarray) -> np.ndarray:
    """Abramowitz-Stegun 7.1.26; absolute error below 1.5e-7."""
    sign = np.sign(x)
    x = np.abs(x)
    t = 1.0 / (1.0 + 0.3275911 * x)
    poly = (
        (((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t
        + 0.254829592
    ) * t
    return sign * (1.0 - poly * np.exp(-x * x))


def _gelu(x: np.ndarray, approximate: bool) -> np.ndarray:
    if approximate:
        return 0.5 * x * (1.0 + np.tanh(0.7978845608 * (x + 0.044715 * x * x * x)))
    return 0.5 * x * (1.0 + _erf(x * np.float32(0.7071067811865476)))


def _softmax(x: np.ndarray, beta: float) -> np.ndarray:
    z = (x - x.max(axis=-1, keepdims=True)) * np.float32(beta)
    np.exp(z, out=z)
    z /= z.sum(axis=-1, keepdims=True)
    return z


def _apply(name: str, a: list[Any], opts: dict[str, Any]) -> np.ndarray:
    if name == "ADD":
        return a[0] + a[1]
    if name == "SUB":
        return a[0] - a[1]
    if name == "MUL":
        return a[0] * a[1]
    if name == "SQUARED_DIFFERENCE":
        return np.square(a[0] - a[1])
    if name == "RSQRT":
        return 1.0 / np.sqrt(a[0])
    if name == "MEAN":
        return a[0].mean(axis=tuple(int(x) for x in np.atleast_1d(a[1])))
    if name in ("RESHAPE", "CAST"):
        return a[0]
    if name == "TRANSPOSE":
        return a[0].transpose(a[1])
    if name == "EMBEDDING_LOOKUP":
        return a[1][a[0]]
    if name == "FULLY_CONNECTED":
        out = a[0] @ a[1].T
        return out if a[2] is None else out + a[2]
    if name == "BATCH_MATMUL":
        x = a[0].swapaxes(-1, -2) if opts["adj_x"] else a[0]
        y = a[1].swapaxes(-1, -2) if opts["adj_y"] else a[1]
        return x @ y
    if name == "SOFTMAX":
        return _softmax(a[0], opts["beta"])
    if name == "GELU":
        return _gelu(a[0], opts["approximate"])
    if name == "GATHER_ND":
        return a[0][tuple(np.moveaxis(a[1], -1, 0))]
    if name == "LOGICAL_AND":
        return np.logical_and(a[0], a[1])
    if name == "SELECT_V2":
        return np.where(a[0], a[1], a[2])
    raise ValueError(f"TFLite op {name} is unsupported")


def run_graph(graph: dict[str, Any], feeds: dict[int, np.ndarray]) -> list[np.ndarray]:
    """Evaluate the graph; ``feeds`` maps input tensor index to value."""
    values: dict[int, Any] = dict(graph["consts"])
    values.update(feeds)
    specs = graph["specs"]
    for name, ins, outs, opts in graph["ops"]:
        args = [values[i] if i >= 0 else None for i in ins]
        _, shape, dtype = specs[outs[0]]
        result = np.asarray(_apply(name, args, opts))
        values[outs[0]] = result.astype(dtype, copy=False).reshape(shape)
    return [values[i] for i in graph["outputs"].values()]
