# -*- coding: utf-8 -*-
# === BIZARDS_OBFUSCATED_MODULE ===
# Modulo de produccion ofuscado (no es codigo fuente legible).
# Generado por prod_builder — no editar.
# Decoder portable: CPython 3 + IronPython/pyRevit (str/bytes indexing).
from __future__ import print_function
import base64 as _b64
import zlib as _zlib


def _biz_ord(x):
    # int (Py3 bytes) o char (Py2/IronPython str)
    return x if isinstance(x, int) else ord(x)


def _biz_xor_decode(payload, key):
    klen = len(key)
    n = len(payload)
    out = bytearray(n)
    for i in range(n):
        out[i] = _biz_ord(payload[i]) ^ _biz_ord(key[i % klen])
    return out


def _biz_to_unicode_source(raw):
    # CPython3: bytes. CPython2: str-bytes. IronPython: str==unicode y zlib
    # mapea cada byte a un codepoint (p. ej. C3 A1 se ve como mojibake).
    if not isinstance(raw, type(u"")):
        return raw.decode("utf-8")
    try:
        return raw.encode("latin-1").decode("utf-8")
    except Exception:
        return raw


_K = _b64.b64decode("Qml6YXJkcy5Ub29sLlByb2QuT2JmdXNjYXRpb24udjE=")
_P = _b64.b64decode(
"""
OrOvNLMKpx5E0ohHaLHgL/+Ej2MjYdpJ9DS5eXtUNBl5IMY8xT2o9prFivMWoXuSVegGVO8CyOdi
Yxu3whNF/rG39lI9BIzHanvdiMxq97Wq2RDb+1V/oKquSw9rZS/tlGOkHEufewv/zy/dbFPj/b5M
EVMec5+tUxWRrEE/USVbbQaI3kQ+PjAIeEsx/kwKg+koxRZxUolE4ls0n0VCaMn8zyZWb8iCfFdH
tTK2LvvaFxwKceTZrkPbxPpMBc8PeP74pBtkh5bXspRcqImemoTZqcLrQnRwJrsPIpJdYU4QAlmU
KdI5preu3meCZVRlYYU13r95Ad/wJEgHEtaPsqZwM7Dy2fvPJ/IcIQgTpE5Pdu9DG+xbx9FuXP/c
hv00HSD1bNc9gTx5XEuy+T+/cfutYCFN9Zz9wSYl+z9BBL524ClQy/UvQhzWaA3pe67sFf+tfMeI
FJacBLczpQv6RA8bU8xTIh6TWBvCvPqDZvVPRZMTJ3pCuUX+8hF1d3j40mFZ1qbHOmucW4otVju0
i0dehsCrgHtA70n5ZkeZku3w50m4CbBpVDl5gRzc6x34M3HzJauW8Q0ajK20e5VdqfAZoN3KcPDs
XuJmSjXhEylh3lLkoFQQZT2kpH65AkxNW86FP49S+tEblLJOVsxJFV8B4W3AJUCSueEBDhJJIeiM
ltLyQ7k5E4NjAO9DkS+6/875KWoW8dVtxzgU5oXLQ/d2gsSRH17GssXiGo7/VnhRxXgz6+17yva5
HTcBThjrK3gBfn4IPVutI89XuVS5XSbLTs6XzOqQ603CYMq324TiGviwuKnbQP4a/KArr7dXDwk/
DAKg/oHsbfXD6ar1az1SXQYUeP9LcpQvDbKgGlJJHbLytGu0OZJL2lT/IPrW/stsSSGGxoQN2g3M
M4s4eSCWaH1vfCT0jinMjxXjgkye+pXz7vRFICpsylK+3w==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
try:
    _SRC = _zlib.decompress(_P)
except Exception:
    try:
        _SRC = _zlib.decompress(bytes(_P))
    except Exception:
        _SRC = _zlib.decompress("".join(chr(b) for b in _P))
_SRC = _biz_to_unicode_source(_SRC)
exec(compile(_SRC, 'canvas_paths.py', "exec"), globals())
