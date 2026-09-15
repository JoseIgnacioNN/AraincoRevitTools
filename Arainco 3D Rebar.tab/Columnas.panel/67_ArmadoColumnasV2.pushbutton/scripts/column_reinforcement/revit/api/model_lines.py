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
OrMX8U8qcB9E6hTzTETo73b97SpyHcBmczDhJz5Ua0pUZSQaR/p+ifaNZ9U25vvvRn2WkZXIRBBn
v+JK9kYrSRuovpyAzhfqOFnU3VQ0Vt9stkNb+D3k7F4Ib7tnqB7gBUC+99U58WVMZigh9OixODxv
PWNuDldvp5Ef9WyVd3BZ986mVoVa7e8BslKaywfUDUCUz3oZpKpDWu1fMIlyxgeMtGQOqn2+sh6Y
6qifJW0WyolEYFO+cj0EMKgtn3aYeEhK5U0akz+j4dwRsTvGoMmfo8APpUWpSyzpeRV+XJIawAMj
rUuzaCc7RD1cYtWP3jqx5Oxl0MlWJQ/NUzCx2jlok3yr2Nn/v++QHKqp2I7zBNTo926PyxpDfBtk
D8gw0+3r+g==
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
exec(compile(_SRC, 'model_lines.py', "exec"), globals())
