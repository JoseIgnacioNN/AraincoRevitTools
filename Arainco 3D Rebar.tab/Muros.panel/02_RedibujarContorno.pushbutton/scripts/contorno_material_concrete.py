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
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
Pb1bjiy11+cWKhHeeUBnRrA=
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
exec(compile(_SRC, 'contorno_material_concrete.py', "exec"), globals())
