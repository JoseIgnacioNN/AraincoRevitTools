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
exec(compile(_SRC, 'lap_detail_link_wall_foundation_schema.py', "exec"), globals())
