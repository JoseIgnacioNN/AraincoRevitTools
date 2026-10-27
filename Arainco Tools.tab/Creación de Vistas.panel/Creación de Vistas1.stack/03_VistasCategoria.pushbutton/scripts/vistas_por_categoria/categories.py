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
exec(compile(_SRC, 'categories.py', "exec"), globals())
