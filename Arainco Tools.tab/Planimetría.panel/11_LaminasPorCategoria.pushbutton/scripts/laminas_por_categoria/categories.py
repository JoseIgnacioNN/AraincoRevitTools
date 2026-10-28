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
