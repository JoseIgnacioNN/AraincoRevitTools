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
