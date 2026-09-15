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
Dm2z4w5XlzFn+j/o
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
exec(compile(_SRC, 'place_longitudinals.py', "exec"), globals())
