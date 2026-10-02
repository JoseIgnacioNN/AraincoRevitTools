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
XK1B2hZLaSJsnX1wp7J1
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
exec(compile(_SRC, 'colocar.py', "exec"), globals())
