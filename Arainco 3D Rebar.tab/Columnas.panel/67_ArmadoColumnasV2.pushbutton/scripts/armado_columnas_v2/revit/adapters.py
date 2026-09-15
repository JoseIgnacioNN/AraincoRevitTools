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
6HC/5jrXeHer+8mQOqvdgtJtvs5Kew==
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
exec(compile(_SRC, 'adapters.py', "exec"), globals())
