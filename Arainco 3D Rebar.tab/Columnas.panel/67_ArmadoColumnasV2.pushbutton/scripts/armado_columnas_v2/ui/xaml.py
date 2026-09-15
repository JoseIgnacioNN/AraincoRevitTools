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
exec(compile(_SRC, 'xaml.py', "exec"), globals())
