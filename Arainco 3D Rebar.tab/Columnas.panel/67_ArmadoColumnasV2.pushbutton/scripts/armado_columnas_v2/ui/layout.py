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
cJ8gVsDuNUoN8A42l3EEUuRJ
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
exec(compile(_SRC, 'layout.py', "exec"), globals())
