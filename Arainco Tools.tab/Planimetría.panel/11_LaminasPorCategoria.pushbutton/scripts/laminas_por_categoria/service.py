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
exec(compile(_SRC, 'service.py', "exec"), globals())
