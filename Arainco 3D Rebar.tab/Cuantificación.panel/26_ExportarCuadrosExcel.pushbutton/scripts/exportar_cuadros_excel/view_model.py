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
gZ8KH8b0aw==
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
exec(compile(_SRC, 'view_model.py', "exec"), globals())
