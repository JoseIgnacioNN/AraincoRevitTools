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
exec(compile(_SRC, 'colocar.py', "exec"), globals())
