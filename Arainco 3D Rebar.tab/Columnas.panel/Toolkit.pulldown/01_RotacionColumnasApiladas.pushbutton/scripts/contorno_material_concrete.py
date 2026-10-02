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
exec(compile(_SRC, 'contorno_material_concrete.py', "exec"), globals())
