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
exec(compile(_SRC, 'armado_muros_vecinos_extremos.py', "exec"), globals())
