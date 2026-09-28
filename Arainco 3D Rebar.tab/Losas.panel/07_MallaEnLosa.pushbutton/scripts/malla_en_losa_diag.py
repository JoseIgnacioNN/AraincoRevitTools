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
ZWOrZr8=
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
exec(compile(_SRC, 'malla_en_losa_diag.py', "exec"), globals())
