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
exec(compile(_SRC, 'arearein_exterior_h_l135_rps.py', "exec"), globals())
