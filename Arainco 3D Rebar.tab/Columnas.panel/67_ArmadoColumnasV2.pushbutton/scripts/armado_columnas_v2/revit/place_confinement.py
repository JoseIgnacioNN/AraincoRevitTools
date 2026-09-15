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
exec(compile(_SRC, 'place_confinement.py', "exec"), globals())
