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
