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
exec(compile(_SRC, 'confinement.py', "exec"), globals())
