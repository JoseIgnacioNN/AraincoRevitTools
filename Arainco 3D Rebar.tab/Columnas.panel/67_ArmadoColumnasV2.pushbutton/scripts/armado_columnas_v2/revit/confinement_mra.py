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
exec(compile(_SRC, 'confinement_mra.py', "exec"), globals())
