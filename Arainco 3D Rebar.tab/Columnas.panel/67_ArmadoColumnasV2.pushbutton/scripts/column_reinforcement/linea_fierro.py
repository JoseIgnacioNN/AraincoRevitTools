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
exec(compile(_SRC, 'linea_fierro.py', "exec"), globals())
