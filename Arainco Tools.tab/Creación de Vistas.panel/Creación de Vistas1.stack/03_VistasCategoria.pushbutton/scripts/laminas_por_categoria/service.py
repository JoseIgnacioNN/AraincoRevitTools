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
MOWgnX9tjUkpHwtpu5Inxkd7m1Lz0XevBGHz
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
exec(compile(_SRC, 'service.py', "exec"), globals())
