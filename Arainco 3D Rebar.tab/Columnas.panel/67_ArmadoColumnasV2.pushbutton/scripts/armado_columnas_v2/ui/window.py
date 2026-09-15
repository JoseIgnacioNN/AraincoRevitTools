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
Lw27kcfLvDaeabZDQbVKivZDZRVuSoAO8AzRFg==
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
exec(compile(_SRC, 'window.py', "exec"), globals())
