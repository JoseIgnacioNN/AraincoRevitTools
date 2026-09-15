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
exec(compile(_SRC, 'conf_draw.py', "exec"), globals())
