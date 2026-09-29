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
ZVdcJNHdbjRnbpdgSwstiQ==
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
exec(compile(_SRC, 'stirrup_draw_135.py', "exec"), globals())
