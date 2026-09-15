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
exec(compile(_SRC, 'draw_elev_scheme.py', "exec"), globals())
