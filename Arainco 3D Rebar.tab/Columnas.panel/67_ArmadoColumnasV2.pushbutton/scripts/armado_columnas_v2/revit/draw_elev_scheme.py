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
Bjo=
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
