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
exec(compile(_SRC, 'bimtools_rebar_3d_visibility.py', "exec"), globals())
