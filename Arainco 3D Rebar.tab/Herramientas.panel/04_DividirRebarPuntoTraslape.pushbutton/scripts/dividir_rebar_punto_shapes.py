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
exec(compile(_SRC, 'dividir_rebar_punto_shapes.py', "exec"), globals())
