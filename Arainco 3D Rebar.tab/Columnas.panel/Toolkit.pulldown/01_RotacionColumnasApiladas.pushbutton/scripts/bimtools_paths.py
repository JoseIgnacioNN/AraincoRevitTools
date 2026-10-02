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
exec(compile(_SRC, 'bimtools_paths.py', "exec"), globals())
