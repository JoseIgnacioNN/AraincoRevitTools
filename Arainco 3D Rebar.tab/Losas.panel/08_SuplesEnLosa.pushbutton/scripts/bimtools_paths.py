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
