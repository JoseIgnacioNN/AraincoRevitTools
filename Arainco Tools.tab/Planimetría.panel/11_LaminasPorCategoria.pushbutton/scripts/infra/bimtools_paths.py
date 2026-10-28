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
