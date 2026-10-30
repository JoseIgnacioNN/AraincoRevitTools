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
exec(compile(_SRC, 'bootstrap_paths.py', "exec"), globals())
