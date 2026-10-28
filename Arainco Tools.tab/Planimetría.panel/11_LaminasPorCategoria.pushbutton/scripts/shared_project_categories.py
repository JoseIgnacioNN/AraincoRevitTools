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
exec(compile(_SRC, 'shared_project_categories.py', "exec"), globals())
