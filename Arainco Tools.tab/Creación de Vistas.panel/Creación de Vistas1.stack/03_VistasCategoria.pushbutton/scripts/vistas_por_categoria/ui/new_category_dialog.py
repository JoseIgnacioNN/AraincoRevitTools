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
exec(compile(_SRC, 'new_category_dialog.py', "exec"), globals())
