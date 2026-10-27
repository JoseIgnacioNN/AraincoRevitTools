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
oAX0p4pztny5IQ==
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
