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
Wg==
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
exec(compile(_SRC, 'embed_anchorage_link_schema.py', "exec"), globals())
