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
exec(compile(_SRC, 'lap_detail_link_vigas_schema.py', "exec"), globals())
