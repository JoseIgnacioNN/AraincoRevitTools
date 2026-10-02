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
TZ4SSA7nafujm3ZHRdlz
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
exec(compile(_SRC, 'conf_draw.py', "exec"), globals())
