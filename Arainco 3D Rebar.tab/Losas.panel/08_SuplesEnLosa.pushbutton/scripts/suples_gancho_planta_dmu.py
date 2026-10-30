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
exec(compile(_SRC, 'suples_gancho_planta_dmu.py', "exec"), globals())
