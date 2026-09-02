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
HmxfFtO+RsEH7s0mx9WNZ6nKXkc=
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
exec(compile(_SRC, 'services.py', "exec"), globals())
