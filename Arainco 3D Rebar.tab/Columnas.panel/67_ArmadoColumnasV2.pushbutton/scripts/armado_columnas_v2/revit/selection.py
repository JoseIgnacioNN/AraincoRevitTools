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
YcFTyQYbZsnUAhiG1BBuL3itKA==
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
exec(compile(_SRC, 'selection.py', "exec"), globals())
