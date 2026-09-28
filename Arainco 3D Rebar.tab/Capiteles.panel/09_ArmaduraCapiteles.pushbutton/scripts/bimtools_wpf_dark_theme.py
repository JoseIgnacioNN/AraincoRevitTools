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
exec(compile(_SRC, 'bimtools_wpf_dark_theme.py', "exec"), globals())
