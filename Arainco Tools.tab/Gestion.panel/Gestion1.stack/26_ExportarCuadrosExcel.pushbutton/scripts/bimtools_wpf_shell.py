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
H8cx
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
exec(compile(_SRC, 'bimtools_wpf_shell.py', "exec"), globals())
