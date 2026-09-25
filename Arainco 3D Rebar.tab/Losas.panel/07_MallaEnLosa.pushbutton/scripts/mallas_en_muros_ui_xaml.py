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
L/JSuis=
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
exec(compile(_SRC, 'mallas_en_muros_ui_xaml.py', "exec"), globals())
