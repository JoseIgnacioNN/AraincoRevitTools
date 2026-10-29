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
exec(compile(_SRC, 'bimtools_rebar_hook_lengths.py', "exec"), globals())
