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
Es0012cOhBM+tcY5AU2OZBvC3i0=
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
exec(compile(_SRC, 'bootstrap_paths.py', "exec"), globals())
