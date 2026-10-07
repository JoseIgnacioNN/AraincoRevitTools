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
exec(compile(_SRC, 'armado_muros_run.py', "exec"), globals())
