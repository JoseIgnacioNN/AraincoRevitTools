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
WiKpKco=
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
exec(compile(_SRC, 'arearein_interior_h_l135_rps.py', "exec"), globals())
