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
TlWe0DEJdPDDRwrVBZuYZPbgm5k=
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
exec(compile(_SRC, 'new_category_dialog.py', "exec"), globals())
