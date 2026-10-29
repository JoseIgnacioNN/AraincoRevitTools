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
0gCDsT5PlmlkW30h
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
exec(compile(_SRC, 'bimtools_instruction_dialog.py', "exec"), globals())
