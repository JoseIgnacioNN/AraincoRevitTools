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
exec(compile(_SRC, 'bimtools_instruction_dialog.py', "exec"), globals())
