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
t7KVH6g2nvrjAYnQiUEOnWOgGUgc
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
exec(compile(_SRC, 'armado_columnas_instruction_dialog.py', "exec"), globals())
