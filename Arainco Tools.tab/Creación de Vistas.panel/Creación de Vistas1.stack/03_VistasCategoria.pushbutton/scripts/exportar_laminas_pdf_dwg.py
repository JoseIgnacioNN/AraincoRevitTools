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
+YcGJQ==
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
exec(compile(_SRC, 'exportar_laminas_pdf_dwg.py', "exec"), globals())
