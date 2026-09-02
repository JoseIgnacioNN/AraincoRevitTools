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
RsEgZUY7hdNIyDGpuA==
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
exec(compile(_SRC, 'exportar_cuadros_excel_ui.py', "exec"), globals())
