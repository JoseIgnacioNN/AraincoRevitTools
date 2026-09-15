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
OrPPNakKqB5YErg7Pok53VHwH7an5dJg02/56Hvtpt8XPBcbJwM4sHBoghvDQ4gMxzqK9aJcxWKy
90ZLYVc+y7pwtv69RaCXALD4vJYyjLU/9itQP5DRFQlYvbWuqGjVoUZ/PgNHLuwM16yFhYPBeLR2
TrUxP8mo5m6+lps1XrTH5VQJDyyJvHpk1ymt6SIfT4JcktU0sxgo+8U+bFwvoRkMKKGYr+k2/aMP
YJU73vzh4ssrWkS1+64Z3jPQVxBrB98pHAHj6rV5dr+ln1C8SlXNMwjXWT0jtYGQ6jg2Vc2z5/6N
AorObacJ9PGF9sguJz+OqGv/7yS+4ClcuPttyZjLxxu8WAcBqj/M4X9w55zdvJezI11KCEKFV6y0
77FBYZYSb830/WVxpD06xue6dFsrCGmER3zXG/hjzEV2Q3+lAQVFqXEMVlPdzbkVuESkvHD2rPMJ
nQkUcPXJh70/uNA+L/20IucjhoQwfZYiwNvH28xFH9pmrr0pgxrCyStHePErMBiF1UOt/O5fAqaa
rx7v8j8drzR4va2WKPWxXuBdX6CzipMF1SaccLRgO1GV7+OZdhn1OGTZKsvILqULYXhjiLrSe/Ve
7Pxxez+I+wBEklck9AsxFKZBvq7tysp4VMSPDEJL5dXfU4dG0ms0ro1/Pxhmq3hbenCVaRWNpSo=
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
exec(compile(_SRC, 'adapters.py', "exec"), globals())
