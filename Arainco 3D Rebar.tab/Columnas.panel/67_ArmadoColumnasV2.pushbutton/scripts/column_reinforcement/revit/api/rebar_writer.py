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
OrMP8DMqsB5EKphS6/geNfCoYsiid8Y60d15+g7HzsfmXEIXRvonXnrhr6meg9hKzt4hHJiNTkhW
RNby7SCwy59rAA+Uk0o/x6aLvYbe469ETpsG31Lm5GwnxpyOznOGyCNbV64LUPIVKLdFfIqME58r
C4nnxTIhztnjL9NJrRPq3dd9tHYt/aEa0saw+4fFPgHQyCus6dlPRzdnz7P8wEoQMR3WDauQ94m2
Lm5q29tosAYTrivpkym6KjXPqroHwvbT8jqM0yNIZ3OCulnCTAG8B+n4xG/rmCLqYiZHy1Rrzms8
IeR6lEzBCPgj1SSF+iBXLXyMlz34Hpn/pc2OS3Pc1KJ5mYm1oC6BL9bYVc4ZphOGUT+NiVHNZMJs
UtZnexA6nLXpQ9N4kPCn2oyJlOpM22zeJAEl7qb6qUaVXmyw8IJkJJ632g==
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
exec(compile(_SRC, 'rebar_writer.py', "exec"), globals())
