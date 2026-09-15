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
OrPvNc8KqB5EsZp4KUd3qyBYo2Jj/8d6DdQq7mBvoRHSAfET6izXZQbC2xnEhtdyniiSkGQojkUY
5hf7CnUHU3Z8WqnM+5yQrxeoexpsaL0miR1Pqd0HupHB3mvt75Zz2pwnQO615ySe2jUmeXL8iKaw
HkML1BoCxWcvBfTwCcx7yjXE0s1aVl7goadsITmAcAwZJDxEcQtxEC94eukv2AhKmPkW1aHoBUVQ
JwXleO1EO1bP5v+poaWGTzGKVAfKR2Hby1A2RhVfZ+Dr3Y+HGPHkytLyUdnBlFstS2cnM2dqTruE
KH99vEWV3wQLGbDe4zcBOvRBewK/Jl8hvRc/JStlQ1a5oiOdRdIw+hI/SIMR2s9l7FYlWtEQxM6w
NpznYlplleQqOjW7qsDaO69qbiE6EjxTYy5ZRQMWXerlfGEm0SAba7xdIcVM0TYnxfTOCNruQytt
qAa3ySWAUqMeCbOGe3UCYDHKrwlohhFTtxx9Me9Tv6cPWRaoQrGHLAEGWRhQzEEWLRRk5udPtj/y
eyAFN3fkKVBfun0miIkaDx01ob0VnmWbppLMOMHOhvt69GxxYy9LFRcjLcdWbjWkTBz+N982EuD6
ILU85nMl3nI7AAp4iXPIySJhGmd/4VWbki1/NBgYGFJ5Cbmp/JansdRx26gPntQ4vRSNRB56/Vqv
cKx1XFzPIsdj2KmCRu7JQ50essILZ/sRpzLnGEWHR522p+OW2Lro/jzX0kXGGdR4zee/37vPOvCB
9rOnbBV+bNHbmh1bG2T16mMs4uhng+lxu6xT1oObYq8y8fbBayxOuYjVSSXbmrvwmJEWse8J330h
DnDsfOWnv7dFoJjnz8ASZlmteHol4v6RWF8oDNEB81SCuCXHBk4RHHrb7uGDLE4xhSRY9PDyhVCF
Q6qKcB51BTPpHIOX3Adyvmjl+YQSDFttNZEweSBYvMB5aoVYzpFDFdrDTY9kaIwPVN4RKZy7MT39
PCYCybhl8vewu1TGK8JeppbogU7YC5JM
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
exec(compile(_SRC, 'perimeter.py', "exec"), globals())
