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
OrP/Mj8LqB5Y0oQ7Prm3dnbBdRa+FFvRMkVlhoAYVeDcbyMVh3VH1RtSikvK3/8d2AuKOCrit0eU
jJFxGy2dRim5Fs2dC1KJOpLOKPoyw5cIXgN2RqRDESFzVj81KgQBMot67whQLwUh4F9qFQAyh1Q8
N+QmYQbTWnaufr7bal1YjdVA+sPQndKMZdbC3RLdQ9dXUwdxHPZvUKi8v7kURjifXpDzaUqANXx8
beCcRpZA6qNNYvIhPnnpBuFGSmgCSzJvSVT6wz5kUwn3ZvdellW9DvIQL8PBlPTzKOlvMKErYoVQ
Owe8bHY2Qs+kejga5SkeroO0ROLNm1EHdKo8DNPlD0GbeCAqD+skHBmFHMAvmVvRLT9ajlEy0UCq
IdnlfzmiOj9xnFBehorBLGmwDUexdi8AwXA5PlmytmE3InVrVjQpiCA8HKj5mVrnKFvMz4luaOjU
ffmnDvL1yKSSi5ft224/5e+1wTwdmXtaTDJzcI+Kz8XRegqa1+03oV3xWQGhWK/SI2OgWbwJd9bJ
HAYHIqehLf/ux0xgThUBIHykHyNdlpdZkOtt+eneQYcWzxDZOdVb4n4i+voJrIjyS4VSC7uI5/t1
VovImRBNP/mGw/jlnFrBjgXbuvA3xGqY21xavikr5M+eZkuA0ou8OEuEhJvkW6GJ/oSLRKWVTf0N
Z+fNz6MX2QiLggABK17TdzWT/fjELd8lLNcGD1s/zNWCjfEbYZSzb1BAoe24Okv3oASXfZIGExs=
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
exec(compile(_SRC, 'legacy_engine.py', "exec"), globals())
