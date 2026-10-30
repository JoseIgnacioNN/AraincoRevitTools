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
OrPPMjMKqB5E04R5JaVQVtIPaE82bSJM5Cr5zE6mwoVwATLFWDb+9XS4UcwRuHGV7tLrTY/n4Q5I
OTiLPnCSbIyHto2gCxObDC7OJg+nKYcpmzUNVeucmBbMBj90i9d1mpjUgqLOfWLzZTt8GhJFc1AY
AF3CZHbkLklrDNVwolQcmWcflxHdqUr13hxjfU/AjnpMGziA1YxRuyhbuv026ksAUkp5R+faaMCr
ZqgGthMKPYee+KLGPWNCRMB14vMdB+hAvQgktHyMgCgyXCiIhorm016n9CCcrn5/ZE9rV8kgaMdF
NdgiRZs6SWushvjjrNZSAQFXPvoVx+304h+OxK/SeVIEc7ZHwDZzkfCU9teqbpIpihVABc6XnQmX
2OKMIwVH8hjeXK0iFEuo/qARf0mvqRMUfkeWSXh0ZeErwDdQntswvw9QOd3l7TekfuVytpJqxr8V
4cP7sucKDkcyOVc4coP2hKpOeRj14crmXHpzXi3LRD5Wg3cQHwk5ioZ+sO6kdtfbdlyDlLm1p9QA
2rQIp9q6FlOgsXa100Is+YekGJIvXSUoOzdLpcXF1ARue4jxTxq4qHUEbfX8mBK6EzkvC7UAnbHz
i4r9d+Ub0DBMVSkAaAxPJIAxWVbUpBfSUB+qZkmM8ocSdksNViE1+WteSnl8iwPU/IH2BcgfOwVl
gFvSwDmGR++BFrk0D6UrKoSItyk+la6J2wHrgt/J3iVzhBpdqrrFAlzRBYWPiWI5hyQP
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
exec(compile(_SRC, 'armado_muros_v3_segments.py', "exec"), globals())
