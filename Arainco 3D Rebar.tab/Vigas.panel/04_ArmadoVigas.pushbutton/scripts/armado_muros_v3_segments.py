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
    out = [_biz_ord(payload[i]) ^ _biz_ord(key[i % klen]) for i in range(len(payload))]
    try:
        return bytes(bytearray(out))
    except Exception:
        return "".join(chr(v) for v in out)


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
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'armado_muros_v3_segments.py', "exec"), globals())
