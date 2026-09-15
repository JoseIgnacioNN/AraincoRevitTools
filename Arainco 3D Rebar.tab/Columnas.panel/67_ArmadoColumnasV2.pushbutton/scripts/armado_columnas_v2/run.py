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
OrPfN8MKqB5Espp5quq6wkeOpjB+vRPDT+SsQnSKXXAJDuFsN25XC5OSwgJzoxFSNrxquOjpRzll
5dUYGcGoLLfnJtzNRurk7bOjJEuHJfC5ZFullzcprhMyG0/sv1tZj2N0s0BWpCufejXlhMC+//v5
Zvzp18Y1ZNUdMrdkdZTGft+z5PDpZd/EXd9ymVrlFnZcyKddImDMTEKkKm27HJ9hzwOwKBU/lejA
JZ4cyk4vC+67E/p1l0WJyzDdVKle4JDseBTJff/zS/EPCmWusBp/FjcGP4LHVm+JaQjQdg0kOSDT
gks/iGNE9l2JYYqLrGPH9pygEv+sRSUgRFXi02L1pI13BVMEYzN1p0pWtWuAEvzxLSQdo0fFtuVJ
QX9HGs3cl8j2tD/B+rbCBDvny135eIwWbtMw9w8WersBo4i48eiOTh+bC6zuz9JcxlN2hJAl3Jfe
MXod961qyqxGZjq/+D1R7CVVnEO8aFv+eEHdXDVlZ0ARXpzlhL4hB++edr6en+RAar1h7V1uCDOw
WOxp0GaXjxUD3ycBeEiD2WJBrxr1LxevA7qBAdMDFz9F+pQFiKxCY3Gl8LMbbqu0LORAuXmmQesz
w9a1GYiZ0x88RUXt/nH9YhseI3TVx5eNm8OnxAH3WzTeBNuNjTMf6t4kPS66q+c3kCZ4bMYyanQo
DDvjWavPAlGg908kWxHfUsux5ZXgfTd9PNsakxx8o4sQn9b6h88KatIM3TEouDriErZRj4T3iWgG
zkF++szATgSjf+e584YSkAEemV1DJS7jLoVlN4hVwevRu+vMwW5boMpqNRlMiTVAdIAPUcx4KrsT
gD7VMTrBrKV2GGRZ69hg5NIUo5Egaa9r7vvgRPeSJx+A7007tgi0dPJTsfJL2vu4DWKICvlZBbW5
V7pDyPryj/cczbNSW26IXNa7fBkUySfzn5AybLlELR5hN6ukcvEoTUGXdZMCm1RnqzUe3EL5H/Qx
ZvyTx3kFpSQIGsKQF3TGRMV6qmh4c8msi64yUU+KVjcBE95+r/6Cl6BlL36pSbpcTRmPbE6GWKgJ
/RodIoONzadvIu+dHwVnzCcgG4v0zfyLtehv12bUglKKZD7mAnQkI49Yaxxh0udIteqQ0TOEFsg9
5Naf3dIy2ZvUKOmDCLAzceKEsmdHQUrlkGRPOTU/
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
exec(compile(_SRC, 'run.py', "exec"), globals())
