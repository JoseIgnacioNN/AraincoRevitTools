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
OrPPNLMKqB5Y0pRHJr9gYO1ZUmiK5UYOMBiC7dp2x6KDsw5LvvJPZBr0kqEQME3fD1Ej9oNRPfTD
bLlmXiKnTaaKWUTN5rrFXYvJPFQpdsJPCSDfQ/+QW1tCs97HjlZhmFOu2DAci4MB4bHaDH1E64/X
Zs54C1Iv+hIPBToU+5NxDa4vjySL8EY7Z9vq3MedfHafKD8QRzqxcvP5ZCrKUIf/2lEkRpRKd8SF
2lxfTVoFA394pW9DFAvxBvfqYL3ZQQYSsc/WvdXVOg9RS++x208LZbuy5RYGii7PQ217QbdfXVYP
ELFh+bxHhaFXleNTjZpndDJn4K0IYKZKwyGfbA3MYttBmF5gWhcxR3PISZR6A4aV9okBTpWeJI3D
eLW+K+ahWfeJ1krwUec55YYZJ33ZVcTfnAmA19yAE/DjSE7Y9yUFbSrNaAU7P2Wghyt/K57SNrSF
IhF19UXA/lHpNosT7xjaiqlZHH4ULRWPd+b29MliGCgCuChrk2/D8PdyfgAlSFwFIGbIYp14BYsz
CnxAPp9+QwlXXAo3B10BP4sV6O7YEQCT8zRJ4De4vsoRE2gh5GgyiAFEL8VFa4BabYr9lEOIoS7P
1ZuZr7OC4KPAeYcpFGsO6Hfpiv53bWqdM2ueakenVzAsX6I/6n99fXtZXvCB/J1qtQe5qe/endlI
dlsdHNu7Xok7e+Mtbqto2SKORS+D+cXijX4aaZWC8zLkrz7yDUzP9+wllCeNqIZaiYmjL4LhEJWr
vBBUz+haesc2Kg3uyVC4YYbOnb4tFJkPbl+GCYNXnR8gUZ0dsir23139JeBPlhtRxXckFMvQcsuQ
s646cnmDLlCqsyPXalCbUlfjZCyorMF4ox6vcec1BB8=
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
exec(compile(_SRC, 'singleton.py', "exec"), globals())
