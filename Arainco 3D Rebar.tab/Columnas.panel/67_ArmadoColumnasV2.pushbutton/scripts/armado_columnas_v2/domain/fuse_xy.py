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
OrP3NqkKqGhEEog7IsRlvvF6npclVTdm86BoZUpHJzbHNLvBJbChdN9Nuokr8vmb69SbLa76hq9P
Ufh793o+da+n49LBAxSov3hb0jYD4sMw0nfBWn60TBiycGNowYRlP5BNx+fAW6dLOvlvCnLrvUr/
Oxmh8y5Hgzi5X74E5WtnHWeCy7dZu7qrHlIIxDj9v7OfkC8lV0s1P0QHrZAPHSJkeNI69LdCYulL
v0wSRSYAkrw6oMElMvNOKDFIGsosk/RxZmdjGQsz23t2utWm9m5qS+Bepnp9cNl60juh1jbZSzbl
C/46vm5dotgihhDrIX2E8nzi7rc+irAlCFaiEwRdoT0YILPRdr3f5NafLkl3uIcTJOgfqpa3VaTb
2gtjGeHpsDhxWXsdg1eSw82NvJH10F3aCwwgEs4bMb4p419H6/F/AWB9HC7k0mpsg534V9SU+bQD
uIg4BU7pjP1Z3r8e5SGSKVR9rvER9HAdLFWaMbGXyUKl45hfrh44HNJMqwYeWK7r2zI1OO62pHjv
uEdfKGyH9p6HW9LIIZxWQz3hLR6iz19ROzOs7POlIa8x+veSCD/dWJ7GEPBhl4UR1FMu7UTjexXd
h6nFFHYCDFFYHoyhpXkxZTge26yAyA+7vbfya05vT21Fp35NP8Afg521ljMMnItwyGBg5R0aXzWk
BooBpy1QuDFxA0hMyKEw++MLwDEnENvZQ6p/q97CtWatS6J4pxpXQXUMJjHDKvHC18wvA1+9hhBG
ugp9n3mAzjyHCEsd5/dkcINr/ut8RV9uKnmSSU5kXDumPzApvFrh43TPkOjYxrRFdTBcT7oyIvmp
4zyM9Vl6hNjnk1zTgJp0bVwU/bhJJHOePeoZFPc0/qdNzu6MejOxNhhum+/M01VDl1yrgrkkX+dn
3Tiq+Ixsscjb2DAtqHmO+yeYUJwztIpAJs0a+sF0mGCruArAZLF8FgwJi8+XM3SO0s7aijUA9uVO
4+UDPAbOCSnEt2+RKzAfoeuLNhcxZIvFV9Mu2Fl3StvOScx2DCfJjovPVCo29otSaOY8PKLmHb1z
3n02AdNAyCMLRDkBsnHLo3j6UkUavK6jZhT6ky+JeqJZazqy6uFdKAQUjtcuDdgIOVN6luMmWA1O
W6OkqIVGInFfMeLmcDrIWVnO5Dj/pZ1P2N8dplppE6YCDmK6FBn6dxgpedV68KpATTXVI3K17Fw8
pXAKxTG0LSbGY4L7rLQ0vkdmwCKQqiyPDpmg4b/5tea8w2s87sDhgDemu506jqaCKzNibCqA3qT/
mIUN/LVGHqvf0aftCG6MOMge4UGUjjqJNg6RZTmPq7FAVZSD/oV81eyKYtnuewd7EwvCt1iCgQdM
wUSzqzWIKrqQSEJfO/0D4xgWMUKeWkPHXFm7ZFuE9w+Iwi/3/iwQi5aM49Al3Fjy8HoXXU9Quqik
/onj1qcSBNw/bTpAvOVW6biLaVzmkvuZamr/b5HXNb04iGWtx7kjBggpGRxXS73W2+Zfi7Q/xMjm
i1iLIOeC/FJZDMjnt/4BzD+yfJb7bteZZWOGw4qTljgtxIGSdNgMVdrtZ9wWTQcPjZJ09O2zjkJS
6MA4C4OVr6pPpQ0GaHkBdViv66fFc5UoJI/3AMCURiLEAxKIe1jT0x9nnra1vKFrQI45OeXvPQwr
KAZiaFgP3TApxPLjtm3fEseF
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
exec(compile(_SRC, 'fuse_xy.py', "exec"), globals())
