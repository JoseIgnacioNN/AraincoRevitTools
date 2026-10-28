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
OrOvNr0KqBhM0IRFEPhlBhcauFQCgHOsEHUKBlz2K7BgZH4n0HdFHP3+O/aOuAdg1BlX0q1FZ6Bx
FO4OWZG/ryZ2UytkAdTg2qRO2pFvzYCbab1regkFY/i8cNBSqjweTddO9VJosMkYwefusvrubUkO
KPAALztFMQo8fcd3WKFPTO8rAnD+tDumNJCmQ3Zm5otbSS6mFsZKPbp7toRiYwFAY7Y6gepI1mju
TyVFdjW6M0niGWAUi3h5gQIoZXsWJizdM3iNAk9QJSQLiXBZ6ZJ9iyhV15pG2z/7Ix6Lf+boWGQ4
8ovV3ZGGplSzz2wgc61hC/dLbHx+C1rxEhV+02IReE7qPmK+wUmfL2sxs0FPaPbEVtJEL9CBVOhB
zmNyaWOqYybOUbb8SycLpZWFvCFUe/nF1iabm4Sk9I8eK32lxhwm1RdLtJaDC14VypDQlQxDsOty
wIAIDlsRX5sMbyQqmbP2cCp7+IBo2SG7CdfnX2anYwsAlsBo2cP6HVy0uGP0uTeRZZfBTd+u4kVW
ePYEO7ACVQ4aC8j1ITCwuG7hpI0U4Ke/9EVOGSsWO6Kt5c0FeVnPx/gs+CpLFXZWBny8mk5o8+PX
dWe9nUpiTxbdxvfQmKx58TcPlMRL/nvQ4UzKbJO6YRhSuBA1bOVTk7Ypss9pjwh8lbtXh7FTpg5u
6K3Adl0umAf9+Shx+VrB/zYIRrx5pWPpuMe346YHn0iZ8L9rXkh2EtTyeN03dS9Rs9ZU8zxn9lZY
qlX5euI0D7SVB2TY9fVf2Vomta58NQFnPrQf8ka7ALFkmrpQVsUjNkeMmG8HMaf503miF2wIo+qL
n/Luj7rdQwAgX8l1AVYF1R8X5v090ou/ProgSUi22hlZTsN3kGrs63yvxyApl8cgVM5Kr83m7dSi
p2vd+cbm9Jwuh30AuwZfLit4MEAQ4Q01yEQ2iwWZg7DQCJCxroA/AX79/j71PSE2PLAx/ApP7TU1
vaYukQxANzY7L04wRFI2rrS+vpSlnd6cAx1AL74vVtIg5iuHIqIAPP2mx5pabjcHZ7XO5V3uDgBa
LLIPcex+RaGQbIODaEf55nBl+TNzXwlGPOlj47IaFH/+3Xj72MG8MThuY9jhf7VpWWd83bRt02bp
zVutObvohiYDLLr5MlqLP1R42+qtozRmLG/7AvCvA9xmzQ6ybLNK78HP6+VObhe+VhE0jVmTEzpM
jBvfYVcph64Zym6o5gpS/mE9kyHqJOncTEyemuaEQGUToJ5/hzwM8nIk7jHyhJveP4qu4KkFf8Uf
xS3LoreGeWfp9CiyuM+kbnNZ2ys2mAKY7mdXP4xAsJW/udhAXLCcetGkc3LhpI9Zqm6AtPk8Pj7w
3PJPRn2O2PMbt8wLzeKAyEP5hsTzpGu1YjfdYJe7Hlkn8LB6Fvanf8KcOHXy+9AGv3bthA/SJni3
QddRt11NUa3BpJ6lUoKOBk6fdoT9sMS3DL15fYK43I5RiPfS4FWHVPp0veQWpz6zvPgbvOM5laDO
HM4Uf9in6ALrtFeGSuWNogdkXp5Jf3fhmHEJZgYQSx5Y99Ti9rTPkUDdAZiDOqc0pEIS1cbXNgrz
CqdlP+s9qYtKrju7w+sYH1eqCPui5AqQEPa1UkCpmsa/UIyKDnZEZfpuEQbQ1XVQ+NJ8owqtGK/E
jNVRAQyycLBcXvLFNa/AzcoSFBqomZ2mqGX20dvZeFYeWZJQUJIqQWXEb1LnlvAi77dWugqVnt1C
8OVzpNNecGSO5b7Vg3s1XFRQKZ6+3ViTbqiidipSnnkUHgH/GKX99ACuP/3OUKRQku5kxo0rWX4q
3UQ5YdMaeaZNBL+s7vxnVduhoURSIqP1AvY6uZaGHMcXOtu4QKia5vofja36+lBo4NohwOAw+dOH
9udHkWhoXbWyVDVG1B/mtGkN+twyftBRbXt0jA==
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
exec(compile(_SRC, 'people.py', "exec"), globals())
