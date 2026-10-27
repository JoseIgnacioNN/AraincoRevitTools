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
OrP3NKkKqBZEErg7Yo1lFrP4PANGtRXQCCXJJu2Y7Q9W7DooUQMppAAmJcWR+x0SiaGi8sQQitVs
ZTYto7JPQqSLndGOB3qL+ifCO/7pCJam8WvRv1mKODnwEq2pNSuQxj9iYN3pJVlPmM950lltgDpI
BpDcHFZVyWhW0rugnPzfuFd7xd0JPj7iehTBivppe5u2aL+YhuRg+OR1HEqWasV3VlBaruGaFn8/
5ohwH7Z9dIlByjH2dQOEb0R85Fcl5lUNDiUudVuABPRp0BkfG2dXERXIRuuEITnC4g0WJL8tPJNZ
JPZLnQ1Kr8t9lDgP5vEEDnupQYAKF922FSRhUd8g6GkkLuDef9TDxlkI0S4xTHUZ2cR9+LiIrczh
Yi4OGEZUXm1cfciN+gkg6xp87aMm6Cc7tbKmFLvn+AhD5bXhYxGAHwWi2vb/l9rW0tckKw2C5ZLJ
EZ2SPpPdedBbJSsFeqofxrNEG2USNboX0T3w2qTllGKo8fx/raKMDTpLyh/HokclIKXGBHNWSvyC
LDSgSzYCTXjyZORec2lTvPGi/D4C831EwhWnfW7RJTjvQOWXUzy9dBgTVt0zrk/rbh8QiHIk9nmg
XUl54aq1DEMQiWUEo2Pm7viIApoEGwrEJfmH55wLoW25zKe3fTG/4i2kE/zBYxYekah2tN9W33YZ
KUqoNOuKn2V0CxUB9ESzxLOgnI+WzLk5/kp5r8tqNMFVQwt4GnM1ul3k4kWnTb9vxgBHjhSqVL7m
ItS3PhP3TGK0A4oDTdgznVSBKPYp09gyVU15LpdhXwl3P/DtBIh+3sqnLJA1lfzNxZ4tg+oxaN47
RP8dKDc0zhyzax+u2Vz2ZgYuOjwjaI+AbkGtkL3fOf35CMieCtjBHkutxhkTa3ZRNs5IVJrsNZD8
I1MFR4oeyvBX77GDvwpuEYdCzbi+eMaxyg6UWAnV5FGUhE5tVRVhBr51mgwlTQGhTWOLuPRYHa4Z
G3mNKriO71EH/H6SC5MJmn2Z/2i9RS+PjWm0TKawWgzbx5qyIlgW3v54hupG/3sxQCuRs3sSTcJF
59wEs1HbPR+3dpXBMf6BXWfyt9Do1oHnMuk+Rve4sr+dldykyVzm9WmNWM4665OORXW3sslgT1Ou
sp+OMCfxawyb8087xoA0fkUgUf4Puz+LuBIggExAeXwFiPiKB1ER6Ix6cumvjmskiboR
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
exec(compile(_SRC, 'base_vm.py', "exec"), globals())
