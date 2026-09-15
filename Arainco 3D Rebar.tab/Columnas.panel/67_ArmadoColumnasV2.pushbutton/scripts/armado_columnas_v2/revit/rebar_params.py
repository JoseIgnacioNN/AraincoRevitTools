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
OrPvNbnqqB5Y05RHpqTg26qMostulbB9uy/qzUYkgpVqW2fJ4kwwoRGYHJp1prUlrZ/cDlC9tXS6
VlBzoP4INomeMTHcscd1FeWynmVu2wnX90CL+epfJY523Wf0bOKV9m1emfs/mm8t+8Lv8CnKmFIR
ZQX0UGqDlpNS6caEuApJKw5IVOsxruLzcs37X0gHI3sa1750ULg/bC5Ucz0loARlMsPlOqkeo9Yw
xI8lWDbA9uN/hxcmYj7NJtEBGPN1sDumGtstyuNSdZfvcidmlEqj3kfUA4aoo8EXLz2s4q6CV0Vk
IozmClO9L8O/5UJ/FLkmkorraIo/7tJ8QJR1Q1NvoGK/GApKTWdR+q24bWD1jNwQkkJCyblRtOE8
I6z508hmJ9k77fVoOe6vUVqJZzvlwldYQQp+y1woF0eFdi6l9YlPyyqr6v9NpaBVj7R2gn2xSWv/
ocGkpMsPE+venNrd18aAoR49CuhRcqKoF9/CzcKUPSuKUGatsXNmixxEcH/BdkfB8L5zLYLnAQiy
7x1xgSNJrkWUum53XlqRCvOH/ETi+9MHB0E84AP81Nkiuo65T2rbsHYbiCIHqPG0GAHrdWwO/v0A
O6L2F9W/5sWbE5r8Wg==
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
exec(compile(_SRC, 'rebar_params.py', "exec"), globals())
