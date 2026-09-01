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
OrMH8LMucB9E6YAVpPEsxdCFW3OP2HshwPPLx6lC7Pw0bgzyF/YHx/mAz0POUwqdhPPxK/mxqgY0
8AtP6jeIzZ32s2j+q7WOCekZUatIgJ4R9V7hpzW/K8ChB2Xl0x20FBDSfPMysoJJgrB4MiZdDwgO
2kse7VQhrSINr/q+TZsgwHJ8TkkcT1/14+5pZL8SBWIv/IM4Fyl0BIpVsbUkSI44Uey8QhGfKzAc
rdDA1MqO1t0X1JLQWS4ncRUtNaJXljOAwR6UUKAmpTaC4lZxI0oI2OW7Rmw0kY1oHBxYzzxHmv2U
xPfDmCnDp6V7cBdxJwpsRwawx2NU1WSz2xQGnxwfTi5gVRr+MsGW0sVbpafHdFYzbuTqf0muSBpX
iHRJM/8f1bMlFVAy9HbzWyPfzEt4fmZ4JejatIuih/NVIvSdkP2UuXIsTQDaGg==
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
exec(compile(_SRC, 'laminas_por_categoria_ui.py', "exec"), globals())
