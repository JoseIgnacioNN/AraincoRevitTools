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
OrP/8DMKaB9YKphWK7Ecg9CwYdhveReCJONoZdlVe5CjrX5V4CQ6AIg6YL21pHrWLWAdCHpCBeNB
Lvu7ic2JLiIUyWdlY30hT8cKMPHm+DZIOBAhd2lj/R+L6MJOKWEGMuakAnRuQ3vd1lwKaeXCiDsm
qwPsUHlvxWthDikysBnMT3CEJvkDpiSw0a7P2yuZWeBqWc0svOKxJ4fLaBW7mrh3RKPZ1VdaNJUG
R0Un+MUctHpMYL0IkT8tge/gMYk3Bx+BeVqnFOaFKB9E7aNvxMozJa5OBHEIpi7D/9vuJbsQSXPm
0LGTJJMP43VG1PT56qPvTk0WOcjsdRR44Bd6i9p0kKwrq53VuOs8hfwB8xpCGeB8zPcD+GAA5x2C
ND2i2BjUT6TBFIYYykkXw1fkEpqK7obgCqYxFt5zBSmPygbxkyFjubGvUAMKVitUpzFFyfCPX4EN
FR3oh+1P5A6Xk76Fnuq3y97zLjm4vl+HDrhAWZDcmcKYIYxOcZxhwCUwtg==
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
exec(compile(_SRC, 'mallas_en_muros_run.py', "exec"), globals())
