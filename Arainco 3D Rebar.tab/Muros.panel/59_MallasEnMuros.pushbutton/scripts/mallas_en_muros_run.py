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
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'mallas_en_muros_run.py', "exec"), globals())
