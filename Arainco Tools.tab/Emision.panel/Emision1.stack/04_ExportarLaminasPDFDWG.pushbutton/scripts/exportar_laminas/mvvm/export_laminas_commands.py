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
OrPPNMMKqG5Yspp5auo67smXP4NPPNTVSf5fxC6xHtTYuucbBRbhg6u/QIT28SFQpVGdTU9cqunF
TCk+zjRa2AZX6O79kHh2ykAu1XyLGGoJ8mE9rDwBFRfiNQ/tfmRO77jsTzLnbe90s/+bdjvAfoXE
8RNY1/RUmpvc2XYOGwG4lFVqZcpl6Xgg/WNKoJOfapHwHGUoROULNTk1AgXEyl03UeV1VMkMWpvK
SaA+2oUG/CgwCxq4dfm0nY1yNZkHRV0026bFAng7275JuNDSv3zyV2gmguLfFFXuJGSGKlKybdi4
jc5XMZLdO1WLTda4/sCRWChf4ACi4CF3CIsCr8OMdoAo4/typ8iPTMAvRDXde0jMjGelU8dWdO28
E+Em67Gg1LMeIMxeGRoSEvX8i1B1Q5RsGAKhps5BEVQyKonp1DW3j47eCsXg879Cpfqm+6isLGc8
0W/nRKS2A5ZjwKJOLBPNjEb8o6Ea0Q3rGaaIJhVZ9loOEeGesSd5984K6TOOxKJCG8HNvfXUZfEJ
Mmp5fYdsI3CRMbGCuZjG5HJiuQyGyukbVF3mO6UyZe9IK+dT40i1RjJa/49GEv5g4RXE8vn3tNjb
BJ6pVd3aX+8h34E7oHzpO8FwdhIl17ZttTmod/gmKDPfTWIImNsj6u9X6GcIoUDSrrOgwCngRZob
8JF4hgdFTSnnhEQwNw8WyHwli+xpm+fHUXftTf83dYJ1PuOoEU/LE9JzuFZiQ7L3mNUuwtKrGbqx
daHNUU5YZp3dxtK3XzUZ881K3K7gdJPWthJlWhAtdwSnsmFSTY7/+t1IOzNOKzcB98XnsPtXYHHL
AO4g9h4vyAAD0WJ4AnOKNgfOT523ZREjk1fylN0tyYqeQzh+DOZTlXtSDEGVLMCGBtRIn+a/QoAz
EXY/kT+S45CmZtQO3rLaZi/RweV9E35fHg==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'export_laminas_commands.py', "exec"), globals())
