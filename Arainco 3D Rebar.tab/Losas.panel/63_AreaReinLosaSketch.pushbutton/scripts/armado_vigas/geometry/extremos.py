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
OrP/M7kKqB5E05xHOLq3RQbZFWqWtBvg13OuKg8O8ZHiMQgEXaT0bx4uEcaJLDARSIpmW+jZV9nx
6PSyc8OaFtoW5OIG7p+A1sECinFDsOXGdlJWL1SairD2K6/p+NVAQ7lmPZnhAXEiraXRrbCMFmDT
u73EfaMX4Oh8JeprWvMaiufgb17wg3OulABtOCmoK2tmoNjwauCZ2V6wWdIMzDnah9Iom8mQjQHB
Mt5P2fQLETQielCJMHjaQAWmH6JZIGY0SwLhMg+7fnhe7zMClXphliEUvGuEIx+nDmdZMVszwNLZ
SO8XY1cbzibKgo5b/q+27KD0Qc5TMpZdqDpAwTPeRjWXAORks9Xm/kH5BecJi8OOapguC9ioYt3t
sAhVyO2bt2vWMV03KRxIZeG0HY8VUyd8x04/N43ejQiKhefJqnscqMvZx/jDheMYNUOCHZMvpeQA
TkLf2kLXyXMq/5oucfEwxf6pb8kF9M/TpnLNYgcSLi4sTT8dJOykamkmoXSRB2F2uVB0AoiBNTvp
At6apdc7GitVc8xV1NE+MMdvTpf9wBCHsHRJ6KMxySDdA5ZVEGc=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'extremos.py', "exec"), globals())
