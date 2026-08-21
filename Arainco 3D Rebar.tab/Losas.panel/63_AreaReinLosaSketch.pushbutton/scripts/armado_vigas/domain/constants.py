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
OrMH8b8O9x5A6pgVpNuvB7WM1tMu1PiLxTFZRZCuzSGWZyhGfDGI0plXGXSGJAN4gxVo20EMc4iw
9hMHVFcKsDuTyt15JemIUexzSw/3BOvCqylaE8doI3YJTjS/pmd1Tc7bn+6rIyWtQimPZ80w5/Wk
tEaF+bJIWazfAhq/N2OEbB0A3kzNCfA9qy8vGTvRnuDgONqo2W2V+f0+39YnI/2aOSlqK6EMAycY
3CeewWa8eKzIlzmrAmKl1acsEZdQ3rqYO7jRHBobJn/Ghf4J2OiLjbyYnt8WWdn/JZEKdCsdSwJC
wNLsyCewhTNsTikSx7RGmYM83QLB5pB25xYTGMtxx56Xe6LIL+o=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'constants.py', "exec"), globals())
