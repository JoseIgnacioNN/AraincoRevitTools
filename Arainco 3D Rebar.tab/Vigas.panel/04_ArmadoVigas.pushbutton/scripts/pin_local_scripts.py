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
OrP3Nb8KqB5Y0YRFmryQY+2ZrDLG8t1I0XYsNVrDHylWsOhyc+pXfia0g/HYRAneNbhSnWTeXNFJ
hrqo4YpI47fOT6z/NAlRFUQwwJFWeTZIOBi4LR7Hx9rpoi43wUjeNN1OWF2N+pR+pFY0x+nRLfKa
q0U1DBluGer8XSsuNDtyLrA5I515Cjc7cq8X5dLbI4WB31Tv46UNKJKIIKBmRzUWBxS10qSh42Ag
N9ithP46/EJwRJmSGpqXspy9/YuXLU8+qVB2FyM3SXtI5Vswnr+5aSGMJquZNWFaSDqXbAtA7bWi
27/3QcLKHsXLruKNibRQB7SbCoWj7slLxGo3p4la7B25wYZa14FYbCiOBaQ61LM3QmR1R6vYjn3i
5K77X9uwGtGDFzc+FbkICqU2/w0DiaupYRerFgAaS6WC3FjgO5odYHRo7JwldAAM5CJskVE0xxR0
zYS6KSpJu1VLUU9L9rPTgRmi9YHCWSspRkBWbNw82xuZZtrM4tyyF0Q5FLMzkbvzZqOlAumhgmBC
qd+KGkqKb4lVIkmuCgVrtnLsLCLjdtda4mWn+T8HuktSt0cJ2/8wlruJELzWKLHuMXuqW9gbsC9Q
l32iryCzrMQZrbeRMa6Iuyy3L+MT+2bYfR6H9t8BxzRq7Uhm0r/tsJHtmPfBC4AJuHH/AsjDiWIR
ajjjftU7lTMnive5Clb84j+EIwa/3m4fHPaZpqk9lZdgT2sb39Hu1X3O+27uOh82CsIByRXQqnX3
CBns0nX9T5CJ
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'pin_local_scripts.py', "exec"), globals())
