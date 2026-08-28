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
OrO/NznqqGhEspx4KzYxSX4MDlSnbSkLH/XpIwJWsPmEad0k5AJJbSK8kv3CaHQTbHEvrPZWvy4T
5j0QSr81OPRedHZgoVmQ/hWSMpnu8cyLoe/wU2LEpAG5m4d52VJ4kZ32XkxQgXgWRAS67k9dQYTx
wP0mqhKwYR6al05E0NEmFwNMF8dShexfY3ojQ2fse479ylx8ITQ3vdVPqgovG5urtnsrx80BCo3l
98yFbhe/ZyqC/JyIUK/cAf36SYXoLaYF5Fu8B4hMePpQqMkYNknP3WJT2Q1FIde3cklu4XYcmazZ
LC5dmgXZXTCk4oTn/zTWrRDizfCKxizCtwOwlYtZUo657vXvMn+hf1ILEwiE8jAz76WyfV2elxhH
Hk+5+iAy82jzI/D5KjLWZxq5NrNLsVFphV21pq8n0fNwTwuJHM1NpTinevn89dfQZsHo4LiHlvPn
nogdfdd7HiM0Kx0Nfb/ZmncPYodcm72y3OENwzBKUiN6a9SJrfH6q5O8kTbTUN68+c3uO5eqXNUH
iRpjDR01oiuH4oBdNB8D94WLbNXZdqnSTxRgJuqi8CTjwg/moC/ZczTA9n4MAYJ19V2J4xhn8t59
Nebp0584av62fr6q0rTEJfe4nDiUUDBytKGlCAcmPTpS6s7yEF+T9p0ytl8Sd75l6h66HCn7axuH
M/Svzz5FPG7kWm7Q2edaKEf1dyZA+gEgEM8gXiJRps1wpAeXv8jzJIuG6Uwwd2SkobBqFlAY3ZVf
oEc/lVHfV0dWMmP0XNmuA8ACKXulvLswz0NJm/mYrhB9h0H5XilldX2pigoyiRDWpMRXfaEe8rnu
lb6hru1y9mKTiR7OErvZbmcsuJ6bkXqe3Q0wRtcp8bCah8ALTr+LnU5DuOTlLdsHw79vCejvQ7CA
vW28wPHr8YVTQ6KPXFD0gh2Xia/FxVKgvBW3GmSIGGM7dDZVwe1FK6Rsnml6dWpZ3r7wI7Cr+fyT
rhGdFtfWHVONrtf5rd3OQNQCxJ4yWvdXLO9LBDBtzaNVXkKLw3cY2VWpp7Sped9sKKBbhXZTSs9T
CuGinm57uQlCNZ2AM//bhpZqkabUfA8sui4fY81wTWAehQIiq1LdKVyWNGLA0OeMOY6Y4I2B2e25
yspMxyBX+SNda3YCEAbMDpNP8C0QbW80ZZOWfhRyhj1gcp6W2oO5iLB4/BdKt+1GqDhahmlw8sTi
bYnAE8So9tGpBVWYhMzk/7uQOtZP1ohn8zTmKuxnsy5/rS3hLoSPfZ9KcWOi2QEiAJ/hN3y3mkZi
bSSSLVlFOHo=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'armado_muros_instruction_dialog.py', "exec"), globals())
