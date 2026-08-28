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
OrPXNr8KrxhE0YRFKIv5O8v7jxfqt7ZlEe51mXw6f6EJODLVthH6uxzuy/hC03JWkS/VFW7lEPxA
G4tIh2kcZn2S7RRiZfWasKNYU9I29JHl4DBewTmWIb8pOgJ5ycmJ1wlI+E5LC+fx8ogY0Z/NS8Cr
FztXXh1oL1sQTEJzSov61YWRY8tYJAvczYraQ7/KgVkTmfkkf8NBLAE6PLxaWxkC5EVWrjlHH/pp
Rng1FOr2by+XjE9TXz+Kz9lEyRAoOTKL0tV/xuDl79ChVtV3kGZq5L9dS9B0K4ZJZxyQVtwi+/wK
ca0X681hHPlr/Xyz2Ne9/ZRbFsEDcYz0Tm1h1JasgBk4k2wQEbddtGOD6cVOaon+g40G6JrWRwDX
YCUu9WpsavBxYejvvXUfoba66sBpMcNtqU7Py5RR+4wvyO6fdpQQRqmtLln3DGORSb1vqXZSYFH/
gkLLRBNndtxXpGe6FTa+/GWPZOxrQbqWZOE/Zn3TgCCyUX+i+oYKZIMV3KlA0xEnjPP/2uc8taml
p6HjTH1sWzDUfqUFKfnqEu/IYunqScQdD1qtbwl1nTwP8DG4zgap+k5rsQUHzDOdGjhbiaQBtLmE
j1fzu3+iPMu0CPyuPJhB6JUZEa4kwCWKvJTOdzfa7vfJPxTm8yGZLyf20Rie3BSfWYKKrppOyF31
+5wJ8bkuuxgmRAgTBgIuPpursJUSMLsMbjgozRNx7+imoJYsRn7msOppAETF/L8NZts0H9q7EteY
DKKPCgFdqyPIDP/Ynrbdp3zHPLdjeB3xVp8di/8HPAqQwyuIvob9x7XDr2EtiM9lOAv5QIJMRGMw
xxOmlmXIxTVky5cHM4E4myH0kmMWyFe3VBXWMsEL1t8UUmwYj/ECZOLdxELS+wUiQIAKkk4x4ceV
A5VLh0febQyBkdpkFngAWxaRddHLFAdRELvDjCRLHWzA69U//pSQ04U/BP4xFOk8cGE9FrWcKb6V
kl2lg1lwcqYEyGt50KkrPx6Wc/YIfI0BhV5PaTG56Gmvr4t8Ig54L9bNwIhxRb9mBvEyT0YC3+mV
vTfvQuKG2nM1JezG40eEEtsKOT+VLDE//KxJmCNmuYVCTPdb7Pac3NwA5bvba5cNPzTGERACb80A
zgWfd50x13n4L90pGvqR8olUSg7oan5WVUja6BbNIcK2m3J9HIcS0Dzbeu4ukhzwGO8k4WIv9umS
+U0vze/75ahomvWtXOnjgYZhvxTOw6lleguhsxz61wh3ILm/3KHdvuyAVOF4bWb6ORBBz11zlvSW
4QCHW+/eBROpNZVYrT+whKDMxn2kc6kn7+ksXMUDjiYpaGHv7Zsx/9MMTJlxIVD9IDGx83uY7kYW
7uwdHfBIh7nOBxrzQyrFAdUzGyjtJqKA5QmneVG+dRUqk4IQSubxIU2AzdpsV7Gh1595zgOdyQ==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bimtools_element_id.py', "exec"), globals())
