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
OrPnNr/2rxpE0ZRFepO36T0ph+c8Ldkh0LFs/ToOKmMZc6OL63UYZRL9CN7VHu4X4JJXTz+HjI8z
hQipONHeTv3/xEElOy2/Bh5OY/n0x5TPwZMo1KwfHTO+c8Q3bSjY2Bg/L5CkvxuFwgoKZcokcOQu
XjDyN2JtKeQ/u/6NGqOu3NBo15+lmE5HSs5AxMTLtWwVPvPQZilHPTE5SJDvNtFNjtZFW8zRoZkY
lulLiJmZlRAMr3g6cOO7z+c4Q8DsAEpezCatWUQeQnUHDPGVy2mdiBy/mLHpaxyscTpFlcNDPXc3
eSJyIxXeNeya2rAoz48ZBK8i2Sl5AOuV1gKf2WYO3wJa/RV6nFIHzdrCo6YTEukPWnqWAzghqOUf
xucZ1pg48H5DJZs1xDnzNj1CfKyNRQE/rCXglLgM3bvz6Dd1BHyqR3A/ETNap9Eu/cNzzBf938My
Fe33aZPP5fRcZFuzE1pYETwFxOg+bkh8XVZ+TQtDTBRsdV2TDyBX/YHlXMj8ZAv2KiNzDtKodxjc
odB/CdpU/n92t/biw8ZGU9U/YfhII5Fj5e5vIfUC3zjcZ1rDMXDBKcA91V1o7ZsRKX8R+wibKBGA
PLIFiQwOskyyBMmuqm0RegQ9+x+2TfWoIbgg6tuEMYmN2oja1PEJHusoo0zvqODYIZChSTcjn5Qt
P2KW7QYoJDhv9Qm3B3dwL0ZjG+vJxDlKyd+ub0IP1ytdFLMD+xjxhvGgWalmaHX08lgmjAvxWgM7
zS2272OT9hb12nmHvzClYVSUYX8BLILoNePD3pRA26SgZ4d+unpPHY2lp7hS5F0pxe7lLxBwq1OJ
bFO8Nzkpd7PJMUE/lmLvWmMp49d6b5MkD6SMJZ/XB5/PHwC2C0DlIIGQctpnHzBJxSpt+jFVXNfh
XRBqVcdChshlPNxt7C8jwpcZ3ae9eO0QfAqOMfapJWasln0YRcg0pn9CUiObNVkBacnKSvsQ10V2
pwMeuhkW8pG4NzFgv1gBaq2eK+u860fNhDwgwSBSNFE8pFrN8MAiMzLPrv6QbhNJIpLIGKlK7tRK
HBdN+QSSn0whC/I0HwpPBJkQTQ+cLJvjH4E9ua/G9linZXIDpq/CSShhhNOsFtc85m7pvnXQVjc6
BIWw26cVPXQFn9qqoZxjlE1cVM4dvA2sLf4bVzf7FX6NUWMGshI6+b2uH0Vto5vCYbkBmFHIwJ42
ATfOaNGPyGZ26Dlq8dDjH0uRXSukQ6TJTmI0bfS0i3sch+igeemHHwLb9seGJkJ5MC7AKTBK7JtF
hMOKcRl5Yso+nHQe3RsTY+/XRKd5boZjPYLIq8MMMKVTNuFE6X3khTbcszDblzSyh2W91/fb8Guq
HG+DUA+mG8+IUjpsz2vKEoR0RNkC8KmlRw0ETsuPZnKbtYdJhFNfSL9qVkpqKP2ywVzY97CSiUlW
R2H+zBGn4m99o76rGs84vz7aArcxovnvqJYlEmXeb1PBs8DWJtlHij5aoJLS0Hj7xbgKg9nmcMBp
yTZidICx/AxeQWT9OsFyAxxydYMHcZjUNE5DExRsOF+xkAsvB/panXAwqlc6trrgRUy/X0nydnkt
CpQXddXpzylPVlq/c91aCa8Xq3KN4XhGSw9wp2hasmPI7BnNWWSGYZ5lMvbyAgxUwyULePEbks3n
8F7D1j5b6Xsaxg76LGe4JWwfQg19LJ7qXRgAIkLNkoT+NragSn/BNyWqvB+m8s3bp18UT93iwR7K
NiMuvLtozCvrIIU8K+XqzDT2eWzyOq9bFbS82xpr0XeR4UoG
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bootstrap.py', "exec"), globals())
