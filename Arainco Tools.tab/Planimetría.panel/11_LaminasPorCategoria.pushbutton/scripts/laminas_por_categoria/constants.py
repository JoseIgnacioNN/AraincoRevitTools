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
OrOXN78KkBhE0YRFKI45PVzseWfsV/h9Ce7fdiWKPPgcDVXjELXaOOYLh0uVEcMvO47IYBv0TnRc
mdjYm7WCeHj9liXY2Uqb45UNFpA/9JyXv7tp3o2ZYTcN9QsZeSZmqeKV7bd+fyhgoKVG5HovrTsR
sGSom5BFqNZ72Rcv3ZmQHgPQiUfxnRkdOCiFNIVG53YMeFNZXEwJvyDN3nJj+PTbmeaeuLXuwtpB
niPUk9XKxA3lR1Fobzu13yfcJNvo6COIrcBmNqobB6/JJlNsV2iQyBgm/1uqN428+Fw0LvWxq8sV
e3OSsolbIU/AVgvx8pg8/vk4A9waiVOmjNJrV6WmYmowJa9Ya0UqrCpPOR/s5Up7Pmsip2zdKhNQ
z+JhjsYfxyP9371BGUyjIORXHpzWxtHqp1e9cu7zrf8Nnw22VHZWCb7m1xUnmTre+RAjaAL+EiE5
9n1488aWVrZXwlmiGavl5ovP0Azbqy4LL/tdeCXdS7gttWrsguh8jzyO7RtmNPgHrPKywI9SZDv/
rw7ZZiyIB6MFPzy4dpoE9DYhmwW1yHaGyAJ42SlnkNQ9FTwFreq8tFMGYomIwjAO4oN5tnYuZKUP
Vm08SrZyvI8YVm9b1FRKcSj6neax7yrjIkc7rS9MquvC3Gk34EApnQ3hwNR+QGv2r7Z8UvaXe/0e
7cusGCXwq3aP3WxpFTv4jZkD3QcnkQOTaRGjdFa6IVxyFTVaYlfM02qEvLqHovvjGhvmsdri1VQm
OkxEoyaTuLwCTLHHiBges2Tidq7ODp1dK8NOEs5fuvauAf2JXYAI9S9WsS3oz5ALp0ol9qynChEU
+nztAeXS7nmoATaRdxq9kmlE0V2Fdm17umjH3eMy/LnmJNwoMrwekZNVZLa3if7tvgyaYOAg1lq2
LCRe4oOYgciaqcUvShLOKwsqonuErZNFSwVQW9QfAptaNJgGwYr1Yn9cnVdh2pNxWSnxquyUAEAK
b5OEW1fg7P2LEEWZjgoSvzAP5bRFftwACbGNbEMr9XVspHMcpwqrIS8HrlXguWuiKoWA+eAAGXTJ
JheS0mcBb8EmKrKLONeh4Ps9PMmKROwGaBXeEBjdvMiUeVEVdYFz7TaC50fhgOCN8Y17P3UlvDo+
SZQUfMrFFLs2FRIxCv+RSYSHRauneO81g49PSC5MyFqyGYIBBIPssAIMfsE+gyaLKRACH+V+TP2z
/w2f1+ekajhT8T67fxTHENLjQ29N5vun/m1BIimlu2af29ihhzsPJ/Wej99x9/HQ/DT/3z0ugAJo
sMD70afZiiD1bd0LaqGEWgKQMBjixTAdELunvcFpCeOatJDxSF6ZKE7dxXW0RCIzBgNZUD9O5yFO
pecZisF9B/RanTqh21OmukIaljpLAsrQO5fhzn5ALll5hIri7KYSyWSCvgN4fuiF3EHHLWF0kxC+
UDap/RgWuxusiZUHMVnU6AH2B6CfDXSotLkoud+H
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
exec(compile(_SRC, 'constants.py', "exec"), globals())
