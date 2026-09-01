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
OrPXNr/qqBhE0ZRFKL43AdhE/79K/14/AaeFDkFv5zP7ff+k54xeT917JyanS3anTSGf3CX0bZht
DfoS/FBSXL+z9FbfZLAcbZyiY5HJ09DGneCtTiDq85YjPba79wFfJ1vnLx0WLV49AFwW1qh1Zg0P
Gfaq/8j/sl83ZCXJOf5td0Sd+8s4pArc047aXa8lr8H2lFkSQuNH5/w5awPhmog/4E7lQPl8PSTK
agQGFwXhvNzrs+7YaoCFm7j7e28GqbGh6y1YrIdh4SKU9cTlN5Rxz2q82JBoG91NcGzRygDUfkFV
yh8KX1SmgFkGw/7ALBUo9FOidTGqeQ5VpHpMOrBhjJuhW7kpwuX9ibq9IHrdcoI/oOu+ioC3NlwC
N5oXixwjh1ABvUSoVY/SgtpJwffqKK1n/maia32KMmfoqpJ7W1CI2zsbjCBgRCcWXtfX58E4leEP
by7SQylsCPG1HoRaOlIkI7OzkRd+CRorfneWD9pHtRvkr/n0+fQgrCB7IWK9TVb9K91mCEbjhLkO
VcNEVS0HaKrS8AsGumwMoie8wp2DnOV8F8P+tPP/MjN4oEpZpuO8t14xc+HyXm6KkZthjpOR+n6i
b946wrbdvgMAUwxdSgl6fdXGAFlgycOm/npT1blIwHjvo00RFQ7C7Z/6IR1Z2jf9bEdSZduLrgz1
Z07jn930sS+pmA8RTj5kguU58lIyaI/vITtKw74jboNsE7ToL7yxSQ7oWIzWNBRqEudNpIItQRS8
jepUQnx7BYGuRDg0rQUm80IXdyNKpml11fBq2OLk3899tryE08LK4taM/sXQkHQ2evPQhlA6gPUQ
scuknUU/0vBC+WHYEaJuM/MOK6iIn2I9D5dGOByHKvVhofQHjkAEUR7eX0A9wzVBV1KJG/yDd/4a
Jd3BBVDNFDvVKReuo7Wjgz/dPbvxuj2BN3jVw/1oGGCHZ8yQY+5cmRwSt90ADYcnii6qSgnIXhDN
B7S7sj3NrC73rvDyYKGRwcu8jlD/3llxMf5l+GlnH5Wv0fyeTsGVCJIVP3ARqc6Asf4Gjx78S16k
dpLdX4xW9wrzF5VHfMUbiLHyAamZdqp4meawfgdFjaRcZig07+XEt+f850h7WYvUOTX2qhQyitpL
O8u3KGHFR1bLr1Ds4QnP5iK6TpDJUcGuoSKbyDp5DHkzczIhiyyMtXM9Jw0rl+el+fX3JIsfPbrz
xFJNpQ3BqlJYbm33wNkBv8zUABfGOD1es3v5VYAUdylMALpPsysCQlKK/z5LE+1B8obMxASBHYq7
Htt0rAispL9QMttOwmUl3xm+ktyEqiyotIGoHP7dyXoi4jwCkF68+UI9b+N25rBaHGGBI9mg3W0b
h0rijp2Klp0XOpevdpFMRZCQXLhVt8AfXZJfa0GsiNN+d9dy0teu7yxUE9HwrrsHjBaeeqIuPBun
ExgED5ZMI2T/Tgft57lazlKL4+3pa9GNPcKmVYVxQJul/8IqXL2qUg0TgZ+/Y/dMDdBbh5IyxA==
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
exec(compile(_SRC, 'bimtools_element_id.py', "exec"), globals())
