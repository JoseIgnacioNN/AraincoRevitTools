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
OrPnNy8KrxhEEYhFJqU3QwD7wOe3cQOxTKS/GuTWRzlKx0K7QrLhLB2WaClxK/KxRCmnSxrm7cdS
4N39ODWumRxN9Af0uw/dFxGso5Btkfq36KFMHmKY0I49q91CVd3NDBh33UNF2W4LQpeMQqp83AHX
Vj1ach7Bu0JJJdpI2scspt/NybwKvsb0EodStN9YUoybNLdNuLtBMmE/618zP4UGbXn5moKsJRNV
B2bDsyxhQ6u1DHpojPwrlYrBVv5PArnS3DPoYVOdJSq0Fz8mNl2lDrt+MysaE4SSiZPjzdoJkIS8
p4LVtt8BiyO5ByUASvGxGEqLiglC20YQtTo6AYRdrmYsQPYPD3pQiYxXmm6tnlcwL6eB0FJCdni0
Vv3QXr8QeuBNIUZOSz/kzJTcjPo8+HmkulBfpRUqPB10ByK63MmxeQM7FJh7QmMqn42iuYvQYY3u
h6GN2SX/mG8LcRdc2s1GdYUPo4URKGs7vr3joSZ7ZlmWbwNaL7u4Px6i8D17uGRY1GR4F4M2AmLD
tVhXAqC8DH11ApavHVqzIrgjtRCvJDZnCxZYgtzRXHF2KMi7Ao8xEFhZfAaOf3PkU5vRAZE5Q4c5
BT/jo9Zx5z0Ns+yrqnA2NsvCbMQd9DKBF6V5uN+WcoojN0K36AXFUg+bJfjL3x0YuDRyhcWA/+Ma
ZHXpbB9BcX75VCBpx19716zQaeiR7egkakBAyePld9QFCBA6ijg9z229+IGzSPVX8iRPr2jilp6O
iwIgwtp8VzbIgEj4OkaJEYRt6691us7XSZGviEFg4y7wZR33uNIxQ9U+B1GS0y4MmMCqYJc0YrLY
68LhhUdyn0hduQOfdCeUcMaDMlgwenlCdOqIS6Cpt/q1JSReIJ719iA46dqtXWMLf+LWiWanWqlk
GM8Z5oOSgXkVxVr1yzfSWSBHTy9Xinvwf5eLaeyEz9WtBNJzP7SwmbGqDyODyG/E2iUxvweg6JD/
pRHJ648m7bvSFg3+rnT+gbxi55mA7LDHxJ8utfTLFj07aY0JrtTM5lxZ1jGjOYrclNYd1ZFr9zh/
F50WFNbO/5z6BK6r/lu8bGwHQv4O9LWC/6Y0hbZ0mvYpNgjuKU5rcE/dbBLdGG1peQnjo8+V9o83
ZJV0YCgnOpHECaKGWcaRWmpOxPZdtq2DSH/cUFJHyBNf6hxNvE+ezm8Od7yo+yLY9ZJo3ZEWa5in
HwHrU9NviudHAVKmqhqLJuUvleSPY749XDc=
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
exec(compile(_SRC, 'people.py', "exec"), globals())
