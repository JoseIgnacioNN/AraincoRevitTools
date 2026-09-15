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
OrPvNL/qkBhY0YhFprlXsRXyK9PhF0lQoWoTh8eliG10YXak1HcSQ53/iwy8Z5hEiFsRzsJZKK2+
dmRfbaXS9Q12RTNNpZHtVmUyJpsr8KExpRC6VA4LG5nXLExkh2NrKlJZTSrZRiaa5yk3FvAD3XQb
lQ5Qjr5ZkyP0ZZ9k9V0YBOYxgkNtitegMBLRxXw/5Bv7JFAdM94eWwO7jv5MJZlz20XVpTk6GvTj
XeI2575pz+TZIJB1xsvAQN8vy8R3EdG3TZwcGm0Wl8QzUc2/CUF4D+sQ7SiRAe1+8t1diKZEUsxX
C/5y03y2UGYHQ3c1TP5bILdYaMJ9XuDFuo47uEjtVwk3fUNN1NP4cSgbCu5a4fPI0wWAQWA2Ymom
K+oFs6hEULXKRFAZfwdNmGae5aTUt5Pu5dbkqkswfqFATuHF1Py0CNWp7nfwcq9Noj7bSfPOYhzS
MrR4YRXuWPWg46fOptcwgagNZ/hwKPPx7/CWdOgSrIJB5gk6qIZUXI90UniqyCxJgzsOqcg1CzUS
duosB10jIvgxueUlUfAmn/iiLGitCV1CwTEAGYKibLS/X6ZTkFu2eXTK+/hrccEbctZWgpjXq4hm
VV5XfG3VgEnehrvn7XkfpXPpDUcdzvXa87e+0jwhNzHkYG/UQjMEf85KnBFpGQJlCTBYO55qUw0V
vM+/j2Jhkiz0vJ0xsCHZwEz2qjX42TQ87Tm14fmBLhgDKBtwghg9/L7y56Q2SAud6DN6xwNL/Ecs
qUXnN7Kx+jWp778MrLCN5vc52sVVHQrHb1MD5N/MUaiG818sQcF3tz7Fxkch7ikrDpPduuEvvjJ1
ZbkdE/fY0GgOdALTQRbnSsKUpzQk4Qccjku3hEbpavbxUootnVdpr3GUcn2CUb0tRH/G6dCWAZF+
2xGBRxnREVkveCWzBdlsKsezAmyLJDz8FU24Shgi5MjGw5ld6mHgtfhlqVbrQYsjyOs7KE4+YjQ1
4dFlCvF71cqpNKHI/qIT+EWuGADxGunZiobxtO7+DxEV7/C4NUBbEJ0I2l1F1+NIqlnpq8qbMOLU
w94v1hyATvfsOiPBaPe/CGwwHO2j5CAZAdjRkMfgVxgtHVsLVEvYFJ0psbOAdKWLPUkp6opAHa5E
p1cT7lBL6G8NqgbisgmCtwomxPK/42RKN8rNVMugOWBmcMwrb7zlvkIaE4udPt+KwzI65p0fbvw9
jgk569CYrZzUiiQZKDnedmPs/HicHwROUq62N7vvKZjaxBMRyUC4QXpGBMGKV7hDcIz4Jnqpz97+
4ZNIB3zBCny1y4raU+QpG+5D6zH7fdzSf9Fo5g0SvEvxvu+1SGSY6mfIaTsV7cmCs9n49aIoQwUm
SE3cALzw7xjVzZlLpZ2Sn2vVkKnwIDgJqWIuv9ZUos7tmnd/4BipYLFGYWY=
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
exec(compile(_SRC, 'segments.py', "exec"), globals())
