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
OrP3NzkKkGhEspx4IykxatCM4Md+pTHxRPtohG1lsEkFXaMk2wXh46Rc3n8qoFOJZp3WH/ZWi0Vf
PcP0ehHAnYpZVsqYsdhBI5p2VdG17rfmXzAZL+erWn484DnF+0DTU5XyHZE1MXPcpyEUTFtMHyI3
NhmfO08CXlJofu/j815w2nYJUuXVF/Wfp5YX2rMEGMtxcRkx9xjgfApdfVCKJLofbSZ0vZ+WqYcs
eDPiGiN0NEVVUWj48yD8P596oSs9YBo+Iv1WElC3NHVZPO71qKuvl+RSrU5we90GaI8jaVPx2cij
Vr7qFSz9KQrQHaHbqfMn3FKiuLqniFvVM7xX2PZdl+3OGQzvC90hLqmKKsCQlXncUDLYdsPmszbd
8RyHRFxTOv9yoPUYJTR7w/bgsUh4uPlMwxMtklrh+Bl5pYpWtRQUXR7gdHJciJu1uSfJBfMhzfEe
OW07pIAb1UtxquA7zBE3hRlYORHWJYMwkht5IkQnle6/pA5JAg+1ZNDbw8vxD81fw+SwxP8XGpRs
t1LWfrNFBbLVmJ/WUZe7TvNUnMfWowBBczAu2beCUyKJPa0In/6uN9RJryv79YPRQjFTJewcg6D6
gt5Dy6NZOKqeUDr57xwMA1hAfHjgAHbLFFbvL21W0k3nIfgJ/rUQGddmqh+gT31Fze07XHBzK3qs
1k7jTS1eZFZ4Bftqn290wFMeVVRAC2IUyXX/60KWJad5I8DPOvV2BY27i5hfJ/pJUsiZ39D5W2KI
5Pyl7r/XeMlE2UgR8DNJ2RTR1PVQZ0VgDd1SmIkdMXnL35skfzzAulIRjGthXmsnjb8GhK5xQJ+b
y17wNv7hV3XJ+eq2fyQjNbhD2uDdf/Kir7sGFYlgiLuKepED1Yf69i/P4IxWTncdxiOErt5N6EQi
jegv0Qd+eCRKAj41PnEmQHxCKqTpO8Jbkx8ATlCzZ8btgxiaahbulIVDZ+quOh6bABkK23VUBawY
CVHeIKOyOFk/5JS1DPJliZyDA9u7wGC98JefQbSwF5Ra68/YDVtqJ3jhxBsWh/6Gr+1LauaRs48Y
GHng4bQMaFbPh3c4q6ipVEXdvt+dSkMa/5DjlQw3oJ0ocOvmyzVRiIxUrr5dEwXwB5Hb8OVCMRwl
duWXY3WFd8jEVjeQL+iaFPdI9uUhtpDMn3pTS3H5XxPJvNDSu2z5VmNwtZXk0LQQpxUFgJmpDBo3
DNZy47EKdz2+ZgP/B5RcPXCo75aElsCww1ME15tlJ4hV6i13U8NjiLPTKJ1lLwFxFEifFepZVdrG
nw+5lFYbdiDQYkFNwFpuhc/BRUi7wE72AE+abYU0VUnooCn93s5bQwy7Q9rW/0QwpQ2Ky4I0JqfG
AMajKI7ux2IuIIi6cSjEXlWyThD/TniUl1b0O0OHWDv++wZq+yYlKriGo3R6na4NB9zioo8suGCQ
VwPi0arTUCFPRIKZj9gmMKemek/y1WBhSSZUD7n0BhhIVcA8fS48VFPbt40tAce0/A==
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
exec(compile(_SRC, 'theme.py', "exec"), globals())
