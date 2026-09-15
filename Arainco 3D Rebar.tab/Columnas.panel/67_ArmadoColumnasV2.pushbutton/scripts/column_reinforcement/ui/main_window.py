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
OrPfN7kWqBhAspxHHrr0BQFYlCtBfirVU+4x2Sh87FXH6LJKaXV7hlRHNvO9X5rF+85QiWON5StG
6htO/5bZsjF2uk6WjqQFRsTrp+knVty/hXC+LEVrb0c4Vl2cV0d0M9Md9x1gk4VwXMKIzDTt+5hT
9jr3Zqr89wSJWSyNVtGrNxMq5s0tTfXsbxOMIeBDBSVVBmqUE8Kg9N2HNuffV2t1779/4kz+yEdi
8WTPyyOw/xNRhZGyNcrgxSGobDw7XoqHnzTB/+XyX/VRuqJbeLQ8WkzFTKDYDSvDEg+HLjMJKW0w
EEHBegZiMwK73lMqyP/OoVcgemtyzfnblACrff68yAiGc5niySF/NFMqZX8YUstvKdghobGXFyR9
ajXcUTPuz9wDgNuXrLLm8NzlcI1EspeIAzbcrZtrF+BPGfjvljmTAPMOxLBSpjAo+6r4QNVz7XHG
PFCFzOtYoALUB/lBVeI912H2zvayYUkST5kkh1pSnsjhc2V1NKKAEars7HaUiLW/3JZqdwQIHgYj
SLVd24fvLxJuasQjjXCscofGY79ziOiKjELxAlTi00yOCv7iUHcjhjD1BETelOyEqi/lgW5buRFv
BFWjG8SNSi6n2tSAVAGdJQWU2hOablSWGXke/l9+y/zeM77gqwhJR/GZEynluMcXq+Xv0g6gHXt5
ej9m0bp/TIkMCFxN6thEQ0ftxrM/e7PArhXstjjSicSmP00j/tcBqyfxRrxQY+N4lv6OKYJVPZ8J
jV4dxWQvF6Vva30kxTc80OWfDM9xoc1hTYf0LjkgYhZSTTsxykarR/jate2YgBWhhRvqoyjY4O35
cWsKFAMNLThCN7BWJZc0ZkzqUmNwQKTfpxzqE/ldjI+F03jJDuGxye6sL9crxjFpsy4T6SNGG0mF
6JqgHHVdxR2R2zdkqjg3CLi7DrxsD6L7kDHApwPW34j1elJYIASka82kR3MkZ6iMc9fbuQJ8zABU
/+00ZBXyuRRNjB0FCaZxx3lXDgzLvQ9baSJlW3dt+at1sg2Y87g0c06fgyF4niBwg5e9Lt5M9yIW
hYSMyWUPWYGm2IrrwMNSn/Mj/R4gI4TcCH4oAW/7IVsOd9ElmQgS79w7DLh9DVU6/VOStIqMjypN
GP0sXCudivgFs1JGDv2b/oSVovSsu08YGLJNZGb9ElAPlOGz8tt/SNEZYLzQhFdmdsx5yA/rwFsg
37zkGPa67mR9swcd1VXlVQZ4uqhrn0ZWkh7hh25cVUiYGLmHXRRobbtGVtNBloAXWnkNvjWd9GCe
NMs4cCp1qyGs1KPFBB+a4C0B/x2PRWQzD7R1LCuDa8Zk2+GVGV2xOuKiqeta4NMyvc7D1s7KEiH0
mJAp+AXiWCmSWlmcLwuU8IM/7P5Pge5Q6agUqFGCzhOW2pqfYm3JQ1on3SVF+g1PvVGR5hK3fXGq
j1FfFdft
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
exec(compile(_SRC, 'main_window.py', "exec"), globals())
