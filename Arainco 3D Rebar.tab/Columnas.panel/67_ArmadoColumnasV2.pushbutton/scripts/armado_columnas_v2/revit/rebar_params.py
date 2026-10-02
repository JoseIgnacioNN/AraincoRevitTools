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
OrPHN78KkBhE0YRFKLk3AXmPostulfCwGcNvfU1mrTE5yxPoBL0hip5wzBWyLeQS1SAGK/DZW+ng
GRz+KuAKwo9m/ZJ/OAxVWdZF1BKM8tyfxdCeLAcGFpQ8nPihC8VrW0FiERtI5hfzr0tQn02n/w9o
RGfzp8O3xIul+h6L0jfcFrZxUC2FdMEvp7J392tOpGXb1DaFUFQ9vU0mZJhxdWxCWPVOFJi+XnpL
Tr/aGDX+ejQZAHv2qTy9zTICuhKC2W2KkASYCJVZS/PKSvlb+N30IQNn6muJFIJ3x34aml2WlSYQ
ses/qhmG0MzSoqsvqloBx2Rf04FtF0iv3q3RsPwCTqVZTsEwQ9FoxQIUqP7CUXuhVj2LXU0slB20
Jwf715WpQFhdS3BgzuXdj10xdTm10M3xq95jpoqgNgtPY29aiunwpcUxCnB0/GbuDetOiw7kKo/F
JyuRzac0jW2s+/0LWiXocL862A8TJT0LoKkS8OWtSIYDeduf8avxtuwMFN2kALtrQabdmnRKRo2G
s6bCFky7XOPeIaGjTEduNm5tKqqQxsEJGdN4CXqRObUJd1gtcEnEYVc7b3bHXr0w71OEMXk+p3ZU
jH15IXW4QfuU800cDuK1yAfk1JXyGJuY5NvZ56cZ+rxr+94F0kkfJH97oRYKbcL+GXAh2LUhvba3
lF0TyZgGa7o4oMdVosoTPhAj1VWqGlIywXyfOJw4YsuCuZvZbch0RpYPQRn8aArbFJ9+chyn7ufG
0kGqWmaSLBLapx76wjNi+8rfJEONOJLqwb4AAsQuQs0cM7abCsWf7CPHldRh4kczeCUlAe+Xe7Zq
zxYP3f7skAi5U8Aj8l3ZU5mQycFx/BbJrrPFgvYzDxW6oGLb01VZpNtpwHMLsigEEIBNKOV3etME
OXZf5QdiM5T3rClIJKD3PvoXBXhe2nfnFnYOfnji44gWEsLborfy0TofHgAUKxozxe3fCnuJm7wB
fc8zTXOTLQxDtVrYeYfkiQWXGiAbQHG2eTjVD5uvZ/z76AWrjICHlXCWUD61FZHkOz1KwNmdBpDD
7pYY2HiGR2VCbuCL0e3QN9Gl//kzUd7JVrAiCTpdMmGF+POLh/Ixq6w2gqoXj1Hl7gmb4VsTl00D
iqyKPz0+QUY3xknHjkhrFe2DONDG0VUajrna
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
exec(compile(_SRC, 'rebar_params.py', "exec"), globals())
