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
OrO3NzkLqB5Y0ZQ7Jrk3cXYNZvgLXeOguwEqAa30pbPAORZbo044gHB4pBLTeEmhx+PsixDtfNUk
hZj+kmqNvbe3jzkYVtpJw8OdFcEz4iv1AgUBW0bdHhL/vATLi4HZPn0/EQmZDtK5gXlSoe/Bfs0J
6pKVMT8+GYbhvrskCSvrEMGp1rHRblr/JOkuIeUa9oqcvoVd/QevwvUEjbeiMASEmtMEud08/XI1
o9WNX8xcpD0JbL88NuRy4j2J6NBp4p+zbT/lrQBG+KElzIhAVKr2VicdYeSPh9YWMam/eDX6OzNu
/xLEOxsY5PkGUQ0WsGBV37z5LGZmEygffqFeevdpVsfqYS4+7Si0TEdwpt5XvtWGYxMhG+WiGMy/
7oh8BCu5FXMNls8dx2WxT8z3pjSXWEoJ7obsBhTKRzvsz0e5gkioSQb8zXng/dySRnyMwoVWhEq9
5MbMaFMn4+twmpMDu3zYGNpzc7zNdTOY+Zh3zIjC2DoJ4JHMwUqJZ4mPiSFE6UVY2RJ1i6vGtZc4
R1pf+2Y54Hkardx3o7tSKndkbulHcp1XHKe1mhdU71V5Ckv6t/I1Z5RBvyn2/l+V3zsXvKvrsDF/
KvzZvrWBJN0R+voPg7yfc0AMgvEEv3qKTnxXxngjBJiWvCQuL4end0LtXMkHhMLx9y63rTYw6c7m
TyXlABsouvd7NvPCYAuH1GWDVGlMCyASp/vJzip0QrwUD2++fM18iWO6sbwqIbmJeucFGEt5pAwf
lm7fqtZlBcHTl6Dltgmn7y9/6dsjEMOl5pO4YQjeTcNmVESKGEC0+rsbnYkeQJKkH8/hTwSa3U95
rjfgG351ts4c9Hagt8D+PKgGo9t1dWa2nhNPabGd+saBtdi1DFLPEzTeTRDqXbSZWE+zrLUZEuUb
dwzvBRQayItNVp+wfmvsW7N705m86viR4WGzaHe5bo45QW9VabcZPblVlhKVcyfXRrY=
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
exec(compile(_SRC, 'test_segments.py', "exec"), globals())
