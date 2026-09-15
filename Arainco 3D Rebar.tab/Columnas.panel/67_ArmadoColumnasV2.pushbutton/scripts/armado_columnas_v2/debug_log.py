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
OrPXMrMKpx5E0pRHaI45PfZaDiCmPwgjwzR4tUwDlLVW0ItFA/Grv7k/xcbRgY2Pd9Qs7tWsB1Tc
IGFwDOGO/Nle6ZBwtuA6Xkz6wniAiV0ZNXPuFcuKkwQP2G7nDuOZtqplhicFk8+quzSNNIFT8zaK
Yfdw2DGObDBWMas3Gj3Kr9l08mnHaDYMXY5N/MbJ4RJgGbkuvxAxGQ7DSdQp+2vKY3pvokK4O6DZ
t12azen6vc9UM8+VT54OmEaR2IjKJZth9gempLdi2WLrplHf0pPkZ3+LeNpP6LCgeflA/hFzw1tI
A/gjtHiMaPfqm8wluwIQ3U1FGC+kF39cqArgcRR3TiSHKRRS2SHWMOfUjuqUegZgS09USS00Y0C4
NdUAHGSIkxnPT8hXFttDHNGVfnnIQP+Z8o9oOSzFZ21a20ziE9DKleVOmxZak5X0Fgi7Frkudh2x
n0vl2m/NkQmo6tl0Ws0D7N0yffeW0ETBP6Eq8nskLu/XuKhjP6Apu5Gj7Yk6kH6R21FUcfhf7b75
TLh6qG5UIz7rHIlaL4Yoq7dP5f+6yV1TurPTB1dseZ/MBEmucaMFcUz2vwyTdlxBrOz81aAGzNDH
IeqWexJWk8/IeQgLrE9oHQlgc/OnkJttdh9IYKZKT0O1Zcq6MRWRX+I5i5c1hq4fgZtrVGFJtXcP
MWedf4PtoIQw4wNF0g==
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
exec(compile(_SRC, 'debug_log.py', "exec"), globals())
