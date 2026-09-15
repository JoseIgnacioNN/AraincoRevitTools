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
OrMH8z8KqB5E6phQ6xCvtyGDI2llL3ciYgA4zx78ABjOzesq8Vu5viCPscppeAhrO4IMJxeXN252
9i3RF1KpvZyokJx4PqxH4zIs5tO6EbE1ax4esQGMphYQKVSzREIBYkpi5e2g/3jH5GeeDk4qaOag
i6pCnkP3M9s9oLBgLQTRH8eG5MgnumKU6f7HfiPnZZpSXiHNadb6k7ui7Wpf5FhlGy8AbsnemeG1
1HJ74kp1EMrq+/JTlQe4To6ZF8ByupiP2ybHg6iUzeSuvWU0WKD2LhRAbYGWpksX8iUrpftbnVLV
cTdT8B9oYzQxbfHOCZCiUvwDx3gagYQcOMi4RLq5OoqFoIrbrLQ9CDHQe2Jhp6aitFqSd86XtsIA
P0CFldgGhSP3BiEhQizlW+16C54nd82aLDV+ZMPrOzIcHOYIkzHXS7+mGsuN00yKttuJvorSHVgi
ajWr7X+sD6IFEO8yxBE65Znii3EFHnW9wLpJCXU197a1M3xySrZprTiZnbnrDNqvAFZA8p/PgfcN
iAEt1wAEgrX81UwgnBS7zauFkAo3R6ci01sf2Z5qBr+d+vlL/WLUWY3myhFVN4QJzPpTpw==
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
exec(compile(_SRC, 'main_viewmodel.py', "exec"), globals())
