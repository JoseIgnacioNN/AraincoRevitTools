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
OrMH8L0qcR9E6IARpFksv6Ssi3qleGBXRbB5jfFNwUX5ObmvHgIoJsoXaBEjH2zTr7W/RhBC2CUi
mMymmsWKjwzm2hygYj6nkRZw0GWLZQohGlVvVxUWMjJrhwKPencIcwGfre4chIQ00BJtrd4nzYRA
W2tX7cJi0vD9V51ls6eiCAhkabQfei3PZnMTavCBzTKRvijypOCqhZkgseBlkMitApZE00lyF+eM
OQtwcUcP/uUsejbAU93Gs9uk+HdwFd34qgP06rTBGr+H2J6EtIOe0qGhTv/7EcYYZCxGkKpIYBaY
JPVBS3QKhfFJc1e3SIp2qr/XeDa1unu5tg/r+ZkBF1DTkaT7Ni4EFSbFrGVFPNxgHRkBBrBjeTle
tyIyCKWNCWE3V/sMHDBASI5LrcwaYZuZV5mlFCZ/0Zx5on72kX5SzNa/
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
exec(compile(_SRC, 'vistas_por_categoria_ui.py', "exec"), globals())
