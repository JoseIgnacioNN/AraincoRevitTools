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
    out = [_biz_ord(payload[i]) ^ _biz_ord(key[i % klen]) for i in range(len(payload))]
    try:
        return bytes(bytearray(out))
    except Exception:
        return "".join(chr(v) for v in out)


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
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'vistas_por_categoria_ui.py', "exec"), globals())
