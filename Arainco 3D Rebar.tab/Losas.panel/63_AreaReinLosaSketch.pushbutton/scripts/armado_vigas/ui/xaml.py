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
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
7A==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'xaml.py', "exec"), globals())
