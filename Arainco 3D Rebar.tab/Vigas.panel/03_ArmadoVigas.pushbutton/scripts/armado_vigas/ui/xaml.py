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
Md8ExwldfOfo
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'xaml.py', "exec"), globals())
