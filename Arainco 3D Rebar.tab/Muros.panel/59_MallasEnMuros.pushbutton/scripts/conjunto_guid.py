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
GbRsqzs=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'conjunto_guid.py', "exec"), globals())
