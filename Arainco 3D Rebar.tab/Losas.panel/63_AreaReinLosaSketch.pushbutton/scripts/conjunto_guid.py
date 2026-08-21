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
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'conjunto_guid.py', "exec"), globals())
