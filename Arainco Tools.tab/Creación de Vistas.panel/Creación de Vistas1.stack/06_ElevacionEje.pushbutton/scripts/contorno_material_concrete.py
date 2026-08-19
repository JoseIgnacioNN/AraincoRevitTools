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
exec(compile(_SRC, 'contorno_material_concrete.py', "exec"), globals())
