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
exec(compile(_SRC, 'sheet_service.py', "exec"), globals())
