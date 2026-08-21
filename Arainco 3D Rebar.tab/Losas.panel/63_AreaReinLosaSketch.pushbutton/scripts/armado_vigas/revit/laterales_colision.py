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
exec(compile(_SRC, 'laterales_colision.py', "exec"), globals())
