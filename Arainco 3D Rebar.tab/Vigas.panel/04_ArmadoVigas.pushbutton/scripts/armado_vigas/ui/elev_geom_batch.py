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
exec(compile(_SRC, 'elev_geom_batch.py', "exec"), globals())
