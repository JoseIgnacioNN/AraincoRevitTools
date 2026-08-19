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
exec(compile(_SRC, 'elevacion_eje_collect.py', "exec"), globals())
