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
exec(compile(_SRC, 'elevacion_eje_wall_tags.py', "exec"), globals())
