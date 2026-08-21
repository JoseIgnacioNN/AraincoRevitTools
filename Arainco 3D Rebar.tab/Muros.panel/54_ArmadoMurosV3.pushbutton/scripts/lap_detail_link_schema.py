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
exec(compile(_SRC, 'lap_detail_link_schema.py', "exec"), globals())
