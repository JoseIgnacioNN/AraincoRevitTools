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
WF3ccDgFdm8hor3RRVlP4w==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'adapters.py', "exec"), globals())
