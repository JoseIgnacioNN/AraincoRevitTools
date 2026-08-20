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
exec(compile(_SRC, 'extremos.py', "exec"), globals())
