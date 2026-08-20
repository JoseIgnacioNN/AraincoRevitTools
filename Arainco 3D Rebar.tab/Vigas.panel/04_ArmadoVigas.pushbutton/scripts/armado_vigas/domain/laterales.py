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
exec(compile(_SRC, 'laterales.py', "exec"), globals())
