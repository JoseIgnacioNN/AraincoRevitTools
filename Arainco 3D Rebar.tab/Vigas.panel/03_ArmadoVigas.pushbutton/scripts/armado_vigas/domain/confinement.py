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
exec(compile(_SRC, 'confinement.py', "exec"), globals())
