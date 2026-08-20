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
qMXL0XeR7P5g
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'xaml.py', "exec"), globals())
