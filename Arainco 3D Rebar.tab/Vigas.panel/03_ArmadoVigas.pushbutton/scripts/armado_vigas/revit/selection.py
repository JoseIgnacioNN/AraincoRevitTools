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
exec(compile(_SRC, 'selection.py', "exec"), globals())
