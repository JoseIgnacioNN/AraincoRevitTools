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
