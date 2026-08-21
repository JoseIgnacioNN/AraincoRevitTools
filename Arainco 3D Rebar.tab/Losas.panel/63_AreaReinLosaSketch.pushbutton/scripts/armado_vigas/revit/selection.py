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
NaxA1liZ+QOi
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'selection.py', "exec"), globals())
