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
exec(compile(_SRC, 'bar_ends.py', "exec"), globals())
