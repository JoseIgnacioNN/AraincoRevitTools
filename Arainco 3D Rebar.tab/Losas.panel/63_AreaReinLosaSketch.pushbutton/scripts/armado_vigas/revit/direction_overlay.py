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
exec(compile(_SRC, 'direction_overlay.py', "exec"), globals())
