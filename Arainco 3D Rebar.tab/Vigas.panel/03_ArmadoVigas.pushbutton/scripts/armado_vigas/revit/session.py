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
exec(compile(_SRC, 'session.py', "exec"), globals())
