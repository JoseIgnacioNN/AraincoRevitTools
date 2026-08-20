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
55MFwyMI0tqzuDgeSBuRA+wQb0kWlG4=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'colocar_laterales.py', "exec"), globals())
