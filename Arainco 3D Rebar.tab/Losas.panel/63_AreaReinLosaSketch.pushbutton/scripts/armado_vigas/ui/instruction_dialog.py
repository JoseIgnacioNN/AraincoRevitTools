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
ZZALISf9Ifm+8TEfAvIQbPgsGxY=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'instruction_dialog.py', "exec"), globals())
