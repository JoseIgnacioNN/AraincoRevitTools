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
exec(compile(_SRC, 'instruction_dialog.py', "exec"), globals())
