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
