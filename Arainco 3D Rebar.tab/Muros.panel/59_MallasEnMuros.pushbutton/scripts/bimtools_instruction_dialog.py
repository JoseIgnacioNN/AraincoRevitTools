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
4A==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bimtools_instruction_dialog.py', "exec"), globals())
