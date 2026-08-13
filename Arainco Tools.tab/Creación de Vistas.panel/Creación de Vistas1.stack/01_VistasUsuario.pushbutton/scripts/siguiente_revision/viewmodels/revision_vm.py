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
WRUKEwXBLEjYUtZ6B9IogCbW/ax9IpP5/Q==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'revision_vm.py', "exec"), globals())
