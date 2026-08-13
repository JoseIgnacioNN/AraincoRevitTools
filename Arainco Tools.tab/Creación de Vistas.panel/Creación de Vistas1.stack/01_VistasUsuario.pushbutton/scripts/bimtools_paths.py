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
exec(compile(_SRC, 'bimtools_paths.py', "exec"), globals())
