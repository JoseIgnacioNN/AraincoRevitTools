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
ajdAmWVnZkJD
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bootstrap_paths.py', "exec"), globals())
