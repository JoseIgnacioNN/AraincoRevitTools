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
exec(compile(_SRC, '__init__.py', "exec"), globals())
