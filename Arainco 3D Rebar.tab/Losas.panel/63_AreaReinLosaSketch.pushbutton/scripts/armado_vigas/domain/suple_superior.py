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
AKe3D4WDKufETZGsDuSbdIG0hLA=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'suple_superior.py', "exec"), globals())
