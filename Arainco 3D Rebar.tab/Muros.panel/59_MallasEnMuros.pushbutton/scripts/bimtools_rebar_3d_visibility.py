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
exec(compile(_SRC, 'bimtools_rebar_3d_visibility.py', "exec"), globals())
