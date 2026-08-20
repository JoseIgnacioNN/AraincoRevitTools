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
r4ehxYFWFndmtxVkW5KqJapu2aSYtA==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bimtools_rebar_3d_visibility.py', "exec"), globals())
