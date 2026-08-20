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
exec(compile(_SRC, 'net_ui.py', "exec"), globals())
