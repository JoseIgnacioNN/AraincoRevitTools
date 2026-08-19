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
exec(compile(_SRC, 'contorno_hormigon_vista_ui.py', "exec"), globals())
