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
exec(compile(_SRC, 'rebar_failures.py', "exec"), globals())
