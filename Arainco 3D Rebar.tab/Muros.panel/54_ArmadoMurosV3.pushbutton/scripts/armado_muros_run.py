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
DJN8Dr7QBQirb3miSpE=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'armado_muros_run.py', "exec"), globals())
