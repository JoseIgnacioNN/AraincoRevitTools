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
exec(compile(_SRC, 'armado_muros_run.py', "exec"), globals())
