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
iXx6OzUNx+NfvTQKk+89rxBtXdJIbg==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'laterales_colision.py', "exec"), globals())
