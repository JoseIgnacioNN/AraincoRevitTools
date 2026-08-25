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
exec(compile(_SRC, 'bimtools_wpf_shell.py', "exec"), globals())
