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
