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
