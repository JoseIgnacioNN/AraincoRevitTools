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
exec(compile(_SRC, 'bimtools_rebar_hook_lengths.py', "exec"), globals())
