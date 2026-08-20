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
UeArAxU=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'armado_muros_rebar_layout.py', "exec"), globals())
