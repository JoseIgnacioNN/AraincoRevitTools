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
exec(compile(_SRC, 'elevacion_eje_beam_tags.py', "exec"), globals())
