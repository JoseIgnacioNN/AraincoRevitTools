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
CNd/jICGStg8Ckw=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'seleccion_caras_elemento.py', "exec"), globals())
