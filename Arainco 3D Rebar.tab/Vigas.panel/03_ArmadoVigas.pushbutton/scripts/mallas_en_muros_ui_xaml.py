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
dNJtS/8hDHj2V3CNS27MkG9uguLd
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'mallas_en_muros_ui_xaml.py', "exec"), globals())
