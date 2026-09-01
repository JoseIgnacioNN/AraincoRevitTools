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
6yWnkjS6sNsyULb4on7CPniK
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'export_laminas_naming_schema.py', "exec"), globals())
