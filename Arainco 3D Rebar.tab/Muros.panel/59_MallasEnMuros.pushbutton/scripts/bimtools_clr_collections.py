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
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bimtools_clr_collections.py', "exec"), globals())
