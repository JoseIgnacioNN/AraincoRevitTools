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
exec(compile(_SRC, 'run.py', "exec"), globals())
