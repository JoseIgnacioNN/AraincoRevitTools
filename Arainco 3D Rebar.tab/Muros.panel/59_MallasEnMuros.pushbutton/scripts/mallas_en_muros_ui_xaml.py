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
IxjjbhM7wrCZdVqtsoY=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'mallas_en_muros_ui_xaml.py', "exec"), globals())
