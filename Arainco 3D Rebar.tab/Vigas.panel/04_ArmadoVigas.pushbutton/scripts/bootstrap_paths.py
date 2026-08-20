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
OrMP70kKsB5Y6RjxrABlVOwdWGNcYeoQZy7fg/iuUjtWHIVJ3Fd3aG1+kI8NbxJXaR78xZIG9smK
hJA74puA93NS3+hwHZSBtG3WblDRA6PwHr+OZOMqSwptLod86T8n6qwD575mrDzLhNJmua6g56KC
jXN2otk2dzDJt+wkLxkrOqnNtg38M9rTvfGvZh15Hry8RHkvy/F1qaQ1EInc+6vGqnrxjDFGyxMu
CyEYCWmfIT/RuAXFVG4vwBwd
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bootstrap_paths.py', "exec"), globals())
