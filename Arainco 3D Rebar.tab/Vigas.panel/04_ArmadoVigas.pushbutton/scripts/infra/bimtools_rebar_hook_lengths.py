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
rOdRFR69VCHVDoI8pc31jLUkoslRipxhHTN6wQ==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bimtools_rebar_hook_lengths.py', "exec"), globals())
