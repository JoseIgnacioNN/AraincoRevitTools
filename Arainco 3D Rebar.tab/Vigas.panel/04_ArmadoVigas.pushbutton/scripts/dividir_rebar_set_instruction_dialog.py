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
exec(compile(_SRC, 'dividir_rebar_set_instruction_dialog.py', "exec"), globals())
