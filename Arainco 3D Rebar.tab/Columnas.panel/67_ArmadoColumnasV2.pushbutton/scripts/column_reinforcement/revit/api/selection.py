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
    n = len(payload)
    out = bytearray(n)
    for i in range(n):
        out[i] = _biz_ord(payload[i]) ^ _biz_ord(key[i % klen])
    return out


def _biz_to_unicode_source(raw):
    # CPython3: bytes. CPython2: str-bytes. IronPython: str==unicode y zlib
    # mapea cada byte a un codepoint (p. ej. C3 A1 se ve como mojibake).
    if not isinstance(raw, type(u"")):
        return raw.decode("utf-8")
    try:
        return raw.encode("latin-1").decode("utf-8")
    except Exception:
        return raw


_K = _b64.b64decode("Qml6YXJkcy5Ub29sLlByb2QuT2JmdXNjYXRpb24udjE=")
_P = _b64.b64decode(
"""
OrP3MLMqsB5Y0ohHgiFrB0/MyeReVaVgc256/NaYKNtMKuE1/D8a8eEgt+O/JKHbRxCKFLiGgJBs
uIaJDlOzX/XpJIqLB9m2nQudPk/3Vj91KiWqPEttfx2mqfV33fjhXmZB12q1vSR9LSi+jHM5L2qG
zrTNYhsL5uHuxXlaPOFivWJpbPrzbJg+g1shjsae5SdaizIlThdJbFDpxo+Qd4duoiLLAojV5ZJO
x4/UsgGYoWJZ8t0gZdeYKnvW/wgqondAwDZQ2XHYb872qswXY0STHAzFf9heV2Bbl+TNvEPbxom0
QDvOkLH0w8HkYjn5fE0j4Ue1VXrDmLS2yodpJJ9HdqXazPWMNKg1eP24L1LnsGQNqZhnyvv5LBdQ
G7s0mKx36mEFS+0bNs6XoDbVBMDPCtDjEWUUlRnOOJrfaoz6wlO4NxUFS7SbENDIX0HK2NI=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
try:
    _SRC = _zlib.decompress(_P)
except Exception:
    try:
        _SRC = _zlib.decompress(bytes(_P))
    except Exception:
        _SRC = _zlib.decompress("".join(chr(b) for b in _P))
_SRC = _biz_to_unicode_source(_SRC)
exec(compile(_SRC, 'selection.py', "exec"), globals())
