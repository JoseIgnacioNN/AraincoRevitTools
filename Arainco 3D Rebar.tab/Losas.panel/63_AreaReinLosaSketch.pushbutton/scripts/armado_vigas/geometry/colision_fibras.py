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
OrP3M0nqYG9EopR4C8Oq+YkLDiBFGMPh8Xz4Trp9rWxx/a5+2VtltSVnQSxmnGOZQrLL+HgOKowm
P05cO9OcZJcSh2vEj162PFbjcfpiHrAqN/o6SqHIi1mHhIBuAPCgCXxqW8JgcHFzo5+L1V+C8xkq
3No2l3vaj1tRYx4teBTB54gv25OXRhDPb/Gh5pko7rAFALKRhNJFgCXOKmOpe2B2CW65ePBrLJx1
0PGQawXHPEJQr2k+x4rB4MGUDmtDC7zuS/+zGHd/r3Fn7FfT983vhooSijWa+1Y1L9Lgk8flwFSj
OoWFeTZ8dIKCqRi4wH8M6HjBbcmVVsTua7p8aXzhHlBuZesSjsdgUEyOb/ja99LUWmCSvdC1HdjT
aJsJUmOBpfdy5vxO8vxSuaDLQSor6XtytsypB7iL5DzkpDFtiexrlrwbhMO6WUG0Sr0QCJ2BgP1E
IVrGuD4LT19C0LH1ye0s7zQDHPvMV36CbjeuAHDHLHodp9gXE7rbtT//qxKv5/cK2+8gkqT25ks7
W15XyTUrEBtJ/mrJl7p2SaNafYi9GF4OkTgrdHwdoEqHVb6WNBsKRTYLgZYOv9zY8vvJncNMR59d
tlF9lAfBbasrLmiviXJAB8SKa+o9eTg=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'colision_fibras.py', "exec"), globals())
