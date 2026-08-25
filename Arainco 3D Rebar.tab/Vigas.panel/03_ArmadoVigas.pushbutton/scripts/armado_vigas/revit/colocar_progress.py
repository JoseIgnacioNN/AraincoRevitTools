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
OrPfN7/qoB5E0YhFaLHgL1NCGComMHxmjKopBy/M9mGmTxRQAIkTFR9kLY4PKHNx5D3A0gm34FOf
HRf1GjnyHql0tS3MnZWd8xiPiMBZZt0ATHLhvNk9fTg/faiffdQlw/s4xJY+IgP1tTK+BTbSw3E/
yQB+4nZKuYFGDpOmrCE5bbOjnEAvn2G+BuqX8hrKgCi9HjPuKbJdc9/GyhzuqkvqLDBqQjY6EUoy
LXB9+6XNDgQsrVHuJuG/wfKx9pODF1KfgfYwWIdidefq8r+SYWzX3q6Gl7KK1f09e28mDpwogxYV
KbS57u2U7HxI1+x44EWVYh06R65qLWL8BoRDN0YRIoV3+hoCDlyrwa9akmvyoIrmdTZm0gctfHcr
lpH0xGPiojHTEAUMPCzgf6uefMzNaZS6RmEsXdzdqkTOqC0++HvXl9aTaj8XidaKRtNUS1N4F0Ti
bAsALpGgStgCWiIqfjgh4ZeznGVbB/Tu/SsFzm4yC9sMI6gATOrD9oHoZBouhD5U4MkMZzAhjh/k
MrKwqqPoefe+oM346R4Sa1S8qe3o4Aqi2aCvMUH/qDBlHk7Yd9NqfW4P9TqLLyuSLIKxZkC/xUtv
/nG7Sx49MVWZnnHv29KWuJRoZ9PAjShR5avEOqSNJfxfHm/0zyRsH79TPfgocy1mcHGMszgUAR23
NKi/DefptO4xgtiPy55ikmkxsL8GKVHTW7vlPWOWpKbFDprVzqXB9y7NEbOJV1Egm2XKGC60Ivn/
y6/4l7bJX3sL8R48znO6U9QsUseZCpZWdeOPlsBq/en04n7/2cW6+ss+Z7rI/2IKwtb8QfXPrMg+
2JsetHcfMIwMzMJcTpnQyJmTuJ8Evhg46CmDwBBDFBBDxBlyUoza4ipblSsZH+tG6ZH7jLEys6+l
H3qJHMQPhD19OuYYGKaoh8bN8xMOfQhnr+17cRrQsZMcrWgIkohdsXFdvaClk+oUXrN6/3jmNybb
DfgiezR6HNcXYUBUy1IWnjtczemMJc94KiDZQFJa5NttFEZguFAh4/Xvh5/Xmu7MK7we8vNkmbuZ
iHkqYqb2RJ4BGXejwcmrZTVl9QK3n4tOB2qKSjVDg/kPzdMsmollwtgBlOG41gQhsomFecHstdT3
JUSv+oIvvPKIpIhWvCiNpeMPhLVaZkU2k37LMczyIZB8Wv6LQP/y+tCtR6mJ195KgrNVx83vvcjm
eZnCnfG5DGMgxm5C
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'colocar_progress.py', "exec"), globals())
