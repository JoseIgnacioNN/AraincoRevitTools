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
OrP3MbkKcR9Y04A7arIwxDng510F7P5UOVAYi1Z7jcZPYb4qxpGyKZvjCIC7y6xIU6WoflgpwAa7
ak6szLxqWwBGiqhzj1W+cF2Ssv7HW0rNoBxzU8JtP05Y+Ag6ltaxUlBpo6BOOGjbjOb7fl8iv14i
SMBk4s6AkYLyBPoyrmuN4Ot/zhP+jdM0aejozlhDxh6g85WNfMHXiC5ngzMRspbdJfnXxIUapYpA
xp8yEKVC+Yj+ECu2O6ZPuLaNaIbVNCWDYgu70qw=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, '__init__.py', "exec"), globals())
