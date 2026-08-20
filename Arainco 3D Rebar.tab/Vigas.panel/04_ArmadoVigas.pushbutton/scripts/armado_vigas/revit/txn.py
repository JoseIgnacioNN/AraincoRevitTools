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
OrOvNjkKlxhEsoR4q540pRBFzWNkcXXf0bP6C/vrrmgRBfO81CNRO8bdslVuXuJ0gCJ2oLpRVCrw
jjhQ1dPvMATppwvAnIvEC5JyAJCPMkVXALlFNuyuhUbPe8M4hZZ7E8gT+wBN66imn/6xXFOdERzZ
rHlO3NETx37zJtbLDZuhSaIh4PyJopVcC8NKBWDyrrxqa5qCpLzZ20UpEvs9zCDsClpOt+k5aMjp
rQcEcBFzt8OAb51QE3c9oFOx/IIiptVS2HZOY72JyMMg17cNauYo2N61MjsDuwiXaFyAPRi6Xx0K
Ok1wpKwI+v03g+0NQzUQ/qkkDHhFvfktRCvGLUjHcQjNxHFokqHHrWCZcGe9e+kw6mUz7TNtQXp5
KCG0gQG+R6XTofAhBHc/RF4hiHl8NMVQ4RE4y7znnXLBoBxaXPzHpcCtusG+Uw/KChgskDOCER+y
J3NlbbJuRyF4ufqWGxVgZE1oSYV4l5b7ad6SrLVpszWnDrE9Q+l4Phj8PBx0U4KlDo5+SAyt8WB+
NHq7h/qsemvX5UJfTYhw8hx9FRWp06ngSrxYG7y4jjqExRUKqfTcRXu3lD8kS0Oq+a9kSXNG43Cr
iOkmABM4I5ZqkuTiDbId/RNXplp59CagcPafoQm1d+7HGzRXh/zJexQvLtckrpKtC3m8pe1fqYze
MAPZcTuqKcGrE7cdMxPoqMKFuI9lBiT3yb3sG3RBq5k8zvHioh5PagnRWV4ru2X6Cx9I/nr1RR3P
MUoolC7xfUtvCC649O0NGTxlbf3RWqu3B1vA/PHv4L/posBM/F9w38tI8jFTiWpUe85XR3Vi/w19
SorhKW+/ZQgvPO1TNlPIoAg80fWrDgRipnyDWYNavavs/R7rzN7VyWo4kzqXqpzUb+062qQRJn27
F6vTxMRjvp/uAIV/iliPPOjoqpro/WpDOjFDr93s+pmh2hd70peLdYyfyibbsIGChR/KcO6XrMfE
B3T+RvzByHSSd7yzGipIiCEBTM9geUFeTj+ZG9cRNdrRagWG6Pp5o1n4DdTMsLPcSP54lxfdLJM4
w6RwZoH0GkDoOLcjx7vjLgwyuifC99+cL8SRR6RanfsyoiAFrRsI3hMnFIYeDdiDvSUcx3iuSPpf
qSx3Ck9SmqeL9Yq6uRE/3ilMvjMemXyQMalGgQx1D23g/6ydBxJH3gNRpcusH/x3HYBpAdTZ8n/s
/LYMq7quaa+3B2U10B2Yk7sOjs6Dw7sQ5Gr0Py97NAjJZwCQiN7wCNBTiWKD073lJSY0ba2y+x7O
p1KwZIp1nDsheXo0CnHwr3kUL6AleFm1qpJtgwFlPciY3UQoqYpbgTn6WLCztDxUH4JBAgaB0yeu
f2eu7KoCgeGN6bP9h7f8oXz5Vi9Zwe5hu7Zqth2VjQHbHIWgFaeJw4SVDhuJjG3JyqVdHaudfw1i
eIl/GiggnrnA1mpSTiKdv0njoXV4ekBHG2Zyp3wkQlB6EfxunMHuFb3mtwTK+n5Qxb9miGNitZgI
z3gFuC6Tx1XUnZQRvATOtwLS+qAmrB0ON04M310gBT4FyFQ3ptGw1tfTo7s0+gxsnsiQiKI2A4ch
e07EiB9cGKGTNjbLNxCOvs9Gj4OgGY+VlKQyaVWCwSUu17UdryoDkh1UJphj5kSpWpq2rRAbqgyG
C/oHlf3P5GMOndHCuihFewEfjG0OtaCwQ9W+ZrVdPqk=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'txn.py', "exec"), globals())
