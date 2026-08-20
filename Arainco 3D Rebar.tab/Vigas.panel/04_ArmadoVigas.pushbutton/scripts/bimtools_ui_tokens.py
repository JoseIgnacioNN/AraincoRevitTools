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
OrMn873q6B5EqYBQq7UutS5HfuNK2KegueytRXYq0OC0C8toySY+7T4mmbx9T9Ag3GH9hDnjhQGo
y4b0n+IOJyqQG+iEqdIQaRMu+92toIDsq340q7qZt8OT09L66dx0zaRpWyaTptCIz/hTUUFn3g60
LOlYVAtJiqfXLoHYu1IfH7QSCxPf8CtbBkAFHNa9Rept2CMqomaGoSo8z2yWcIusdc9bgK+mkPZe
SXqEMRBi3c08X+xUK+OvlrY4ZqMAOXfVIWedMynojA3fe4elQfwYiGGYMvjjkxzFkjkTniefATKy
7EqfNTZ8cBsCVUsVc1gGtOiINVbujB4/UjYTf6qjk4Q/8meHih1nFBjY9le6BKcnR6Yx7kDqYJ3V
6eXdKBzy2oR8bfeIIQ9pQIAPrgL4kOCd1Gj5TfLnDF3exXY/rwB5ZbNp+brihLPbZON1fM09SGN9
S8KoSSnyCAzwXQ6PKwd1eDx0SubYBrMwt+Igu5tTrne1AliU2MEPBfNxre1286+UhPMpdcW1obsn
3EBgzu9ZueEv+5JOzFWvCINExVcKOs6qGB8NUffp6M8XVsoUEW7cZoe+XB0K6cvHnGRQmZxQYU+q
dnXpreQDJNqRwUU5yYwJg/4=
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'bimtools_ui_tokens.py', "exec"), globals())
