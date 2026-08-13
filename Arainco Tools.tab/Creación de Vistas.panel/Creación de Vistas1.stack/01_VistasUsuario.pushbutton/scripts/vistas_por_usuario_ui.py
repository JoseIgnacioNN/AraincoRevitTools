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
OrMP8LMucB9E6YAVpPEsxbybW/O0rglzS/bMN1NIpRWZLwxSuUAax/mAzwPMkcnRLbGLDj5gWoVT
RJD9geGSNUQOefHa6C5jdJcdXEBr3riPn9Pm12dUkwnofmF0Cn4fSlty5o9PG8UdR27nR5Yq0UAj
9eB03H02A+oPWN4tP2v2cmiMAdRtT2yNYJ/Sas3SKxmZyhJQZzGnmlewOGgw6wA4FvPmI/dZhuTN
7JCDxl42sQzB19wRPi07yhHAGbDyKlHAXZ47EZfW4Bj0PIqrocv/Dd2/UFfRjKcxJjYtkVSy6E6G
Eo5b3IT/3lJY/YZpCalSRa/3utWnXrURcVod7bAK6maxfCCUUGXZIweR9i44H4pO6zaa/CBehDt0
TGXtOHDGJunRGnWqEC4bQ7fbjPvTXp/srGQ0jd0G
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'vistas_por_usuario_ui.py', "exec"), globals())
