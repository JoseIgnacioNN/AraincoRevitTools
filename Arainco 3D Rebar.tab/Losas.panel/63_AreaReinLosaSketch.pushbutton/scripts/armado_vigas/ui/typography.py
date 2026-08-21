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
OrMX8bMqsW5E6YCXOkPKPzd+dsPv/NY1QN0BJ2HwvuFsDOkX1YodJ4oXZzPu2/toi/C0RFr88HMC
/P6yirG6xr2azcAJtva6ZsqcJ4aN0FS8B+23DDBFpmO1wOUc/cocFuPKxeCWlW0f64KKOwKmRNDe
jWKLeWaxIxQjo2Z4XHaZCvBxH2wMiR5UBlxbOjd5Xg/aEXynqWSKEDwmja3VDW3NXaMWfrxAS0mm
gN9Z3I/GWfn1IKSLDRVkkhI4N2kItzUxpjNWur+e5c+/UqkCgiemWjiYu8T4Fjutrl4U0MM8uwbb
yMTs2y0uo+59+CKk5g0Fjxr9pZ9FrnqmVqfVTkobl2neKBiz9T6DM1zOfnLV193t0GUwOpskAvoP
CNsWrMpemXF+n291oQqmI008XtzFWe83Ke17vDjmw41OZCwD13v2AvYQ
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'typography.py', "exec"), globals())
