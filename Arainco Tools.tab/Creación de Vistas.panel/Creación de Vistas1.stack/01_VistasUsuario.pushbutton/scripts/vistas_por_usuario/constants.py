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
OrPX9L/qqBhE74AWpNGEDXNwjNkmpbEjb/hFeCRKL6H2YOhqs/GbI1fmOykc2WNrrG27e5x3ypbb
G/tIh37kgrRPtMiEpaFeMlecoaF/kp3TZWc3Q5Gucc2tNJGPCujeuejou0INy6g7MJHquVKtW9CE
zHkLLFR3LW4YuN98qk7XOQs8AzyeCz+7x/eEPqbL3sILCYCgLS2GnLOcGLMcsZD0T9eR8PFGhN7u
2j2C0jvyaGeLHzQ/kQlpbEvmmH8NR4JVzxv14kmwlrmb8DxhLTPP1b6iMAybBuztdudTGSSnKO6n
I3aRVxXN94pPAH5RyhNAX0JMmmbV/ONXgoM/9sZFr5NKMUgK1H4wzLYvKtwi+rQAuYeNO8uS34xe
/yAqI4XkIiqX9WVHOp/6W0Xwish8VZOypdSM3Ch+W0+15jcTcRWa+Reg90Yu33b8yjv3YKyYCIxK
m8UwCi9ae9QyQv/UMzfWXJMUhjBERwW2ypnCaXlBxECqJ5CfeCqXeG4PGMfCANBaQTY18dikVw+y
cf+Vg+rcM9H5VppoSU0ZiKDXL1v97Elnkj0ECTwAUQva+fTsfVO2jafHbu0sYY+trQEMoTvmLPLv
FzS/IyINzXi8uQUcwKrNn4IZITHqN1vzD3Ppo/d73VzYlfxr1uEMhhTVhH3UvYvNpxM3Lq5r+yJw
6eHsFEnRwfkRbpWOC39n88mXfJKBew0YYoCjlcLeN+gA4zhwz3asIeuAQGC07A4rDCgOOnNeiHmo
t5DF0cSNILplWR4shQyYc+bzPCKl695K4RDICgTJChi0IJAWkaLICl/ktnKc1EahHDjJiJY+nKSp
Y+826doHAB4j8AIQwbn6Yz1flJLEjlJWvbvDL0ByfTi9YHLVFY/MNy2lFdsVkZoNmZHS/kBMg0Pi
OVc+ZhbiFpYJbGU/2km92CheiYRQHgc+53zP9jX8YzfEeQ4AatkVwL2jOeBiTgCIrvG70LtVu07I
lSG8xWIqckDCVBIZCJlteGnJAaG97/WgIByI/0RlCFQugcdYj41YqLbinoc+8dl8f3wIMG+38lY8
Y15MAxXxZnptwRBjCBbmswZitsw6eJjqGp5b/pgfS1pfQSZMPfTBT4nG4vAIgrxyGyUz9ULrdI9R
f1ek2gc9BiPau4oF/kslRMGtRE5H25GFkWdkZFBa
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'constants.py', "exec"), globals())
