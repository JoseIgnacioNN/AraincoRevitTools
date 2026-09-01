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
OrOXOb8KkBhE0YRFprRQeOp+W19yW3HXuRDqbbdr5lX5Lvj3tGoECwZKxkjg+Z3fK56owMGpxdSU
750mVG0kpy5QpP5U89GXEH1aX/GcDpMi2nFZnpPzhtvkA1EPo+0svik188/ZzsqMVUxw4vXCw2Qo
+NdQO5cRGNsxNWVJntNg1cl5wz5r7x6uwaw2ATs/Kp6sAgWtmpEcUfzfN7zqJPGZwlQXkITZu4C1
ZIIB8JJOz6mAbLVC2Qa/ZiWItOssf6bowumdZBRsFRKu/EcSiTT+LFms5yAK/Fyd6U1cRDZH8XP1
3micTkJWsQ0nmupq134N/C3YN4ISDu+Qd4Kc9AKM2Nc/jXh8FGLFrGnkAoSkDvW2JX6iPGlRjw+o
Z/r2qTtSDEs0K7A/bEza3ulbn81nyiwAl30MhTjyZUiWfEihSb+t2hG64rzFQuvP9XqqS8t70XLt
9EY/8bkjOFq9t+JxoQY4m1UIuJeUA5Eb4BKSfBgxZe0n6jQEJILlZQGEk2Ky/DRyunyKPo68P36/
h02YopdvykvdHQGv4OCxic2oneh8wAKcJjLNRnFsFq1o69nIvGAxCIdpPTdlBakXY38U+Uj9Qnrf
RrQesI011j8EG53bAwMQ5qP+ZFT+1wijZcVegi5OmhYb/u7fHxNRhoAzv/tcCfqX6UUKs7CEEcfu
OBuWUPECI/wqF8VSF18kEcNmP5j97sb3tvUfYA/7DeRfakC7HEz4NXC4qlWFro9eT6JHcxxr8EsA
t6Qt4AynkBMp584+TuKkkskul/Enynn3OIT95PS6xYwX0Pyccg9Xv60V4r8by4xPluHELAWG0FU+
96vtpqbtY7vjP3v2kTAw9rhYtgn5ArhCAeTeRz1xHGjPd+tIwzu7MN1cteYhFddwjNHHxWyMghzU
uH6pE1+UvnSp8dqi6LT7PQyawrFrZHq1Ve2u8WCB/S63WKXQcHBgGVRJwznE+yg4zOzy+O2NT5Cg
Hg2IAkXA9uRCB5SlbUiO4ijpZcA+C8J5wZMM6x6OtUV0oSzqgGrRIAszz/Ic0zqu3a9PngDdeCfj
XNz+EdfszwgYsXZ8yTa5OdYso4IedMNjWP8f9m/MST638PIVYHBu4IpiRXMjioFFMHUhcsZo2nsN
KXgPy7u06zk5l+TNcW91La4nlxII+ZcuNQMikQaKlmwS5wmVxKngyC98QcTMkgJ8TqeNjOvCMW1G
dl1/LpnikEzxl4sJkPYEfQYvQ+/sm94yE1iQBRcuOi3TXyrp+Bf6sbwJdZgZfrGl5TmjzNXt1H+O
ws5qRCuAduO259Zp60kVa1rUr0VzQtSEF5NMCrnYuKMX0k8iWhtsIEBzDx5AMolJOc7Gi1Ugecmu
ZFi3TCgQbtKWDTEbspR3pDtvoFgli6TB+couGWjPeIvszkaRUHZ+PMO3CN4eZAtVJ8CO2FgJcHLt
eF18hpyI4wTmiUoH42KnESUpemP1vkRjQCP5GPP9LcWkgy4V54AAIRPRqSjyLgBUEKo8l+s+e/SS
t5EDPQwBtoMrTITiOrHGTJ+JrW1mAyea6Jwz6hhejwmH27UvVlo3lj2Es/xOkO+BiH7M/jBu1EpX
l4qavjuSfH7fweRQnftWqe3cOGUctJxQg3CLCGziqsld2NxM9yoeeH6SaFFEBzwCrQzSDLO5FnWF
fREXexWMD8TUbInt3ul7r801JCABOx5hBI4CDitKfulciV97abg3AiMuI2yAhstH7EkPgGSZpjeK
hYPV8TCK5ilQ/K9rm/BZqrMsqVjpXmDPopMKjkeG/2PuLBbVacvDaCqodUWZ6JFNr+TJEg==
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
# type(u"") = unicode en Py2/IronPython, str en Py3. No usar isinstance(..., str):
# en IronPython zlib devuelve str (bytes UTF-8) y compile() lo leeria como cp1252.
if not isinstance(_SRC, type(u"")):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'export_laminas_strategies.py', "exec"), globals())
