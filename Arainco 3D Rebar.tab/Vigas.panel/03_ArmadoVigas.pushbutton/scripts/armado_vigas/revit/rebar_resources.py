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
OrO/NL8K6G5E0Zx4Dadw/umGFTcvYSRBEO4tAoLmuJzomWDnGnGra+8ku1Bgz9PYOt0t/4O6Y4SO
JMBzYuKaTXGat1aKeT/Ycn+k6KXLMUcpFJbVcBJHeUwlosPi17VeQ2D4gPeww+pvAa8xzNbbDMKA
oGmS1RPpYCdQJEryhyrkFHAWbvpkfE3nXIkWpJHoLZyLdnP/a74w8X0xM6JK9gv8tkBe6T9+i5X6
q7FNMWyUqyjkxOegbxJgjysIN7B36A64nec1FkIe6kIskQq1Xo7Qn2Pc3wMGHCUz5errKrtqzyRD
XMw/AAFkH9J2LP+hpteGm0RPe8G0AVHw6+UvBU1GNXyOfRBPtNvVfWpdjD1sYu5+y3nQ/0v1YHNq
aoa8Df1DaNR0qqRGoR3ImPNO4AiP3sOEKJjSbir068n+LLEp8THgx6ctJN1sbi1LuOjZMzzq5Aw9
QKotc4B8L13vQoPbI/Kpj1EGBXobWKsw+dznW8zkf9ZJr1D4YZr8yKVQzK9qGTKwdfC3JdZcgMJG
ZtMskuFptDQcvofTg/EPdv12X0dcP5jT286/4wcdcCd8z0BtgbSfGm9Iv9ti5UHRp4LvIDOC6EzS
9ozlCNGgNZ/8wCH8HQNqtYrKz7kFbjekA2JTpqYF8axCzbHpqJNEvoqeZU+wXme2SuCh8x/PsWMq
O/Cw33LMA7IIvoAM7P55M/lkBe2lESx+awBotjqBnRAlIwWa0kvbLFs9C33qa9hSBpsHAFKNnqba
2YOiG4+vmbH0Er071wlNvBWUCympN/uEfEcjGNtWMKO5vBSwJXBtxO54uPi9dcE05rZJEsZzaGsA
IVcp57t4mXjwS4oqIwrGJKrnq/ksEaAqNapk/5gp83oi/JOs7sAD0RkVo6ykvwelAKZKkn6/+Wve
/lPU48bGNpiYpoUYk94jw5+nvwIHojUtHfPL0HBblkYU6/7xk/TnZR7Is4mhctwaG7uuRtgH4Jf/
mzK2CYEQwugHlriQ4Hm4zPmNMguObz81vOz4FErLjjU/rB4b8DgX9SHgEgfUNdYgklw8AuW9
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'rebar_resources.py', "exec"), globals())
