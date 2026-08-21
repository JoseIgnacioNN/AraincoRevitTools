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
OrPvNznroB5E0Yg7drXgzwhXErIsSCe5z8C7bXxzzp3SjKzyEICru8l5OTlRRXWBlOlXuQvv9LAj
O8CI9akRj1umSROHmPkCYXgFGWNmJx6R+NJGxaQbBUqyXKWkG0WkgiYq9HOW5gd6rjde1QHzPijg
+UpG3DZwc17EZ2lMt7t9Vi9XNCc6fEwyfQpk5DXYBuuj+o3nelb3eCRfZ7j+Q+9kICh6mp2JW9zg
v9WJ7rG4Y5Cz/ayel7KJKoELWMVrw8TSw8dWBti9bTcF3XRpmTI3GYi8oBFARdNxx3t0pxGJ4D8w
QvFX5wZdBKX6y4iDSONMWLP+HNPNxgJwf6n0YWj90kk/SbvMJIHMqZEzcXnOGgQIEk37IvjgfUv6
ticb+kxpk4rB/gZkP6OYP0Xx64+oFD52prYwomvrGJUYboG4237uBHUilQ1MQzDGbCDHRqD7Ce/M
SYNwhglZvme1a30AMcQUeC964yO1RAoizz+tbdNhpGUpP3llSWJEdBTcm6O/2jMqxB1+Y72t/xgd
rizYezLdSx1IgQ8hkxWD1tq3qZX4EG+a3YSg+9aBCxotOluS09Rc/dQ0Lw6uYlsMsSj/7IRD1umV
DCADdje+k1b4PeH1fTLSf/gpMcUj47it+nI0/ftiptVJBQ9HhGGRA4nSUTmSHVZX2cf0/E3Ri+BT
maX3KVmTY0mDiYju7Y3ci8Xnb2EGh+PyN8XtaVTk1++a36ccMw+DEaCoZhqGsmeC1AtZAH9/JzKM
6FzWH7MrgesI0W7qQsWfLnQc+5YTy1NINvinJ1O42kghBQXgyFIR1kKIQmvBpXjkFcEDNuWpgLdz
xliXn4F5+VvExXL5rFokeNmra93cja+PZjG3dbAmRdbkpfmToHp6y1im4Egzwe01Cke7kluwPnAV
Gbg0DcZSABnXAQBRuXtYuiig1Ii/vt3MouJLPkXFwBTzmnoXOQlXzZZcaH4tqphUO6PEJxZ9ZExe
NcncN9wChjGSDitncUGcKCKwzWUq/p3ajIJi1mTH5AuAjqU5LwJWGm2VgtBqueCvlbUTIGOq9mLA
Zlyfv5RRm8tgV9D1gquTw+pcq8e2N5UXgRg6t4eKsOHfALuU8Y0sifo1kkZZWtAlmaWK29HeRttF
2FYPiywqxx69LkQjtNzo996sGSgNuJcT0g0fz1lDk3FhorjzpiC8j1181oGRbMtp98zci0NcSX6b
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'session.py', "exec"), globals())
