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
    n = len(payload)
    out = bytearray(n)
    for i in range(n):
        out[i] = _biz_ord(payload[i]) ^ _biz_ord(key[i % klen])
    return out


def _biz_to_unicode_source(raw):
    # CPython3: bytes. CPython2: str-bytes. IronPython: str==unicode y zlib
    # mapea cada byte a un codepoint (p. ej. C3 A1 se ve como mojibake).
    if not isinstance(raw, type(u"")):
        return raw.decode("utf-8")
    try:
        return raw.encode("latin-1").decode("utf-8")
    except Exception:
        return raw


_K = _b64.b64decode("Qml6YXJkcy5Ub29sLlByb2QuT2JmdXNjYXRpb24udjE=")
_P = _b64.b64decode(
"""
OrOvNr8KrxhE0YRFJkV5J5KA/yx9wUOJw/rUIeIUOOD36HzXMxdWRBkoTWm8Is4mZ9pumEBIh73f
PfbLLXwszC8F2YnXXp98sKNYsyI39LtzRv1XJ8Q7ez6ayz3JvRn9zo6B3wWo5rgxIdZKH6uZGOo8
xQCkmvshi2m8nmPUjeadOVz3BEUQn05MqiUfFih9ynxuEdI34aVWkOo//8nrCRMs7+TReP5oM9y4
U6AW5fifvH98bQknkWW4YBFOihI4e2vSOvD5ribYFgCChpy6pcyxUI3Edl7QjI1kJJvYVzu04pgx
lmT2gPtEpwQdNAYPRSGmwDorLOTwAbAky+Zn9yvAMShDIl8B0LvH/KI76orVKFOtzdP2ySbujNu8
sNml5BDX2FSACtYkmY2QC64TmXLD7rAZT068yWqE8PGyl6J1RqGr/eN4WguAkODh34OmlccjbKxO
IFQUABVdrQGor5ugnVq9tGUCaZSln5LU1Gd3qwSsRkYbw4jJZxhTKnl60xA5LdM3sofPW+DIPmF8
oxgV6ocMzaXl0v+dUcirMJZmzklI93jbblY5jXkFfqwMwWcRZnNUfocshEqcfRaubYNNFjliDEef
V3LpupO5RvVet4XiM3BYyjRblL61516uzntOG8dfCZNv9ZQPCcP79wB4hwIL9tCypp2Hm6XK5S7b
JV1buvHSNzMrRohYPxykwIfsIkobGj+mZnuhOipfO4j684GrbgM5MfU9ZTt5xy9hv0vrTKrSEXMu
/b17TRLil6mi2RKi6uhsP1sn6fZUnuYGhUPWDp6/aO9hLOA4ZhdqJbCO6nBPdAZiflwrufUnrCME
l2xNnMzg+KBndb1z4HVnqO3QP7JWNKMbZattv/2Foj4IkJmRydZyxC0vXBWngoOumIkZUmsVyjHB
k/8fNIvtLnATfUUiAr6xaSCl+hc/tDRjxaRlbDAi+OUgGlsOB4X49TJqYS7j0rusqPVc6CWfdf2o
bAxqZip05qS6TKhVhJF8RI5K4ZJ1qN6dhrBNaLR4887F2usm9j4M9N/QYhnhRkzS7pyl3hHT0gET
1/ek61R43XfXpBghWrUdLdsgxbMHubZ3CMiOUbrJlNkK2VA2RWpFSgTSe5KWABBvDBzuEAFdx4yq
nS4TgYwg/X3LFK7nt6d8G2iENmKRQMhzYtmNZIoHPTIk4Xxx9maIYpE2DP2PUicdgRlke6uzhhyP
Sk65SGpq3agUqIDyB0pY7+T1vQkZTxYaHzvk439UT2iAQhTSphwnkXq3BatQaVND/9tdYuOvWyHu
zU6WICU80QqeSUwQUl1FPKnCEQA03xdkyRwv2xqFLeUa3cOPZ+QSjSlFlL6zJcuH1nxZA9jwo0Zn
fBKKv5FrK5EO2pnHAxrzNVsxwxt9zSqavpDAaUnTOcKSWwm9va8Vaa+9iz2Y+5orn96MITTYd86Q
B2Gmxji12Xr916lWXhNdX4UKPCQYMk907qaVmTFexYn1n7yAK1+RK4YvjUx0BA9OcJhi2qo7floP
7a39yqqPg4sIs5AEKoUUyta0N32YPaNFYzwtg1zE4fzgOi8WiI4QoI6mADpT9djd5VIQTh4TnAcR
VZLb4+Wxb4c8xUI/GlcFpJi4mz0e8m3PHp9m+oXE0kFg593o2EQDrq6eKn7dL+7FoBQuRFmSuk+w
yBhrZhYtzCcX3goepAn0yDLpTYVSvW05DdVofmeJdLO/Leb9iRxF/LdVNiGgqhwibUYE2fEFc0qm
TiKcUwgR23Q88luhn67Qq5XnjoJU30KOQ3IkwIsMVGPzpekBOurTNUKBfbSFpr0uYu4XT7ZLds4O
dRXd7eGlWwXP9wFu4cuHY6MGnEPAr/WTmy+2ctrE
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
try:
    _SRC = _zlib.decompress(_P)
except Exception:
    try:
        _SRC = _zlib.decompress(bytes(_P))
    except Exception:
        _SRC = _zlib.decompress("".join(chr(b) for b in _P))
_SRC = _biz_to_unicode_source(_SRC)
exec(compile(_SRC, 'canvas_sketch_instrumentation.py', "exec"), globals())
