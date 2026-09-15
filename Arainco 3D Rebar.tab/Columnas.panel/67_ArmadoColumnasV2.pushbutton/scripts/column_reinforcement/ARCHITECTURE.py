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
OrP3NLkKWGtEsqA7usL/fEGYz28sI1W06V+MZaUZIYHxQkcE1ac/2YgR9/Q7T7NglC/j0AXuTIIq
TLxyqadys5+lDhSnzQBggXO8lcdpa7A+mfbC+qOUPIrjj4eGhu+3jbBcUFtj1XhudZcfuHTGNX3x
qrxL9ivUl7UhQECDIAVI2+xBvDw1Hjpqng3tnoytohBy2PUEoK9DRp0g4UOYDXsmCE5UQ9XAvR4M
4ThyVXQLgvDGjmSgGfJBVvyPr1XcUpM9bhJ1qqDbhiXTGJ1L8vqQvN8To6ALqbW663698l55UENc
z2Sf8PfiVz2KfVXA3ayLvkMbDl9YN60pEgFf3VK6auyORHL+6fwnvRTij48nzA9pv683sJgP3cv7
iTU2ssyzO1guq3R3yrczka40cQAsAm/fXWeVXM8f2ryc612aheyDRLJ6EPVZPcjuGUQ5BiMV0S9L
57Q7gNQ0aWikoekvfzVfe4dfluxCoTS/4FNAUQkqj1foubNtIFuvxCucWsxpMgdeCGfl+n99Iism
tYhwr8osaHv74jx8hnsrzye9CISMlqOaqb/9GwfLt2V6juPWYc7j7BIGk03j6fNFXKqQdOZGiDld
x6qTjwtkQWM/P2PLosBDcjjD6jzXPnH2opPeIUpYXLnxzIp+wbKMqqDlwRVY0hZ7PyyXj6/UyOHq
nrQWRJ2RBlqJoQLTL5VFaS+UHxC8YMjZVBcnpOxTmH4/5m8kq6WjGu9/Qkm640+MCi4GEnWS+oPL
d76nDaksgTvufKFeZCUUiq856mpvALTBJ43i0yrOz/+AwWTcs1qidtEWtZbTjV1QRY/YBiJndKVf
0RRxED+08IFhzcyd9Rjej+d90q0ZISbK66TQQmBOQl+/cnUkdWdeneXWp08XenoLrlXJs5kqh18Q
QRwuibZeMAkXtzwcW1nbOMdchv8sBC4reO2IHTwGVCAhwmI46vs+5fNhGYlaRVYaWAiDCoAFsYw7
t79flemQxZKqX4m/yeJ1TT5VAJ2JLLwSH9DowTxbntuyEtHVwn6qvzZtAHKx4QxYWNWomIPi9IcZ
oYIzEzIUN4MLWc989a5yLkvJD3pkhg4gEmREn3PmuqUPy58w+XJCdQjo2gUqVQw8AzyzTXbqCWw0
hc6yMdTEMDFZbN2fr5hZhJJMTh39L+dP4qeYexaeTDP1jv4auqjg4ny2I2JZe88iovyOUi19rUMP
k1jKv9wWOdiKWraUXDjrmZiToZtpUi9amw==
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
exec(compile(_SRC, 'ARCHITECTURE.py', "exec"), globals())
