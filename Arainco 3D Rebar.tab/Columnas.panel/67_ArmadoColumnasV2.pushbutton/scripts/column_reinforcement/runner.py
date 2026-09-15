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
OrPfNqkKqBZEErg7Prr50xS6klxzW48DxX15D9810vyZa2hMoVLEPP2fjajVZ4luGcRohqEPO4kH
Mx/lsw7+XTNR5lgBggPc2jOTHm9ay40GHWy0ylpN21YldahZmq7jV0xRJ6q7355REJZCZu3NCF9A
YY2JS+D5Sz7BxYZkTJoA/8xzJsMiWjPq/iiAKbMGvcBuKwuABQrs5QtdTVfYOocC6beJWFQrAXQl
9BjcJkC7efn2llITp8pOREozSyxrF9lG7oTAX0cR3snA5rrPHaqzRA35EI0Y95FQbTe68BgDZKUw
4OCAZp1ZQY5sBqC0M7pEBYgWeSqhc34E591cFw4/qNzGQiCbGmDh4HicS0yQ//Ee/WS2P8H0y4Mc
JhF8S34BN3ZgJJzs27ORwxLvYXUo3RuIF/GcvL9e+InBhmOuxgYJhae7EGfajGI69JhFdg5H+a/r
UUEMLQCLwi99KCIoCeBn68sM++0iMd/zetcvSwca8gmq3fHOlMuZFPznCD8ebcbQ4S5W794rpeKY
qXc3hoBY7e+ShAUjBqN9ElmaGmF2g6DJUFIX6SdfRWSwEJosKKVp7yyghHN7h+xdJeHvH+w3sl+x
pccsZRpy2UVOeCiQ2n9lXILYnXqhtq5BgzPtDqNE+W6cWq5CQ0W5bblb+YBM7L9B4+uoRFxpX+LY
8xfrHeYlAm2a8kKbK1Q20shqqdS8F4HchZzuLQ8nJ6MBm9b3d46HmfT7Zwo44DDD7TKj8Re1SVAI
tgntpNnIhEyu4XAGqrxxaLDKCuRgwo8IGvKeo0b+wAmufkhwhQwKZuSh6oBu8aekKdp8h3yhx3+0
XMGVHL1+70dlxiINib9NcYr4tuowUPXZd9Oj9qKkDCK/X8GwPc9JvRAmHnUshJXdcuxe+jfShQK1
XZFS6bPQTsiR3KY3imIVk8+Y8sXHfzcbApvfJbgN9QyvAIby4uWVpzIngROaZUOsxHQX/KMrzzZv
VMWccTozGoZm91BJJQbt35tLc3sXlyo3GJvn28kxkSGzVtz4JOWE6H0q2gAFOKa3B8sBh0WxaIzq
F1CQC5tlhu3oQ5MspwD80UhK5/lvmjBLFiyUDhcsP4xwh2E8Oyy0gdVhLS/AdkdFpvqufy4pU0A1
bq77bkmfyv4r46IglO7reEm++8bR8Nbc4IYlgXjX98DU5V1wL71lMbISJsY3M4MU6mcGIfEicIBr
DhjoroGx4UW0NxxmFaEbl2FVFcpM3Jy9RebWemjVUyaH8KGFfQdffcLWhTiQ7dv2N887MSV12hyY
T1Qz2pwAuLkDjzzpuirdK5z8pRvfW52DHNTe1EXYqY9Z+0N6L4/rRWEhwOUojpiVR2PycgKMV8x+
cBVaVOgClmcWueLfgxug1hUhrsI/SjsFCSy/hSZB1oOfPsCGaeyXvrYiziduyb6jHth7zFfChEsH
25dlslLDRUumuEQnogx+CTIjd/9AuO6PfdkbMjvX9xVyaPelwPbHDLsIoK+AEyz47NCNV/viYaL3
ccpC/qfnvni7pqUB5V83WGu6dSUcZWjv9+UgPekWdr4svUjTjeWnp6OpgY7TYgjK1tyotXl1j1GX
POlPUO62GRoN8mw6gk46WTMj7ANmefxhPg6EzV5HWNFgyqZe4kV9b0G14+EKkCavORLtcm5XTP2L
x9EhDCtTLghRwDE=
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
exec(compile(_SRC, 'runner.py', "exec"), globals())
