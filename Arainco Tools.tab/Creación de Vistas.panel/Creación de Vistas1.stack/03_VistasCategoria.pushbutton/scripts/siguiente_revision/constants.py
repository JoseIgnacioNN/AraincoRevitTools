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
OrMHNb8WqBhE0YRFWMT5vNV2Fpv/sUBeIVB9rI/x5mVuHkiiM0w7T9Jte/xcu+yVTpO7tDVrsgyz
U4Bt+gZHvymnItr5mU0CooB24FBKM87yixM9pNsq7r9sog9WMnhtZ4L8qjSww4CASDHNnkQiHpwC
+4T9J7vlIlaHpdR6JOWMCZrDO5fZ0tcszPqSFpV1UlISdDZCaYzJlXeBFosC+hdBw0ulX76lpo70
6AOHbisxQAWrkzX49rE2f0ocAr1LaR33c5382ZxFtSlJCvLSF5yYY7WsxLJsNy4ET6F5zVflLOQi
y71Z85R53mF0NaMPtK2OUX74NfNDFUx1S88REZnE3VfO8JVKFI1Xm4zQwl8bYSCDca1+xE28/cUE
CbPhD3u+QHVqMvhvHjymi6ofF0The/VTWqlEe6OvXQmlShVD3O4lkRn2lGimCWuh5SgM3sRmtXBi
PCrzE0WcBxBTWE1m2X/zkVeAH6xDI/Ado5yIw4pbOsnu/6IQWxYQlvf5MOpvUVLc+kkk2cbz9fuu
xEAkJ3zCfe7ff3uB7QvX+sm8/cKOKZf4GTxhXPZSNw1WejGnL4fhl9aScRZDQWWlCA8a9UREnqjl
23CojxL1u2iCa0Yaci2bUe6WVv2MFR6jBEgzkmunNFkdhJ4acClglh0+NrZB87hbZx5ARdB6zN6G
qw1VFlD6gkf8BE1l8Sf4nFr4SNldqMS57GYQA2rVaveJJSWBlPyqZwOPV/3yxNs2yMQn+LsUj1se
euecohespejHAg4Ft8+Vg77YGx/+F1DBnFAYFLdwkoAB8WAvlijuGXBKQXssWVbL7THjdVKH5Gen
5XCGoet8YeCYWZJEUyX3wzbfncY/ZWBu3Tp8Fz78iTH2vzObaXoDMF59BgGumMis13cpNeCAudt5
gYoeq0APDoz0RRB9kE6PEnhQFt4IiJlcyncD/CtAxTFHZnn5cBlPTM7QwYcVlZMc4cifJOfwF3Pe
oEfjjvi5N/IU7c5Kml+QLr6cts8NPQ4SYiFQAYTpuhrqUDrSt+642XzRqQ6BBTNLrjlHObDf04mD
Glo4Z4WfixP06adL1e043vipQS+5rYZafwXuy4GOmf+4dbPwoeFO9qJSQkpupQKJAIZgrzwdWRk8
XxPvjzzkIBBhHPRiqH5DZlQRC92oIwtvJTNBNud/SNgSnoFvw0MPAO++U7iwE1zpkPmztqySMEXT
Emyz5IEORhU9iFP6Voc1v+iCfq+nNT6zP1sbBsyU1TezpwG3TSGRwItPujzCJ1UL6HDk/O4qbtFg
4+oPEATJMowno/MUzqzl3BrfgMmH7YGaVM6Mk5UCbCl53/wKl3YP0y3BNeI+Yr5aT2R3mFdy7t8q
SlY76frxLY15Dx22xKjchScQFK1RbF1KK9yQE7L8pK0d3sQUYRGyalCua7vTiyFhv9g3/VBXFg+l
mcolof2uYEAaSMbUwuznfPY8yPVrKjyTpdM7opzFfQfuzuHNLv3i/4yy1Hx/y5ffAjB33HCP3Y4d
I2fiqA==
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
exec(compile(_SRC, 'constants.py', "exec"), globals())
