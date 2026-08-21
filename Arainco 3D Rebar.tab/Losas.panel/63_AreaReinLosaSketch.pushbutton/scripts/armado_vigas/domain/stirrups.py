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
OrOfN78KqGhE0ZxFKLyzyykf++B+8ONDKSlMXwwVM9VAKBcCw0UkClVBJkB54YGMg+bshvjJKwT3
H/vJS/lfvRzD+dmu/apBsvQAOtRVVLR3puRGSaZqz5UGm0SDFtjavDzv3fMg80bgpbA7I5xdmsdT
HgjOXnZ8k+DqyCp0Bqk8luyYU6JL4AY6fdS8bNt33RouMC/r5AI94eSJy0BWplrGG99KHQjej3QZ
HJLI7d/Hje+5P4mCAZ6UkXqVU4jDh/b9Rh4bDVLmtlUTmZYKGxTCdyxE4yMQtHyfMR7nbCCp1VBc
j9VlIt+dCGLOl5TM4Ffckmri0odbulXcLTGTSyWg9JYOEmXMhxCB8xuL1VDacPG5WLKmq4dyNLPE
4udtkyh4U8UWbGlr0YEpgI+Ilgk4BPLVazyL9F4INjpyDQMOPM/LzzmsIRI0PNOj5ozgxx5FAmnK
HpbXEP/mUuNCojA1RmpSp/nBNgZn/LmKMY9W33264Ch3oWUuQXyiu5QuG0iwNRXgNbVCpvFM1/X1
VpIp+mf3FutpjM1xg2hOOAdSeFcAcDcPNT/JbOMFCd1bgbVk18ys1Ml1uXnyqme5hmRirhRsj2pm
uGIXN4a/yH8/cfl4enYvQ6ySaUecaHHhrNe0qnbO6Za1KhePfnpnJthjOT1zxJq5B2EtfJRBEP36
XE9ymiH3g8UYJA0h0CHll5fnG+qsyV1omefmnxX+yeVJEp2d4TIdPy4LMnwtZOjOJENS00J5xPz7
IT3E2X+YlbkO4B1b2j2AbOqGpdrjw8NPy9RQ56uOJ9kErj3FTrAfWKPaRxnslVm8A36la3HY8e0v
quvAZGqYsKG8ts3muAmkfu5BkPtATSKnwOM8FlsGapSLq/n/LjTxoGUBtGeRc0oxJgCZwujqO1My
Rv56FVhzv6o0Ii+b58BifysN/sJPTmcDPUiliLSyHpz7KY0RgfZOQ8AB2dLk0ioZWc2zI/XLimgE
sdMiiOY3vnYOCDIYTszMgftL1VtAXgr1x7Lb5lYdwWF3gZpHniFL+9fm5HAEWhWM0LwO0gQBnlF5
OCIjEr24ai8Ij2vsvEJKhN0EXr6oE0Hb2otO1+g7kr+FoVUlnd8+uBOPh5wSzoHSDUUN8yLWyFOB
2NJqZmS0N7j4/UK1raan59jEsnU7SQ5+SJtzWK3hiuZdFuMhlGErCd4st3B6OHivJD+1tLMktguK
mCOpwCwX
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'stirrups.py', "exec"), globals())
