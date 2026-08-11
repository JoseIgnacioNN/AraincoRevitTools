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
OrP/8D8qcB9Y6phVqwVEQrRPjmQTonYoRSBn/98caFtr+j6YW05YIJUiV24ZsVZXtzgVckQ+iIaY
h5p2ouGRUT6ulBBzQYTdzzU+/+e0pIrsM42lrV5fQNNHF24ZRlsXYm456ExROYP8C9mQohBkaftT
y+yCdlLP6SmGzM+vKu8XbpZdrrHcnu05vNrr4dYoU0OWczw/Pk8ve4rH2Kt9yWNi61c+K8meIGhA
4hX9QLEldNzXpDdoV52Kj2u3gsuSRj7q3YcEQ1ivUgkhcwTbUZ4v+4wtW+LXZ2rr4HdsEqbngzYH
nfAhM6E/vq6YQpYGBSIu1DZVnEtQ9ydp3W+vCIuzF/o+gOO15HoPvYrKpwx1PuF4JhZbmsfGPcR/
yRJSWxP6K+Yvo4ktDzHLZWk4xZwJYivBPdPZ7uC/BRrRvVF4cqGeyIgGkXVguFe+Lm0eUxaFgo77
uZidyuqSd8akgyZ8zMro0q9BIZCfefzBqYZd+mjj84iK
""".replace("\n", "").replace("\r", "")
)
_P = _biz_xor_decode(_P, _K)
_SRC = _zlib.decompress(_P)
if not isinstance(_SRC, str):
    _SRC = _SRC.decode("utf-8")
exec(compile(_SRC, 'mallas_en_muros_run.py', "exec"), globals())
